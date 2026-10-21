// Cloudflare Worker: nimmt Lexware-Webhooks entgegen, prüft die Signatur und
// löst per repository_dispatch den GitHub-Workflow "Run FeWo Skripte" aus.
//
// Secrets/Vars (wrangler secret put ...):
//   GITHUB_TOKEN  Fine-grained PAT, nur dieses Repo, Permission "Contents: read & write"
//   GITHUB_REPO   z. B. "DR1987-dev/FeWo-Google-App-Scripts"
//   ORGANIZATION_ID (optional) Lexware organizationId; fremde Events werden verworfen
//   LODGIFY_SECRET_<EVENT>  Signatur-Secret je Lodgify-Webhook (siehe subscribe-lodgify.sh)

// https://developers.lexware.io/webhookSignature/public/public_key.pub
const LEXWARE_PUBLIC_KEY_PEM = `-----BEGIN PUBLIC KEY-----
MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAtCkZXZqT2zFRmA83KsBC
tQSv9t+AaWiZpRWvzE4wiQEh0aHzNYkT/DlP419pngqoLxnqW9SsgWfKOG3utoEV
z+Lru9odZntW/n/yaRK3f0AcMtuWs/Z1PZ4BbN/RlFYyqlPg0VfuAvZMAa+J9WAl
Fuy0T4Y0EdO0d+EL8BpzSLqj3XLI+YuK4hUhlTOxHmdDtNUkugnT9b9khb+oPDmB
3yxRP6Rmo2qo6waSc+pihCdOiZKzwA4fA1HT8DM54OaNQ4qd9NRf5uECijFlpQu6
vZdW7Z+LZ8HTgGQVSsJdCzUtM1wiWOERrHvd8IpBMvUAvGM/qaJnAmwsEr4J/El4
yCmur8wwYBtSUH0Xue42YysdmSS0pOMMUAAvr70HQMDHagYkErMtu6FAyAmPz8p/
tAqJJpwAhxITbnc5WsmZOu0Ke255nXQ/2Go1TKv+45mSb7RLsvpmhVZ4nFByFt24
vG6IwNReuLhJJL468jrdTpxyZRn0QqkxiQV7jDd9Dp1NVd/W3+n86Aos9LnJa+ut
kN+jOOWmJdJjnQetJlbDVg3ex+XHVmCjiRYggCGPlpXVJwdEUNGzxhPGKjxmeiQT
ahwFA9iHtwBw7yK5VfM/hA+JeF2FXYhTdehfClAQt1YCYXkgUEFxm9idRdBoCY6U
bmFUXQLdi7tZyDor8Rxoq2MCAwEAAQ==
-----END PUBLIC KEY-----`;

function b64ToBytes(b64) {
    const bin = atob(b64.replace(/\s+/g, ""));
    return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}

async function importKey() {
    const der = b64ToBytes(
        LEXWARE_PUBLIC_KEY_PEM.replace(/-----(BEGIN|END) PUBLIC KEY-----/g, "")
    );
    return crypto.subtle.importKey(
        "spki",
        der,
        { name: "RSASSA-PKCS1-v1_5", hash: "SHA-512" },
        false,
        ["verify"]
    );
}

export async function verifySignature(rawBody, signatureB64) {
    if (!signatureB64) return false;
    let signature;
    try {
        signature = b64ToBytes(signatureB64);
    } catch {
        return false;
    }
    const key = await importKey();
    const enc = new TextEncoder();
    const candidates = [rawBody];
    try {
        // Lexware signs the compact JSON; tolerate pretty-printed transport.
        candidates.push(JSON.stringify(JSON.parse(rawBody)));
    } catch {
        /* ignore */
    }
    for (const c of candidates) {
        if (await crypto.subtle.verify("RSASSA-PKCS1-v1_5", key, signature, enc.encode(c))) return true;
    }
    return false;
}

const LODGIFY_EVENTS = [
    "booking_new_any_status",
    "booking_change",
    "booking_status_change_booked",
    "booking_status_change_tentative",
    "booking_status_change_open",
    "booking_status_change_declined",
    "availability_change",
];

async function verifyLodgifySignature(rawBody, header, secret) {
    if (!header || !secret) return false;
    const given = header.replace(/^sha256=/i, "").trim().toLowerCase();
    const key = await crypto.subtle.importKey(
        "raw",
        new TextEncoder().encode(secret),
        { name: "HMAC", hash: "SHA-256" },
        false,
        ["sign"]
    );
    const mac = new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(rawBody)));
    const expected = Array.from(mac, (b) => b.toString(16).padStart(2, "0")).join("");
    if (expected.length !== given.length) return false;
    let diff = 0;
    for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ given.charCodeAt(i);
    return diff === 0;
}

async function dispatchToGithub(env, eventType, clientPayload) {
    const res = await fetch(`https://api.github.com/repos/${env.GITHUB_REPO}/dispatches`, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${env.GITHUB_TOKEN}`,
            Accept: "application/vnd.github+json",
            "User-Agent": "webhook-relay",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        body: JSON.stringify({ event_type: eventType, client_payload: clientPayload }),
    });
    return res.ok;
}

export default {
    async fetch(request, env) {
        const url = new URL(request.url);
        const lodgifyMatch = url.pathname.match(/^\/lodgify\/([a-z_]+)$/);

        if (request.method === "HEAD") return new Response(null, { status: 200 }); // Lexware prüft die URL per HEAD
        if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });

        const rawBody = await request.text();

        if (lodgifyMatch) {
            const event = lodgifyMatch[1];
            if (!LODGIFY_EVENTS.includes(event)) return new Response("Not Found", { status: 404 });
            const secret = env[`LODGIFY_SECRET_${event.toUpperCase()}`];
            if (!(await verifyLodgifySignature(rawBody, request.headers.get("ms-signature"), secret))) {
                return new Response("Invalid signature", { status: 401 });
            }
            let bookingId = null;
            try {
                let parsed = JSON.parse(rawBody);
                if (Array.isArray(parsed)) parsed = parsed[0];
                bookingId =
                    parsed?.booking?.id ??
                    parsed?.booking_id ??
                    parsed?.bookingId ??
                    parsed?.id ??
                    null;
            } catch {
                /* ignore */
            }
            // Lodgify erwartet exakt 200 OK.
            const ok = await dispatchToGithub(env, "lodgify-event", { eventType: event, bookingId });
            return new Response(null, { status: ok ? 200 : 502 });
        }

        if (!(await verifySignature(rawBody, request.headers.get("X-Lxo-Signature")))) {
            return new Response("Invalid signature", { status: 401 });
        }
        let event;
        try {
            event = JSON.parse(rawBody);
        } catch {
            return new Response("Bad Request", { status: 400 });
        }
        if (env.ORGANIZATION_ID && event.organizationId !== env.ORGANIZATION_ID) {
            return new Response("Ignored", { status: 202 });
        }
        // Non-2xx lässt Lexware nach seiner Retry-Strategie erneut zustellen.
        const ok = await dispatchToGithub(env, "lexware-event", {
            eventType: event.eventType,
            resourceId: event.resourceId,
            eventDate: event.eventDate,
        });
        return new Response(null, { status: ok ? 204 : 502 });
    },
};
