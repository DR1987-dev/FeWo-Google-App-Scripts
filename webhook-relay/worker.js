// Cloudflare Worker: nimmt Lexware-Webhooks entgegen, prüft die Signatur und
// löst per repository_dispatch den GitHub-Workflow "Run FeWo Skripte" aus.
//
// Secrets/Vars (wrangler secret put ...):
//   GITHUB_TOKEN  Fine-grained PAT, nur dieses Repo, Permission "Contents: read & write"
//   GITHUB_REPO   z. B. "DR1987-dev/FeWo-Google-App-Scripts"
//   ORGANIZATION_ID (optional) Lexware organizationId; fremde Events werden verworfen

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

export default {
    async fetch(request, env) {
        if (request.method === "HEAD") return new Response(null, { status: 200 }); // Lexware prüft die URL per HEAD
        if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });

        const rawBody = await request.text();
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

        const res = await fetch(`https://api.github.com/repos/${env.GITHUB_REPO}/dispatches`, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${env.GITHUB_TOKEN}`,
                Accept: "application/vnd.github+json",
                "User-Agent": "lexware-webhook-relay",
                "X-GitHub-Api-Version": "2022-11-28",
            },
            body: JSON.stringify({
                event_type: "lexware-event",
                client_payload: {
                    eventType: event.eventType,
                    resourceId: event.resourceId,
                    eventDate: event.eventDate,
                },
            }),
        });
        // Non-2xx lässt Lexware nach seiner Retry-Strategie erneut zustellen.
        return new Response(null, { status: res.ok ? 204 : 502 });
    },
};
