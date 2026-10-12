// Worker: führt die asynchronen Google-/HTTP-Aufrufe aus, während der Haupt-Thread synchron (Atomics.wait) wartet.
import { parentPort, workerData } from "node:worker_threads";
import { google } from "googleapis";

const { port, signal, env } = workerData;
const sig = new Int32Array(signal);
let authPromise = null;

function getAuth() {
    if (!authPromise) {
        const credentials = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT);
        const auth = new google.auth.JWT({
            email: credentials.client_email,
            key: credentials.private_key,
            subject: env.GOOGLE_IMPERSONATE_USER || undefined,
            scopes: [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.readonly",
                "https://www.googleapis.com/auth/gmail.send",
            ],
        });
        authPromise = auth.authorize().then(() => auth);
    }
    return authPromise;
}

async function sheetsApi() { return google.sheets({ version: "v4", auth: await getAuth() }); }

const handlers = {
    async fetch({ url, method, headers, body, timeoutMs }) {
        const res = await fetch(url, {
            method: (method || "get").toUpperCase(),
            headers,
            body: body === undefined || body === null ? undefined : body,
            signal: AbortSignal.timeout(timeoutMs || 120000),
        });
        const outHeaders = {};
        res.headers.forEach((v, k) => { outHeaders[k] = v; });
        return { status: res.status, headers: outHeaders, body: await res.text() };
    },

    async "sheets.load"({ spreadsheetId }) {
        const api = await sheetsApi();
        const res = await api.spreadsheets.get({
            spreadsheetId,
            includeGridData: true,
            fields: "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount,frozenRowCount)),data(rowData(values(effectiveValue,effectiveFormat(numberFormat(type))))))",
        });
        return res.data.sheets || [];
    },

    async "sheets.batch"({ spreadsheetId, requests }) {
        const api = await sheetsApi();
        const res = await api.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
        return res.data.replies || [];
    },

    async "values.batchUpdate"({ spreadsheetId, data }) {
        const api = await sheetsApi();
        await api.spreadsheets.values.batchUpdate({
            spreadsheetId,
            requestBody: { valueInputOption: "USER_ENTERED", data },
        });
        return true;
    },

    async "values.batchClear"({ spreadsheetId, ranges }) {
        const api = await sheetsApi();
        await api.spreadsheets.values.batchClear({ spreadsheetId, requestBody: { ranges } });
        return true;
    },

    async "drive.findFolders"({ name }) {
        const drive = google.drive({ version: "v3", auth: await getAuth() });
        const q = `mimeType='application/vnd.google-apps.folder' and trashed=false and name='${String(name).replace(/[\\']/g, "\\$&")}'`;
        const res = await drive.files.list({ q, fields: "files(id,name)", pageSize: 100 });
        return res.data.files || [];
    },

    async "drive.listFiles"({ folderId }) {
        const drive = google.drive({ version: "v3", auth: await getAuth() });
        const out = [];
        let pageToken;
        do {
            const res = await drive.files.list({
                q: `'${folderId}' in parents and trashed=false and mimeType!='application/vnd.google-apps.folder'`,
                fields: "nextPageToken,files(id,name)",
                pageSize: 1000,
                pageToken,
            });
            out.push(...(res.data.files || []));
            pageToken = res.data.nextPageToken;
        } while (pageToken);
        return out;
    },

    async "drive.download"({ fileId }) {
        const drive = google.drive({ version: "v3", auth: await getAuth() });
        const res = await drive.files.get({ fileId, alt: "media" }, { responseType: "text" });
        return typeof res.data === "string" ? res.data : JSON.stringify(res.data);
    },

    async "mail.send"({ from, to, subject, body }) {
        const gmail = google.gmail({ version: "v1", auth: await getAuth() });
        const encSubject = "=?UTF-8?B?" + Buffer.from(subject, "utf8").toString("base64") + "?=";
        const raw = [
            `From: ${from}`, `To: ${to}`, `Subject: ${encSubject}`,
            "MIME-Version: 1.0", 'Content-Type: text/plain; charset="UTF-8"',
            "Content-Transfer-Encoding: base64", "", Buffer.from(body, "utf8").toString("base64"),
        ].join("\r\n");
        await gmail.users.messages.send({ userId: "me", requestBody: { raw: Buffer.from(raw).toString("base64url") } });
        return true;
    },
};

port.on("message", async ({ op, args }) => {
    let reply;
    try {
        if (!handlers[op]) throw new Error("Unbekannte Operation: " + op);
        reply = { ok: true, value: await handlers[op](args || {}) };
    } catch (e) {
        const apiErr = e?.response?.data?.error;
        reply = { ok: false, error: (apiErr && (apiErr.message || JSON.stringify(apiErr))) || e?.message || String(e) };
    }
    port.postMessage(reply);
    Atomics.store(sig, 0, 1);
    Atomics.notify(sig, 0);
});
parentPort.postMessage("ready");
