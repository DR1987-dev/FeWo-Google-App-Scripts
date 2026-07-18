import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { globSync } from "glob";
import { google } from "googleapis";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.resolve(__dirname, "..");
const OWNER_EMAIL = "d.rybosch@davidsapartment.com";

async function main() {
    const scriptId = await resolveScriptId();
    if (!scriptId || /AbCdEfGhIj/i.test(scriptId)) {
        throw new Error("SCRIPT_ID fehlt oder ist ein Platzhalter. Setze das GitHub Secret SCRIPT_ID.");
    }

    console.log(`Deploy target scriptId: ${scriptId}`);
    console.log(`Owner hint: ${OWNER_EMAIL}`);

    const auth = await createAuth();
    const script = google.script({ version: "v1", auth });

    const files = await collectAppsScriptFiles();
    if (!files.length) {
        throw new Error("Keine Apps-Script-Dateien gefunden (.js, .gs, appsscript.json).");
    }

    console.log(`Uploading ${files.length} file(s) to Apps Script project...`);
    await script.projects.updateContent({
        scriptId,
        requestBody: {
            files,
        },
    });

    const versionDescription = buildVersionDescription();
    const versionRes = await script.projects.versions.create({
        scriptId,
        requestBody: {
            description: versionDescription,
        },
    });

    const versionNumber = versionRes.data.versionNumber;
    if (!versionNumber) {
        throw new Error("Versionsnummer konnte nicht erstellt werden.");
    }

    console.log(`Created version: ${versionNumber}`);

    const deploymentName = "CI Deployment";
    const deploymentsRes = await script.projects.deployments.list({ scriptId });
    const allDeployments = deploymentsRes.data.deployments || [];

    // @HEAD is the special "latest code" development deployment that Google creates automatically.
    // It is not a versioned deployment and does not count against the 20-deployment limit.
    const versionedDeployments = allDeployments.filter((d) => d.deploymentId !== "@HEAD");

    // Due to a previous bug (d.description vs d.deploymentConfig.description), multiple
    // "CI Deployment" entries may have accumulated. Find them all and clean up duplicates.
    const ciDeployments = versionedDeployments.filter(
        (d) => d.deploymentConfig?.description === deploymentName
    );

    if (ciDeployments.length > 0) {
        // Keep the first entry as the canonical deployment; delete any extras
        const [primary, ...extras] = ciDeployments;
        for (const extra of extras) {
            await script.projects.deployments.delete({
                scriptId,
                deploymentId: extra.deploymentId,
            });
            console.log(`Deleted duplicate CI deployment ${extra.deploymentId}`);
        }

        await script.projects.deployments.update({
            scriptId,
            deploymentId: primary.deploymentId,
            requestBody: {
                deploymentConfig: {
                    versionNumber,
                    manifestFileName: "appsscript",
                    description: deploymentName,
                },
            },
        });
        console.log(`Updated deployment ${primary.deploymentId} -> version ${versionNumber}`);
    } else {
        // No CI deployment exists yet. If we're at the 20-deployment limit, remove the oldest
        // versioned deployments (sorted by updateTime ascending) to make room for the new one.
        const DEPLOYMENT_LIMIT = 20;
        if (versionedDeployments.length >= DEPLOYMENT_LIMIT) {
            const slotsNeeded = versionedDeployments.length - DEPLOYMENT_LIMIT + 1;
            // Sort oldest-first so we reliably delete the least-recently-updated deployments
            const sortedOldestFirst = [...versionedDeployments].sort(
                (a, b) => new Date(a.updateTime) - new Date(b.updateTime)
            );
            const toDelete = sortedOldestFirst.slice(0, slotsNeeded);
            for (const dep of toDelete) {
                await script.projects.deployments.delete({
                    scriptId,
                    deploymentId: dep.deploymentId,
                });
                console.log(`Deleted old deployment ${dep.deploymentId} to stay within limit`);
            }
        }

        const created = await script.projects.deployments.create({
            scriptId,
            requestBody: {
                deploymentConfig: {
                    versionNumber,
                    manifestFileName: "appsscript",
                    description: deploymentName,
                },
            },
        });
        console.log(`Created deployment ${created.data.deploymentId} -> version ${versionNumber}`);
    }

    console.log("Deploy finished successfully.");
    console.log("If deployment/auth fails, ensure service account has Editor access to the Apps Script project.");
}

async function createAuth() {
    const credentialsPath =
        process.env.GOOGLE_APPLICATION_CREDENTIALS ||
        path.resolve(PROJECT_ROOT, "credentials.json");

    const credentials = JSON.parse(
        await fs.readFile(credentialsPath, "utf8")
    );

    const auth = new google.auth.JWT({
        email: credentials.client_email,
        key: credentials.private_key,

        // <-- GANZ WICHTIG
        subject: "d.rybosch@davidsapartment.com",

        scopes: [
            "https://www.googleapis.com/auth/script.projects",
            "https://www.googleapis.com/auth/script.deployments",
            "https://www.googleapis.com/auth/drive"
        ]
    });

    await auth.authorize();

    return auth;
}

async function resolveScriptId() {
    if (process.env.SCRIPT_ID && String(process.env.SCRIPT_ID).trim()) {
        return String(process.env.SCRIPT_ID).trim();
    }

    const claspPath = path.resolve(PROJECT_ROOT, ".clasp.json");
    try {
        const raw = await fs.readFile(claspPath, "utf8");
        const parsed = JSON.parse(raw);
        return String(parsed.scriptId || "").trim();
    } catch (_err) {
        return "";
    }
}

function buildVersionDescription() {
    const sha = String(process.env.GITHUB_SHA || "").slice(0, 7);
    const ref = String(process.env.GITHUB_REF_NAME || "local");
    const stamp = new Date().toISOString();
    return `CI ${ref}${sha ? ` (${sha})` : ""} @ ${stamp}`;
}

async function collectAppsScriptFiles() {
    const patterns = ["*.js", "*.gs", "appsscript.json"];
    const matches = patterns.flatMap((pattern) =>
        globSync(pattern, {
            cwd: PROJECT_ROOT,
            nodir: true,
            ignore: ["node_modules/**", ".git/**", ".github/**"],
        })
    );

    const unique = Array.from(new Set(matches)).sort((a, b) => a.localeCompare(b));
    const files = [];

    for (const relPath of unique) {
        const absPath = path.resolve(PROJECT_ROOT, relPath);
        const source = await fs.readFile(absPath, "utf8");
        const parsed = path.parse(relPath);

        if (parsed.base === "appsscript.json") {
            files.push({
                name: "appsscript",
                type: "JSON",
                source,
            });
            continue;
        }

        files.push({
            name: parsed.name,
            type: "SERVER_JS",
            source,
        });
    }

    return files;
}

main().catch((err) => {
    console.error("Deploy failed:", err?.message || err);
    process.exit(1);
});
