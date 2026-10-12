import { Worker, MessageChannel, receiveMessageOnPort } from "node:worker_threads";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
let worker, port1, sig;

export async function startBridge() {
    const channel = new MessageChannel();
    port1 = channel.port1;
    const signal = new SharedArrayBuffer(4);
    sig = new Int32Array(signal);
    worker = new Worker(path.join(here, "worker.js"), {
        workerData: {
            port: channel.port2,
            signal,
            env: {
                GOOGLE_SERVICE_ACCOUNT: process.env.GOOGLE_SERVICE_ACCOUNT || "",
                GOOGLE_IMPERSONATE_USER: process.env.GOOGLE_IMPERSONATE_USER || "",
            },
        },
        transferList: [channel.port2],
    });
    await new Promise((resolve, reject) => {
        worker.once("message", resolve);
        worker.once("error", reject);
    });
}

export function stopBridge() { if (worker) worker.terminate(); }

/** Synchroner Aufruf einer Worker-Operation. */
export function call(op, args) {
    Atomics.store(sig, 0, 0);
    port1.postMessage({ op, args });
    Atomics.wait(sig, 0, 0);
    const msg = receiveMessageOnPort(port1);
    if (!msg) throw new Error("Keine Antwort vom Worker für " + op);
    if (!msg.message.ok) throw new Error(msg.message.error);
    return msg.message.value;
}

export function sleepMs(ms) {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}
