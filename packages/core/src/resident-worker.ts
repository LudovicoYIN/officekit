import { loadDocument } from "./document-store.js";
import { removeSessionRecord, writeSessionRecord } from "./session-registry.js";

const filePath = process.argv[2];

if (!filePath) {
  console.error("resident worker requires <file>.");
  process.exit(1);
}

let residentDocument: Awaited<ReturnType<typeof loadDocument>> | null = null;

async function main() {
  residentDocument = await loadDocument(filePath);
  await writeSessionRecord("resident", filePath, {
    kind: "resident",
    filePath,
    pid: process.pid,
    startedAt: new Date().toISOString(),
    format: residentDocument.format,
  });
}

async function shutdown(exitCode = 0) {
  await removeSessionRecord("resident", filePath);
  residentDocument = null;
  process.exit(exitCode);
}

process.on("SIGINT", () => {
  shutdown(0).catch(() => process.exit(1));
});

process.on("SIGTERM", () => {
  shutdown(0).catch(() => process.exit(1));
});

main()
  .then(() => {
    setInterval(() => {
      void residentDocument;
    }, 60_000);
  })
  .catch(async (error) => {
    console.error(error instanceof Error ? error.message : String(error));
    await removeSessionRecord("resident", filePath);
    process.exit(1);
  });
