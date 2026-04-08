import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";

const SESSION_DIR = path.join(tmpdir(), "officekit-sessions");

function sessionFileName(kind, filePath) {
  const digest = createHash("sha1").update(path.resolve(filePath)).digest("hex");
  return `${kind}-${digest}.json`;
}

export function getSessionFilePath(kind, filePath) {
  return path.join(SESSION_DIR, sessionFileName(kind, filePath));
}

export async function writeSessionRecord(kind, filePath, record) {
  await mkdir(SESSION_DIR, { recursive: true });
  const sessionPath = getSessionFilePath(kind, filePath);
  await writeFile(sessionPath, JSON.stringify(record, null, 2), "utf8");
  return sessionPath;
}

export async function readSessionRecord(kind, filePath) {
  try {
    const sessionPath = getSessionFilePath(kind, filePath);
    const raw = await readFile(sessionPath, "utf8");
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export async function removeSessionRecord(kind, filePath) {
  const sessionPath = getSessionFilePath(kind, filePath);
  await rm(sessionPath, { force: true });
}

export async function waitForSessionRecord(kind, filePath, timeoutMs = 2000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const record = await readSessionRecord(kind, filePath);
    if (record) {
      return record;
    }
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  return null;
}
