/**
 * Handle-based document operations for @officekit/ppt.
 *
 * Provides functions to open, close, and manage PPTX documents
 * in memory for efficient multiple operations.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { readStoredZip } from "../../core/src/zip.js";
import { err, ok, okVoid, notFound, invalidInput, fail } from "./result.js";
import type { Result, DocumentHandle as ResultHandle } from "./types.js";
import {
  registerDocument,
  unregisterDocument,
  hasDocument,
  getDocument,
  markDirty,
  markClean,
  type DocumentHandle,
} from "./registry.js";
import type { Buffer } from "node:buffer";

// Re-export the DocumentHandle type
export type { DocumentHandle } from "./registry.js";

// ============================================================================
// Open Document
// ============================================================================

/**
 * Opens a PPTX file and keeps it in memory for subsequent operations.
 *
 * Use the returned handle for all subsequent operations instead of the file path.
 * The file stays loaded in memory as a zip buffer until close() is called.
 *
 * @param filePath - Absolute path to the PPTX file
 * @returns Result containing the document handle on success
 *
 * @example
 * const result = await open("/path/to/presentation.pptx");
 * if (result.ok) {
 *   const handle = result.data.handle;
 *   // Use handle for subsequent operations...
 * }
 */
export async function open(filePath: string): Promise<Result<{ handle: DocumentHandle; filePath: string }>> {
  // Validate filePath
  if (!filePath) {
    return invalidInput("File path is required");
  }

  // Check if it's a valid pptx file
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".pptx") {
    return invalidInput(`Expected a .pptx file, got '${ext}'`, "Only .pptx files are supported");
  }

  try {
    // Read the file into memory
    const buffer = await readFile(filePath);

    // Parse the zip
    const zip = readStoredZip(buffer);

    // Register the document
    const handle = registerDocument(filePath, zip);

    return ok({ handle, filePath });
  } catch (e) {
    if (e instanceof Error) {
      // Provide more helpful error messages for common issues
      if (e.message.includes("ENOENT") || e.message.includes("no such file")) {
        return notFound("File", filePath, "Check that the file path is correct and the file exists");
      }
      if (e.message.includes("Invalid zip")) {
        return invalidInput(`File is not a valid PPTX: ${e.message}`, "Ensure the file is a valid PowerPoint presentation");
      }
      return fail(e, "operation_failed");
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Close Document
// ============================================================================

/**
 * Closes an open document, optionally saving changes to disk.
 *
 * @param handle - The document handle returned by open()
 * @param save - If true, writes changes back to the original file
 * @returns Result indicating success or failure
 *
 * @example
 * // Close without saving
 * await close(handle);
 *
 * // Close and save changes
 * await close(handle, true);
 */
export async function close(handle: DocumentHandle, save?: boolean): Promise<Result<void>> {
  // Check if document is open
  if (!hasDocument(handle)) {
    return notFound("Document", handle, "The document may already be closed or was never opened");
  }

  const document = getDocument(handle);
  if (!document) {
    return notFound("Document", handle);
  }

  try {
    // If save is requested and document is dirty, write to disk
    if (save && document.dirty) {
      await saveToFile(handle);
    }

    // Unregister the document (releases memory)
    unregisterDocument(handle);

    return okVoid();
  } catch (e) {
    if (e instanceof Error) {
      return fail(e, "operation_failed");
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Saves the document to disk.
 */
async function saveToFile(handle: DocumentHandle): Promise<Result<void>> {
  const document = getDocument(handle);
  if (!document) {
    return notFound("Document", handle);
  }

  try {
    // Import createStoredZip dynamically to avoid circular dependencies
    const { createStoredZip } = await import("../../core/src/zip.js");

    // Convert Map entries to array format for createStoredZip
    const entries: Array<{ name: string; data: Buffer }> = [];
    for (const [name, data] of document.zip) {
      entries.push({ name, data });
    }

    // Create the zip buffer
    const zipBuffer = createStoredZip(entries);

    // Write to file
    await writeFile(document.filePath, zipBuffer);

    // Mark as clean after saving
    markClean(handle);

    return okVoid();
  } catch (e) {
    if (e instanceof Error) {
      return fail(e, "operation_failed");
    }
    return err("operation_failed", String(e));
  }
}

// ============================================================================
// Check if Document is Open
// ============================================================================

/**
 * Checks if a document is currently open.
 *
 * @param handle - The document handle to check
 * @returns True if the document is open, false otherwise
 *
 * @example
 * if (isOpen(handle)) {
 *   // Document is still in memory
 * }
 */
export function isOpen(handle: DocumentHandle): boolean {
  return hasDocument(handle);
}

// ============================================================================
// Get Document Info
// ============================================================================

/**
 * Gets information about an open document.
 *
 * @param handle - The document handle
 * @returns Result containing document info or error if not found
 */
export function getInfo(handle: DocumentHandle): Result<{
  handle: DocumentHandle;
  filePath: string;
  dirty: boolean;
  openedAt: Date;
}> {
  const document = getDocument(handle);
  if (!document) {
    return notFound("Document", handle, "The document may already be closed or was never opened");
  }

  return ok({
    handle: document.handle,
    filePath: document.filePath,
    dirty: document.dirty,
    openedAt: document.openedAt,
  });
}

/**
 * Gets the zip contents for a document.
 *
 * @param handle - The document handle
 * @returns Result containing the zip Map or error if not found
 */
export function getZip(handle: DocumentHandle): Result<Map<string, Buffer>> {
  const document = getDocument(handle);
  if (!document) {
    return notFound("Document", handle, "The document may already be closed or was never opened");
  }

  return ok(document.zip);
}

/**
 * Gets the file path for a document.
 *
 * @param handle - The document handle
 * @returns Result containing the file path or error if not found
 */
export function getFilePath(handle: DocumentHandle): Result<string> {
  const document = getDocument(handle);
  if (!document) {
    return notFound("Document", handle, "The document may already be closed or was never opened");
  }

  return ok(document.filePath);
}

/**
 * Marks a document as modified.
 *
 * @param handle - The document handle
 */
export function setDirty(handle: DocumentHandle): void {
  markDirty(handle);
}

/**
 * Checks if a document has unsaved changes.
 *
 * @param handle - The document handle
 * @returns True if document has unsaved changes
 */
export function isDirty(handle: DocumentHandle): boolean {
  const document = getDocument(handle);
  return document?.dirty ?? false;
}

// ============================================================================
// Internal: Update zip contents (used by mutations)
// ============================================================================

/**
 * Updates the zip contents for a document.
 * This is used internally by mutation operations.
 *
 * @param handle - The document handle
 * @param zip - The new zip contents
 */
export function updateZip(handle: DocumentHandle, zip: Map<string, Buffer>): void {
  const document = getDocument(handle);
  if (document) {
    document.zip = zip;
    markDirty(handle);
  }
}
