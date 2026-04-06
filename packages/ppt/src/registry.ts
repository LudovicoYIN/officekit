/**
 * Document Registry for @officekit/ppt.
 *
 * Maintains a registry of open documents in memory for efficient
 * multiple operations without repeated file I/O.
 */

import type { Buffer } from "node:buffer";

/**
 * Represents an open document in the registry.
 */
export interface OpenDocument {
  /** Unique handle ID */
  handle: DocumentHandle;
  /** Absolute path to the source file */
  filePath: string;
  /** The zip contents as a Map (entry name -> content buffer) */
  zip: Map<string, Buffer>;
  /** Whether the document has been modified since opening */
  dirty: boolean;
  /** When the document was opened */
  openedAt: Date;
}

export type DocumentHandle = string;

/**
 * Internal registry storage.
 */
const registry = new Map<DocumentHandle, OpenDocument>();

/**
 * Handle counter for generating unique IDs.
 */
let handleCounter = 0;

/**
 * Generates a unique document handle.
 */
function generateHandle(): DocumentHandle {
  handleCounter++;
  return `doc-${Date.now()}-${handleCounter}`;
}

/**
 * Gets the registered document for a handle.
 */
export function getDocument(handle: DocumentHandle): OpenDocument | undefined {
  return registry.get(handle);
}

/**
 * Checks if a document is registered under the given handle.
 */
export function hasDocument(handle: DocumentHandle): boolean {
  return registry.has(handle);
}

/**
 * Registers a new document in the registry.
 *
 * @param filePath - The path to the document
 * @param zip - The zip contents
 * @returns The assigned document handle
 */
export function registerDocument(filePath: string, zip: Map<string, Buffer>): DocumentHandle {
  const handle = generateHandle();
  const document: OpenDocument = {
    handle,
    filePath,
    zip,
    dirty: false,
    openedAt: new Date(),
  };
  registry.set(handle, document);
  return handle;
}

/**
 * Marks a document as dirty (modified).
 */
export function markDirty(handle: DocumentHandle): void {
  const doc = registry.get(handle);
  if (doc) {
    doc.dirty = true;
  }
}

/**
 * Marks a document as clean (saved).
 */
export function markClean(handle: DocumentHandle): void {
  const doc = registry.get(handle);
  if (doc) {
    doc.dirty = false;
  }
}

/**
 * Removes a document from the registry.
 *
 * @param handle - The document handle to remove
 * @returns The removed document, if it existed
 */
export function unregisterDocument(handle: DocumentHandle): OpenDocument | undefined {
  const doc = registry.get(handle);
  registry.delete(handle);
  return doc;
}

/**
 * Gets all registered document handles.
 */
export function getAllHandles(): DocumentHandle[] {
  return Array.from(registry.keys());
}

/**
 * Gets the number of open documents.
 */
export function getOpenCount(): number {
  return registry.size;
}

/**
 * Clears all documents from the registry.
 * Use with caution - this does not save changes.
 */
export function clearRegistry(): void {
  registry.clear();
}
