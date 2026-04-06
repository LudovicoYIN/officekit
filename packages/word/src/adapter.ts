/**
 * Word adapter for @officekit/word.
 *
 * This module provides Get and Query functions for Word documents.
 * It reads docx files (ZIP archives containing XML) and parses the XML
 * to extract document structure, text, and formatting.
 *
 * @example
 * import { getWordNode, queryWordNodes } from "./adapter.js";
 *
 * // Get a specific node by path
 * const result = await getWordNode("document.docx", "/body/p[1]", 1);
 *
 * // Query nodes using a selector
 * const paragraphs = await queryWordNodes("document.docx", "p");
 * const boldRuns = await queryWordNodes("document.docx", "r[bold=true]");
 */

import { readFile } from "node:fs/promises";
import JSZip from "jszip";

import { err, ok } from "./result.js";
import { parsePath, buildPath } from "./path.js";
import { parseSelector } from "./selectors.js";
import type { Result, DocumentNode, PathSegment } from "./types.js";

// ============================================================================
// ZIP Helpers
// ============================================================================

async function readDocxZip(filePath: string): Promise<JSZip> {
  const buffer = await readFile(filePath);
  return await JSZip.loadAsync(buffer);
}

async function getXmlEntry(zip: JSZip, entryName: string): Promise<string | null> {
  const entry = zip.file(entryName);
  if (!entry) return null;
  return await entry.async("string");
}

// ============================================================================
// XML Text Extraction Helpers
// ============================================================================

/**
 * Extracts all text content from an XML string.
 */
function extractTextFromXml(xml: string): string {
  const texts: string[] = [];
  const regex = /<[^>]*:t[^>]*>([^<]*)<\/[^>]*:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    texts.push(match[1]);
  }
  return texts.join("");
}

/**
 * Extracts text content from w:t elements in a more robust way.
 */
function extractTextSimple(xml: string): string {
  const texts: string[] = [];
  const regex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    texts.push(match[1]);
  }
  return texts.join("");
}

/**
 * Gets all paragraph texts from document XML.
 */
function getParagraphsInfo(xml: string): Array<{ index: number; text: string; style?: string; paraId?: string }> {
  const paragraphs: Array<{ index: number; text: string; style?: string; paraId?: string }> = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;
  while ((match = paraRegex.exec(xml)) !== null) {
    idx++;
    const paraXml = match[0];
    const text = extractTextSimple(paraXml);

    let style: string | undefined;
    let paraId: string | undefined;

    const styleMatch = paraXml.match(/<w:pStyle[^>]*w:val="([^"]*)"/);
    if (styleMatch) style = styleMatch[1];

    const paraIdMatch = paraXml.match(/<w:paraId[^>]*w:val="([^"]*)"/);
    if (paraIdMatch) paraId = paraIdMatch[1];

    paragraphs.push({ index: idx, text, style, paraId });
  }

  return paragraphs;
}

/**
 * Gets all table info from document XML.
 */
function getTablesInfo(xml: string): Array<{ index: number; rows: number; cols: number }> {
  const tables: Array<{ index: number; rows: number; cols: number }> = [];

  const tblRegex = /<w:tbl[\\s\\S]*?<\\/w:tbl>/g;
  let match;
  let idx = 0;
  while ((match = tblRegex.exec(xml)) !== null) {
    idx++;
    const tblXml = match[0];
    const rows = (tblXml.match(/<w:tr[\\s\\S]*?<\\/w:tr>/g) || []).length;
    const firstRow = tblXml.match(/<w:tr[\\s\\S]*?<\\/w:tr>/);
    const cols = firstRow ? (firstRow[0].match(/<w:tc[\\s\\S]*?<\\/w:tc>/g) || []).length : 0;
    tables.push({ index: idx, rows, cols });
  }

  return tables;
}

// ============================================================================
// Document Node Helpers
// ============================================================================

function createDocumentNode(path: string, type: string, text?: string, format?: Record<string, unknown>): DocumentNode {
  return {
    path,
    type,
    text,
    format: format || {},
  };
}

function createErrorNode(path: string, message: string): DocumentNode {
  return {
    path,
    type: "error",
    text: message,
    format: {},
  };
}

// ============================================================================
// Get Word Node
// ============================================================================

/**
 * Gets a node at the specified path from a Word document.
 *
 * @param filePath - Path to the .docx file
 * @param path - Path to the node (e.g., "/body/p[1]", "/body/tbl[1]/tr[1]/tc[2]")
 * @param depth - How deep to fetch children (0 = just this node, 1 = one level, etc.)
 * @returns Result containing the DocumentNode or error
 *
 * @example
 * const result = await getWordNode("document.docx", "/body", 1);
 * if (result.ok) {
 *   console.log(result.data.path);  // "/body"
 *   console.log(result.data.children?.length);  // number of children
 * }
 */
export async function getWordNode(filePath: string, path: string, depth = 1): Promise<Result<DocumentNode>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const parsed = parsePath(path);
    if (!parsed.ok) {
      return err("invalid_path", parsed.error?.message || "Invalid path");
    }

    const segments = parsed.data?.segments || [];
    const result = navigateToElement(documentXml, zip, segments, depth);

    if (!result) {
      return err("not_found", `Path not found: ${path}`);
    }

    return ok(result);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Navigates to an element based on path segments.
 */
function navigateToElement(
  documentXml: string,
  zip: JSZip,
  segments: PathSegment[],
  depth: number,
  parentPath = "",
): DocumentNode | null {
  if (segments.length === 0) {
    return createDocumentNode("/", "document");
  }

  const first = segments[0];
  let currentPath = "/" + first.name + (first.index !== undefined ? `[${first.index}]` : "");
  let currentNode: DocumentNode | null = null;

  switch (first.name) {
    case "body": {
      if (segments.length === 1) {
        const paras = getParagraphsInfo(documentXml);
        const tables = getTablesInfo(documentXml);
        const children: DocumentNode[] = [];

        for (let i = 0; i < paras.length; i++) {
          children.push(createDocumentNode(
            `/body/p[${i + 1}]`,
            "paragraph",
            paras[i].text,
            { style: paras[i].style, paraId: paras[i].paraId }
          ));
        }
        for (let i = 0; i < tables.length; i++) {
          children.push(createDocumentNode(
            `/body/tbl[${i + 1}]`,
            "table",
            undefined,
            { rowCount: tables[i].rows, columnCount: tables[i].cols }
          ));
        }

        currentNode = createDocumentNode("/body", "body");
        if (depth > 0) {
          currentNode.children = children;
          currentNode.childCount = children.length;
        }
      }
      break;
    }

    case "p":
    case "paragraph": {
      const paras = getParagraphsInfo(documentXml);
      const idx = (first.index || 1) - 1;
      if (idx >= 0 && idx < paras.length) {
        const para = paras[idx];
        currentPath = `/body/p[${idx + 1}]`;
        currentNode = createDocumentNode(
          currentPath,
          "paragraph",
          para.text,
          { style: para.style, paraId: para.paraId }
        );

        if (depth > 0) {
          const runs = getRunsFromParagraph(documentXml, idx + 1);
          currentNode.children = runs;
          currentNode.childCount = runs.length;
        }
      }
      break;
    }

    case "tbl":
    case "table": {
      const tables = getTablesInfo(documentXml);
      const idx = (first.index || 1) - 1;
      if (idx >= 0 && idx < tables.length) {
        const table = tables[idx];
        currentPath = `/body/tbl[${idx + 1}]`;
        currentNode = createDocumentNode(
          currentPath,
          "table",
          undefined,
          { rowCount: table.rows, columnCount: table.cols }
        );

        if (depth > 0) {
          const rows: DocumentNode[] = [];
          for (let i = 0; i < table.rows; i++) {
            rows.push(createDocumentNode(
              `/body/tbl[${idx + 1}]/tr[${i + 1}]`,
              "row",
              undefined,
              { cellCount: table.cols }
            ));
          }
          currentNode.children = rows;
          currentNode.childCount = rows.length;
        }
      }
      break;
    }

    case "header": {
      const headerIdx = (first.index || 1) - 1;
      const headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
      if (headerEntry) {
        const headerXml = await headerEntry.async("string");
        const text = extractTextSimple(headerXml);
        currentNode = createDocumentNode(
          `/header[${headerIdx + 1}]`,
          "header",
          text
        );
      }
      break;
    }

    case "footer": {
      const footerIdx = (first.index || 1) - 1;
      const footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
      if (footerEntry) {
        const footerXml = await footerEntry.async("string");
        const text = extractTextSimple(footerXml);
        currentNode = createDocumentNode(
          `/footer[${footerIdx + 1}]`,
          "footer",
          text
        );
      }
      break;
    }

    case "styles": {
      const stylesXml = await getXmlEntry(zip, "word/styles.xml");
      if (stylesXml) {
        const styles = parseStyles(stylesXml);
        currentNode = createDocumentNode("/styles", "styles");
        if (depth > 0) {
          currentNode.children = styles;
          currentNode.childCount = styles.length;
        }
      }
      break;
    }

    case "numbering": {
      const numberingXml = await getXmlEntry(zip, "word/numbering.xml");
      if (numberingXml) {
        currentNode = createDocumentNode("/numbering", "numbering");
      }
      break;
    }

    case "settings": {
      const settingsXml = await getXmlEntry(zip, "word/settings.xml");
      if (settingsXml) {
        currentNode = createDocumentNode("/settings", "settings");
      }
      break;
    }

    default: {
      break;
    }
  }

  if (!currentNode) {
    return null;
  }

  if (segments.length > 1 && currentNode) {
    const remainingPath = segments.slice(1);
    const childPath = buildChildPath(currentNode.path, remainingPath);

    if (remainingPath.length === 1 && remainingPath[0].name === "tr") {
      const rowIdx = (remainingPath[0].index || 1) - 1;
      const rowPath = `${currentNode.path}/tr[${rowIdx + 1}]`;
      return createDocumentNode(rowPath, "row");
    }

    if (remainingPath.length === 2 &&
        (remainingPath[0].name === "tr" || remainingPath[0].name === "row") &&
        (remainingPath[1].name === "tc" || remainingPath[1].name === "cell")) {
      const rowIdx = (remainingPath[0].index || 1) - 1;
      const cellIdx = (remainingPath[1].index || 1) - 1;
      const cellPath = `${currentNode.path}/tr[${rowIdx + 1}]/tc[${cellIdx + 1}]`;
      return createDocumentNode(cellPath, "cell");
    }
  }

  return currentNode;
}

function buildChildPath(parentPath: string, segments: PathSegment[]): string {
  if (segments.length === 0) return parentPath;

  const seg = segments[0];
  let path = parentPath;
  if (seg.index !== undefined) {
    path += `/${seg.name}[${seg.index}]`;
  } else if (seg.stringIndex !== undefined) {
    path += `/${seg.name}[${seg.stringIndex}]`;
  } else {
    path += `/${seg.name}`;
  }

  if (segments.length > 1) {
    path = buildChildPath(path, segments.slice(1));
  }

  return path;
}

/**
 * Gets runs from a specific paragraph.
 */
function getRunsFromParagraph(documentXml: string, paraIndex: number): DocumentNode[] {
  const runs: DocumentNode[] = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let idx = 0;

  while ((match = paraRegex.exec(documentXml)) !== null) {
    idx++;
    if (idx !== paraIndex) continue;

    const paraXml = match[0];
    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    let runMatch;
    let runIdx = 0;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      runIdx++;
      const runXml = runMatch[0];
      const text = extractTextSimple(runXml);

      const format: Record<string, unknown> = {};
      if (runXml.includes("<w:b/>") || runXml.includes("<w:b ")) format.bold = true;
      if (runXml.includes("<w:i/>") || runXml.includes("<w:i ")) format.italic = true;
      if (runXml.includes("<w:u ")) format.underline = "single";
      if (runXml.includes("<w:strike/>") || runXml.includes("<w:strike ")) format.strike = true;

      const fontMatch = runXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
      if (fontMatch) format.font = fontMatch[1];

      const sizeMatch = runXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
      if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

      const colorMatch = runXml.match(/<w:color[^>]*w:val="([^"]*)"/);
      if (colorMatch) format.color = colorMatch[1];

      runs.push(createDocumentNode(
        `/body/p[${paraIndex}]/r[${runIdx}]`,
        "run",
        text,
        format
      ));
    }
    break;
  }

  return runs;
}

/**
 * Parses styles from styles.xml.
 */
function parseStyles(stylesXml: string): DocumentNode[] {
  const styles: DocumentNode[] = [];

  const styleRegex = /<w:style[^>]*>([\\s\\S]*?)<\\/w:style>/g;
  let match;
  let idx = 0;

  while ((match = styleRegex.exec(stylesXml)) !== null) {
    idx++;
    const styleXml = match[0];

    const styleIdMatch = styleXml.match(/w:styleId="([^"]*)"/);
    const styleId = styleIdMatch ? styleIdMatch[1] : `style${idx}`;

    const nameMatch = styleXml.match(/<w:name[^>]*w:val="([^"]*)"/);
    const name = nameMatch ? nameMatch[1] : styleId;

    const typeMatch = styleXml.match(/w:type="([^"]*)"/);
    const type = typeMatch ? typeMatch[1] : "paragraph";

    const format: Record<string, unknown> = { id: styleId, name, type };

    const fontMatch = styleXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
    if (fontMatch) format.font = fontMatch[1];

    const sizeMatch = styleXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
    if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

    if (styleXml.includes("<w:b/>") || styleXml.includes("<w:b ")) format.bold = true;
    if (styleXml.includes("<w:i/>") || styleXml.includes("<w:i ")) format.italic = true;

    const colorMatch = styleXml.match(/<w:color[^>]*w:val="([^"]*)"/);
    if (colorMatch) format.color = colorMatch[1];

    styles.push(createDocumentNode(
      `/styles/${styleId}`,
      "style",
      name,
      format
    ));
  }

  return styles;
}

// ============================================================================
// Query Word Nodes
// ============================================================================

/**
 * Queries nodes using a selector from a Word document.
 *
 * @param filePath - Path to the .docx file
 * @param selector - CSS-like selector (e.g., "p", "p[style=Heading1]", "r[bold=true]")
 * @returns Result containing an array of DocumentNodes or error
 *
 * @example
 * const result = await queryWordNodes("document.docx", "p");
 * if (result.ok) {
 *   console.log(result.data.length);  // number of paragraphs
 *   console.log(result.data[0].text);  // first paragraph text
 * }
 *
 * @example
 * // Query all bold runs
 * const boldRuns = await queryWordNodes("document.docx", "r[bold=true]");
 *
 * @example
 * // Query paragraphs containing specific text
 * const matches = await queryWordNodes("document.docx", 'p:contains("Hello")');
 */
export async function queryWordNodes(filePath: string, selector: string): Promise<Result<DocumentNode[]>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const parsed = parseSelector(selector);
    if (!parsed.ok) {
      return err("invalid_selector", parsed.error?.message || "Invalid selector");
    }

    const selectorData = parsed.data!;
    const results: DocumentNode[] = [];

    const elementType = selectorData.element || "p";

    switch (elementType) {
      case "p":
      case "paragraph": {
        const paras = getParagraphsInfo(documentXml);
        for (let i = 0; i < paras.length; i++) {
          const para = paras[i];
          if (!matchesSelectorAttributes(para, selectorData.attributes, documentXml, i + 1)) {
            continue;
          }
          if (selectorData.containsText && !para.text.includes(selectorData.containsText)) {
            continue;
          }

          const node = createDocumentNode(
            `/body/p[${i + 1}]`,
            "paragraph",
            para.text,
            { style: para.style, paraId: para.paraId }
          );
          results.push(node);
        }
        break;
      }

      case "r":
      case "run": {
        const runs = getAllRuns(documentXml);
        for (let i = 0; i < runs.length; i++) {
          const run = runs[i];
          if (!matchesRunAttributes(run, selectorData.attributes)) {
            continue;
          }
          if (selectorData.containsText && !run.text.includes(selectorData.containsText)) {
            continue;
          }

          results.push(createDocumentNode(
            run.path,
            "run",
            run.text,
            run.format
          ));
        }
        break;
      }

      case "tbl":
      case "table": {
        const tables = getTablesInfo(documentXml);
        for (let i = 0; i < tables.length; i++) {
          const table = tables[i];
          results.push(createDocumentNode(
            `/body/tbl[${i + 1}]`,
            "table",
            undefined,
            { rowCount: table.rows, columnCount: table.cols }
          ));
        }
        break;
      }

      case "tr":
      case "row": {
        const tables = getTablesInfo(documentXml);
        for (let t = 0; t < tables.length; t++) {
          for (let r = 0; r < tables[t].rows; r++) {
            results.push(createDocumentNode(
              `/body/tbl[${t + 1}]/tr[${r + 1}]`,
              "row",
              undefined,
              { cellCount: tables[t].cols }
            ));
          }
        }
        break;
      }

      case "tc":
      case "cell": {
        const tables = getTablesInfo(documentXml);
        for (let t = 0; t < tables.length; t++) {
          for (let r = 0; r < tables[t].rows; r++) {
            for (let c = 0; c < tables[t].cols; c++) {
              results.push(createDocumentNode(
                `/body/tbl[${t + 1}]/tr[${r + 1}]/tc[${c + 1}]`,
                "cell"
              ));
            }
          }
        }
        break;
      }

      case "header": {
        let headerIdx = 0;
        let headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
        while (headerEntry) {
          const headerXml = await headerEntry.async("string");
          const text = extractTextSimple(headerXml);
          if (!selectorData.containsText || text.includes(selectorData.containsText)) {
            results.push(createDocumentNode(
              `/header[${headerIdx + 1}]`,
              "header",
              text
            ));
          }
          headerIdx++;
          headerEntry = zip.file(`word/header${headerIdx + 1}.xml`);
        }
        break;
      }

      case "footer": {
        let footerIdx = 0;
        let footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
        while (footerEntry) {
          const footerXml = await footerEntry.async("string");
          const text = extractTextSimple(footerXml);
          if (!selectorData.containsText || text.includes(selectorData.containsText)) {
            results.push(createDocumentNode(
              `/footer[${footerIdx + 1}]`,
              "footer",
              text
            ));
          }
          footerIdx++;
          footerEntry = zip.file(`word/footer${footerIdx + 1}.xml`);
        }
        break;
      }

      case "style":
      case "styles": {
        const stylesXml = await getXmlEntry(zip, "word/styles.xml");
        if (stylesXml) {
          const styles = parseStyles(stylesXml);
          for (const style of styles) {
            if (!selectorData.containsText || style.text?.includes(selectorData.containsText)) {
              results.push(style);
            }
          }
        }
        break;
      }

      case "bookmark": {
        const bookmarks = getBookmarks(documentXml);
        for (const bookmark of bookmarks) {
          results.push(createDocumentNode(
            bookmark.path,
            "bookmark",
            bookmark.text,
            { name: bookmark.name, id: bookmark.id }
          ));
        }
        break;
      }

      case "sdt":
      case "contentcontrol": {
        const sdts = getContentControls(documentXml);
        for (const sdt of sdts) {
          results.push(createDocumentNode(
            sdt.path,
            "sdt",
            sdt.text,
            sdt.format
          ));
        }
        break;
      }

      default: {
        break;
      }
    }

    return ok(results);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Checks if a paragraph matches selector attributes.
 */
function matchesSelectorAttributes(
  para: { index: number; text: string; style?: string; paraId?: string },
  attrs: Record<string, string>,
  documentXml: string,
  paraIndex: number,
): boolean {
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "empty") {
      if (value === "true" && para.text.trim().length > 0) return false;
      if (value === "false" && para.text.trim().length === 0) return false;
      continue;
    }

    if (key === "style") {
      const styleMatch = para.style === value;
      if (!styleMatch) return false;
      continue;
    }

    if (key === "index") {
      if (parseInt(value) !== para.index) return false;
      continue;
    }

    if (key.startsWith("@")) {
      continue;
    }

    if (key === "numId" || key === "numid") {
      continue;
    }
  }
  return true;
}

/**
 * Checks if a run matches selector attributes.
 */
function matchesRunAttributes(
  run: { text: string; format: Record<string, unknown> },
  attrs: Record<string, string>,
): boolean {
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "empty") {
      if (value === "true" && run.text.trim().length > 0) return false;
      if (value === "false" && run.text.trim().length === 0) return false;
      continue;
    }

    if (key === "bold") {
      const isBold = run.format.bold === true;
      const shouldBeBold = value === "true";
      if (isBold !== shouldBeBold) return false;
      continue;
    }

    if (key === "italic") {
      const isItalic = run.format.italic === true;
      const shouldBeItalic = value === "true";
      if (isItalic !== shouldBeItalic) return false;
      continue;
    }

    if (key === "underline") {
      const hasUnderline = run.format.underline !== undefined;
      if (value === "true" && !hasUnderline) return false;
      if (value !== "true" && hasUnderline && run.format.underline !== value) return false;
      continue;
    }

    if (key === "strike") {
      const hasStrike = run.format.strike === true;
      const shouldBeStruck = value === "true";
      if (hasStrike !== shouldBeStruck) return false;
      continue;
    }

    if (key === "font") {
      const font = run.format.font as string | undefined;
      if (!font || !font.toLowerCase().includes(value.toLowerCase())) return false;
      continue;
    }

    if (key === "size") {
      const size = run.format.size as string | undefined;
      if (!size) return false;
      const sizeNum = parseFloat(size);
      const targetNum = parseFloat(value);
      if (isNaN(sizeNum) || isNaN(targetNum)) return false;
      if (Math.abs(sizeNum - targetNum) > 0.1) return false;
      continue;
    }

    if (key === "color") {
      const color = run.format.color as string | undefined;
      if (!color) return false;
      if (color.toLowerCase() !== value.toLowerCase()) return false;
      continue;
    }
  }
  return true;
}

/**
 * Gets all runs from the document.
 */
function getAllRuns(documentXml: string): Array<{
  path: string;
  text: string;
  format: Record<string, unknown>;
  paraIndex: number;
  runIndex: number;
}> {
  const runs: Array<{
    path: string;
    text: string;
    format: Record<string, unknown>;
    paraIndex: number;
    runIndex: number;
  }> = [];

  const paraRegex = /<w:p[\\s\\S]*?<\\/w:p>/g;
  let match;
  let paraIdx = 0;

  while ((match = paraRegex.exec(documentXml)) !== null) {
    paraIdx++;
    const paraXml = match[0];
    const runRegex = /<w:r[\\s\\S]*?<\\/w:r>/g;
    let runMatch;
    let runIdx = 0;

    while ((runMatch = runRegex.exec(paraXml)) !== null) {
      runIdx++;
      const runXml = runMatch[0];
      const text = extractTextSimple(runXml);

      const format: Record<string, unknown> = {};
      if (runXml.includes("<w:b/>") || runXml.includes("<w:b ")) format.bold = true;
      if (runXml.includes("<w:i/>") || runXml.includes("<w:i ")) format.italic = true;
      if (runXml.includes("<w:u ")) format.underline = "single";
      if (runXml.includes("<w:strike/>") || runXml.includes("<w:strike ")) format.strike = true;

      const fontMatch = runXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
      if (fontMatch) format.font = fontMatch[1];

      const sizeMatch = runXml.match(/<w:sz[^>]*w:val="([^"]*)"/);
      if (sizeMatch) format.size = `${parseInt(sizeMatch[1]) / 2}pt`;

      const colorMatch = runXml.match(/<w:color[^>]*w:val="([^"]*)"/);
      if (colorMatch) format.color = colorMatch[1];

      runs.push({
        path: `/body/p[${paraIdx}]/r[${runIdx}]`,
        text,
        format,
        paraIndex: paraIdx,
        runIndex: runIdx,
      });
    }
  }

  return runs;
}

/**
 * Gets bookmarks from the document.
 */
function getBookmarks(documentXml: string): Array<{ path: string; name: string; id: string; text: string }> {
  const bookmarks: Array<{ path: string; name: string; id: string; text: string }> = [];

  const bookmarkStartRegex = /<w:bookmarkStart[^>]*w:id="([^"]*)"[^>]*w:name="([^"]*)"[^>]*>/g;
  let match;

  while ((match = bookmarkStartRegex.exec(documentXml)) !== null) {
    const id = match[1];
    const name = match[2];

    const startIdx = match.index;
    const endIdx = documentXml.indexOf("</w:bookmarkEnd>", startIdx);
    const bookmarkContent = documentXml.slice(startIdx, endIdx > 0 ? endIdx + 16 : undefined);
    const text = extractTextSimple(bookmarkContent);

    bookmarks.push({
      path: `/bookmark[${name}]`,
      name,
      id,
      text,
    });
  }

  return bookmarks;
}

/**
 * Gets content controls (SDT) from the document.
 */
function getContentControls(documentXml: string): Array<{
  path: string;
  text: string;
  format: Record<string, unknown>;
}> {
  const sdts: Array<{ path: string; text: string; format: Record<string, unknown> }> = [];

  const sdtRegex = /<w:sdt[\\s\\S]*?<\\/w:sdt>/g;
  let match;
  let idx = 0;

  while ((match = sdtRegex.exec(documentXml)) !== null) {
    idx++;
    const sdtXml = match[0];
    const text = extractTextSimple(sdtXml);

    const format: Record<string, unknown> = {};
    const tagMatch = sdtXml.match(/<w:tag[^>]*w:val="([^"]*)"/);
    if (tagMatch) format.tag = tagMatch[1];

    sdts.push({
      path: `/body/sdt[${idx}]`,
      text,
      format,
    });
  }

  return sdts;
}

// ============================================================================
// Document Info
// ============================================================================

/**
 * Gets basic document information without deep traversal.
 */
export async function getDocumentInfo(filePath: string): Promise<Result<DocumentNode>> {
  try {
    const zip = await readDocxZip(filePath);
    const documentXml = await getXmlEntry(zip, "word/document.xml");

    if (!documentXml) {
      return err("not_found", "Document.xml not found in docx archive");
    }

    const paras = getParagraphsInfo(documentXml);
    const tables = getTablesInfo(documentXml);

    let headerCount = 0;
    let footerCount = 0;

    let entry = zip.file(`word/header${headerCount + 1}.xml`);
    while (entry) {
      headerCount++;
      entry = zip.file(`word/header${headerCount + 1}.xml`);
    }

    entry = zip.file(`word/footer${footerCount + 1}.xml`);
    while (entry) {
      footerCount++;
      entry = zip.file(`word/footer${footerCount + 1}.xml`);
    }

    const node = createDocumentNode("/", "document");
    node.childCount = 1;
    node.children = [createDocumentNode("/body", "body", undefined, {
      paragraphCount: paras.length,
      tableCount: tables.length,
      headerCount,
      footerCount,
    })];
    node.format = {
      paragraphCount: paras.length,
      tableCount: tables.length,
      headerCount,
      footerCount,
    };

    return ok(node);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
