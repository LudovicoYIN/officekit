import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { renderHtmlPreview } from "./src/html-preview/index.js";

const filePath = "/Users/llm/Desktop/OpenClaw医学培训班大纲-批注版本.docx";
const buffer = await readFile(filePath);
const zip = await JSZip.loadAsync(buffer);

const documentXml = await zip.file("word/document.xml").async("string");
const stylesXml = await zip.file("word/styles.xml").async("string");

const html = await renderHtmlPreview(zip, documentXml, stylesXml);

await writeFile("/Users/llm/Desktop/preview.html", html);
console.log("HTML preview saved to /Users/llm/Desktop/preview.html");
