import { extname } from "node:path";

export const supportedFormats = ["word", "excel", "powerpoint"] as const;
export type SupportedFormat = (typeof supportedFormats)[number];

const extensionToFormat: Record<string, SupportedFormat> = {
  ".docx": "word",
  ".xlsx": "excel",
  ".pptx": "powerpoint",
};

export function detectFormat(filePath: string): SupportedFormat | null {
  return extensionToFormat[extname(filePath).toLowerCase()] ?? null;
}

export function assertFormat(filePath: string): SupportedFormat {
  const format = detectFormat(filePath);
  if (!format) {
    throw new Error(`Unsupported file extension for '${filePath}'. Expected .docx, .xlsx, or .pptx.`);
  }
  return format;
}
