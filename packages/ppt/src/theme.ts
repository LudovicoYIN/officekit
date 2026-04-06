/**
 * Theme operations for @officekit/ppt.
 *
 * Provides functions to manage themes in PowerPoint presentations:
 * - Get theme colors and fonts
 * - Get/set theme color values
 * - Get theme fonts
 * - Apply a different theme
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createStoredZip, readStoredZip } from "../../core/src/zip.js";
import { err, ok, invalidInput } from "./result.js";
import type { Result } from "./types.js";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents theme colors.
 */
export interface ThemeColors {
  /** Dark 1 color */
  dark1?: string;
  /** Dark 2 color */
  dark2?: string;
  /** Light 1 color */
  light1?: string;
  /** Light 2 color */
  light2?: string;
  /** Accent 1 color */
  accent1?: string;
  /** Accent 2 color */
  accent2?: string;
  /** Accent 3 color */
  accent3?: string;
  /** Accent 4 color */
  accent4?: string;
  /** Accent 5 color */
  accent5?: string;
  /** Accent 6 color */
  accent6?: string;
  /** Hyperlink color */
  hyperlink?: string;
  /** Followed hyperlink color */
  followedHyperlink?: string;
}

/**
 * Represents theme fonts.
 */
export interface ThemeFonts {
  /** Major font (headings) */
  major?: string;
  /** Minor font (body text) */
  minor?: string;
}

/**
 * Represents a complete theme.
 */
export interface ThemeInfo {
  /** Theme name */
  name?: string;
  /** Theme colors */
  colors?: ThemeColors;
  /** Theme fonts */
  fonts?: ThemeFonts;
}

/**
 * Theme color index.
 */
export type ThemeColorIndex =
  | "dk1" | "dark1"
  | "dk2" | "dark2"
  | "lt1" | "light1"
  | "lt2" | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedhyperlink";

/**
 * Theme font type.
 */
export type ThemeFontType = "major" | "minor" | "heading" | "body";

// ============================================================================
// Helpers
// ============================================================================

/**
 * Reads an entry from the zip as a string.
 * Throws an error if the entry is not found.
 */
function requireEntry(zip: Map<string, Buffer>, entryName: string): string {
  const buffer = zip.get(entryName);
  if (!buffer) {
    throw new Error(`OOXML entry '${entryName}' is missing`);
  }
  return buffer.toString("utf8");
}

/**
 * Gets an entry from the zip as a string, or null if not found.
 */
function getEntry(zip: Map<string, Buffer>, entryName: string): string | null {
  const buffer = zip.get(entryName);
  if (!buffer) {
    return null;
  }
  return buffer.toString("utf8");
}

/**
 * Parses theme colors from theme XML.
 */
function parseThemeColors(themeXml: string): ThemeColors {
  const colors: ThemeColors = {};

  // Helper to extract color value
  const extractColor = (elementName: string): string | undefined => {
    // Try srgbClr first
    const srgbMatch = new RegExp(`<a:${elementName}[^>]*>[\\s\\S]*?<a:srgbClr\\s+val="([^"]*)"[^>]*>[\\s\\S]*?</a:${elementName}>|<a:${elementName}\\s+val="([^"]*)"`, "i").exec(themeXml);
    if (srgbMatch) {
      return srgbMatch[1] || srgbMatch[2];
    }

    // Try sysClr
    const sysMatch = new RegExp(`<a:${elementName}[^>]*>[\\s\\S]*?<a:sysClr[^>]*lastClr="([^"]*)"[^>]*>[\\s\\S]*?</a:${elementName}>`, "i").exec(themeXml);
    if (sysMatch) {
      return sysMatch[1];
    }

    return undefined;
  };

  // Extract all colors
  if (extractColor("dk1")) colors.dark1 = extractColor("dk1");
  if (extractColor("lt1")) colors.light1 = extractColor("lt1");
  if (extractColor("dk2")) colors.dark2 = extractColor("dk2");
  if (extractColor("lt2")) colors.light2 = extractColor("lt2");
  if (extractColor("accent1")) colors.accent1 = extractColor("accent1");
  if (extractColor("accent2")) colors.accent2 = extractColor("accent2");
  if (extractColor("accent3")) colors.accent3 = extractColor("accent3");
  if (extractColor("accent4")) colors.accent4 = extractColor("accent4");
  if (extractColor("accent5")) colors.accent5 = extractColor("accent5");
  if (extractColor("accent6")) colors.accent6 = extractColor("accent6");
  if (extractColor("hlink")) colors.hyperlink = extractColor("hlink");
  if (extractColor("folHlink")) colors.followedHyperlink = extractColor("folHlink");

  return colors;
}

/**
 * Parses theme fonts from theme XML.
 */
function parseThemeFonts(themeXml: string): ThemeFonts {
  const fonts: ThemeFonts = {};

  // Extract major font (heading font)
  const majorMatch = /<a:majorFont[^>]*>[\s\S]*?<a:latin[^>]*typeface="([^"]*)"[^>]*>[\s\S]*?<\/a:majorFont>/.exec(themeXml);
  if (majorMatch) {
    fonts.major = majorMatch[1];
  }

  // Extract minor font (body font)
  const minorMatch = /<a:minorFont[^>]*>[\s\S]*?<a:latin[^>]*typeface="([^"]*)"[^>]*>[\s\S]*?<\/a:minorFont>/.exec(themeXml);
  if (minorMatch) {
    fonts.minor = minorMatch[1];
  }

  return fonts;
}

/**
 * Parses theme name from theme XML.
 */
function parseThemeName(themeXml: string): string | undefined {
  const nameMatch = /<a:theme[^>]*name="([^"]*)"[^>]*>/.exec(themeXml);
  return nameMatch ? nameMatch[1] : undefined;
}

/**
 * Normalizes a color to 6-character hex format.
 */
function normalizeColor(color: string): string {
  // Remove # prefix if present
  color = color.replace(/^#/, "");

  // Expand 3-character hex to 6-character
  if (color.length === 3) {
    color = color[0] + color[0] + color[1] + color[1] + color[2] + color[2];
  }

  return color.toUpperCase();
}

/**
 * Validates a hex color.
 */
function isValidHexColor(color: string): boolean {
  return /^[0-9A-Fa-f]{6}$/.test(color);
}

// ============================================================================
// Theme Operations
// ============================================================================

/**
 * Gets the theme information from a presentation.
 *
 * @param filePath - Path to the PPTX file
 *
 * @example
 * const result = await getTheme("/path/to/presentation.pptx");
 * if (result.ok) {
 *   console.log(result.data.theme.colors);
 *   console.log(result.data.theme.fonts);
 * }
 */
export async function getTheme(
  filePath: string
): Promise<Result<{ theme: ThemeInfo }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Try to find theme in standard location
    let themeXml: string | null = null;
    let themeEntry = "ppt/theme/theme1.xml";

    try {
      themeXml = requireEntry(zip, themeEntry);
    } catch {
      // Try alternative locations
      const entries = Array.from(zip.keys());
      const themeEntryAlt = entries.find(e => e.match(/^ppt\/theme\/theme\d*\.xml$/i));
      if (themeEntryAlt) {
        themeEntry = themeEntryAlt;
        themeXml = requireEntry(zip, themeEntry);
      }
    }

    if (!themeXml) {
      return invalidInput("Theme not found in presentation");
    }

    const theme: ThemeInfo = {
      name: parseThemeName(themeXml),
      colors: parseThemeColors(themeXml),
      fonts: parseThemeFonts(themeXml),
    };

    return ok({ theme });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets a specific theme color value.
 *
 * @param filePath - Path to the PPTX file
 * @param colorIndex - Color index (dk1, dk2, lt1, lt2, accent1-6, hyperlink, followedhyperlink)
 *
 * @example
 * const result = await getThemeColor("/path/to/presentation.pptx", "accent1");
 * if (result.ok) {
 *   console.log(result.data.color); // e.g., "4472C4"
 * }
 */
export async function getThemeColor(
  filePath: string,
  colorIndex: ThemeColorIndex
): Promise<Result<{ color: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find theme entry
    let themeXml: string | null = null;
    const entries = Array.from(zip.keys());
    const themeEntry = entries.find(e => e.match(/^ppt\/theme\/theme\d*\.xml$/i));

    if (!themeEntry) {
      return invalidInput("Theme not found in presentation");
    }

    themeXml = requireEntry(zip, themeEntry);
    const colors = parseThemeColors(themeXml);

    // Map color index to color value
    const colorMap: Record<string, string | undefined> = {
      dk1: colors.dark1,
      dark1: colors.dark1,
      dk2: colors.dark2,
      dark2: colors.dark2,
      lt1: colors.light1,
      light1: colors.light1,
      lt2: colors.light2,
      light2: colors.light2,
      accent1: colors.accent1,
      accent2: colors.accent2,
      accent3: colors.accent3,
      accent4: colors.accent4,
      accent5: colors.accent5,
      accent6: colors.accent6,
      hyperlink: colors.hyperlink,
      followedhyperlink: colors.followedHyperlink,
    };

    const color = colorMap[colorIndex.toLowerCase()];
    if (!color) {
      return invalidInput(`Theme color '${colorIndex}' not found`);
    }

    return ok({ color });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Sets a theme color value.
 *
 * @param filePath - Path to the PPTX file
 * @param colorIndex - Color index (dk1, dk2, lt1, lt2, accent1-6, hyperlink, followedhyperlink)
 * @param value - New color value as 6-character hex (e.g., "FF0000")
 *
 * @example
 * const result = await setThemeColor("/path/to/presentation.pptx", "accent1", "4472C4");
 */
export async function setThemeColor(
  filePath: string,
  colorIndex: ThemeColorIndex,
  value: string
): Promise<Result<void>> {
  try {
    // Validate color format
    const normalizedColor = normalizeColor(value);
    if (!isValidHexColor(normalizedColor)) {
      return invalidInput("Color must be a 6-character hex value (e.g., 'FF0000' or 'F00')");
    }

    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find theme entry
    const entries = Array.from(zip.keys());
    const themeEntry = entries.find(e => e.match(/^ppt\/theme\/theme\d*\.xml$/i));

    if (!themeEntry) {
      return invalidInput("Theme not found in presentation");
    }

    let themeXml = requireEntry(zip, themeEntry);

    // Map color index to element name
    const elementMap: Record<string, string> = {
      dk1: "dk1",
      dark1: "dk1",
      dk2: "dk2",
      dark2: "dk2",
      lt1: "lt1",
      light1: "lt1",
      lt2: "lt2",
      light2: "lt2",
      accent1: "accent1",
      accent2: "accent2",
      accent3: "accent3",
      accent4: "accent4",
      accent5: "accent5",
      accent6: "accent6",
      hyperlink: "hlink",
      followedhyperlink: "folHlink",
    };

    const elementName = elementMap[colorIndex.toLowerCase()];
    if (!elementName) {
      return invalidInput(`Invalid color index '${colorIndex}'`);
    }

    // Replace the color value using regex
    // Match either srgbClr with val attribute or sysClr with lastClr attribute
    const colorPattern = new RegExp(
      `(<a:${elementName}[^>]*>)[\\s\\S]*?(</a:${elementName}>|<a:${elementName}\\s+)`,
      "i"
    );

    // For srgbClr format: <a:accent1><a:srgbClr val="XXXXXX"/></a:accent1>
    const srgbPattern = new RegExp(
      `<a:${elementName}[^>]*>[\\s\\S]*?<a:srgbClr\\s+val="[^"]*"[^>]*>[\\s\\S]*?</a:${elementName}>`,
      "i"
    );

    if (srgbPattern.test(themeXml)) {
      // Replace the existing srgbClr value
      themeXml = themeXml.replace(
        new RegExp(`(<a:srgbClr\\s+val=")[^"]*(")`, "i"),
        `$1${normalizedColor}$2`
      );
    } else {
      // Try sysClr format: <a:dk1><a:sysClr val="windowText" lastClr="XXXXXX"/></a:dk1>
      const sysPattern = new RegExp(
        `<a:${elementName}[^>]*>[\\s\\S]*?<a:sysClr[^>]*lastClr="[^"]*"[^>]*>[\\s\\S]*?</a:${elementName}>`,
        "i"
      );

      if (sysPattern.test(themeXml)) {
        themeXml = themeXml.replace(
          new RegExp(`(lastClr=")[^"]*(")`, "i"),
          `$1${normalizedColor}$2`
        );
      } else {
        // Create new color element
        const newColorXml = `<a:${elementName}><a:srgbClr val="${normalizedColor}"/></a:${elementName}>`;

        // Try to find and replace existing or insert new
        const existingPattern = new RegExp(`<a:${elementName}[^>]*>[\\s\\S]*?</a:${elementName}>|<a:${elementName}\\s+[^>]*/>`, "i");
        if (existingPattern.test(themeXml)) {
          themeXml = themeXml.replace(existingPattern, newColorXml);
        } else {
          // Insert after the last color element in clrScheme
          const lastColorMatch = themeXml.match(/<\/a:folHlink>/i);
          if (lastColorMatch) {
            themeXml = themeXml.replace(/<\/a:folHlink>/i, `${newColorXml}</a:folHlink>`);
          }
        }
      }
    }

    // Build new zip with updated theme
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, dataEntry] of zip.entries()) {
      if (name === themeEntry) {
        newEntries.push({ name, data: Buffer.from(themeXml, "utf8") });
      } else {
        newEntries.push({ name, data: dataEntry });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));

    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Gets a theme font value.
 *
 * @param filePath - Path to the PPTX file
 * @param fontType - Font type (major/minor or heading/body)
 *
 * @example
 * const result = await getThemeFont("/path/to/presentation.pptx", "major");
 * if (result.ok) {
 *   console.log(result.data.font); // e.g., "Calibri Light"
 * }
 */
export async function getThemeFont(
  filePath: string,
  fontType: ThemeFontType
): Promise<Result<{ font: string }>> {
  try {
    const buffer = await readFile(filePath);
    const zip = readStoredZip(buffer);

    // Find theme entry
    const entries = Array.from(zip.keys());
    const themeEntry = entries.find(e => e.match(/^ppt\/theme\/theme\d*\.xml$/i));

    if (!themeEntry) {
      return invalidInput("Theme not found in presentation");
    }

    const themeXml = requireEntry(zip, themeEntry);
    const fonts = parseThemeFonts(themeXml);

    // Map font type to font value
    let font: string | undefined;
    switch (fontType.toLowerCase()) {
      case "major":
      case "heading":
        font = fonts.major;
        break;
      case "minor":
      case "body":
        font = fonts.minor;
        break;
    }

    if (!font) {
      return invalidInput(`Theme font '${fontType}' not found`);
    }

    return ok({ font });
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}

/**
 * Applies a different theme to the presentation.
 *
 * @param filePath - Path to the PPTX file
 * @param themePath - Path to a theme file (.thmx) to apply
 *
 * @example
 * const result = await applyTheme(
 *   "/path/to/presentation.pptx",
 *   "/path/to/new_theme.thmx"
 * );
 */
export async function applyTheme(
  filePath: string,
  themePath: string
): Promise<Result<void>> {
  try {
    // Read the theme file
    const themeBuffer = await readFile(themePath);

    // Check if it's a valid theme file (thmx is a zip)
    let themeZip: Map<string, Buffer>;
    try {
      themeZip = readStoredZip(themeBuffer);
    } catch {
      return invalidInput("Invalid theme file - must be a valid .thmx file");
    }

    // Find the theme XML in the thmx (usually in ppt/theme/theme1.xml)
    let themeXml: string | null = null;
    let themeXmlEntry: string | null = null;

    for (const [name, data] of themeZip.entries()) {
      if (name.match(/^ppt\/theme\/theme\d*\.xml$/i)) {
        themeXml = data.toString("utf8");
        themeXmlEntry = name;
        break;
      }
    }

    if (!themeXml || !themeXmlEntry) {
      return invalidInput("Theme file does not contain a valid theme");
    }

    // Read the presentation
    const presBuffer = await readFile(filePath);
    const presZip = readStoredZip(presBuffer);

    // Find existing theme entry
    const entries = Array.from(presZip.keys());
    const existingThemeEntry = entries.find(e => e.match(/^ppt\/theme\/theme\d*\.xml$/i));

    if (!existingThemeEntry) {
      return invalidInput("Presentation does not contain a theme");
    }

    // Build new zip with the new theme
    const newEntries: Array<{ name: string; data: Buffer }> = [];

    for (const [name, dataEntry] of presZip.entries()) {
      if (name === existingThemeEntry) {
        // Replace with new theme
        newEntries.push({ name, data: Buffer.from(themeXml!, "utf8") });
      } else {
        newEntries.push({ name, data: dataEntry });
      }
    }

    await writeFile(filePath, createStoredZip(newEntries));

    return ok(void 0);
  } catch (e) {
    if (e instanceof Error) {
      return err("operation_failed", e.message);
    }
    return err("operation_failed", String(e));
  }
}
