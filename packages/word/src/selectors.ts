/**
 * Selector grammar parser for @officekit/word.
 *
 * Selectors are used to query elements within a Word document using a CSS-like syntax.
 * They can filter by element type, position, attributes, and text content.
 *
 * Selector Syntax:
 * ---------------
 * [basic]     := elementtype[index]
 * [type]      := elementtype[index][@attr=value]  (attribute filter)
 * [text]      := elementtype:contains("text")
 * [combinator]:= ancestor descendant
 *              | parent > child
 * [empty]     := elementtype:empty
 * [compound]  := selector:contains("text")[@attr=value]
 *
 * Examples:
 * ---------
 * - `p` - All paragraphs
 * - `p[1]` - First paragraph
 * - `p[style=Heading1]` - Paragraphs with style Heading1
 * - `p:contains("Hello")` - Paragraphs containing "Hello"
 * - `p:empty` - Empty paragraphs
 * - `p > r` - Runs that are direct children of paragraphs
 * - `r[bold=true]` - Bold runs
 * - `r[font=Arial]` - Runs with Arial font
 * - `tbl` - All tables
 * - `tr` - All table rows
 * - `tc` - All table cells
 * - `sdt` - All content controls
 * - `bookmark` - All bookmarks
 * - `field[fieldType=page]` - Page number fields
 * - `header` - All headers
 * - `footer` - All footers
 * - `style` - All styles
 * - `toc` - All TOCs
 * - `formfield` - All form fields
 * - `editable` - All editable content
 *
 * Attribute Filters:
 * ------------------
 * - `[@style=Name]` - Match by style name or ID
 * - `[@alignment=left]` - Match by alignment
 * - `[@bold=true]` - Match bold runs
 * - `[@italic=true]` - Match italic runs
 * - `[@font=Arial]` - Match by font
 * - `[@size=12pt]` - Match by font size
 * - `[@color=FF0000]` - Match by color
 * - `[@underline=single]` - Match underlined runs
 * - `[@numId=1]` - Match list paragraphs by numId
 * - `[@paraId=XXX]` - Match by paragraph stable ID
 * - `[@textId=XXX]` - Match by text stable ID
 * - `[@sdtId=XXX]` - Match by content control ID
 * - `[@commentId=XXX]` - Match by comment ID
 *
 * Negation:
 * ---------
 * Use `!=` instead of `=` to negate:
 * - `p[style!=Heading1]` - Paragraphs NOT using Heading1 style
 */

import { err, ok, andThen } from "./result.js";
import type { ParsedSelector, Result } from "./types.js";

// ============================================================================
// Selector Pattern Constants
// ============================================================================

const SELECTOR_PATTERNS = {
  INDEXED: /^([a-zA-Z]+)\[(\d+)\]/,
  NAMED: /^([a-zA-Z]+)\[([a-zA-Z][a-zA-Z0-9_]*)\]/,
  ATTRIBUTE: /^([a-zA-Z]+)\[@([a-zA-Z]+)=([^\]]+)\]/,
  CONTAINS: /^:contains\(\s*(["']?)(.*?)\1\s*\)/,
  EMPTY: /^:empty\b/,
  CHILD_COMBINATOR: /^\s*>\s*/,
  DESCENDANT_COMBINATOR: /^\s+/,
} as const;

const KNOWN_ELEMENT_TYPES = new Set([
  "p",
  "paragraph",
  "r",
  "run",
  "tbl",
  "table",
  "tr",
  "row",
  "tc",
  "cell",
  "sdt",
  "contentcontrol",
  "bookmark",
  "header",
  "footer",
  "footnote",
  "endnote",
  "style",
  "styles",
  "toc",
  "tableofcontents",
  "field",
  "formfield",
  "editable",
  "media",
  "picture",
  "image",
  "img",
  "chart",
  "hyperlink",
  "comment",
  "section",
  "oMath",
  "oMathPara",
  "math",
  "equation",
]);

const PATH_ALIASES: Record<string, string> = {
  paragraph: "p",
  table: "tbl",
  row: "tr",
  cell: "tc",
  contentcontrol: "sdt",
  tableofcontents: "toc",
};

// ============================================================================
// Selector Parsing
// ============================================================================

/**
 * Parses a selector string into a ParsedSelector structure.
 *
 * @example
 * parseSelector("p[1]")
 * // Returns: { element: "p", attributes: { index: "1" } }
 *
 * @example
 * parseSelector('p:contains("Hello")')
 * // Returns: { element: "p", attributes: {}, containsText: "Hello" }
 */
export function parseSelector(selector: string): Result<ParsedSelector> {
  if (!selector || typeof selector !== "string") {
    return err("invalid_selector", "Selector must be a non-empty string");
  }

  const result: ParsedSelector = {
    attributes: {},
  };

  let remaining = selector.trim();

  let element: string | undefined;
  let attributes: Record<string, string> = {};

  while (remaining.length > 0) {
    let matched = false;

    const indexedMatch = remaining.match(SELECTOR_PATTERNS.INDEXED);
    if (indexedMatch && !element) {
      const type = normalizeElementName(indexedMatch[1]);
      const index = indexedMatch[2];
      element = type;
      attributes.index = index;
      remaining = remaining.slice(indexedMatch[0].length);
      matched = true;
    }

    const attrMatch = remaining.match(SELECTOR_PATTERNS.ATTRIBUTE);
    if (attrMatch) {
      const type = attrMatch[1].toLowerCase();
      const attrName = attrMatch[2];
      const attrValue = attrMatch[3];

      if (!element) {
        element = normalizeElementName(type);
      }
      attributes[attrName] = attrValue;
      remaining = remaining.slice(attrMatch[0].length);
      matched = true;
    }

    const containsMatch = remaining.match(SELECTOR_PATTERNS.CONTAINS);
    if (containsMatch) {
      result.containsText = containsMatch[2];
      remaining = remaining.slice(containsMatch[0].length);
      matched = true;
    }

    if (SELECTOR_PATTERNS.EMPTY.test(remaining)) {
      attributes.empty = "true";
      remaining = remaining.replace(SELECTOR_PATTERNS.EMPTY, "");
      matched = true;
    }

    if (!element) {
      const typeMatch = remaining.match(/^([a-zA-Z]+)/);
      if (typeMatch) {
        element = normalizeElementName(typeMatch[1]);
        remaining = remaining.slice(typeMatch[0].length);
        matched = true;
      }
    }

    if (SELECTOR_PATTERNS.CHILD_COMBINATOR.test(remaining)) {
      result.attributes.childCombinator = ">";
      remaining = remaining.replace(SELECTOR_PATTERNS.CHILD_COMBINATOR, "");
      matched = true;

      const nextSelectorResult = parseSelector(remaining);
      if (nextSelectorResult.ok && nextSelectorResult.data) {
        result.childSelector = nextSelectorResult.data;
        remaining = "";
        matched = true;
      }
      continue;
    }

    if (SELECTOR_PATTERNS.DESCENDANT_COMBINATOR.test(remaining)) {
      result.attributes.childCombinator = " ";
      remaining = remaining.replace(SELECTOR_PATTERNS.DESCENDANT_COMBINATOR, "");
      matched = true;

      const nextSelectorResult = parseSelector(remaining);
      if (nextSelectorResult.ok && nextSelectorResult.data) {
        result.childSelector = nextSelectorResult.data;
        remaining = "";
        matched = true;
      }
      continue;
    }

    if (!matched) {
      break;
    }
  }

  if (element) {
    result.element = element;
  }
  result.attributes = { ...result.attributes, ...attributes };

  return ok(result);
}

/**
 * Normalizes element names to their canonical form.
 */
function normalizeElementName(name: string): string {
  const lower = name.toLowerCase();
  return PATH_ALIASES[lower] || lower;
}

// ============================================================================
// Selector Building
// ============================================================================

/**
 * Builds a selector string from a ParsedSelector.
 */
export function buildSelector(parsed: ParsedSelector): string {
  const parts: string[] = [];

  if (parsed.element) {
    parts.push(parsed.element);
  }

  if (parsed.attributes.index) {
    const lastIdx = parts.length - 1;
    parts[lastIdx] += `[${parsed.attributes.index}]`;
  }

  for (const [key, value] of Object.entries(parsed.attributes)) {
    if (key !== "index" && key !== "empty" && key !== "childCombinator") {
      const lastIdx = parts.length - 1;
      parts[lastIdx] += `[@${key}=${value}]`;
    }
  }

  if (parsed.attributes.empty === "true") {
    const lastIdx = parts.length - 1;
    parts[lastIdx] += ":empty";
  }

  if (parsed.containsText !== undefined) {
    const lastIdx = parts.length - 1;
    parts[lastIdx] += `:contains("${parsed.containsText}")`;
  }

  if (parsed.attributes.childCombinator && parsed.childSelector) {
    parts.push(parsed.attributes.childCombinator);
    parts.push(buildSelector(parsed.childSelector));
  }

  return parts.join(" ");
}

// ============================================================================
// Selector Validation
// ============================================================================

/**
 * Validates that a selector is well-formed.
 */
export function validateSelector(selector: string): Result<void> {
  return andThen(parseSelector(selector), () => ok(void 0));
}

/**
 * Checks if a selector is valid without returning detailed error.
 */
export function isValidSelector(selector: string): boolean {
  return parseSelector(selector).ok;
}

// ============================================================================
// Selector Helpers
// ============================================================================

/**
 * Creates a selector for a specific element type.
 */
export function typeSelector(elementType: string): ParsedSelector {
  return {
    element: normalizeElementName(elementType),
    attributes: {},
  };
}

/**
 * Creates a selector for paragraphs with a specific style.
 */
export function styleSelector(styleName: string): ParsedSelector {
  return {
    element: "p",
    attributes: { style: styleName },
  };
}

/**
 * Creates a selector with a text filter.
 */
export function textSelector(elementType: string, text: string): ParsedSelector {
  return {
    element: normalizeElementName(elementType),
    attributes: {},
    containsText: text,
  };
}

/**
 * Creates an indexed selector.
 */
export function indexedSelector(elementType: string, index: number): ParsedSelector {
  return {
    element: normalizeElementName(elementType),
    attributes: { index: String(index) },
  };
}

/**
 * Checks if a parsed selector has a text filter.
 */
export function hasTextFilter(parsed: ParsedSelector): boolean {
  return parsed.containsText !== undefined;
}

/**
 * Checks if a parsed selector has an attribute filter.
 */
export function hasAttributeFilter(parsed: ParsedSelector): boolean {
  return Object.keys(parsed.attributes).length > 0;
}

/**
 * Checks if a parsed selector targets a specific element type.
 */
export function isElementType(parsed: ParsedSelector, type: string): boolean {
  return parsed.element === normalizeElementName(type);
}

/**
 * Checks if a parsed selector targets a paragraph.
 */
export function isParagraphSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "p") || isElementType(parsed, "paragraph");
}

/**
 * Checks if a parsed selector targets a run.
 */
export function isRunSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "r") || isElementType(parsed, "run");
}

/**
 * Checks if a parsed selector targets a table.
 */
export function isTableSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "tbl") || isElementType(parsed, "table");
}

/**
 * Checks if a parsed selector targets a table row.
 */
export function isRowSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "tr") || isElementType(parsed, "row");
}

/**
 * Checks if a parsed selector targets a table cell.
 */
export function isCellSelector(parsed: ParsedSelector): boolean {
  return isElementType(parsed, "tc") || isElementType(parsed, "cell");
}

/**
 * Checks if a parsed selector has a child combinator.
 */
export function hasChildCombinator(parsed: ParsedSelector): boolean {
  return parsed.childSelector !== undefined;
}
