/**
 * MCP Tool Definitions for @officekit/ppt.
 *
 * This module defines all the tools that the MCP server exposes to AI assistants
 * for operating on PowerPoint files.
 */

import type { McpTool, McpToolInputSchema } from "./mcp-server.js";

// ============================================================================
// Tool Input Schemas
// ============================================================================

const filePathSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
  },
  required: ["filePath"],
};

const slidePathSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    slideIndex: {
      type: "number",
      description: "1-based slide index",
    },
  },
  required: ["filePath", "slideIndex"],
};

const shapePathSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    pptPath: {
      type: "string",
      description: "PPT path to the element (e.g., '/slide[1]/shape[1]')",
    },
  },
  required: ["filePath", "pptPath"],
};

const moveSlideSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    fromIndex: {
      type: "number",
      description: "Current 1-based position of the slide",
    },
    toIndex: {
      type: "number",
      description: "New 1-based position for the slide",
    },
  },
  required: ["filePath", "fromIndex", "toIndex"],
};

const swapSlidesSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    index1: {
      type: "number",
      description: "1-based index of the first slide",
    },
    index2: {
      type: "number",
      description: "1-based index of the second slide",
    },
  },
  required: ["filePath", "index1", "index2"],
};

const copySlideSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    sourceIndex: {
      type: "number",
      description: "1-based index of the slide to copy",
    },
    targetIndex: {
      type: "number",
      description: "1-based index where the copy should be inserted (-1 for end)",
    },
  },
  required: ["filePath", "sourceIndex", "targetIndex"],
};

const addSlideSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    layoutId: {
      type: "number",
      description: "Optional 1-based layout index to use",
    },
  },
  required: ["filePath"],
};

const removeSlideSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    index: {
      type: "number",
      description: "1-based slide index to remove",
    },
  },
  required: ["filePath", "index"],
};

const setShapeTextSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    pptPath: {
      type: "string",
      description: "PPT path to the shape (e.g., '/slide[1]/shape[1]')",
    },
    text: {
      type: "string",
      description: "The new text content",
    },
  },
  required: ["filePath", "pptPath", "text"],
};

const addShapeSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    slideIndex: {
      type: "number",
      description: "1-based slide index",
    },
    shapeType: {
      type: "string",
      description: "Preset shape type (e.g., 'rectangle', 'ellipse', 'triangle')",
    },
    x: {
      type: "number",
      description: "X position in EMUs",
    },
    y: {
      type: "number",
      description: "Y position in EMUs",
    },
    width: {
      type: "number",
      description: "Width in EMUs",
    },
    height: {
      type: "number",
      description: "Height in EMUs",
    },
  },
  required: ["filePath", "slideIndex", "shapeType", "x", "y", "width", "height"],
};

const removeShapeSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    pptPath: {
      type: "string",
      description: "PPT path to the shape (e.g., '/slide[1]/shape[1]')",
    },
  },
  required: ["filePath", "pptPath"],
};

const swapShapesSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    path1: {
      type: "string",
      description: "PPT path to the first shape",
    },
    path2: {
      type: "string",
      description: "PPT path to the second shape",
    },
  },
  required: ["filePath", "path1", "path2"],
};

const copyShapeSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    sourcePath: {
      type: "string",
      description: "PPT path to the source shape",
    },
    targetSlideIndex: {
      type: "number",
      description: "1-based index of the target slide",
    },
  },
  required: ["filePath", "sourcePath", "targetSlideIndex"],
};

const rawGetSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    pptPath: {
      type: "string",
      description: "PPT path to the element",
    },
  },
  required: ["filePath", "pptPath"],
};

const rawSetSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    pptPath: {
      type: "string",
      description: "PPT path to the element",
    },
    xml: {
      type: "string",
      description: "Raw XML to set",
    },
  },
  required: ["filePath", "pptPath", "xml"],
};

const batchSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    operations: {
      type: "array",
      description: "Array of operations to execute",
      items: {
        type: "object",
        properties: {
          op: {
            type: "string",
            description: "Operation type (rawSet, set, remove, swap, copyFrom)",
          },
          params: {
            type: "object",
            description: "Operation parameters",
          },
        },
      },
    },
  },
  required: ["filePath", "operations"],
};

const viewOptionsSchema: McpToolInputSchema = {
  type: "object",
  properties: {
    filePath: {
      type: "string",
      description: "Path to the PPTX file",
    },
    slideIndex: {
      type: "number",
      description: "Optional 1-based slide index to view specific slide",
    },
  },
  required: ["filePath"],
};

// ============================================================================
// Tool Definitions
// ============================================================================

/**
 * All MCP tools exposed by the PPT server.
 */
export const pptTools: McpTool[] = [
  // ========== Slide Management ==========
  {
    name: "Add",
    description: "Adds a new slide to the presentation. Use layoutId to specify which layout to use (1-based index).",
    inputSchema: addSlideSchema,
  },
  {
    name: "AddPart",
    description: "Adds a new shape to a slide. Requires shape type (rectangle, ellipse, etc.), position (x, y in EMUs), and size (width, height in EMUs).",
    inputSchema: addShapeSchema,
  },
  {
    name: "Remove",
    description: "Removes a slide from the presentation at the given index.",
    inputSchema: removeSlideSchema,
  },
  {
    name: "Move",
    description: "Moves a slide from one position to another. Use fromIndex for current position and toIndex for new position.",
    inputSchema: moveSlideSchema,
  },
  {
    name: "Swap",
    description: "Swaps two slides in the presentation. Both slides must exist.",
    inputSchema: swapSlidesSchema,
  },
  {
    name: "CopyFrom",
    description: "Copies a slide or shape from one location to another. For slides, use sourceIndex and targetIndex. For shapes, use sourcePath and targetSlideIndex.",
    inputSchema: {
      type: "object",
      properties: {
        filePath: { type: "string", description: "Path to the PPTX file" },
        sourceIndex: { type: "number", description: "1-based index of slide to copy (for slide copy)" },
        targetIndex: { type: "number", description: "1-based target index (for slide copy)" },
        sourcePath: { type: "string", description: "PPT path to source shape (for shape copy)" },
        targetSlideIndex: { type: "number", description: "Target slide for shape copy" },
      },
      required: ["filePath"],
    },
  },

  // ========== Query Operations ==========
  {
    name: "Get",
    description: "Gets detailed information about a specific element at the given PPT path (slide, shape, table, chart, or placeholder).",
    inputSchema: shapePathSchema,
  },
  {
    name: "Query",
    description: "Queries the presentation for elements matching criteria. Returns all slides or shapes depending on parameters.",
    inputSchema: {
      type: "object",
      properties: {
        filePath: { type: "string", description: "Path to the PPTX file" },
        selector: { type: "string", description: "Selector string (e.g., 'slide', 'shape', '/slide[1]')" },
      },
      required: ["filePath"],
    },
  },

  // ========== Mutation Operations ==========
  {
    name: "Set",
    description: "Sets the text content of a shape at the given path.",
    inputSchema: setShapeTextSchema,
  },
  {
    name: "Raw",
    description: "Gets the raw XML for an element at the given path.",
    inputSchema: rawGetSchema,
  },
  {
    name: "RawSet",
    description: "Sets the raw XML for an element at the given path. Use with caution as this bypasses safety checks.",
    inputSchema: rawSetSchema,
  },
  {
    name: "Batch",
    description: "Executes multiple mutations in a single batch operation for efficiency.",
    inputSchema: batchSchema,
  },

  // ========== View Operations ==========
  {
    name: "ViewAsText",
    description: "Extracts plain text from the presentation. Use slideIndex to get text from a specific slide only.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsAnnotated",
    description: "Gets an annotated view showing element types, positions, names, and properties.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsOutline",
    description: "Gets an outline/summary view of the presentation structure.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsStats",
    description: "Gets statistics about the presentation (shape counts, text lengths, etc.).",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsIssues",
    description: "Finds potential problems in the presentation (missing titles, text overflow risks, etc.).",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsHtml",
    description: "Renders the presentation as a self-contained HTML document.",
    inputSchema: viewOptionsSchema,
  },
  {
    name: "ViewAsSvg",
    description: "Renders the presentation as SVG vector graphics.",
    inputSchema: viewOptionsSchema,
  },

  // ========== Validation Operations ==========
  {
    name: "CheckShapeTextOverflow",
    description: "Checks if text overflows in a specific shape and returns details about the overflow.",
    inputSchema: shapePathSchema,
  },
];

/**
 * Map of tool names to their definitions.
 */
export const toolByName = new Map<string, McpTool>(
  pptTools.map((tool) => [tool.name, tool])
);
