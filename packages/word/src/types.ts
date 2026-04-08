/**
 * Shared TypeScript types for Word elements.
 */

// ============================================================================
// Result Envelope
// ============================================================================

export interface Result<T> {
  ok: boolean;
  data?: T;
  error?: ResultError;
}

export interface ResultError {
  code: string;
  message: string;
  suggestion?: string;
}

// ============================================================================
// Word Document Model
// ============================================================================

export interface WordDocumentModel {
  filePath: string;
  metadata: WordDocumentMetadata;
}

export interface WordDocumentMetadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  description?: string;
  category?: string;
  lastModifiedBy?: string;
  revision?: string;
  created?: string;
  modified?: string;
}

// ============================================================================
// Document Node
// ============================================================================

export interface DocumentNode {
  path: string;
  type: string;
  text?: string;
  style?: string;
  preview?: string;
  childCount?: number;
  children?: DocumentNode[];
  format?: Record<string, unknown>;
}

// ============================================================================
// Paragraph and Run Models
// ============================================================================

export interface ParagraphModel {
  index: number;
  path: string;
  text: string;
  style?: string;
  alignment?: "left" | "center" | "right" | "justify" | "both";
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
  firstLineIndent?: number;
  leftIndent?: number;
  runs: RunModel[];
  childCount?: number;
}

export interface RunModel {
  index: number;
  text: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: string;
  color?: string;
  highlight?: string;
}

// ============================================================================
// Table Model
// ============================================================================

export interface TableModel {
  path: string;
  name?: string;
  rowCount?: number;
  columnCount?: number;
  rows: TableRowModel[];
}

export interface TableRowModel {
  index: number;
  path: string;
  cellCount?: number;
  cells: TableCellModel[];
}

export interface TableCellModel {
  index: number;
  path: string;
  text: string;
  gridSpan?: number;
  rowSpan?: number;
  fill?: string;
  valign?: "top" | "center" | "bottom";
}

// ============================================================================
// Header/Footer Model
// ============================================================================

export interface HeaderModel {
  index: number;
  path: string;
  type?: string;
  text?: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  alignment?: string;
}

export interface FooterModel {
  index: number;
  path: string;
  type?: string;
  text?: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  alignment?: string;
}

// ============================================================================
// Footnote/Endnote Model
// ============================================================================

export interface FootnoteModel {
  id: string;
  path: string;
  text: string;
}

export interface EndnoteModel {
  id: string;
  path: string;
  text: string;
}

// ============================================================================
// Field Model
// ============================================================================

export interface FieldModel {
  index: number;
  path: string;
  type: string;
  instruction: string;
  result: string;
  dirty?: boolean;
}

// ============================================================================
// TOC Model
// ============================================================================

export interface TocModel {
  index: number;
  path: string;
  instruction: string;
  levels?: string;
  hyperlinks?: boolean;
  pageNumbers?: boolean;
}

// ============================================================================
// Style Model
// ============================================================================

export interface StyleModel {
  id: string;
  name: string;
  type?: string;
  basedOn?: string;
  next?: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  underline?: string;
  strike?: boolean;
  alignment?: string;
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
}

// ============================================================================
// Path Segment Types
// ============================================================================

export interface PathSegment {
  name: string;
  index?: number;
  stringIndex?: string;
}

export interface ParsedPath {
  isAbsolute: boolean;
  segments: PathSegment[];
  original: string;
}

// ============================================================================
// Selector Types
// ============================================================================

export interface ParsedSelector {
  element?: string;
  attributes: Record<string, string>;
  containsText?: string;
  childSelector?: ParsedSelector;
}

// ============================================================================
// Advanced TOC Types
// ============================================================================

export interface TocOptions {
  title?: string;
  levels?: string;
  hyperlinks?: boolean;
  pageNumbers?: boolean;
  tabLeader?: "dot" | "dash" | "underline" | "none";
  styles?: string[];
}

export interface TocEntry {
  text: string;
  level: number;
  page: number;
  style: string;
  path: string;
}

export interface TocInfo {
  index: number;
  path: string;
  instruction: string;
  title?: string;
  entries: TocEntry[];
}

// ============================================================================
// Advanced Style Inheritance Types
// ============================================================================

export interface StyleWithInheritance {
  id: string;
  name: string;
  type: string;
  basedOn?: string;
  next?: string;
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  underline?: string;
  strike?: boolean;
  alignment?: string;
  spaceBefore?: string;
  spaceAfter?: string;
  lineSpacing?: string;
  indentation?: string;
  firstLineIndent?: string;
  shading?: string;
  border?: Record<string, unknown>;
  // Inherited properties
  inheritedFont?: string;
  inheritedSize?: string;
  inheritedBold?: boolean;
  inheritedItalic?: boolean;
  inheritedColor?: string;
  inheritedUnderline?: string;
  inheritedAlignment?: string;
  inheritedSpaceBefore?: string;
  inheritedSpaceAfter?: string;
  inheritedLineSpacing?: string;
  // Chain info
  inheritanceChain: string[];
}

// ============================================================================
// Advanced Header/Footer Types
// ============================================================================

export interface HeaderFooterOptions {
  text?: string;
  alignment?: "left" | "center" | "right";
  font?: string;
  size?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  field?: string;
  image?: string;
  logoPath?: string;
  linkToPrevious?: boolean;
}

export interface SectionInfo {
  index: number;
  path: string;
  properties: {
    pageWidth?: number;
    pageHeight?: number;
    orientation?: string;
    marginTop?: number;
    marginBottom?: number;
    marginLeft?: number;
    marginRight?: number;
    columns?: number;
    columnSpace?: number;
    sectionType?: string;
    titlePage?: boolean;
  };
  headers: number[];
  footers: number[];
}

// ============================================================================
// Advanced Watermark Types
// ============================================================================

export interface WatermarkOptions {
  text?: string;
  color?: string;
  font?: string;
  size?: string;
  rotation?: number;
  opacity?: number;
  position?: "center" | "diagonal";
}

// ============================================================================
// Advanced Form Field Types
// ============================================================================

export interface FormFieldInfo {
  type: "text" | "checkbox" | "dropdown";
  name: string;
  value: string;
  enabled: boolean;
  editable: boolean;
  path: string;
  defaultValue?: string;
  maxLength?: number;
  checked?: boolean;
  items?: string[];
  defaultIndex?: number;
}

export interface AdvancedFormFieldOptions {
  name?: string;
  label?: string;
  defaultValue?: string;
  maxLength?: number;
  checked?: boolean;
  enabled?: boolean;
  calculated?: boolean;
  expression?: string;
  items?: string[];
  defaultIndex?: number;
}

// ============================================================================
// Advanced Document Protection Types
// ============================================================================

export interface DocumentProtection {
  mode?: string;
  enforced?: boolean;
}

export interface AdvancedProtectionOptions {
  mode?: "readOnly" | "comments" | "trackedChanges" | "forms" | "full";
  enforcement?: boolean;
  password?: string;
  editExceptions?: string[];
}
