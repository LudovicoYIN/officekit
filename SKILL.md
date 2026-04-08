---
name: officekit
description: Node.js/Bun toolkit for creating, inspecting, previewing, and modifying Office documents (.docx, .xlsx, .pptx).
---

# officekit

Office document manipulation toolkit with CLI and programmatic API.

## Capabilities

- **Word** (.docx): Create, edit, query documents with styles, TOC, headers/footers
- **Excel** (.xlsx): Spreadsheet operations with formulas (50+ functions), pivot tables, charts
- **PowerPoint** (.pptx): Slides, animations, morph transitions, themes
- **Preview**: Live HTML preview server for all document types

## Command Families

- `create` - Create new documents
- `view` - View document content (text/outline/annotated/stats/html)
- `get` - Get specific elements
- `query` - Query document structure (CSS selectors)
- `set` - Set element properties
- `add` - Add new elements
- `remove` - Remove elements
- `move` - Move elements to new positions
- `swap` - Swap two elements
- `copy` - Copy elements
- `batch` - Batch operations
- `raw` - View raw XML
- `watch` - Live preview with file watching
- `check` - Layout checking
- `validate` - Document validation (OpenXML schema)
- `import` - Import CSV/TSV data
- `about` - Show version info
- `contracts` - Show capability summary

## Path Syntax

```
Word:  /body/p[1]/r[2]       # paragraph 1, run 2
Excel: /Sheet1/A1:B10        # range
       /Sheet1/$A$1          # absolute reference
PPT:   /slide[1]/shape[2]    # slide 1, shape 2
```

Selectors: `:contains(text)`, `:has(selector)`, `:eq(n)`

## Usage

```bash
# Create documents
officekit create demo.docx
officekit create spreadsheet.xlsx
officekit create presentation.pptx

# View content
officekit view demo.docx
officekit view spreadsheet.xlsx --sheet Sheet1

# Query structure
officekit query demo.docx /body/p
officekit query spreadsheet.xlsx /Sheet1/A1:B10

# Add content
officekit add demo.docx /body --type paragraph --prop "text=Hello"

# Modify
officekit set demo.docx /body/p[1] --prop "bold=true"
officekit move demo.docx /p[1] /to /p[3]
officekit swap demo.docx /p[1] /p[2]

# Batch operations
officekit batch demo.docx '[{"op":"add","path":"/body","type":"paragraph"}]'

# Validate and check
officekit validate demo.docx
officekit check presentation.pptx

# Preview
officekit watch demo.docx
```

## API

```typescript
import { createWordDocument, getWordNode, setWordNode } from "@officekit/word";
import { createExcelWorkbook, setExcelCell } from "@officekit/excel";
import { createPresentation, addSlide } from "@officekit/ppt";

// Word
await createWordDocument("output.docx");
await setWordNode("doc.docx", "/body/p[1]", { props: { text: "Hello" } });

// Excel
await createExcelWorkbook("output.xlsx");
await setExcelCell("sheet.xlsx", "/Sheet1/A1", { value: 42, formula: "=SUM(B1:B10)" });

// PowerPoint
await createPresentation("output.pptx");
await addSlide("slides.pptx");
```
