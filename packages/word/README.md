# @officekit/word

Word document adapter for officekit.

## Features

- Document creation and editing
- Style management with inheritance
- Table of Contents (TOC) support
- Header/Footer management
- Document protection and watermarks
- Form fields
- HTML preview

## Usage

```typescript
import { createWordDocument, getWordNode, setWordNode } from "@officekit/word";

// Create document
await createWordDocument("output.docx");

// Query content
const result = await queryWordNodes("document.docx", "/body/p[1]");

// Set content
await setWordNode("document.docx", "/body/p[1]", { props: { text: "Hello" } });
```
