# @officekit/ppt

PowerPoint document adapter for officekit.

## Features

- Slide creation and manipulation
- Shape operations (alignment, distribution)
- Animations and morph transitions
- Master views and layouts
- Theme management
- Text overflow checking
- Table, chart, and media support
- HTML/SVG preview

## Usage

```typescript
import { createPresentation, getSlide, addSlide } from "@officekit/ppt";

// Create presentation
await createPresentation("output.pptx");

// Add slide
await addSlide("presentation.pptx");

// Get slide content
const slide = await getSlide("presentation.pptx", 1);
```
