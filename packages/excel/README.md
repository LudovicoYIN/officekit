# @officekit/excel

Excel document adapter for officekit.

## Features

- Workbook and sheet management
- Cell operations with formula support (150+ functions)
- Named ranges with chain resolution
- Pivot tables
- Data validation and conditional formatting
- Charts
- CSV/TSV import

## Usage

```typescript
import { createExcelWorkbook, queryExcelNodes, setExcelCell } from "@officekit/excel";

// Create workbook
await createExcelWorkbook("output.xlsx");

// Query cells
const result = await queryExcelNodes("spreadsheet.xlsx", "/Sheet1/A1");

// Set cell value
await setExcelCell("spreadsheet.xlsx", "/Sheet1/A1", { value: 42, formula: "=SUM(B1:B10)" });
```
