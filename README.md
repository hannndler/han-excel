# Han Excel Builder

🚀 **Advanced Excel file generator with TypeScript support, comprehensive styling, and optimized performance**

A modern, fully-typed library for creating complex Excel reports with multiple worksheets, advanced styling, and high performance.

**📖 [Read in Spanish / Leer en Español](README.es.md)**

---

## 📑 Table of Contents

- [✨ Features](#-features)
  - [📊 Data Structure](#-data-structure)
  - [📈 Data Types](#-data-types)
  - [🎨 Advanced Styling](#-advanced-styling)
  - [🔧 Advanced Features](#-advanced-features)
- [🌐 Browser & Node.js Compatibility](#-browser--nodejs-compatibility)
  - [Compatibility Table](#compatibility-table)
  - [Environment-Specific Details](#environment-specific-details)
- [💾 Exporting Files: Browser vs Node.js](#-exporting-files-browser-vs-nodejs)
  - [🌐 Browser Environment](#-browser-environment)
  - [🖥️ Node.js Environment](#️-nodejs-environment)
  - [📊 Comparison Table](#-comparison-table)
  - [💡 Best Practices](#-best-practices)
- [📦 Installation](#-installation)
- [🚀 Quick Start](#-quick-start)
- [📚 API Documentation](#-api-documentation)
  - [Core Classes](#core-classes)
  - [Data Types](#data-types)
  - [Styling](#styling)
- [🎯 Advanced Examples](#-advanced-examples)
- [🧪 Testing](#-testing)
- [🛠️ Development](#️-development)
- [📋 Migration from legacy-excel](#-migration-from-legacy-excel)
- [📚 Additional Resources](#-additional-resources)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)
- [🆘 Support](#-support)

---

## ✨ Features

### 📊 Data Structure
- ✅ **Multiple Worksheets** - Create complex workbooks with multiple sheets
- ✅ **Multiple Tables per Sheet** - Create multiple independent tables in a single sheet
- ✅ **Nested Headers** - Full support for headers with multiple nesting levels
- ✅ **Hierarchical Data** - Support for data with children structure (nested data)

### 📈 Data Types
- ✅ **STRING** - Text values
- ✅ **NUMBER** - Numeric values
- ✅ **BOOLEAN** - True/false values
- ✅ **DATE** - Date values
- ✅ **PERCENTAGE** - Percentage values
- ✅ **CURRENCY** - Currency values
- ✅ **LINK** - Hyperlinks with customizable text
- ✅ **FORMULA** - Excel formulas

### 🎨 Advanced Styling
- ✅ **Fluent API** - StyleBuilder with chainable methods
- ✅ **Fonts** - Full control over name, size, color, bold, italic, underline
- ✅ **Colors** - Backgrounds, text colors with support for hex, RGB and themes
- ✅ **Borders** - Customizable borders on all sides with multiple styles
- ✅ **Alignment** - Horizontal (left, center, right, justify) and vertical (top, middle, bottom)
- ✅ **Text** - Text wrapping, shrink to fit, text rotation
- ✅ **Number Formats** - Multiple predefined and custom formats
- ✅ **Alternating Rows** - Support for alternating stripes in tables

### 🔧 Advanced Features
- ✅ **TypeScript First** - Complete type safety with comprehensive interfaces
- ✅ **Event System** - EventEmitter to monitor the build process
- ✅ **Validation** - Robust data validation system
- ✅ **Metadata** - Full support for workbook metadata (author, title, description, etc.)
- ✅ **Multiple Export Formats** - Direct download, Buffer, Blob
- ✅ **Excel Reading** - Read Excel files and convert to JSON
- ✅ **Hyperlinks** - Create links with customizable text
- ✅ **Cell Merging** - Horizontal and vertical cell merging
- ✅ **Custom Dimensions** - Customizable column width and row height
- ✅ **Cell Comments** - Add comments to cells (read and write support)
- ✅ **Data Validation** - Apply data validation rules to cells (list, whole, decimal, textLength, date, custom)
- ✅ **Auto Filters** - Enable automatic filters for tables and worksheets
- ✅ **Conditional Formatting** - Apply conditional formatting rules to cells based on values or formulas
- ✅ **Freeze Panes** - Freeze rows and columns for easier navigation
- ✅ **Worksheet Protection** - Protect worksheets with password and permission settings
- ✅ **Images/Pictures** - Add images to worksheets (PNG, JPEG, GIF, BMP, WebP)
- ✅ **Row/Column Grouping** - Group rows and columns for collapsible outlines
- ✅ **Named Ranges** - Define named ranges for easy references
- ✅ **Excel Structured Tables** - Create Excel structured tables with automatic styles
- ✅ **Advanced Print Settings** - Headers/footers and repeat rows/columns on each page
- ✅ **Hide/Show Rows & Columns** - Hide or show specific rows and columns
- ✅ **Rich Text in Cells** - Format text with multiple styles within a single cell
- ✅ **Cell-level Protection** - Protect individual cells with lock/hide options
- ✅ **Pivot Tables** - Create pivot tables for data analysis
- ✅ **Slicers** - Visual filters for tables and pivot tables (documented, requires advanced ExcelJS)
- ✅ **Watermarks** - Add watermarks to worksheets (text or image)
- ✅ **Cell-level Page Breaks** - Manual page breaks at row level
- ✅ **Data Connections** - External data connections (documented, requires advanced ExcelJS)
- ✅ **Cell Styles (Predefined)** - Reusable cell styles for consistency
- ✅ **Themes** - Color themes for the entire workbook
- ✅ **Split Panes** - Divide window into panes (different from freeze panes)
- ✅ **Sheet Views** - Multiple views of the same sheet (normal, page break preview, page layout)

## 🌐 Browser & Node.js Compatibility

Han Excel Builder works in both **browser** and **Node.js** environments. Most features are fully compatible with both, but some have limitations depending on the environment.

### Compatibility Table

| Feature | Browser | Node.js | Notes |
|---------|---------|---------|-------|
| **Basic Features** |
| Multiple Worksheets | ✅ | ✅ | Fully compatible |
| Nested Headers | ✅ | ✅ | Fully compatible |
| Hierarchical Data | ✅ | ✅ | Fully compatible |
| All Data Types (STRING, NUMBER, etc.) | ✅ | ✅ | Fully compatible |
| **Styling** |
| StyleBuilder & Fluent API | ✅ | ✅ | Fully compatible |
| Fonts, Colors, Borders | ✅ | ✅ | Fully compatible |
| Conditional Formatting | ✅ | ✅ | Fully compatible |
| Themes | ✅ | ✅ | Fully compatible |
| Predefined Cell Styles | ✅ | ✅ | Fully compatible |
| **Advanced Features** |
| Images/Pictures | ✅ | ✅ | Fully compatible |
| Pivot Tables | ✅ | ✅ | Fully compatible |
| Freeze Panes | ✅ | ✅ | Fully compatible |
| Worksheet Protection | ✅ | ✅ | Fully compatible |
| Data Validation | ✅ | ✅ | Fully compatible |
| Rich Text in Cells | ✅ | ✅ | Fully compatible |
| Cell-level Protection | ✅ | ✅ | Fully compatible |
| Row/Column Grouping | ✅ | ✅ | Fully compatible |
| Named Ranges | ✅ | ✅ | Fully compatible |
| Excel Structured Tables | ✅ | ✅ | Fully compatible |
| Hide/Show Rows & Columns | ✅ | ✅ | Fully compatible |
| Split Panes | ✅ | ✅ | Fully compatible |
| Sheet Views | ✅ | ✅ | Fully compatible |
| **File Operations** |
| Generate & Download | ✅ | ✅ | Browser: Uses Blob/Download. Node: Can write to file |
| Read Excel Files | ✅ | ✅ | Browser: From File/Blob. Node: Also from file path |
| **Features with Limitations** |
| Templates | ⚠️ | ✅ | Browser: Only ArrayBuffer/Blob. Node: Also file paths |
| Streaming | ⚠️ | ✅ | Browser: Limited. Node: Full support |
| Charts (as image) | ✅ | ✅ | Requires chart library compatible with environment |
| Sparklines (as image) | ✅ | ✅ | Requires chart library compatible with environment |

### Legend
- ✅ **Compatible**: Works fully in this environment
- ⚠️ **Limited**: Works but with restrictions or requires special configuration
- ❌ **Not Available**: Does not work in this environment

### Environment-Specific Details

#### ✅ Fully Compatible Features
Most features work identically in both browser and Node.js:
- All styling features (StyleBuilder, themes, conditional formatting)
- All data structure features (worksheets, tables, headers)
- All cell features (merging, protection, validation)
- Image insertion
- Pivot tables
- All export formats (Buffer, Blob, download)

#### ⚠️ Features with Limitations

**Templates:**
- **Browser**: Can only load templates from `ArrayBuffer` or `Blob` (e.g., from `fetch()` or File input)
  ```typescript
  // Browser: Load from fetch
  const response = await fetch('/template.xlsx');
  const buffer = await response.arrayBuffer();
  await builder.loadTemplate(buffer);
  ```
- **Node.js**: Can load from file path or `ArrayBuffer`
  ```typescript
  // Node: Load from file
  await builder.loadTemplate('./template.xlsx');
  // Or from buffer
  await builder.loadTemplate(buffer);
  ```

**Streaming (Large Files):**
- **Browser**: Limited by browser stream capabilities. May require polyfills or alternatives
- **Node.js**: Full support with `ExcelJS.stream.xlsx.WorkbookWriter`
  ```typescript
  // Node: Full streaming support
  const stream = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  ```

**Charts/Sparklines (as images):**
- **Browser**: Requires browser-compatible chart library (Chart.js, D3.js, Plotly.js)
- **Node.js**: Requires Node-compatible chart library (can use canvas/server-side rendering)
- **Note**: The chart library must be compatible with the execution environment

**File Reading:**
- **Browser**: Use `ExcelReader.fromFile()` or `ExcelReader.fromBlob()`
- **Node.js**: Can also use `ExcelReader.fromPath()` for file system access

## 💾 Exporting Files: Browser vs Node.js

The way you export Excel files differs between browser and Node.js environments. Here's how to handle each case:

### 🌐 Browser Environment

In the browser, you have three main options:

#### 1. **Direct Download** (Recommended for Browser)
Automatically triggers a download in the user's browser:

```typescript
import { ExcelBuilder } from 'han-excel-builder';

const builder = new ExcelBuilder();
// ... build your workbook ...

// Automatically downloads the file
const result = await builder.generateAndDownload('report.xlsx');

if (result.success) {
  console.log('File downloaded successfully!');
} else {
  console.error('Download failed:', result.error);
}
```

**Result**: The file is automatically downloaded to the user's default download folder.

#### 2. **Get as Blob** (For Custom Handling)
Get the file as a Blob for custom handling (e.g., upload to server, preview, etc.):

```typescript
// Get as Blob
const result = await builder.toBlob();

if (result.success) {
  const blob = result.data; // Blob object
  
  // Option A: Upload to server
  const formData = new FormData();
  formData.append('file', blob, 'report.xlsx');
  await fetch('/api/upload', { method: 'POST', body: formData });
  
  // Option B: Create object URL for preview
  const url = URL.createObjectURL(blob);
  window.open(url);
  
  // Option C: Manual download
  const link = document.createElement('a');
  link.href = url;
  link.download = 'report.xlsx';
  link.click();
}
```

**Result**: Returns a `Blob` object that you can use for custom operations.

#### 3. **Get as ArrayBuffer** (For Low-Level Operations)
Get the raw binary data:

```typescript
// Get as ArrayBuffer
const result = await builder.toBuffer();

if (result.success) {
  const buffer = result.data; // ArrayBuffer
  
  // Use for low-level operations
  // e.g., send via WebSocket, process with other libraries, etc.
}
```

**Result**: Returns an `ArrayBuffer` with the raw Excel file data.

---

### 🖥️ Node.js Environment

In Node.js, you typically want to save the file to disk. Here are the options:

#### 1. **Save to File** (Recommended for Node.js - Simple & Direct)
**NEW!** Now as simple as browser - just one method call:

```typescript
import { ExcelBuilder } from 'han-excel-builder';

const builder = new ExcelBuilder();
// ... build your workbook ...

// Save directly to file - automatically creates directories if needed
const result = await builder.saveToFile('./output/report.xlsx');

if (result.success) {
  console.log('File saved successfully!');
} else {
  console.error('Save failed:', result.error);
}
```

**Result**: File is saved to the specified path on the file system. Parent directories are created automatically.

**Options:**
```typescript
await builder.saveToFile('./output/report.xlsx', {
  createDir: true,  // Create parent directories (default: true)
  encoding: 'binary' // File encoding (default: 'binary')
});
```

#### 2. **Save to Stream** (For Large Files)
For very large files, you can write directly to a stream:

```typescript
import { ExcelBuilder } from 'han-excel-builder';
import fs from 'fs';

const builder = new ExcelBuilder();
// ... build your workbook ...

const writeStream = fs.createWriteStream('./output/report.xlsx');
const result = await builder.saveToStream(writeStream);

if (result.success) {
  console.log('File streamed successfully!');
  writeStream.end();
}
```

**Result**: File is written to disk using streams (better for large files).

#### 3. **Manual Save** (Using toBuffer + fs)
If you need more control, you can still use the manual approach:

```typescript
import { ExcelBuilder } from 'han-excel-builder';
import fs from 'fs/promises';

const builder = new ExcelBuilder();
// ... build your workbook ...

// Get as buffer
const result = await builder.toBuffer();

if (result.success) {
  const buffer = result.data; // ArrayBuffer
  
  // Write to file
  await fs.writeFile('./output/report.xlsx', Buffer.from(buffer));
  console.log('File saved successfully!');
} else {
  console.error('Build failed:', result.error);
}
```

**Result**: File is saved to the specified path on the file system.

#### 4. **Send as HTTP Response** (For Web Servers)
If you're building a web server, you can send the file directly:

```typescript
import { ExcelBuilder } from 'han-excel-builder';
import express from 'express';

const app = express();

app.get('/download-report', async (req, res) => {
  const builder = new ExcelBuilder();
  // ... build your workbook ...
  
  const result = await builder.toBuffer();
  
  if (result.success) {
    res.setHeader('Content-Type', 
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 
      'attachment; filename="report.xlsx"');
    res.send(Buffer.from(result.data));
  } else {
    res.status(500).json({ error: result.error });
  }
});
```

**Result**: File is sent as HTTP response for download.

#### 5. **Using generateAndDownload() in Node.js**
While `generateAndDownload()` works in Node.js, it's not recommended as it uses browser-specific APIs. Use `saveToFile()` instead:

```typescript
// ❌ Not recommended in Node.js
await builder.generateAndDownload('report.xlsx');

// ✅ Recommended in Node.js
await builder.saveToFile('report.xlsx');
```

**Result**: `saveToFile()` is the Node.js equivalent of `generateAndDownload()` - simple and direct!

---

### 📊 Comparison Table

| Method | Browser | Node.js | Result Type | Use Case |
|--------|---------|---------|-------------|----------|
| `generateAndDownload()` | ✅ Recommended | ⚠️ Works but not ideal | `void` | Direct download (browser) |
| `saveToFile()` | ❌ Not available | ✅ **Recommended** | `void` | **Direct save (Node.js)** - Simple! |
| `saveToStream()` | ❌ Not available | ✅ Good | `void` | Stream to file (large files) |
| `toBlob()` | ✅ Good | ✅ Works | `Blob` | Custom handling, uploads |
| `toBuffer()` | ✅ Works | ✅ Good | `ArrayBuffer` | Manual save, HTTP response |

### 💡 Best Practices

**Browser:**
- Use `generateAndDownload()` for simple downloads
- Use `toBlob()` when you need to upload to a server or handle the file programmatically
- Use `toBuffer()` for low-level operations

**Node.js:**
- **Use `saveToFile()` for simple file saving** - Just like `generateAndDownload()` in browser!
- Use `saveToStream()` for very large files
- Use `toBuffer()` + HTTP response for web servers
- Avoid `generateAndDownload()` in Node.js (use `saveToFile()` instead)

### 🔄 Example: Universal Export Function

Here's a helper function that works in both environments:

```typescript
async function exportExcel(
  builder: ExcelBuilder, 
  filename: string
): Promise<void> {
  const isBrowser = typeof window !== 'undefined';
  
  if (isBrowser) {
    // Browser: Direct download
    await builder.generateAndDownload(filename);
  } else {
    // Node.js: Save to file - Now as simple as browser!
    const result = await builder.saveToFile(filename);
    
    if (result.success) {
      console.log(`File saved: ${filename}`);
    } else {
      throw new Error(result.error?.message || 'Export failed');
    }
  }
}

// Usage
await exportExcel(builder, 'report.xlsx');
```

**Note**: Both methods (`generateAndDownload()` and `saveToFile()`) are now equally simple - one method call!

## 📦 Installation

```bash
npm install han-excel-builder
# or
yarn add han-excel-builder
# or
pnpm add han-excel-builder
```

## 🚀 Quick Start

### Basic Example

```typescript
import { ExcelBuilder, CellType, NumberFormat, StyleBuilder, BorderStyle } from 'han-excel-builder';

// Create a simple report
const builder = new ExcelBuilder({
  metadata: {
    title: 'Sales Report',
    author: 'My Company',
    description: 'Monthly sales report'
  }
});

const worksheet = builder.addWorksheet('Sales');

// Add main header
worksheet.addHeader({
  key: 'title',
  value: 'Monthly Sales Report',
  type: CellType.STRING,
  mergeCell: true,
  styles: new StyleBuilder()
    .fontName('Arial')
    .fontSize(16)
    .fontBold()
    .backgroundColor('#4472C4')
    .fontColor('#FFFFFF')
    .centerAlign()
    .border(BorderStyle.THIN, '#8EAADB')
    .build()
});

// Add sub-headers
worksheet.addSubHeaders([
  {
    key: 'product',
    value: 'Product',
    type: CellType.STRING,
    colWidth: 20,
    styles: new StyleBuilder()
      .fontBold()
      .backgroundColor('#8EAADB')
      .fontColor('#FFFFFF')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  },
  {
    key: 'sales',
    value: 'Sales',
    type: CellType.CURRENCY,
    colWidth: 15,
    numberFormat: '$#,##0',
    styles: new StyleBuilder()
      .fontBold()
      .backgroundColor('#8EAADB')
      .fontColor('#FFFFFF')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  }
]);

// Add data
worksheet.addRow([
  {
    key: 'product-1',
    value: 'Product A',
    type: CellType.STRING,
    header: 'Product'
  },
  {
    key: 'sales-1',
    value: 1500.50,
    type: CellType.CURRENCY,
    header: 'Sales',
    numberFormat: '$#,##0.00'
  }
]);

// Generate and download
await builder.generateAndDownload('sales-report.xlsx');
```

## 📚 API Documentation

### Core Classes

#### `ExcelBuilder`

Main class for creating Excel workbooks.

```typescript
const builder = new ExcelBuilder({
  metadata: {
    title: 'My Report',
    author: 'My Name',
    company: 'My Company',
    description: 'Report description',
    keywords: 'excel, report, data',
    created: new Date(),
    modified: new Date()
  },
  enableValidation: true,
  enableEvents: true,
  maxWorksheets: 255,
  maxRowsPerWorksheet: 1048576,
  maxColumnsPerWorksheet: 16384
});

// Main methods
builder.addWorksheet(name, config);      // Add a worksheet
builder.getWorksheet(name);              // Get a worksheet
builder.removeWorksheet(name);           // Remove a worksheet
builder.setCurrentWorksheet(name);       // Set current worksheet
builder.build(options);                  // Build and get ArrayBuffer
builder.generateAndDownload(fileName);    // Generate and download
builder.toBuffer(options);               // Get as Buffer
builder.toBlob(options);                // Get as Blob
builder.validate();                      // Validate workbook
builder.clear();                         // Clear all worksheets
builder.getStats();                      // Get statistics

// Event system
builder.on(eventType, listener);
builder.off(eventType, listenerId);
builder.removeAllListeners(eventType);
```

#### `ExcelReader`

Class for reading Excel files and converting them to JSON with 3 different output formats.

**Available formats:**
- `worksheet` (default) - Complete structure with sheets, rows and cells
- `detailed` - Each cell with position information (text, column, row)
- `flat` - Just the data, without structure

```typescript
import { ExcelReader, OutputFormat } from 'han-excel-builder';

// ===== FORMAT 1: WORKSHEET (default) =====
// Complete structure organized by sheets
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.WORKSHEET, // or 'worksheet'
  useFirstRowAsHeaders: true
});

if (result.success) {
  const workbook = result.data;
  // workbook.sheets[] - Array of sheets
  // workbook.sheets[0].rows[] - Array of rows
  // workbook.sheets[0].rows[0].cells[] - Array of cells
  // workbook.sheets[0].rows[0].cells[0].comment - Cell comment (if exists)
  // workbook.sheets[0].rows[0].data - Object with data (if useFirstRowAsHeaders)
}

// ===== FORMAT 2: DETAILED =====
// Each cell with position information
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.DETAILED, // or 'detailed'
  includeFormatting: true
});

if (result.success) {
  const detailed = result.data;
  // detailed.cells[] - Array of all cells with:
  //   - value: cell value
  //   - text: cell text
  //   - column: column number (1-based)
  //   - columnLetter: column letter (A, B, C...)
  //   - row: row number (1-based)
  //   - reference: cell reference (A1, B2...)
  //   - sheet: sheet name
  //   - comment: cell comment (if exists)
  detailed.cells.forEach(cell => {
    console.log(`${cell.sheet}!${cell.reference}: ${cell.text}`);
    if (cell.comment) {
      console.log(`  Comment: ${cell.comment}`);
    }
  });
}

// ===== FORMAT 3: FLAT =====
// Just the data, without structure
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.FLAT, // or 'flat'
  useFirstRowAsHeaders: true
});

if (result.success) {
  const flat = result.data;
  
  // If single sheet:
  if ('data' in flat) {
    // flat.data[] - Array of objects or arrays
    // flat.headers[] - Headers (if useFirstRowAsHeaders)
    flat.data.forEach(row => {
      console.log(row); // { Product: 'A', Price: 100 } or ['A', 100]
    });
  }
  
  // If multiple sheets:
  if ('sheets' in flat) {
    // flat.sheets['SheetName'].data[] - Data by sheet
    Object.keys(flat.sheets).forEach(sheetName => {
      console.log(`Sheet: ${sheetName}`);
      flat.sheets[sheetName].data.forEach(row => {
        console.log(row);
      });
    });
  }
}

// ===== USING MAPPER TO TRANSFORM DATA =====
// The mapper allows transforming the response before returning it
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.WORKSHEET,
  useFirstRowAsHeaders: true,
  // Mapper receives the payload and returns the transformation
  mapper: (data) => {
    // Transform data according to needs
    const transformed = {
      totalSheets: data.totalSheets,
      sheets: data.sheets.map(sheet => ({
        name: sheet.name,
        // Convert rows to objects with transformed data
        rows: sheet.rows.map(row => {
          if (row.data) {
            // Transform each field
            return {
              ...row.data,
              // Add calculated fields
              total: Object.values(row.data).reduce((sum, val) => {
                return sum + (typeof val === 'number' ? val : 0);
              }, 0)
            };
          }
          return row;
        })
      }))
    };
    return transformed;
  }
});

// Example with FLAT format and mapper
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.FLAT,
  useFirstRowAsHeaders: true,
  mapper: (data) => {
    // If flat format from single sheet
    if ('data' in data && Array.isArray(data.data)) {
      return data.data.map((row: any) => ({
        ...row,
        // Add validations or transformations
        isValid: Object.values(row).every(val => val !== null && val !== undefined)
      }));
    }
    return data;
  }
});

// Example with DETAILED format and mapper
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.DETAILED,
  mapper: (data) => {
    // Group cells by sheet
    const groupedBySheet: Record<string, typeof data.cells> = {};
    data.cells.forEach(cell => {
      if (!groupedBySheet[cell.sheet]) {
        groupedBySheet[cell.sheet] = [];
      }
      groupedBySheet[cell.sheet].push(cell);
    });
    return {
      sheets: Object.keys(groupedBySheet).map(sheetName => ({
        name: sheetName,
        cells: groupedBySheet[sheetName]
      }))
    };
  }
});
```

**Reading options:**

```typescript
interface IExcelReaderOptions {
  outputFormat?: 'worksheet' | 'detailed' | 'flat' | OutputFormat; // Output format
  mapper?: (data: IJsonWorkbook | IDetailedFormat | IFlatFormat | IFlatFormatMultiSheet) => unknown; // Function to transform the response
  useFirstRowAsHeaders?: boolean;    // Use first row as headers
  includeEmptyRows?: boolean;        // Include empty rows
  headers?: string[] | Record<number, string>; // Custom headers
  sheetName?: string | number;       // Sheet name or index
  startRow?: number;                 // Starting row (1-based)
  endRow?: number;                    // Ending row (1-based)
  startColumn?: number;               // Starting column (1-based)
  endColumn?: number;                 // Ending column (1-based)
  includeFormatting?: boolean;        // Include formatting information
  includeFormulas?: boolean;          // Include formulas
  datesAsISO?: boolean;               // Convert dates to ISO string
}
```

**Output formats:**

- **`worksheet`** (default): Complete structure with sheets, rows and cells
- **`detailed`**: Array of cells with position information (text, column, row, reference)
- **`flat`**: Just the data, without structure (flat arrays or objects)

#### `Worksheet`

Represents an individual worksheet.

```typescript
const worksheet = builder.addWorksheet('My Sheet', {
  tabColor: '#FF0000',
  defaultRowHeight: 20,
  defaultColWidth: 15,
  pageSetup: {
    orientation: 'portrait',
    paperSize: 9
  }
});

// Main methods
worksheet.addHeader(header);             // Add main header
worksheet.addSubHeaders(headers);        // Add sub-headers
worksheet.addRow(row);                   // Add data row
worksheet.addFooter(footer);             // Add footer
worksheet.addTable(config);              // Create new table
worksheet.finalizeTable();               // Finalize current table
worksheet.getTable(name);                // Get table by name
worksheet.validate();                    // Validate sheet
```

### Data Types

#### `CellType`

```typescript
enum CellType {
  STRING = 'string',        // Text
  NUMBER = 'number',        // Number
  BOOLEAN = 'boolean',      // True/False
  DATE = 'date',            // Date
  PERCENTAGE = 'percentage', // Percentage
  CURRENCY = 'currency',    // Currency
  LINK = 'link',           // Hyperlink
  FORMULA = 'formula'      // Formula
}
```

#### `NumberFormat`

```typescript
enum NumberFormat {
  GENERAL = 'General',
  NUMBER = '#,##0',
  NUMBER_DECIMALS = '#,##0.00',
  CURRENCY = '$#,##0.00',
  CURRENCY_INTEGER = '$#,##0',
  PERCENTAGE = '0%',
  PERCENTAGE_DECIMALS = '0.00%',
  DATE = 'dd/mm/yyyy',
  DATE_TIME = 'dd/mm/yyyy hh:mm',
  TIME = 'hh:mm:ss',
  CUSTOM = 'custom'
}
```

### Styling

#### `StyleBuilder`

Fluent API for creating cell styles.

```typescript
const style = new StyleBuilder()
  // Fonts
  .fontName('Arial')
  .fontSize(12)
  .fontBold()
  .fontItalic()
  .fontUnderline()
  .fontColor('#FF0000')
  
  // Backgrounds and borders
  .backgroundColor('#FFFF00')
  .border(BorderStyle.THIN, '#000000')
  .borderTop(BorderStyle.MEDIUM, '#000000')
  .borderLeft(BorderStyle.THIN, '#000000')
  .borderBottom(BorderStyle.THIN, '#000000')
  .borderRight(BorderStyle.THIN, '#000000')
  
  // Alignment
  .centerAlign()
  .leftAlign()
  .rightAlign()
  .horizontalAlign(HorizontalAlignment.CENTER)
  .verticalAlign(VerticalAlignment.MIDDLE)
  .wrapText()
  
  // Formats
  .numberFormat('$#,##0.00')
  .striped()
  
  // Conditional formatting
  .conditionalFormat({
    type: 'cellIs',
    operator: 'greaterThan',
    values: [1000],
    style: StyleBuilder.create()
      .backgroundColor('#90EE90')
      .fontColor('#006400')
      .build()
  })
  
  .build();

// Alternative static method
const style2 = StyleBuilder.create()
  .fontBold()
  .fontSize(14)
  .build();
```

#### `BorderStyle`

```typescript
enum BorderStyle {
  THIN = 'thin',
  MEDIUM = 'medium',
  THICK = 'thick',
  DOTTED = 'dotted',
  DASHED = 'dashed',
  DOUBLE = 'double',
  HAIR = 'hair',
  MEDIUM_DASHED = 'mediumDashed',
  DASH_DOT = 'dashDot',
  MEDIUM_DASH_DOT = 'mediumDashDot',
  DASH_DOT_DOT = 'dashDotDot',
  MEDIUM_DASH_DOT_DOT = 'mediumDashDotDot',
  SLANT_DASH_DOT = 'slantDashDot'
}
```

## 🎯 Advanced Examples

### Multiple Tables in a Sheet

```typescript
import { ExcelBuilder, CellType, StyleBuilder, BorderStyle } from 'han-excel-builder';

const builder = new ExcelBuilder();
const worksheet = builder.addWorksheet('Complete Report');

// ===== FIRST TABLE =====
worksheet.addTable({
  name: 'Sales',
  showBorders: true,
  showStripes: true,
  style: 'TableStyleLight1'
});

worksheet.addHeader({
  key: 'header-sales',
  type: CellType.STRING,
  value: 'SALES SUMMARY',
  mergeCell: true,
  styles: new StyleBuilder()
    .fontBold()
    .fontSize(16)
    .backgroundColor('#4472C4')
    .fontColor('#FFFFFF')
    .centerAlign()
    .build()
});

worksheet.addSubHeaders([
  { key: 'product', type: CellType.STRING, value: 'Product' },
  { key: 'sales', type: CellType.CURRENCY, value: 'Sales' }
]);

worksheet.addRow([
  { key: 'p1', type: CellType.STRING, value: 'Product A', header: 'Product' },
  { key: 'v1', type: CellType.CURRENCY, value: 1500, header: 'Sales' }
]);

worksheet.finalizeTable();

// ===== SECOND TABLE =====
worksheet.addTable({
  name: 'Employees',
  showBorders: true,
  showStripes: true,
  style: 'TableStyleMedium1'
});

worksheet.addHeader({
  key: 'header-employees',
  type: CellType.STRING,
  value: 'TOP EMPLOYEES',
  mergeCell: true,
  styles: new StyleBuilder()
    .fontBold()
    .fontSize(16)
    .backgroundColor('#70AD47')
    .fontColor('#FFFFFF')
    .centerAlign()
    .build()
});

worksheet.addSubHeaders([
  { key: 'name', type: CellType.STRING, value: 'Name' },
  { key: 'sales', type: CellType.CURRENCY, value: 'Sales' }
]);

worksheet.addRow([
  { key: 'e1', type: CellType.STRING, value: 'John Doe', header: 'Name' },
  { key: 've1', type: CellType.CURRENCY, value: 150000, header: 'Sales' }
]);

worksheet.finalizeTable();

await builder.generateAndDownload('multiple-tables.xlsx');
```

### Nested Headers

```typescript
worksheet.addSubHeaders([
  {
    key: 'sales',
    value: 'Sales',
    type: CellType.STRING,
    children: [
      {
        key: 'sales-q1',
        value: 'Q1',
        type: CellType.STRING
      },
      {
        key: 'sales-q2',
        value: 'Q2',
        type: CellType.STRING
      }
    ]
  },
  {
    key: 'expenses',
    value: 'Expenses',
    type: CellType.STRING,
    children: [
      {
        key: 'expenses-q1',
        value: 'Q1',
        type: CellType.STRING
      },
      {
        key: 'expenses-q2',
        value: 'Q2',
        type: CellType.STRING
      }
    ]
  }
]);
```

### Cell Comments

Add comments to cells for additional context or notes:

```typescript
// Add comment to a header
worksheet.addHeader({
  key: 'header-1',
  type: CellType.STRING,
  value: 'Sales Report',
  comment: 'This is the main title of the report'
});

// Add comment to a data cell
worksheet.addRow([
  {
    key: 'product-1',
    value: 'Product A',
    type: CellType.STRING,
    header: 'Product',
    comment: 'Best selling product this month'
  },
  {
    key: 'sales-1',
    value: 1500.50,
    type: CellType.CURRENCY,
    header: 'Sales',
    comment: 'Sales increased 15% from last month'
  }
]);

// Comments are also supported in subheaders and footers
worksheet.addSubHeaders([
  {
    key: 'product',
    type: CellType.STRING,
    value: 'Product',
    comment: 'Product name column'
  }
]);
```

When reading Excel files, comments are included in the output:

```typescript
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.WORKSHEET
});

if (result.success) {
  result.data.sheets.forEach(sheet => {
    sheet.rows.forEach(row => {
      row.cells.forEach(cell => {
        if (cell.comment) {
          console.log(`Cell ${cell.reference}: ${cell.comment}`);
        }
      });
    });
  });
}
```

### Hyperlinks

```typescript
worksheet.addRow([
  {
    key: 'link-1',
    type: CellType.LINK,
    value: 'Visit site',
    link: 'https://example.com',
    mask: 'Click here', // Visible text
    header: 'Link'
  }
]);
```

### Data Validation

Apply data validation rules to cells to restrict input values:

```typescript
// List validation (dropdown)
worksheet.addRow([
  {
    key: 'status-1',
    value: 'Active',
    type: CellType.STRING,
    header: 'Status',
    validation: {
      type: 'list',
      formula1: 'Active,Inactive,Pending',
      showErrorMessage: true,
      errorMessage: 'Please select a valid status',
      showInputMessage: true,
      inputMessage: 'Select status from the list',
      allowBlank: false
    }
  }
]);

// Number range validation
worksheet.addRow([
  {
    key: 'age-1',
    value: 25,
    type: CellType.NUMBER,
    header: 'Age',
    validation: {
      type: 'whole',
      operator: 'between',
      formula1: 18,
      formula2: 100,
      showErrorMessage: true,
      errorMessage: 'Age must be between 18 and 100',
      allowBlank: false
    }
  }
]);

// Date validation
worksheet.addRow([
  {
    key: 'date-1',
    value: new Date(),
    type: CellType.DATE,
    header: 'Date',
    validation: {
      type: 'date',
      operator: 'greaterThan',
      formula1: new Date('2020-01-01'),
      showErrorMessage: true,
      errorMessage: 'Date must be after 2020-01-01',
      allowBlank: true
    }
  }
]);

// Text length validation
worksheet.addRow([
  {
    key: 'name-1',
    value: 'John Doe',
    type: CellType.STRING,
    header: 'Name',
    validation: {
      type: 'textLength',
      operator: 'lessThanOrEqual',
      formula1: 50,
      showErrorMessage: true,
      errorMessage: 'Name must be 50 characters or less',
      allowBlank: false
    }
  }
]);

// Custom formula validation
worksheet.addRow([
  {
    key: 'value-1',
    value: 100,
    type: CellType.NUMBER,
    header: 'Value',
    validation: {
      type: 'custom',
      formula1: '=A1>0',
      showErrorMessage: true,
      errorMessage: 'Value must be greater than 0',
      allowBlank: false
    }
  }
]);
```

### Auto Filters

Enable automatic filters for tables and worksheets to allow users to filter data:

```typescript
// Enable auto filter for a table
worksheet.addTable({
  name: 'Sales',
  autoFilter: true, // Enable auto filter for this table
  showBorders: true
});

// Enable auto filter at worksheet level
const worksheet = builder.addWorksheet('Sales Report', {
  autoFilter: {
    enabled: true,
    startRow: 2, // Start from row 2 (skip header)
    endRow: 100, // End at row 100
    startColumn: 1,
    endColumn: 5
  }
});

// Or use a range
const worksheet = builder.addWorksheet('Sales Report', {
  autoFilter: {
    enabled: true,
    range: {
      start: { row: 2, col: 1, reference: 'A2' },
      end: { row: 100, col: 5, reference: 'E100' }
    }
  }
});
```

### Data with Children (Hierarchical Structure)

```typescript
worksheet.addRow([
  {
    key: 'row-1',
    type: CellType.STRING,
    value: 'Main Category',
    header: 'Category',
    children: [
      {
        key: 'child-1',
        type: CellType.STRING,
        value: 'Subcategory 1',
        header: 'Subcategory'
      },
      {
        key: 'child-2',
        type: CellType.NUMBER,
        value: 100,
        header: 'Value'
      }
    ]
  }
]);
```

### Conditional Formatting

Apply conditional formatting rules to cells based on values, formulas, or conditions:

```typescript
worksheet.addRow([
  {
    key: 'sales-1',
    type: CellType.NUMBER,
    value: 1500,
    header: 'Sales',
    styles: new StyleBuilder()
  .conditionalFormat({
    type: 'cellIs',
    operator: 'greaterThan',
    values: [1000],
    style: StyleBuilder.create()
      .backgroundColor('#90EE90')
      .fontColor('#006400')
      .build()
  })
      .conditionalFormat({
        type: 'cellIs',
        operator: 'lessThan',
        values: [500],
        style: StyleBuilder.create()
          .backgroundColor('#FFB6C1')
          .fontColor('#8B0000')
          .build()
      })
      .build()
  }
]);

// Multiple conditional formats with different types
worksheet.addRow([
  {
    key: 'status-1',
    type: CellType.STRING,
    value: 'Active',
    header: 'Status',
    styles: new StyleBuilder()
      .conditionalFormat({
        type: 'containsText',
        operator: 'equal',
        values: ['Active'],
        style: StyleBuilder.create()
          .backgroundColor('#90EE90')
          .build()
      })
      .build()
  }
]);
```

### Freeze Panes

Freeze rows and columns to keep headers visible while scrolling:

```typescript
const worksheet = builder.addWorksheet('Sales Report', {
  freezePanes: {
    row: 2,        // Freeze from row 2 (header row)
    col: 1,        // Freeze from column 1
    reference: 'A2' // Optional: cell reference
  }
});
```

### Worksheet Protection

Protect worksheets with password and configure permissions:

```typescript
const worksheet = builder.addWorksheet('Protected Sheet', {
  protected: true,
  protectionPassword: 'mypassword123',
  // Other protection options are handled by ExcelJS defaults
});
```

### Images/Pictures

Add images to worksheets (logos, charts, signatures, etc.):

```typescript
// From ArrayBuffer (e.g., from fetch)
const response = await fetch('https://example.com/logo.png');
const arrayBuffer = await response.arrayBuffer();

worksheet.addImage({
  buffer: arrayBuffer,
  extension: 'png',
  position: {
    row: 1,
    col: 1
  },
  size: {
    width: 200,
    height: 100
  },
  description: 'Company Logo'
});

// From base64 string
worksheet.addImage({
  buffer: 'data:image/png;base64,iVBORw0KGgoAAAANS...',
  extension: 'png',
  position: {
    row: 'A2', // Can use cell reference
    col: 1
  },
  size: {
    scaleX: 0.5, // Scale to 50%
    scaleY: 0.5
  },
  hyperlink: 'https://example.com'
});
```

### Row/Column Grouping

Group rows and columns to create collapsible outlines:

```typescript
// Group rows 2-10 (collapsible)
worksheet.groupRows(2, 10, true);

// Group columns A-C
worksheet.groupColumns(1, 3, false);

// Nested grouping
worksheet.groupRows(2, 5, false);   // Level 1
worksheet.groupRows(2, 3, false);   // Level 2 (nested)
```

### Named Ranges

Define named ranges for easy references in formulas:

```typescript
// Using string range
worksheet.addNamedRange('SalesData', 'A1:D100');

// Using ICellRange
worksheet.addNamedRange('HeaderRow', {
  start: { row: 1, col: 1, reference: 'A1' },
  end: { row: 1, col: 5, reference: 'E1' }
});

// Named range with scope (worksheet-specific)
worksheet.addNamedRange('LocalRange', 'A1:A10', 'Sheet1');
```

### Excel Structured Tables

Create Excel structured tables with automatic styles and features:

```typescript
// First, add data to the worksheet
worksheet.addSubHeaders([
  { key: 'product', value: 'Product', type: CellType.STRING },
  { key: 'sales', value: 'Sales', type: CellType.NUMBER },
  { key: 'revenue', value: 'Revenue', type: CellType.CURRENCY }
]);

// Add data rows...
worksheet.addRow([...]);

// Then add Excel structured table
worksheet.addExcelTable({
  name: 'SalesTable',
  range: {
    start: 'A1',
    end: 'C10'
  },
  style: 'TableStyleMedium2',
  headerRow: true,
  totalRow: true,
  columns: [
    { name: 'Product', filterButton: true },
    { name: 'Sales', filterButton: true, totalsRowFunction: 'sum' },
    { name: 'Revenue', filterButton: true, totalsRowFunction: 'sum' }
  ]
});
```

### Advanced Print Settings

Configure headers, footers, and repeat rows/columns:

```typescript
const worksheet = builder.addWorksheet('Report', {
  printHeadersFooters: {
    header: {
      left: 'Company Name',
      center: 'Sales Report',
      right: 'Page &P of &N'
    },
    footer: {
      left: 'Confidential',
      center: 'Generated on &D',
      right: '&F'
    }
  },
  printRepeat: {
    rows: [1, 2], // Repeat header rows on each page
    columns: 'A:B' // Repeat first two columns
  }
});
```

### Hide/Show Rows & Columns

Hide or show specific rows and columns:

```typescript
// Hide single row
worksheet.hideRows(5);

// Hide multiple rows
worksheet.hideRows([3, 4, 5, 10]);

// Show rows (if previously hidden)
worksheet.showRows([3, 4]);

// Hide columns by number or letter
worksheet.hideColumns(1);        // Column A
worksheet.hideColumns('B');       // Column B
worksheet.hideColumns([1, 2, 3]); // Columns A, B, C
worksheet.hideColumns(['A', 'D']); // Columns A and D

// Show columns
worksheet.showColumns([1, 2]);
```

### Rich Text in Cells

Format text with multiple styles within a single cell:

```typescript
worksheet.addRow([
  {
    key: 'rich-text-1',
    type: CellType.STRING,
    value: '', // Empty value when using richText
    header: 'Description',
    richText: [
      {
        text: 'This is ',
        font: 'Arial',
        size: 11
      },
      {
        text: 'bold',
        font: 'Arial',
        size: 11,
        bold: true,
        color: '#FF0000'
      },
      {
        text: ' and ',
        font: 'Arial',
        size: 11
      },
      {
        text: 'italic',
        font: 'Arial',
        size: 11,
        italic: true,
        color: '#0000FF'
      },
      {
        text: ' text!',
        font: 'Arial',
        size: 11
      }
    ]
  }
]);
```

### Cell-level Protection

Protect individual cells with lock/hide options:

```typescript
worksheet.addRow([
  {
    key: 'protected-1',
    type: CellType.STRING,
    value: 'Locked Cell',
    header: 'Status',
    cellProtection: {
      locked: true,   // Cell cannot be edited
      hidden: false   // Formula is visible
    }
  },
  {
    key: 'unlocked-1',
    type: CellType.STRING,
    value: 'Editable Cell',
    header: 'Notes',
    cellProtection: {
      locked: false,  // Cell can be edited
      hidden: false
    }
  },
  {
    key: 'hidden-1',
    type: CellType.FORMULA,
    value: '=SUM(A1:A10)',
    header: 'Total',
    cellProtection: {
      locked: true,
      hidden: true    // Formula is hidden (shows only result)
    }
  }
]);

// Note: Worksheet must be protected for cell protection to take effect
const worksheet = builder.addWorksheet('Protected Sheet', {
  protected: true,
  protectionPassword: 'password123'
});
```

### Pivot Tables

Create pivot tables for data analysis:

```typescript
// First, create a data sheet
const dataSheet = builder.addWorksheet('Sales Data');
dataSheet.addSubHeaders([
  { key: 'category', value: 'Category', type: CellType.STRING },
  { key: 'product', value: 'Product', type: CellType.STRING },
  { key: 'sales', value: 'Sales', type: CellType.NUMBER },
  { key: 'revenue', value: 'Revenue', type: CellType.CURRENCY }
]);

// Add data rows...
dataSheet.addRow([...]);

// Create a pivot table sheet
const pivotSheet = builder.addWorksheet('Pivot Analysis');
pivotSheet.addPivotTable({
  name: 'SalesPivot',
  ref: 'A1',
  sourceRange: 'A1:D100',
  sourceSheet: 'Sales Data',
  fields: {
    rows: ['Category', 'Product'],
    columns: [],
    values: [
      { name: 'Sales', stat: 'sum' },
      { name: 'Revenue', stat: 'sum' }
    ],
    filters: []
  },
  options: {
    showRowGrandTotals: true,
    showColGrandTotals: true,
    showHeaders: true
  }
});
```

### Slicers

Add visual filters (slicers) to tables and pivot tables:

```typescript
// Note: Slicers require advanced ExcelJS XML manipulation
// This feature is documented but requires manual XML editing for full implementation

worksheet.addSlicer({
  name: 'CategorySlicer',
  targetTable: 'SalesTable',
  column: 'Category',
  position: {
    row: 1,
    col: 'F'
  },
  size: {
    width: 200,
    height: 300
  }
});
```

### Watermarks

Add watermarks to worksheets (text or image):

```typescript
// Text watermark
worksheet.addWatermark({
  text: 'CONFIDENTIAL',
  position: {
    horizontal: 'center',
    vertical: 'middle'
  },
  opacity: 0.3,
  fontSize: 72,
  fontColor: '#CCCCCC',
  rotation: -45
});

// Image watermark
const watermarkImage = await fetch('https://example.com/watermark.png')
  .then(r => r.arrayBuffer());

worksheet.addWatermark({
  image: {
    buffer: watermarkImage,
    extension: 'png',
    position: {
      row: 500,
      col: 10
    },
    size: {
      width: 400,
      height: 300,
      scaleX: 0.3,
      scaleY: 0.3
    }
  },
  position: {
    horizontal: 'center',
    vertical: 'middle'
  },
  opacity: 0.3
});
```

### Cell-level Page Breaks

Add manual page breaks at row level:

```typescript
worksheet.addRow([
  {
    key: 'row-1',
    type: CellType.STRING,
    value: 'Data Row 1',
    header: 'Name'
  }
]);

// Add page break before this row
worksheet.addRow([
  {
    key: 'row-2',
    type: CellType.STRING,
    value: 'Data Row 2',
    header: 'Name',
    pageBreak: true  // Page break before this row
  }
]);

// Page breaks also work in headers and footers
worksheet.addHeader({
  key: 'section-header',
  value: 'New Section',
  type: CellType.STRING,
  pageBreak: true  // Page break before this header
});
```

### Cell Styles (Predefined)

Create reusable cell styles for consistency across your workbook:

```typescript
import { ExcelBuilder, StyleBuilder, CellType } from 'han-excel-builder';

const builder = new ExcelBuilder();

// Define reusable styles
builder.addCellStyle('headerStyle', StyleBuilder.create()
  .font({ name: 'Arial', size: 14, bold: true })
  .fill({ backgroundColor: '#4472C4' })
  .fontColor('#FFFFFF')
  .build()
);

builder.addCellStyle('highlightStyle', StyleBuilder.create()
  .fill({ backgroundColor: '#FFE699' })
  .font({ bold: true })
  .build()
);

// Use predefined styles in cells
const sheet = builder.addWorksheet('Data');
sheet.addHeader({
  key: 'name',
  value: 'Name',
  type: CellType.STRING,
  styleName: 'headerStyle' // Use predefined style
});

sheet.addRow({
  key: 'row1',
  header: 'name',
  value: 'John Doe',
  type: CellType.STRING,
  styleName: 'highlightStyle' // Use predefined style
});
```

### Themes

Apply color themes to the entire workbook:

```typescript
const builder = new ExcelBuilder();

// Set workbook theme
builder.setTheme({
  name: 'Corporate Theme',
  colors: {
    dark1: '#000000',
    light1: '#FFFFFF',
    dark2: '#1F4E78',
    light2: '#EEECE1',
    accent1: '#4472C4',
    accent2: '#ED7D31',
    accent3: '#A5A5A5',
    accent4: '#FFC000',
    accent5: '#5B9BD5',
    accent6: '#70AD47',
    hyperlink: '#0563C1',
    followedHyperlink: '#954F72'
  },
  fonts: {
    major: {
      latin: 'Calibri',
      eastAsian: 'Calibri',
      complexScript: 'Calibri'
    },
    minor: {
      latin: 'Calibri',
      eastAsian: 'Calibri',
      complexScript: 'Calibri'
    }
  }
});

// Theme colors will be applied throughout the workbook
const sheet = builder.addWorksheet('Report');
// ... add data
```

### Split Panes

Divide the window into panes for comparing distant sections:

```typescript
const sheet = builder.addWorksheet('Data', {
  splitPanes: {
    xSplit: 3, // Split after column C
    ySplit: 5, // Split after row 5
    topLeftCell: 'D6', // Top-left cell in bottom-right pane
    activePane: 'bottomRight' // Active pane: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight'
  }
});
```

### Sheet Views

Configure different views of the same sheet:

```typescript
const sheet = builder.addWorksheet('Report', {
  views: {
    state: 'normal', // 'normal' | 'pageBreakPreview' | 'pageLayout'
    zoomScale: 100, // Zoom level (10-400)
    zoomScaleNormal: 100, // Normal zoom level
    showGridLines: true,
    showRowColHeaders: true,
    showRuler: true, // For page layout view
    rightToLeft: false
  }
});
```

### Data Connections

Add external data connections:

```typescript
// Note: Data connections require advanced ExcelJS XML manipulation
// This feature is documented but requires manual XML editing for full implementation

worksheet.addDataConnection({
  name: 'SalesDB',
  type: 'odbc',
  connectionString: 'Driver={SQL Server};Server=server;Database=SalesDB;',
  commandText: 'SELECT * FROM Sales WHERE Year = 2024',
  refresh: {
    refreshOnOpen: true,
    refreshInterval: 60  // minutes
  },
  credentials: {
    username: 'user',
    integratedSecurity: false
    // Password should be set by user in Excel after opening
  }
});
```

### Multiple Worksheets

```typescript
const builder = new ExcelBuilder();

// Sheet 1: Summary
const summarySheet = builder.addWorksheet('Summary');
summarySheet.addHeader({
  key: 'title',
  value: 'Executive Summary',
  type: CellType.STRING,
  mergeCell: true
});

// Sheet 2: Details
const detailsSheet = builder.addWorksheet('Details');
detailsSheet.addSubHeaders([
  { key: 'date', value: 'Date', type: CellType.DATE },
  { key: 'amount', value: 'Amount', type: CellType.CURRENCY }
]);

await builder.generateAndDownload('multi-sheet-report.xlsx');
```

### Export in Different Formats

```typescript
// Direct download (browser)
await builder.generateAndDownload('report.xlsx');

// Get as Buffer
const bufferResult = await builder.toBuffer();
if (bufferResult.success) {
  const buffer = bufferResult.data;
  // Use buffer...
}

// Get as Blob
const blobResult = await builder.toBlob();
if (blobResult.success) {
  const blob = blobResult.data;
  // Use blob...
}
```

### Event System

```typescript
builder.on('build:started', (event) => {
  console.log('Build started');
});

builder.on('build:completed', (event) => {
  console.log('Build completed', event.data);
});

builder.on('build:error', (event) => {
  console.error('Build error', event.data.error);
});

// Remove listener
const listenerId = builder.on('build:started', handler);
builder.off('build:started', listenerId);
```

### Read Excel and Convert to JSON

```typescript
import { ExcelReader } from 'han-excel-builder';

// Read from a file (browser)
const fileInput = document.querySelector('input[type="file"]');
fileInput.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const result = await ExcelReader.fromFile(file, {
    useFirstRowAsHeaders: true,
    datesAsISO: true,
    includeFormatting: false
  });

  if (result.success) {
    const workbook = result.data;
    
    // Process each sheet
    workbook.sheets.forEach(sheet => {
      console.log(`Processing sheet: ${sheet.name}`);
      
      // Convert to array of objects (if using headers)
      const data = sheet.rows.map(row => row.data || {});
      console.log('Data:', data);
    });
  }
});

// Read from ArrayBuffer (from API)
async function readExcelFromAPI() {
  const response = await fetch('/api/excel-file');
  const buffer = await response.arrayBuffer();
  
  const result = await ExcelReader.fromBuffer(buffer, {
    useFirstRowAsHeaders: true,
    sheetName: 'Sales' // Read only 'Sales' sheet
  });

  if (result.success) {
    const sheet = result.data.sheets[0];
    const sales = sheet.rows.map(row => row.data);
    return sales;
  }
}

// Read from path (Node.js)
async function readExcelFromPath() {
  const result = await ExcelReader.fromPath('./report.xlsx', {
    useFirstRowAsHeaders: true,
    startRow: 2, // Skip header
    includeFormulas: true
  });

  if (result.success) {
    console.log(`Processing time: ${result.processingTime}ms`);
    return result.data;
  }
}
```

## 🧪 Testing

```bash
# Run tests
npm test

# Run tests with coverage
npm run test:coverage

# Run tests in watch mode
npm run test:watch
```

## 🛠️ Development

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Run linting
npm run lint

# Format code
npm run format

# Type checking
npm run type-check

# Generate documentation
npm run docs
```

## 📋 Migration from legacy-excel

If you're migrating from the legacy version, here's a quick comparison:

```typescript
// Legacy way
const worksheets: IWorksheets[] = [{
  name: "Report",
  tables: [{
    headers: [...],
    subHeaders: [...],
    body: [...],
    footers: [...]
  }]
}];
await fileBuilder(worksheets, "report");

// New way
const builder = new ExcelBuilder();
const worksheet = builder.addWorksheet('Report');
worksheet.addHeader({...});
worksheet.addSubHeaders([...]);
worksheet.addRow([...]);
worksheet.addFooter([...]);
await builder.generateAndDownload('report');
```

## 📚 Additional Resources

- 📖 [Multiple Tables Guide](./MULTIPLE-TABLES-GUIDE.md)
- 📖 [Implemented Improvements](./IMPROVEMENTS.md)
- 📖 [Test Results](./TEST-RESULTS.md)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m 'Add my feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- 📖 [Documentation](https://github.com/hannndler/-han-excel)
- 🐛 [Issues](https://github.com/hannndler/-han-excel/issues)
- 💬 [Discussions](https://github.com/hannndler/-han-excel/discussions)

---

Made with ❤️ by the Han Excel Team 
