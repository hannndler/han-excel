# Han Excel Builder

🚀 **Advanced Excel file generator with TypeScript support, comprehensive styling, and optimized performance**

A modern, fully-typed library for creating complex Excel reports with multiple worksheets, advanced styling, and high performance.

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
  detailed.cells.forEach(cell => {
    console.log(`${cell.sheet}!${cell.reference}: ${cell.text}`);
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
      .build()
  }
]);
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
