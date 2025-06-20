# Han Excel Builder

🚀 **Advanced Excel file generator with TypeScript support, comprehensive styling, and optimized performance**

A modern, fully-typed library for creating complex Excel reports with multiple worksheets, advanced styling, and high performance.

## ✨ Features

- 📊 **Multiple Worksheets Support** - Create complex workbooks with multiple sheets
- 🎨 **Advanced Styling** - Full control over fonts, colors, borders, and cell formatting
- 📈 **Data Types** - Support for strings, numbers, dates, percentages, and custom formats
- 🔧 **TypeScript First** - Complete type safety with comprehensive interfaces
- ⚡ **High Performance** - Optimized for large datasets with streaming support
- 🧪 **Fully Tested** - Comprehensive test suite with 100% coverage
- 📚 **Well Documented** - Complete API documentation with examples
- 🛠️ **Developer Friendly** - ESLint, Prettier, and modern tooling

## 📦 Installation

```bash
npm install han-excel-builder
# or
yarn add han-excel-builder
# or
pnpm add han-excel-builder
```

## 🚀 Quick Start

```typescript
import { ExcelBuilder, CellType, NumberFormat, StyleBuilder } from 'han-excel-builder';

// Create a simple report
const builder = new ExcelBuilder();

const worksheet = builder.addWorksheet('Sales Report');

// Add headers
worksheet.addHeader({
  key: 'title',
  value: 'Monthly Sales Report',
  type: CellType.STRING,
  mergeCell: true,
  styles: StyleBuilder.create()
    .fontBold()
    .fontSize(16)
    .centerAlign()
    .build()
});

// Add sub-headers
worksheet.addSubHeaders([
  {
    key: 'product',
    value: 'Product',
    type: CellType.STRING,
    width: 20
  },
  {
    key: 'sales',
    value: 'Sales',
    type: CellType.NUMBER,
    width: 15,
    numberFormat: NumberFormat.CURRENCY
  }
]);

// Add data
worksheet.addRow([
  { key: 'product', value: 'Product A', type: CellType.STRING },
  { key: 'sales', value: 1500.50, type: CellType.NUMBER }
]);

// Generate and download
await builder.generateAndDownload('sales-report');
```

## 📚 API Documentation

### Core Classes

#### `ExcelBuilder`
Main class for creating Excel workbooks.

```typescript
const builder = new ExcelBuilder({
  author: 'Your Name',
  company: 'Your Company',
  created: new Date()
});
```

#### `Worksheet`
Represents a single worksheet in the workbook.

```typescript
const worksheet = builder.addWorksheet('Sheet Name', {
  tabColor: '#FF0000',
  defaultRowHeight: 20,
  defaultColWidth: 15
});
```

### Data Types

#### `CellType`
- `STRING` - Text values
- `NUMBER` - Numeric values
- `BOOLEAN` - True/false values
- `DATE` - Date values
- `PERCENTAGE` - Percentage values
- `CURRENCY` - Currency values

#### `NumberFormat`
- `GENERAL` - Default format
- `NUMBER` - Number with optional decimals
- `CURRENCY` - Currency format
- `PERCENTAGE` - Percentage format
- `DATE` - Date format
- `TIME` - Time format
- `CUSTOM` - Custom format string

### Styling

#### `StyleBuilder`
Fluent API for creating cell styles.

```typescript
const style = StyleBuilder.create()
  .fontBold()
  .fontSize(12)
  .fontColor('#FF0000')
  .backgroundColor('#FFFF00')
  .border('thin', '#000000')
  .centerAlign()
  .verticalAlign('middle')
  .wrapText()
  .build();
```

## 🎯 Advanced Examples

### Complex Report with Multiple Worksheets

```typescript
import { ExcelBuilder, CellType, NumberFormat, StyleBuilder } from 'han-excel-builder';

const builder = new ExcelBuilder({
  author: 'Report Generator',
  company: 'Your Company'
});

// Summary worksheet
const summarySheet = builder.addWorksheet('Summary');
summarySheet.addHeader({
  key: 'title',
  value: 'Annual Report Summary',
  type: CellType.STRING,
  mergeCell: true,
  styles: StyleBuilder.create().fontBold().fontSize(18).centerAlign().build()
});

// Detailed worksheet
const detailSheet = builder.addWorksheet('Details');
detailSheet.addSubHeaders([
  { key: 'date', value: 'Date', type: CellType.DATE, width: 12 },
  { key: 'category', value: 'Category', type: CellType.STRING, width: 15 },
  { key: 'amount', value: 'Amount', type: CellType.NUMBER, width: 12, numberFormat: NumberFormat.CURRENCY },
  { key: 'percentage', value: '%', type: CellType.PERCENTAGE, width: 8 }
]);

// Add data with alternating row colors
data.forEach((row, index) => {
  const rowStyle = index % 2 === 0 
    ? StyleBuilder.create().backgroundColor('#F0F0F0').build()
    : undefined;
    
  detailSheet.addRow([
    { key: 'date', value: row.date, type: CellType.DATE },
    { key: 'category', value: row.category, type: CellType.STRING },
    { key: 'amount', value: row.amount, type: CellType.NUMBER },
    { key: 'percentage', value: row.percentage, type: CellType.PERCENTAGE }
  ], rowStyle);
});

await builder.generateAndDownload('annual-report');
```

### Conditional Styling

```typescript
const style = StyleBuilder.create()
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
worksheet.addHeaders([...]);
worksheet.addSubHeaders([...]);
worksheet.addRows([...]);
await builder.generateAndDownload('report');
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- 📖 [Documentation](https://github.com/your-org/han-excel-builder/docs)
- 🐛 [Issues](https://github.com/your-org/han-excel-builder/issues)
- 💬 [Discussions](https://github.com/your-org/han-excel-builder/discussions)

---

Made with ❤️ by the Han Excel Team 