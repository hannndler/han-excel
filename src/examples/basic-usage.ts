/**
 * Basic Usage Example
 * 
 * This example demonstrates the basic usage of Han Excel Builder
 */

import { ExcelBuilder, CellType, NumberFormat, StyleBuilder, BorderStyle } from '../index';

/**
 * Basic usage example demonstrating the core functionality
 */
export async function basicUsageExample() {
  // Create a new Excel builder
  const builder = new ExcelBuilder({
    metadata: {
      title: 'Basic Usage Example',
      author: 'Han Excel Builder',
      description: 'Demonstrates basic Excel generation functionality'
    },
  });

  // Add a worksheet
  const worksheet = builder.addWorksheet('Sales Data');

  // Add header
  worksheet.addHeader({
    key: 'header-1',
    type: CellType.STRING,
    value: 'Sales Report - Q1 2024',
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

  // Add subheaders
  worksheet.addSubHeaders([
    {
      key: 'subheader-1',
      type: CellType.STRING,
      value: 'Product',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'subheader-2',
      type: CellType.STRING,
      value: 'Category',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'subheader-3',
      type: CellType.STRING,
      value: 'Sales',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'subheader-4',
      type: CellType.STRING,
      value: 'Revenue',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Add data rows
  const salesData = [
    { product: 'Laptop', category: 'Electronics', sales: 150, revenue: 225000 },
    { product: 'Mouse', category: 'Electronics', sales: 300, revenue: 15000 },
    { product: 'Keyboard', category: 'Electronics', sales: 200, revenue: 40000 },
    { product: 'Monitor', category: 'Electronics', sales: 75, revenue: 112500 },
    { product: 'Desk', category: 'Furniture', sales: 50, revenue: 25000 },
    { product: 'Chair', category: 'Furniture', sales: 80, revenue: 32000 },
    { product: 'Lamp', category: 'Furniture', sales: 120, revenue: 18000 }
  ];

  salesData.forEach((item, index) => {
    worksheet.addRow([
      {
        key: `product-${index}`,
        type: CellType.STRING,
        value: item.product,
        header: 'Product',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `category-${index}`,
        type: CellType.STRING,
        value: item.category,
        header: 'Category',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `sales-${index}`,
        type: CellType.NUMBER,
        value: item.sales,
        header: 'Sales',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `revenue-${index}`,
        type: CellType.CURRENCY,
        value: item.revenue,
        header: 'Revenue',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      }
    ]);
  });

  // Add footer with totals
  worksheet.addFooter([
    {
      key: 'footer-1',
      type: CellType.STRING,
      value: 'Total',
      header: 'Total',
      mergeCell: true,
      mergeTo: 2,
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .rightAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'footer-2',
      type: CellType.NUMBER,
      value: salesData.reduce((sum, item) => sum + item.sales, 0),
      header: 'Total Sales',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'footer-3',
      type: CellType.CURRENCY,
      value: salesData.reduce((sum, item) => sum + item.revenue, 0),
      header: 'Total Revenue',
      numberFormat: '$#,##0',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Add a second worksheet for summary
  const summarySheet = builder.addWorksheet('Summary');

  // Add header to summary sheet
  summarySheet.addHeader({
    key: 'summary-header',
    type: CellType.STRING,
    value: 'Sales Summary by Category',
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

  // Add subheaders to summary sheet
  summarySheet.addSubHeaders([
    {
      key: 'summary-subheader-1',
      type: CellType.STRING,
      value: 'Category',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'summary-subheader-2',
      type: CellType.STRING,
      value: 'Total Sales',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'summary-subheader-3',
      type: CellType.STRING,
      value: 'Total Revenue',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Calculate summary data
  const categorySummary = salesData.reduce((acc, item) => {
    if (!acc[item.category]) {
      acc[item.category] = { sales: 0, revenue: 0 };
    }
    acc[item.category]!.sales += item.sales;
    acc[item.category]!.revenue += item.revenue;
    return acc;
  }, {} as Record<string, { sales: number; revenue: number }>);

  // Add summary rows
  Object.entries(categorySummary).forEach(([category, data], index) => {
    summarySheet.addRow([
      {
        key: `summary-category-${index}`,
        type: CellType.STRING,
        value: category,
        header: 'Category',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN)
          .build()
      },
      {
        key: `summary-sales-${index}`,
        type: CellType.NUMBER,
        value: data.sales,
        header: 'Total Sales',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN)
          .build()
      },
      {
        key: `summary-revenue-${index}`,
        type: CellType.CURRENCY,
        value: data.revenue,
        header: 'Total Revenue',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN)
          .build()
      }
    ]);
  });

  // Add a third worksheet for performance data
  const performanceSheet = builder.addWorksheet('Performance');

  // Add header
  performanceSheet.addHeader({
    key: 'performance-header',
    type: CellType.STRING,
    value: 'Monthly Performance Data',
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

  // Add subheaders
  performanceSheet.addSubHeaders([
    { key: 'month', type: CellType.STRING, value: 'Month' },
    { key: 'target', type: CellType.NUMBER, value: 'Target' },
    { key: 'actual', type: CellType.NUMBER, value: 'Actual' },
    { key: 'variance', type: CellType.PERCENTAGE, value: 'Variance %' }
  ]);

  // Add performance data
  const performanceData = [
    { month: 'January', target: 100000, actual: 95000, variance: -5 },
    { month: 'February', target: 110000, actual: 115000, variance: 4.5 },
    { month: 'March', target: 120000, actual: 125000, variance: 4.2 },
    { month: 'April', target: 130000, actual: 128000, variance: -1.5 },
    { month: 'May', target: 140000, actual: 145000, variance: 3.6 },
    { month: 'June', target: 150000, actual: 152000, variance: 1.3 }
  ];

  performanceData.forEach((row) => {
    performanceSheet.addRow([
      {
        key: `month-${row.month}`,
        type: CellType.STRING,
        value: row.month,
        header: 'Month',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `target-${row.month}`,
        type: CellType.CURRENCY,
        value: row.target,
        header: 'Target',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `actual-${row.month}`,
        type: CellType.CURRENCY,
        value: row.actual,
        header: 'Actual',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `variance-${row.month}`,
        type: CellType.PERCENTAGE,
        value: row.variance / 100,
        header: 'Variance %',
        numberFormat: '0.0%',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .fontColor(row.variance >= 0 ? '#008000' : '#FF0000')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      }
    ]);
  });

  // Build and download the workbook
  const result = await builder.generateAndDownload('basic-usage-example.xlsx');
  
  if (result.success) {
    console.log('✅ Basic usage example completed successfully!');
  } else {
    console.error('❌ Basic usage example failed:', result.error);
  }
}

/**
 * Multiple worksheets example
 */
export async function createMultipleWorksheetsReport(): Promise<void> {
  const builder = new ExcelBuilder({
    metadata: {
      author: 'Han Excel Builder',
      title: 'Multi-Sheet Report',
      company: 'Example Corp'
    }
  });

  // Summary worksheet
  const summarySheet = builder.addWorksheet('Summary');
  summarySheet.addHeader({
    key: 'title',
    value: 'Executive Summary',
    type: CellType.STRING,
    mergeCell: true,
    styles: StyleBuilder.create()
      .fontBold()
      .fontSize(18)
      .centerAlign()
      .backgroundColor('#2E75B6')
      .fontColor('#FFFFFF')
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  });

  // Details worksheet
  const detailsSheet = builder.addWorksheet('Details');
  detailsSheet.addSubHeaders([
    {
      key: 'date',
      value: 'Date',
      type: CellType.DATE,
      styles: StyleBuilder.create()
        .fontBold()
        .backgroundColor('#D9E1F2')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'category',
      value: 'Category',
      type: CellType.STRING,
      styles: StyleBuilder.create()
        .fontBold()
        .backgroundColor('#D9E1F2')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'amount',
      value: 'Amount',
      type: CellType.NUMBER,
      styles: StyleBuilder.create()
        .fontBold()
        .backgroundColor('#D9E1F2')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Add sample data
  const detailsData = [
    { date: new Date('2024-01-15'), category: 'Revenue', amount: 150000 },
    { date: new Date('2024-01-16'), category: 'Expenses', amount: -45000 },
    { date: new Date('2024-01-17'), category: 'Revenue', amount: 180000 },
    { date: new Date('2024-01-18'), category: 'Expenses', amount: -52000 }
  ];

  detailsData.forEach((row, index) => {
    detailsSheet.addRow([
      {
        key: 'date',
        value: row.date,
        type: CellType.DATE,
        header: 'Date',
        styles: StyleBuilder.create()
          .border(BorderStyle.THIN, '#8EAADB')
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .build()
      },
      {
        key: 'category',
        value: row.category,
        type: CellType.STRING,
        header: 'Category',
        styles: StyleBuilder.create()
          .border(BorderStyle.THIN, '#8EAADB')
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .build()
      },
      {
        key: 'amount',
        value: row.amount,
        type: CellType.NUMBER,
        header: 'Amount',
        styles: StyleBuilder.create()
          .border(BorderStyle.THIN, '#8EAADB')
          .rightAlign()
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .build()
      }
    ]);
  });

  await builder.generateAndDownload('multi-sheet-report');
}

/**
 * Conditional formatting example
 */
export async function createConditionalFormattingExample(): Promise<void> {
  const builder = new ExcelBuilder();

  const worksheet = builder.addWorksheet('Performance Data');

  worksheet.addSubHeaders([
    {
      key: 'employee',
      value: 'Employee',
      type: CellType.STRING
    },
    {
      key: 'sales',
      value: 'Sales',
      type: CellType.NUMBER,
      numberFormat: NumberFormat.CURRENCY
    },
    {
      key: 'target',
      value: 'Target',
      type: CellType.NUMBER,
      numberFormat: NumberFormat.CURRENCY
    },
    {
      key: 'performance',
      value: 'Performance %',
      type: CellType.PERCENTAGE
    }
  ]);

  const performanceData = [
    { employee: 'John Doe', sales: 125000, target: 100000 },
    { employee: 'Jane Smith', sales: 98000, target: 100000 },
    { employee: 'Bob Johnson', sales: 150000, target: 100000 },
    { employee: 'Alice Brown', sales: 85000, target: 100000 }
  ];

  performanceData.forEach((row) => {
    const performance = row.sales / row.target;
    
    worksheet.addRow([
      {
        key: 'employee',
        value: row.employee,
        type: CellType.STRING,
        header: 'Employee'
      },
      {
        key: 'sales',
        value: row.sales,
        type: CellType.NUMBER,
        header: 'Sales',
        numberFormat: NumberFormat.CURRENCY
      },
      {
        key: 'target',
        value: row.target,
        type: CellType.NUMBER,
        header: 'Target',
        numberFormat: NumberFormat.CURRENCY
      },
      {
        key: 'performance',
        value: performance,
        type: CellType.PERCENTAGE,
        header: 'Performance %',
        styles: StyleBuilder.create()
          .conditionalFormat({
            type: 'cellIs',
            operator: 'greaterThan',
            values: [1.1],
            style: StyleBuilder.create()
              .backgroundColor('#90EE90')
              .fontColor('#006400')
              .build()
          })
          .conditionalFormat({
            type: 'cellIs',
            operator: 'lessThan',
            values: [0.9],
            style: StyleBuilder.create()
              .backgroundColor('#FFB6C1')
              .fontColor('#8B0000')
              .build()
          })
          .build()
      }
    ]);
  });

  await builder.generateAndDownload('performance-report');
} 

// basicUsageExample().then(() => {
//   console.log('Excel generado correctamente');
// }).catch(console.error);

createMultipleWorksheetsReport().then(() => {
  console.log('Excel generado correctamente');
}).catch(console.error);

createConditionalFormattingExample().then(() => {
  console.log('Excel generado correctamente');
}).catch(console.error);