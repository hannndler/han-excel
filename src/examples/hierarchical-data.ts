/**
 * Hierarchical Data Example
 * 
 * Demonstrates how to work with childrens in complex nested data structures
 */

import { ExcelBuilder, CellType, NumberFormat, StyleBuilder, BorderStyle } from '../index';

/**
 * Example: Company organizational structure with budgets
 */
export async function createHierarchicalCompanyReport(): Promise<void> {
  const builder = new ExcelBuilder({
    metadata: {
      title: 'Company Organizational Report',
      author: 'Han Excel Builder'
    }
  });

  const worksheet = builder.addWorksheet('Organization');

  // Headers
  worksheet.addHeader({
    key: 'title',
    value: 'Company Organizational Structure & Budget Allocation',
    type: CellType.STRING,
    mergeCell: true,
    styles: StyleBuilder.create()
      .fontBold()
      .fontSize(18)
      .centerAlign()
      .backgroundColor('#2E75B6')
      .fontColor('#FFFFFF')
      .build()
  });

  // Add headers
  worksheet.addHeader({
    key: 'company',
    value: 'Company',
    type: CellType.STRING,
    styles: new StyleBuilder()
      .fontBold()
      .backgroundColor('#4472C4')
      .fontColor('#FFFFFF')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  });

  worksheet.addSubHeaders([
    {
      key: 'department',
      value: 'Department',
      type: CellType.STRING,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'team',
      value: 'Team',
      type: CellType.STRING,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'budget',
      value: 'Budget',
      type: CellType.NUMBER,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'employees',
      value: 'Employees',
      type: CellType.NUMBER,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Complex hierarchical data structure
  const companyData = [
    {
      company: 'TechCorp Inc.',
      departments: [
        {
          name: 'Engineering',
          budget: 800000,
          employees: 45,
          teams: [
            { name: 'Frontend', budget: 300000, employees: 15 },
            { name: 'Backend', budget: 350000, employees: 18 },
            { name: 'DevOps', budget: 150000, employees: 12 }
          ]
        },
        {
          name: 'Sales',
          budget: 500000,
          employees: 25,
          teams: [
            { name: 'Enterprise', budget: 300000, employees: 12 },
            { name: 'SMB', budget: 200000, employees: 13 }
          ]
        },
        {
          name: 'Marketing',
          budget: 300000,
          employees: 15,
          teams: [
            { name: 'Digital', budget: 180000, employees: 8 },
            { name: 'Content', budget: 120000, employees: 7 }
          ]
        }
      ]
    }
  ];

  // Build hierarchical rows with childrens
  companyData.forEach((company) => {
    const companyRow = [
      // Company cell with childrens
      {
        key: 'company',
        value: company.company,
        type: CellType.STRING,
        header: 'Company',
        mergeCell: true,
        rowHeight: 25,
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#8EAADB')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: company.departments.flatMap(dept => 
          dept.teams.map(() => ({
            key: 'company',
            value: company.company,
            type: CellType.STRING,
            header: 'Company',
            styles: new StyleBuilder()
              .backgroundColor('#E7E6E6')
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        )
      },
      // Department cell with childrens
      {
        key: 'department',
        value: '', // Will be filled by childrens
        type: CellType.STRING,
        header: 'Department',
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#70AD47')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: company.departments.flatMap(dept => 
          dept.teams.map(() => ({
            key: 'department',
            value: dept.name,
            type: CellType.STRING,
            header: 'Department',
            styles: new StyleBuilder()
              .backgroundColor('#C6EFCE')
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        )
      },
      // Team cell
      {
        key: 'team',
        value: '', // Will be filled by childrens
        type: CellType.STRING,
        header: 'Team',
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#FFC000')
          .fontColor('#000000')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: company.departments.flatMap(dept => 
          dept.teams.map(() => ({
            key: 'team',
            value: 'Team A',
            type: CellType.STRING,
            header: 'Team',
            styles: new StyleBuilder()
              .backgroundColor('#FFEB9C')
              .border(BorderStyle.THIN, '#8EAADB')
              .centerAlign()
              .build()
          }))
        )
      },
      // Budget cell with childrens
      {
        key: 'budget',
        value: company.departments.reduce((sum, dept) => sum + dept.budget, 0),
        type: CellType.NUMBER,
        header: 'Budget',
        numberFormat: NumberFormat.CURRENCY,
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#8EAADB')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: company.departments.flatMap(dept => 
          dept.teams.map(() => ({
            key: 'budget',
            value: 50000,
            type: CellType.NUMBER,
            header: 'Budget',
            numberFormat: NumberFormat.CURRENCY,
            styles: new StyleBuilder()
              .backgroundColor('#E7E6E6')
              .rightAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        )
      },
      // Employees cell with childrens
      {
        key: 'employees',
        value: company.departments.reduce((sum, dept) => sum + dept.employees, 0),
        type: CellType.NUMBER,
        header: 'Employees',
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#8EAADB')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: company.departments.flatMap(dept => 
          dept.teams.map(() => ({
            key: 'employees',
            value: 5,
            type: CellType.NUMBER,
            header: 'Employees',
            styles: new StyleBuilder()
              .backgroundColor('#E7E6E6')
              .centerAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        )
      }
    ];

    worksheet.addRow(companyRow);
  });

  // Add totals footer
  const totalBudget = companyData[0]?.departments.reduce((sum, dept) => sum + dept.budget, 0) || 0;
  const totalEmployees = companyData[0]?.departments.reduce((sum, dept) => sum + dept.employees, 0) || 0;

  worksheet.addFooter([
    {
      key: 'totalLabel',
      value: 'TOTAL',
      type: CellType.STRING,
      header: 'Total',
      mergeCell: true,
      mergeTo: 2,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'totalBudget',
      value: totalBudget,
      type: CellType.NUMBER,
      header: 'Total Budget',
      numberFormat: NumberFormat.CURRENCY,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'totalEmployees',
      value: totalEmployees,
      type: CellType.NUMBER,
      header: 'Total Employees',
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  await builder.generateAndDownload('company-hierarchical-report');
}

/**
 * Example: Product catalog with categories and subcategories
 */
export async function createProductCatalogReport(): Promise<void> {
  const builder = new ExcelBuilder();
  const worksheet = builder.addWorksheet('Product Catalog');

  // Headers
  worksheet.addHeader({
    key: 'title',
    value: 'Product Catalog by Category',
    type: CellType.STRING,
    mergeCell: true,
    styles: StyleBuilder.create()
      .fontBold()
      .fontSize(16)
      .centerAlign()
      .backgroundColor('#4472C4')
      .fontColor('#FFFFFF')
      .build()
  });

  // Add headers for the second example
  worksheet.addHeader({
    key: 'category',
    value: 'Category',
    type: CellType.STRING,
    styles: new StyleBuilder()
      .fontBold()
      .backgroundColor('#4472C4')
      .fontColor('#FFFFFF')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  });

  worksheet.addSubHeaders([
    {
      key: 'department',
      value: 'Department',
      type: CellType.STRING,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'team',
      value: 'Team',
      type: CellType.STRING,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'budget',
      value: 'Budget',
      type: CellType.NUMBER,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'employees',
      value: 'Employees',
      type: CellType.NUMBER,
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#4472C4')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Product data with nested structure
  const productData = [
    {
      category: 'Electronics',
      subcategories: [
        {
          name: 'Computers',
          products: [
            { name: 'Laptop Pro', price: 1299.99, stock: 25 },
            { name: 'Desktop Gaming', price: 899.99, stock: 15 },
            { name: 'Tablet Air', price: 599.99, stock: 30 }
          ]
        },
        {
          name: 'Phones',
          products: [
            { name: 'Smartphone X', price: 799.99, stock: 40 },
            { name: 'Phone Mini', price: 499.99, stock: 35 }
          ]
        }
      ]
    },
    {
      category: 'Clothing',
      subcategories: [
        {
          name: 'Men',
          products: [
            { name: 'Casual Shirt', price: 29.99, stock: 100 },
            { name: 'Jeans Classic', price: 49.99, stock: 75 }
          ]
        },
        {
          name: 'Women',
          products: [
            { name: 'Summer Dress', price: 39.99, stock: 60 },
            { name: 'Blouse Elegant', price: 34.99, stock: 45 }
          ]
        }
      ]
    }
  ];

  // Add category rows with hierarchical data
  productData.forEach((category) => {
    const categoryRow = [
      {
        key: 'category',
        value: category.category,
        type: CellType.STRING,
        header: 'Category',
        mergeCell: true,
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#70AD47')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: category.subcategories.map(sub => ({
          key: 'subcategory',
          value: sub.name,
          type: CellType.STRING,
          header: 'Subcategory',
          styles: new StyleBuilder()
            .fontBold()
            .backgroundColor('#A9D08E')
            .fontColor('#000000')
            .centerAlign()
            .border(BorderStyle.THIN, '#8EAADB')
            .build(),
          childrens: sub.products.map(() => ({
            key: 'product',
            value: 'Product A',
            type: CellType.STRING,
            header: 'Product',
            styles: new StyleBuilder()
              .backgroundColor('#E2EFDA')
              .centerAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        }))
      },
      {
        key: 'sales',
        value: category.subcategories.reduce((sum, sub) => sum + (sub as any).sales || 0, 0),
        type: CellType.NUMBER,
        header: 'Sales',
        numberFormat: NumberFormat.CURRENCY,
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#70AD47')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: category.subcategories.map(sub => ({
          key: 'sub-sales',
          value: (sub as any).sales || 0,
          type: CellType.NUMBER,
          header: 'Sales',
          numberFormat: NumberFormat.CURRENCY,
          styles: new StyleBuilder()
            .fontBold()
            .backgroundColor('#A9D08E')
            .fontColor('#000000')
            .centerAlign()
            .border(BorderStyle.THIN, '#8EAADB')
            .build(),
          childrens: sub.products.map(() => ({
            key: 'product-sales',
            value: 5000,
            type: CellType.NUMBER,
            header: 'Sales',
            numberFormat: NumberFormat.CURRENCY,
            styles: new StyleBuilder()
              .backgroundColor('#E2EFDA')
              .centerAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        }))
      },
      {
        key: 'revenue',
        value: category.subcategories.reduce((sum, sub) => sum + (sub as any).revenue || 0, 0),
        type: CellType.NUMBER,
        header: 'Revenue',
        numberFormat: NumberFormat.CURRENCY,
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#70AD47')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: category.subcategories.map(sub => ({
          key: 'sub-revenue',
          value: (sub as any).revenue || 0,
          type: CellType.NUMBER,
          header: 'Revenue',
          numberFormat: NumberFormat.CURRENCY,
          styles: new StyleBuilder()
            .fontBold()
            .backgroundColor('#A9D08E')
            .fontColor('#000000')
            .centerAlign()
            .border(BorderStyle.THIN, '#8EAADB')
            .build(),
          childrens: sub.products.map(() => ({
            key: 'product-revenue',
            value: 4500,
            type: CellType.NUMBER,
            header: 'Revenue',
            numberFormat: NumberFormat.CURRENCY,
            styles: new StyleBuilder()
              .backgroundColor('#E2EFDA')
              .centerAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        }))
      },
      {
        key: 'margin',
        value: category.subcategories.reduce((sum, sub) => sum + (sub as any).margin || 0, 0) / category.subcategories.length,
        type: CellType.PERCENTAGE,
        header: 'Margin %',
        numberFormat: '0.0%',
        styles: new StyleBuilder()
          .fontBold()
          .backgroundColor('#70AD47')
          .fontColor('#FFFFFF')
          .centerAlign()
          .border(BorderStyle.THIN, '#8EAADB')
          .build(),
        childrens: category.subcategories.map(sub => ({
          key: 'sub-margin',
          value: ((sub as any).margin || 0) / 100,
          type: CellType.PERCENTAGE,
          header: 'Margin %',
          numberFormat: '0.0%',
          styles: new StyleBuilder()
            .fontBold()
            .backgroundColor('#A9D08E')
            .fontColor('#000000')
            .centerAlign()
            .border(BorderStyle.THIN, '#8EAADB')
            .build(),
          childrens: sub.products.map(() => ({
            key: 'product-margin',
            value: 0.15,
            type: CellType.PERCENTAGE,
            header: 'Margin %',
            numberFormat: '0.0%',
            styles: new StyleBuilder()
              .backgroundColor('#E2EFDA')
              .centerAlign()
              .border(BorderStyle.THIN, '#8EAADB')
              .build()
          }))
        }))
      }
    ];

    worksheet.addRow(categoryRow);
  });

  await builder.generateAndDownload('product-catalog-report');
} 