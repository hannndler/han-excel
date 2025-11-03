/**
 * Multiple Tables Example
 * 
 * Este ejemplo demuestra cómo crear múltiples tablas en una sola hoja de cálculo
 */

import { ExcelBuilder, CellType, StyleBuilder, BorderStyle } from '../index';

/**
 * Ejemplo: Reporte con múltiples tablas en una sola hoja
 */
export async function createMultipleTablesExample(): Promise<void> {
  const builder = new ExcelBuilder({
    metadata: {
      title: 'Reporte con Múltiples Tablas',
      author: 'Han Excel Builder',
      description: 'Demuestra cómo crear múltiples tablas en una sola hoja'
    }
  });

  const worksheet = builder.addWorksheet('Reporte Completo');

  // ===== PRIMERA TABLA: RESUMEN DE VENTAS =====
  worksheet.addTable({
    name: 'ResumenVentas',
    showBorders: true,
    showStripes: true,
    style: 'TableStyleLight1'
  });

  // Header de la primera tabla
  worksheet.addHeader({
    key: 'header-ventas',
    type: CellType.STRING,
    value: 'RESUMEN DE VENTAS - Q1 2024',
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

  // Subheaders de la primera tabla
  worksheet.addSubHeaders([
    {
      key: 'producto',
      type: CellType.STRING,
      value: 'Producto',
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
      key: 'categoria',
      type: CellType.STRING,
      value: 'Categoría',
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
      key: 'ventas',
      type: CellType.NUMBER,
      value: 'Ventas',
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
      key: 'ingresos',
      type: CellType.CURRENCY,
      value: 'Ingresos',
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

  // Datos de la primera tabla
  const ventasData = [
    { producto: 'Laptop', categoria: 'Electrónicos', ventas: 150, ingresos: 225000 },
    { producto: 'Mouse', categoria: 'Electrónicos', ventas: 300, ingresos: 15000 },
    { producto: 'Escritorio', categoria: 'Muebles', ventas: 50, ingresos: 25000 },
    { producto: 'Silla', categoria: 'Muebles', ventas: 80, ingresos: 32000 }
  ];

  ventasData.forEach((item, index) => {
    worksheet.addRow([
      {
        key: `producto-${index}`,
        type: CellType.STRING,
        value: item.producto,
        header: 'Producto',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `categoria-${index}`,
        type: CellType.STRING,
        value: item.categoria,
        header: 'Categoría',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `ventas-${index}`,
        type: CellType.NUMBER,
        value: item.ventas,
        header: 'Ventas',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `ingresos-${index}`,
        type: CellType.CURRENCY,
        value: item.ingresos,
        header: 'Ingresos',
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

  // Footer de la primera tabla
  worksheet.addFooter([
    {
      key: 'total-label',
      type: CellType.STRING,
      value: 'TOTAL',
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
      key: 'total-ventas',
      type: CellType.NUMBER,
      value: ventasData.reduce((sum, item) => sum + item.ventas, 0),
      header: 'Total Ventas',
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
      key: 'total-ingresos',
      type: CellType.CURRENCY,
      value: ventasData.reduce((sum, item) => sum + item.ingresos, 0),
      header: 'Total Ingresos',
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

  // Finalizar la primera tabla
  worksheet.finalizeTable();

  // ===== SEGUNDA TABLA: RENDIMIENTO POR MES =====
  worksheet.addTable({
    name: 'RendimientoMensual',
    showBorders: true,
    showStripes: true,
    style: 'TableStyleMedium1'
  });

  // Header de la segunda tabla
  worksheet.addHeader({
    key: 'header-rendimiento',
    type: CellType.STRING,
    value: 'RENDIMIENTO MENSUAL - Q1 2024',
    mergeCell: true,
    styles: new StyleBuilder()
      .fontName('Arial')
      .fontSize(16)
      .fontBold()
      .backgroundColor('#70AD47')
      .fontColor('#FFFFFF')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  });

  // Subheaders de la segunda tabla
  worksheet.addSubHeaders([
    {
      key: 'mes',
      type: CellType.STRING,
      value: 'Mes',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'objetivo',
      type: CellType.CURRENCY,
      value: 'Objetivo',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'real',
      type: CellType.CURRENCY,
      value: 'Real',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'variacion',
      type: CellType.PERCENTAGE,
      value: 'Variación %',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Datos de la segunda tabla
  const rendimientoData = [
    { mes: 'Enero', objetivo: 100000, real: 95000, variacion: -5 },
    { mes: 'Febrero', objetivo: 110000, real: 115000, variacion: 4.5 },
    { mes: 'Marzo', objetivo: 120000, real: 125000, variacion: 4.2 }
  ];

  rendimientoData.forEach((item, index) => {
    worksheet.addRow([
      {
        key: `mes-${index}`,
        type: CellType.STRING,
        value: item.mes,
        header: 'Mes',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#E2EFDA' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `objetivo-${index}`,
        type: CellType.CURRENCY,
        value: item.objetivo,
        header: 'Objetivo',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#E2EFDA' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `real-${index}`,
        type: CellType.CURRENCY,
        value: item.real,
        header: 'Real',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#E2EFDA' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `variacion-${index}`,
        type: CellType.PERCENTAGE,
        value: item.variacion / 100,
        header: 'Variación %',
        numberFormat: '0.0%',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .fontColor(item.variacion >= 0 ? '#008000' : '#FF0000')
          .backgroundColor(index % 2 === 0 ? '#E2EFDA' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      }
    ]);
  });

  // Footer de la segunda tabla
  worksheet.addFooter([
    {
      key: 'promedio-label',
      type: CellType.STRING,
      value: 'PROMEDIO',
      header: 'Promedio',
      mergeCell: true,
      mergeTo: 2,
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#70AD47')
        .fontColor('#FFFFFF')
        .rightAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'promedio-objetivo',
      type: CellType.CURRENCY,
      value: rendimientoData.reduce((sum, item) => sum + item.objetivo, 0) / rendimientoData.length,
      header: 'Promedio Objetivo',
      numberFormat: '$#,##0',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#70AD47')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'promedio-real',
      type: CellType.CURRENCY,
      value: rendimientoData.reduce((sum, item) => sum + item.real, 0) / rendimientoData.length,
      header: 'Promedio Real',
      numberFormat: '$#,##0',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#70AD47')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'promedio-variacion',
      type: CellType.PERCENTAGE,
      value: rendimientoData.reduce((sum, item) => sum + item.variacion, 0) / rendimientoData.length / 100,
      header: 'Promedio Variación',
      numberFormat: '0.0%',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#70AD47')
        .fontColor('#FFFFFF')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Finalizar la segunda tabla
  worksheet.finalizeTable();

  // ===== TERCERA TABLA: TOP EMPLEADOS =====
  worksheet.addTable({
    name: 'TopEmpleados',
    showBorders: true,
    showStripes: true,
    style: 'TableStyleDark1'
  });

  // Header de la tercera tabla
  worksheet.addHeader({
    key: 'header-empleados',
    type: CellType.STRING,
    value: 'TOP EMPLEADOS - Q1 2024',
    mergeCell: true,
    styles: new StyleBuilder()
      .fontName('Arial')
      .fontSize(16)
      .fontBold()
      .backgroundColor('#FFC000')
      .fontColor('#000000')
      .centerAlign()
      .border(BorderStyle.THIN, '#8EAADB')
      .build()
  });

  // Subheaders de la tercera tabla
  worksheet.addSubHeaders([
    {
      key: 'empleado',
      type: CellType.STRING,
      value: 'Empleado',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#FFEB9C')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'departamento',
      type: CellType.STRING,
      value: 'Departamento',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#FFEB9C')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'ventas-empleado',
      type: CellType.CURRENCY,
      value: 'Ventas',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#FFEB9C')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'comision',
      type: CellType.CURRENCY,
      value: 'Comisión',
      styles: new StyleBuilder()
        .fontName('Arial')
        .fontSize(12)
        .fontBold()
        .backgroundColor('#FFEB9C')
        .fontColor('#000000')
        .centerAlign()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Datos de la tercera tabla
  const empleadosData = [
    { empleado: 'Juan Pérez', departamento: 'Ventas', ventas: 150000, comision: 15000 },
    { empleado: 'María García', departamento: 'Ventas', ventas: 140000, comision: 14000 },
    { empleado: 'Carlos López', departamento: 'Marketing', ventas: 120000, comision: 12000 },
    { empleado: 'Ana Martínez', departamento: 'Ventas', ventas: 110000, comision: 11000 }
  ];

  empleadosData.forEach((item, index) => {
    worksheet.addRow([
      {
        key: `empleado-${index}`,
        type: CellType.STRING,
        value: item.empleado,
        header: 'Empleado',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#FFF2CC' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `departamento-${index}`,
        type: CellType.STRING,
        value: item.departamento,
        header: 'Departamento',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#FFF2CC' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `ventas-empleado-${index}`,
        type: CellType.CURRENCY,
        value: item.ventas,
        header: 'Ventas',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#FFF2CC' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      },
      {
        key: `comision-${index}`,
        type: CellType.CURRENCY,
        value: item.comision,
        header: 'Comisión',
        numberFormat: '$#,##0',
        styles: new StyleBuilder()
          .fontName('Arial')
          .fontSize(11)
          .backgroundColor(index % 2 === 0 ? '#FFF2CC' : '#FFFFFF')
          .border(BorderStyle.THIN, '#8EAADB')
          .build()
      }
    ]);
  });

  // Finalizar la tercera tabla
  worksheet.finalizeTable();

  // Generar y descargar el archivo
  const result = await builder.generateAndDownload('multiple-tables-example.xlsx');
  
  if (result.success) {
    console.log('✅ Ejemplo de múltiples tablas completado exitosamente!');
    console.log('📊 Se crearon 3 tablas en una sola hoja:');
    console.log('   - Resumen de Ventas');
    console.log('   - Rendimiento Mensual');
    console.log('   - Top Empleados');
  } else {
    console.error('❌ Error al generar el ejemplo:', result.error);
  }
}

// Ejecutar el ejemplo
createMultipleTablesExample().then(() => {
  console.log('Excel con múltiples tablas generado correctamente');
}).catch(console.error);
