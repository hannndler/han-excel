/**
 * Ejemplo simple de múltiples tablas
 * 
 * Este ejemplo demuestra cómo usar las nuevas funcionalidades de tablas múltiples
 */

import { ExcelBuilder, CellType, StyleBuilder, BorderStyle } from '../index';

/**
 * Ejemplo básico de múltiples tablas
 */
export async function createSimpleMultipleTablesExample(): Promise<void> {
  const builder = new ExcelBuilder({
    metadata: {
      title: 'Ejemplo Múltiples Tablas',
      author: 'Han Excel Builder'
    }
  });

  const worksheet = builder.addWorksheet('Ejemplo Tablas');

  // ===== PRIMERA TABLA =====
  worksheet.addTable({
    name: 'Tabla1',
    showBorders: true,
    showStripes: true
  });

  // Agregar contenido a la primera tabla
  worksheet.addHeader({
    key: 'header1',
    type: CellType.STRING,
    value: 'PRIMERA TABLA - Datos de Ventas',
    mergeCell: true,
    styles: new StyleBuilder()
      .fontBold()
      .fontSize(14)
      .backgroundColor('#4472C4')
      .fontColor('#FFFFFF')
      .centerAlign()
      .build()
  });

  worksheet.addSubHeaders([
    {
      key: 'producto',
      type: CellType.STRING,
      value: 'Producto',
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .build()
    },
    {
      key: 'precio',
      type: CellType.NUMBER,
      value: 'Precio',
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#8EAADB')
        .fontColor('#FFFFFF')
        .centerAlign()
        .build()
    }
  ]);

  // Datos de la primera tabla
  worksheet.addRow([
    {
      key: 'laptop',
      type: CellType.STRING,
      value: 'Laptop',
      header: 'Producto',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'precio-laptop',
      type: CellType.NUMBER,
      value: 1000,
      header: 'Precio',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  worksheet.addRow([
    {
      key: 'mouse',
      type: CellType.STRING,
      value: 'Mouse',
      header: 'Producto',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'precio-mouse',
      type: CellType.NUMBER,
      value: 25,
      header: 'Precio',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Finalizar la primera tabla
  worksheet.finalizeTable();

  // ===== SEGUNDA TABLA =====
  worksheet.addTable({
    name: 'Tabla2',
    showBorders: true,
    showStripes: true
  });

  // Agregar contenido a la segunda tabla
  worksheet.addHeader({
    key: 'header2',
    type: CellType.STRING,
    value: 'SEGUNDA TABLA - Datos de Empleados',
    mergeCell: true,
    styles: new StyleBuilder()
      .fontBold()
      .fontSize(14)
      .backgroundColor('#70AD47')
      .fontColor('#FFFFFF')
      .centerAlign()
      .build()
  });

  worksheet.addSubHeaders([
    {
      key: 'nombre',
      type: CellType.STRING,
      value: 'Nombre',
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .build()
    },
    {
      key: 'edad',
      type: CellType.NUMBER,
      value: 'Edad',
      styles: new StyleBuilder()
        .fontBold()
        .backgroundColor('#A9D08E')
        .fontColor('#000000')
        .centerAlign()
        .build()
    }
  ]);

  // Datos de la segunda tabla
  worksheet.addRow([
    {
      key: 'juan',
      type: CellType.STRING,
      value: 'Juan Pérez',
      header: 'Nombre',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'edad-juan',
      type: CellType.NUMBER,
      value: 30,
      header: 'Edad',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  worksheet.addRow([
    {
      key: 'maria',
      type: CellType.STRING,
      value: 'María García',
      header: 'Nombre',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    },
    {
      key: 'edad-maria',
      type: CellType.NUMBER,
      value: 28,
      header: 'Edad',
      styles: new StyleBuilder()
        .border(BorderStyle.THIN, '#8EAADB')
        .build()
    }
  ]);

  // Finalizar la segunda tabla
  worksheet.finalizeTable();

  // Generar el archivo
  const result = await builder.generateAndDownload('simple-multiple-tables.xlsx');
  
  if (result.success) {
    console.log('✅ Ejemplo simple de múltiples tablas completado exitosamente!');
    console.log('📊 Se crearon 2 tablas en una sola hoja');
  } else {
    console.error('❌ Error al generar el ejemplo:', result.error);
  }
}

// Ejecutar el ejemplo
createSimpleMultipleTablesExample().then(() => {
  console.log('Excel con múltiples tablas generado correctamente');
}).catch(console.error);

