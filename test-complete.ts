/**
 * Test completo para verificar que han-excel-builder funciona y genera archivos
 */

import { ExcelBuilder, CellType, NumberFormat, StyleBuilder } from './src/index';
import * as fs from 'fs';
import * as path from 'path';

async function testCompleteFunctionality() {
  console.log('🧪 Iniciando test completo de han-excel-builder...');
  
  try {
    // 1. Crear el builder
    console.log('📝 Creando ExcelBuilder...');
    const builder = new ExcelBuilder({
      metadata: {
        title: 'Test Report Completo',
        author: 'Test User',
        description: 'Test completo de funcionalidad',
        keywords: 'test, excel, report',
        category: 'Test Reports'
      }
    });

    // 2. Agregar worksheet principal
    console.log('📊 Agregando worksheet principal...');
    const worksheet = builder.addWorksheet('Datos de Ventas');

    // 3. Agregar header principal
    console.log('📋 Agregando header principal...');
    worksheet.addHeader({
      key: 'main-title',
      value: 'Reporte de Ventas - Q1 2024',
      type: CellType.STRING,
      mergeCell: true,
      styles: StyleBuilder.create()
        .fontBold()
        .fontSize(18)
        .centerAlign()
        .backgroundColor('#2563EB')
        .fontColor('#FFFFFF')
        .build()
    });

    // 4. Agregar sub-headers
    console.log('📋 Agregando sub-headers...');
    worksheet.addSubHeaders([
      {
        key: 'product',
        value: 'Producto',
        type: CellType.STRING,
        styles: StyleBuilder.create().fontBold().backgroundColor('#4472C4').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'category',
        value: 'Categoría',
        type: CellType.STRING,
        styles: StyleBuilder.create().fontBold().backgroundColor('#4472C4').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'sales',
        value: 'Ventas',
        type: CellType.NUMBER,
        styles: StyleBuilder.create().fontBold().backgroundColor('#4472C4').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'revenue',
        value: 'Ingresos',
        type: CellType.NUMBER,
        styles: StyleBuilder.create().fontBold().backgroundColor('#4472C4').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'date',
        value: 'Fecha',
        type: CellType.DATE,
        styles: StyleBuilder.create().fontBold().backgroundColor('#4472C4').fontColor('#FFFFFF').centerAlign().build()
      }
    ]);

    // 5. Agregar datos de prueba
    console.log('📊 Agregando datos de prueba...');
    const salesData = [
      { product: 'Laptop HP', category: 'Electrónicos', sales: 25, revenue: 37500, date: new Date('2024-01-15') },
      { product: 'Mouse Logitech', category: 'Accesorios', sales: 150, revenue: 7500, date: new Date('2024-01-16') },
      { product: 'Teclado Mecánico', category: 'Accesorios', sales: 80, revenue: 16000, date: new Date('2024-01-17') },
      { product: 'Monitor 27"', category: 'Electrónicos', sales: 12, revenue: 18000, date: new Date('2024-01-18') },
      { product: 'Auriculares', category: 'Audio', sales: 200, revenue: 20000, date: new Date('2024-01-19') },
      { product: 'Webcam HD', category: 'Audio', sales: 75, revenue: 11250, date: new Date('2024-01-20') }
    ];

    salesData.forEach((row, index) => {
      worksheet.addRow([
        { 
          key: `product-${index}`, 
          value: row.product, 
          type: CellType.STRING, 
          header: 'Producto',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
            .build()
        },
        { 
          key: `category-${index}`, 
          value: row.category, 
          type: CellType.STRING, 
          header: 'Categoría',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
            .build()
        },
        { 
          key: `sales-${index}`, 
          value: row.sales, 
          type: CellType.NUMBER, 
          header: 'Ventas',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
            .centerAlign()
            .build()
        },
        { 
          key: `revenue-${index}`, 
          value: row.revenue, 
          type: CellType.NUMBER, 
          header: 'Ingresos',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
            .numberFormat('#,##0')
            .build()
        },
        { 
          key: `date-${index}`, 
          value: row.date, 
          type: CellType.DATE, 
          header: 'Fecha',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F2F2F2' : '#FFFFFF')
            .centerAlign()
            .build()
        }
      ]);
    });

    // 6. Agregar worksheet de resumen
    console.log('📊 Agregando worksheet de resumen...');
    const summaryWorksheet = builder.addWorksheet('Resumen');

    // Header del resumen
    summaryWorksheet.addHeader({
      key: 'summary-title',
      value: 'Resumen de Ventas por Categoría',
      type: CellType.STRING,
      mergeCell: true,
      styles: StyleBuilder.create()
        .fontBold()
        .fontSize(16)
        .centerAlign()
        .backgroundColor('#059669')
        .fontColor('#FFFFFF')
        .build()
    });

    // Sub-headers del resumen
    summaryWorksheet.addSubHeaders([
      {
        key: 'category-summary',
        value: 'Categoría',
        type: CellType.STRING,
        styles: StyleBuilder.create().fontBold().backgroundColor('#10B981').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'total-sales',
        value: 'Total Ventas',
        type: CellType.NUMBER,
        styles: StyleBuilder.create().fontBold().backgroundColor('#10B981').fontColor('#FFFFFF').centerAlign().build()
      },
      {
        key: 'total-revenue',
        value: 'Total Ingresos',
        type: CellType.NUMBER,
        styles: StyleBuilder.create().fontBold().backgroundColor('#10B981').fontColor('#FFFFFF').centerAlign().build()
      }
    ]);

    // Calcular resumen por categoría
    const categorySummary = salesData.reduce((acc, item) => {
      if (!acc[item.category]) {
        acc[item.category] = { sales: 0, revenue: 0 };
      }
      acc[item.category].sales += item.sales;
      acc[item.category].revenue += item.revenue;
      return acc;
    }, {} as Record<string, { sales: number; revenue: number }>);

    // Agregar filas de resumen
    Object.entries(categorySummary).forEach(([category, data], index) => {
      summaryWorksheet.addRow([
        { 
          key: `cat-${index}`, 
          value: category, 
          type: CellType.STRING, 
          header: 'Categoría',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F0FDF4' : '#FFFFFF')
            .fontBold()
            .build()
        },
        { 
          key: `sales-sum-${index}`, 
          value: data.sales, 
          type: CellType.NUMBER, 
          header: 'Total Ventas',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F0FDF4' : '#FFFFFF')
            .centerAlign()
            .build()
        },
        { 
          key: `revenue-sum-${index}`, 
          value: data.revenue, 
          type: CellType.NUMBER, 
          header: 'Total Ingresos',
          styles: StyleBuilder.create()
            .backgroundColor(index % 2 === 0 ? '#F0FDF4' : '#FFFFFF')
            .numberFormat('#,##0')
            .build()
        }
      ]);
    });

    // 7. Validar el workbook
    console.log('✅ Validando workbook...');
    const validation = builder.validate();
    if (!validation.success) {
      throw new Error(`Validación falló: ${validation.error?.message}`);
    }

    // 8. Obtener estadísticas
    console.log('📈 Obteniendo estadísticas...');
    const stats = builder.getStats();
    console.log('Estadísticas del workbook:', {
      worksheets: stats.totalWorksheets,
      cells: stats.totalCells,
      memoryUsage: `${(stats.memoryUsage / 1024).toFixed(2)} KB`,
      buildTime: `${stats.buildTime}ms`,
      stylesUsed: stats.stylesUsed
    });

    // 9. Generar buffer
    console.log('🔨 Generando buffer...');
    const result = await builder.toBuffer();
    
    if (!result.success) {
      throw new Error(`Error al generar buffer: ${result.error?.message}`);
    }

    console.log('✅ Buffer generado exitosamente!');
    console.log(`📏 Tamaño del archivo: ${(result.data.byteLength / 1024).toFixed(2)} KB`);

    // 10. Guardar archivo en disco
    console.log('💾 Guardando archivo en disco...');
    const outputPath = path.join(process.cwd(), 'test-report-complete.xlsx');
    fs.writeFileSync(outputPath, Buffer.from(result.data));
    
    console.log(`✅ Archivo guardado en: ${outputPath}`);

    // 11. Verificar que el archivo existe
    if (fs.existsSync(outputPath)) {
      const fileStats = fs.statSync(outputPath);
      console.log(`✅ Archivo verificado - Tamaño: ${(fileStats.size / 1024).toFixed(2)} KB`);
    } else {
      throw new Error('El archivo no se guardó correctamente');
    }

    console.log('🎉 ¡Test completo exitoso!');
    console.log('📁 Puedes abrir el archivo test-report-complete.xlsx en Excel para verificar el resultado');
    
    return true;

  } catch (error) {
    console.error('❌ Error en el test completo:', error);
    return false;
  }
}

// Ejecutar el test completo
testCompleteFunctionality().then(success => {
  if (success) {
    console.log('✅ Test completo pasó correctamente');
    process.exit(0);
  } else {
    console.log('❌ Test completo falló');
    process.exit(1);
  }
}).catch(error => {
  console.error('💥 Error fatal en test completo:', error);
  process.exit(1);
}); 