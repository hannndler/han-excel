/**
 * Test simple para verificar que han-excel-builder funciona
 */

import { ExcelBuilder, CellType, NumberFormat, StyleBuilder } from './src/index';

async function testBasicFunctionality() {
  console.log('🧪 Iniciando test de han-excel-builder...');
  
  try {
    // 1. Crear el builder
    console.log('📝 Creando ExcelBuilder...');
    const builder = new ExcelBuilder({
      metadata: {
        title: 'Test Report',
        author: 'Test User',
        description: 'Test de funcionalidad básica'
      }
    });

    // 2. Agregar worksheet
    console.log('📊 Agregando worksheet...');
    const worksheet = builder.addWorksheet('Test Data');

    // 3. Agregar header
    console.log('📋 Agregando header...');
    worksheet.addHeader({
      key: 'title',
      value: 'Test Report - Funcionalidad Básica',
      type: CellType.STRING,
      mergeCell: true,
      styles: StyleBuilder.create()
        .fontBold()
        .fontSize(16)
        .centerAlign()
        .backgroundColor('#2563EB')
        .fontColor('#FFFFFF')
        .build()
    });

    // 4. Agregar sub-headers
    console.log('📋 Agregando sub-headers...');
    worksheet.addSubHeaders([
      {
        key: 'name',
        value: 'Nombre',
        type: CellType.STRING,
        styles: StyleBuilder.create().fontBold().backgroundColor('#F3F4F6').build()
      },
      {
        key: 'value',
        value: 'Valor',
        type: CellType.NUMBER,
        styles: StyleBuilder.create().fontBold().backgroundColor('#F3F4F6').build()
      },
      {
        key: 'date',
        value: 'Fecha',
        type: CellType.DATE,
        styles: StyleBuilder.create().fontBold().backgroundColor('#F3F4F6').build()
      }
    ]);

    // 5. Agregar datos de prueba
    console.log('📊 Agregando datos de prueba...');
    const testData = [
      { name: 'Producto A', value: 1500.50, date: new Date('2024-01-15') },
      { name: 'Producto B', value: 2200.75, date: new Date('2024-01-16') },
      { name: 'Producto C', value: 1800.25, date: new Date('2024-01-17') }
    ];

    testData.forEach((row, index) => {
      worksheet.addRow([
        { key: `name-${index}`, value: row.name, type: CellType.STRING, header: 'Nombre' },
        { key: `value-${index}`, value: row.value, type: CellType.NUMBER, header: 'Valor' },
        { key: `date-${index}`, value: row.date, type: CellType.DATE, header: 'Fecha' }
      ]);
    });

    // 6. Validar el workbook
    console.log('✅ Validando workbook...');
    const validation = builder.validate();
    if (!validation.success) {
      throw new Error(`Validación falló: ${validation.error?.message}`);
    }

    // 7. Obtener estadísticas
    console.log('📈 Obteniendo estadísticas...');
    const stats = builder.getStats();
    console.log('Estadísticas:', {
      worksheets: stats.totalWorksheets,
      cells: stats.totalCells,
      memoryUsage: stats.memoryUsage,
      buildTime: stats.buildTime
    });

    // 8. Generar buffer
    console.log('🔨 Generando buffer...');
    const result = await builder.toBuffer();
    
    if (!result.success) {
      throw new Error(`Error al generar buffer: ${result.error?.message}`);
    }

    console.log('✅ Buffer generado exitosamente!');
    console.log(`📏 Tamaño del archivo: ${result.data.byteLength} bytes`);

    // 9. Generar y descargar (solo en navegador)
    if (typeof window !== 'undefined') {
      console.log('💾 Descargando archivo...');
      const downloadResult = await builder.generateAndDownload('test-report.xlsx');
      
      if (downloadResult.success) {
        console.log('✅ Archivo descargado exitosamente!');
      } else {
        console.log('⚠️ Error al descargar:', downloadResult.error?.message);
      }
    } else {
      console.log('🌐 Ejecutando en Node.js - omitiendo descarga automática');
    }

    console.log('🎉 ¡Test completado exitosamente!');
    return true;

  } catch (error) {
    console.error('❌ Error en el test:', error);
    return false;
  }
}

// Ejecutar el test
testBasicFunctionality().then(success => {
  if (success) {
    console.log('✅ Todos los tests pasaron correctamente');
    process.exit(0);
  } else {
    console.log('❌ Algunos tests fallaron');
    process.exit(1);
  }
}).catch(error => {
  console.error('💥 Error fatal:', error);
  process.exit(1);
}); 