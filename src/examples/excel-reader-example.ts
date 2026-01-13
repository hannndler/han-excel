/**
 * Excel Reader Example
 * 
 * Este ejemplo demuestra cómo leer archivos Excel y convertirlos a JSON
 */

import { ExcelReader } from '../core/ExcelReader';
import { OutputFormat } from '../types/reader.types';
import type { IJsonWorkbook, IDetailedFormat, IFlatFormat, IFlatFormatMultiSheet } from '../types/reader.types';

/**
 * Ejemplo básico: Leer Excel desde un ArrayBuffer
 */
export async function readExcelFromBufferExample() {
  // Simular un ArrayBuffer (en la práctica vendría de una petición HTTP, FileReader, etc.)
  // const response = await fetch('archivo.xlsx');
  // const buffer = await response.arrayBuffer();
  
  // const result = await ExcelReader.fromBuffer(buffer, {
  //   useFirstRowAsHeaders: true,
  //   includeEmptyRows: false
  // });

  // if (result.success) {
  //   console.log('Workbook:', result.data);
  //   console.log('Total sheets:', result.data.totalSheets);
  //   result.data.sheets.forEach(sheet => {
  //     console.log(`Sheet: ${sheet.name}`, sheet.rows);
  //   });
  // } else {
  //   console.error('Error:', result.error);
  // }
}

/**
 * Ejemplo: Leer Excel desde un File (navegador)
 */
export async function readExcelFromFileExample(file: File) {
  const result = await ExcelReader.fromFile(file, {
    useFirstRowAsHeaders: true,
    includeFormatting: true,
    datesAsISO: true
  });

  if (result.success) {
    const workbook = result.data;
    
    // Acceder a la primera hoja
    const firstSheet = workbook.sheets[0];
    
    if (firstSheet) {
      // Si usamos headers, cada fila tiene un objeto 'data'
      firstSheet.rows.forEach(row => {
        if (row.data) {
          console.log('Row data:', row.data);
        } else {
          // Si no usamos headers, acceder a las celdas directamente
          console.log('Row cells:', row.cells);
        }
      });
    }
  } else {
    console.error('Error reading file:', result.error);
  }
}

/**
 * Ejemplo: Leer una hoja específica
 */
export async function readSpecificSheetExample(buffer: ArrayBuffer) {
  const result = await ExcelReader.fromBuffer(buffer, {
    sheetName: 'Ventas', // Nombre de la hoja
    useFirstRowAsHeaders: true,
    startRow: 2, // Empezar desde la fila 2 (saltar header)
    endRow: 100 // Leer hasta la fila 100
  });

  if (result.success) {
    const sheet = result.data.sheets[0];
    if (sheet) {
      console.log(`Sheet: ${sheet.name}`);
      console.log(`Total rows: ${sheet.totalRows}`);
      console.log(`Total columns: ${sheet.totalColumns}`);
      console.log(`Headers: ${sheet.headers?.join(', ')}`);
    }
  }
}

/**
 * Ejemplo: Leer con headers personalizados
 */
export async function readWithCustomHeadersExample(buffer: ArrayBuffer) {
  const result = await ExcelReader.fromBuffer(buffer, {
    useFirstRowAsHeaders: false,
    headers: ['Producto', 'Cantidad', 'Precio', 'Total'], // Headers personalizados
    startRow: 1,
    includeFormulas: true
  });

  if (result.success) {
    const sheet = result.data.sheets[0];
    if (sheet) {
      sheet.rows.forEach((row, index) => {
        console.log(`Row ${index + 1}:`, row.cells.map(cell => cell.value));
      });
    }
  }
}

/**
 * Ejemplo: Leer Excel desde ruta (Node.js)
 */
export async function readExcelFromPathExample() {
  // Solo funciona en Node.js
  // const result = await ExcelReader.fromPath('./archivo.xlsx', {
  //   useFirstRowAsHeaders: true,
  //   includeFormatting: false
  // });

  // if (result.success) {
  //   console.log('Processing time:', result.processingTime, 'ms');
  //   console.log('Workbook:', JSON.stringify(result.data, null, 2));
  // }
}

/**
 * Ejemplo completo: Procesar Excel y convertir a formato específico
 */
export async function processExcelToCustomFormat(file: File) {
  const result = await ExcelReader.fromFile(file, {
    useFirstRowAsHeaders: true,
    datesAsISO: true
  });

  if (!result.success) {
    throw new Error(result.error.message);
  }

  const workbook = result.data;
  
  // Convertir a formato personalizado
  const customFormat = workbook.sheets.map(sheet => ({
    name: sheet.name,
    data: sheet.rows.map(row => row.data || {})
  }));

  return customFormat;
}

/**
 * Ejemplo: Usar mapper para transformar datos
 */
export async function readExcelWithMapper(file: File) {
  // Ejemplo 1: Mapper con formato WORKSHEET
  const result1 = await ExcelReader.fromFile(file, {
    outputFormat: OutputFormat.WORKSHEET,
    useFirstRowAsHeaders: true,
    mapper: (data: IJsonWorkbook) => {
      // Transformar cada hoja
      return {
        totalSheets: data.totalSheets,
        sheets: data.sheets.map(sheet => ({
          name: sheet.name,
          // Convertir filas a objetos con validación
          items: sheet.rows
            .filter(row => row.data) // Solo filas con datos
            .map(row => ({
              ...(row.data || {}),
              // Agregar campos calculados
              isValid: row.data ? Object.values(row.data).every(
                val => val !== null && val !== undefined && val !== ''
              ) : false
            }))
        }))
      };
    }
  });

  // Ejemplo 2: Mapper con formato FLAT
  const result2 = await ExcelReader.fromFile(file, {
    outputFormat: OutputFormat.FLAT,
    useFirstRowAsHeaders: true,
    mapper: (data: IFlatFormat | IFlatFormatMultiSheet) => {
      if ('data' in data && Array.isArray(data.data)) {
        // Transformar cada fila
        return data.data.map((row: Record<string, unknown> | unknown[]) => {
          if (Array.isArray(row)) {
            return row;
          }
          // Normalizar nombres de campos
          const normalized: Record<string, unknown> = {};
          Object.keys(row).forEach(key => {
            const normalizedKey = key.toLowerCase().replace(/\s+/g, '_');
            normalized[normalizedKey] = row[key];
          });
          return normalized;
        });
      }
      return data;
    }
  });

  // Ejemplo 3: Mapper con formato DETAILED
  const result3 = await ExcelReader.fromFile(file, {
    outputFormat: OutputFormat.DETAILED,
    mapper: (data: IDetailedFormat) => {
      // Agrupar celdas por hoja y crear estructura personalizada
      const grouped: Record<string, Array<{ ref: string; value: unknown }>> = {};
      
      data.cells.forEach(cell => {
        if (!grouped[cell.sheet]) {
          grouped[cell.sheet] = [];
        }
        const sheetGroup = grouped[cell.sheet];
        if (sheetGroup) {
          sheetGroup.push({
            ref: cell.reference,
            value: cell.value
          });
        }
      });

      return {
        sheets: Object.keys(grouped).map(sheetName => ({
          name: sheetName,
          cells: grouped[sheetName] || []
        }))
      };
    }
  });

  return { result1, result2, result3 };
}

