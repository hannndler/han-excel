# Han Excel Builder

🚀 **Generador avanzado de archivos Excel con soporte TypeScript, estilos completos y rendimiento optimizado**

Una biblioteca moderna y completamente tipada para crear reportes Excel complejos con múltiples hojas de cálculo, estilos avanzados y alto rendimiento.

## ✨ Características

### 📊 Estructura de Datos
- ✅ **Múltiples Hojas de Cálculo** - Crea workbooks complejos con múltiples hojas
- ✅ **Múltiples Tablas por Hoja** - Crea varias tablas independientes en una sola hoja
- ✅ **Headers Anidados** - Soporte completo para headers con múltiples niveles de anidación
- ✅ **Datos Jerárquicos** - Soporte para datos con estructura de children (datos anidados)

### 📈 Tipos de Datos
- ✅ **STRING** - Valores de texto
- ✅ **NUMBER** - Valores numéricos
- ✅ **BOOLEAN** - Valores verdadero/falso
- ✅ **DATE** - Valores de fecha
- ✅ **PERCENTAGE** - Valores de porcentaje
- ✅ **CURRENCY** - Valores de moneda
- ✅ **LINK** - Hipervínculos con texto personalizable
- ✅ **FORMULA** - Fórmulas de Excel

### 🎨 Estilos Avanzados
- ✅ **API Fluida** - StyleBuilder con métodos encadenables
- ✅ **Fuentes** - Control completo sobre nombre, tamaño, color, negrita, cursiva, subrayado
- ✅ **Colores** - Fondos, colores de texto con soporte para hex, RGB y temas
- ✅ **Bordes** - Bordes personalizables en todos los lados con múltiples estilos
- ✅ **Alineación** - Horizontal (izquierda, centro, derecha, justificar) y vertical (arriba, medio, abajo)
- ✅ **Texto** - Ajuste de texto, contracción para ajustar, rotación de texto
- ✅ **Formatos de Número** - Múltiples formatos predefinidos y personalizados
- ✅ **Filas Alternadas** - Soporte para rayas alternadas en tablas

### 🔧 Funcionalidades Avanzadas
- ✅ **TypeScript First** - Seguridad de tipos completa con interfaces exhaustivas
- ✅ **Sistema de Eventos** - EventEmitter para monitorear el proceso de construcción
- ✅ **Validación** - Sistema robusto de validación de datos
- ✅ **Metadata** - Soporte completo para metadata del workbook (autor, título, descripción, etc.)
- ✅ **Múltiples Formatos de Exportación** - Descarga directa, Buffer, Blob
- ✅ **Lectura de Excel** - Lee archivos Excel y convierte a JSON
- ✅ **Hipervínculos** - Creación de enlaces con texto personalizable
- ✅ **Merge de Celdas** - Fusión horizontal y vertical de celdas
- ✅ **Dimensiones Personalizadas** - Ancho de columnas y alto de filas personalizables

## 📦 Instalación

```bash
npm install han-excel-builder
# o
yarn add han-excel-builder
# o
pnpm add han-excel-builder
```

## 🚀 Inicio Rápido

### Ejemplo Básico

```typescript
import { ExcelBuilder, CellType, NumberFormat, StyleBuilder, BorderStyle } from 'han-excel-builder';

// Crear un reporte simple
const builder = new ExcelBuilder({
  metadata: {
    title: 'Reporte de Ventas',
    author: 'Mi Empresa',
    description: 'Reporte mensual de ventas'
  }
});

const worksheet = builder.addWorksheet('Ventas');

// Agregar header principal
worksheet.addHeader({
  key: 'title',
  value: 'Reporte Mensual de Ventas',
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

// Agregar sub-headers
worksheet.addSubHeaders([
  {
    key: 'producto',
    value: 'Producto',
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
    key: 'ventas',
    value: 'Ventas',
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

// Agregar datos
worksheet.addRow([
  {
    key: 'producto-1',
    value: 'Producto A',
    type: CellType.STRING,
    header: 'Producto'
  },
  {
    key: 'ventas-1',
    value: 1500.50,
    type: CellType.CURRENCY,
    header: 'Ventas',
    numberFormat: '$#,##0.00'
  }
]);

// Generar y descargar
await builder.generateAndDownload('reporte-ventas.xlsx');
```

## 📚 Documentación de API

### Clases Principales

#### `ExcelBuilder`

Clase principal para crear workbooks de Excel.

```typescript
const builder = new ExcelBuilder({
  metadata: {
    title: 'Mi Reporte',
    author: 'Mi Nombre',
    company: 'Mi Empresa',
    description: 'Descripción del reporte',
    keywords: 'excel, reporte, datos',
    created: new Date(),
    modified: new Date()
  },
  enableValidation: true,
  enableEvents: true,
  maxWorksheets: 255,
  maxRowsPerWorksheet: 1048576,
  maxColumnsPerWorksheet: 16384
});

// Métodos principales
builder.addWorksheet(name, config);      // Agregar una hoja
builder.getWorksheet(name);              // Obtener una hoja
builder.removeWorksheet(name);           // Eliminar una hoja
builder.setCurrentWorksheet(name);       // Establecer hoja actual
builder.build(options);                  // Construir y obtener ArrayBuffer
builder.generateAndDownload(fileName);    // Generar y descargar
builder.toBuffer(options);               // Obtener como Buffer
builder.toBlob(options);                // Obtener como Blob
builder.validate();                      // Validar workbook
builder.clear();                         // Limpiar todas las hojas
builder.getStats();                      // Obtener estadísticas

// Sistema de eventos
builder.on(eventType, listener);
builder.off(eventType, listenerId);
builder.removeAllListeners(eventType);
```

#### `ExcelReader`

Clase para leer archivos Excel y convertirlos a JSON con 3 formatos de salida diferentes.

**Formatos disponibles:**
- `worksheet` (por defecto) - Estructura completa con hojas, filas y celdas
- `detailed` - Cada celda con información de posición (texto, columna, fila)
- `flat` - Solo los datos, sin estructura

```typescript
import { ExcelReader, OutputFormat } from 'han-excel-builder';

// ===== FORMATO 1: WORKSHEET (por defecto) =====
// Estructura completa organizada por hojas
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.WORKSHEET, // o 'worksheet'
  useFirstRowAsHeaders: true
});

if (result.success) {
  const workbook = result.data;
  // workbook.sheets[] - Array de hojas
  // workbook.sheets[0].rows[] - Array de filas
  // workbook.sheets[0].rows[0].cells[] - Array de celdas
  // workbook.sheets[0].rows[0].data - Objeto con datos (si useFirstRowAsHeaders)
}

// ===== FORMATO 2: DETAILED =====
// Cada celda con información de posición
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.DETAILED, // o 'detailed'
  includeFormatting: true
});

if (result.success) {
  const detailed = result.data;
  // detailed.cells[] - Array de todas las celdas con:
  //   - value: valor de la celda
  //   - text: texto de la celda
  //   - column: número de columna (1-based)
  //   - columnLetter: letra de columna (A, B, C...)
  //   - row: número de fila (1-based)
  //   - reference: referencia de celda (A1, B2...)
  //   - sheet: nombre de la hoja
  detailed.cells.forEach(cell => {
    console.log(`${cell.sheet}!${cell.reference}: ${cell.text}`);
  });
}

// ===== FORMATO 3: FLAT =====
// Solo los datos, sin estructura
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.FLAT, // o 'flat'
  useFirstRowAsHeaders: true
});

if (result.success) {
  const flat = result.data;
  
  // Si es una sola hoja:
  if ('data' in flat) {
    // flat.data[] - Array de objetos o arrays
    // flat.headers[] - Headers (si useFirstRowAsHeaders)
    flat.data.forEach(row => {
      console.log(row); // { Producto: 'A', Precio: 100 } o ['A', 100]
    });
  }
  
  // Si son múltiples hojas:
  if ('sheets' in flat) {
    // flat.sheets['NombreHoja'].data[] - Datos por hoja
    Object.keys(flat.sheets).forEach(sheetName => {
      console.log(`Hoja: ${sheetName}`);
      flat.sheets[sheetName].data.forEach(row => {
        console.log(row);
      });
    });
  }
}

// ===== USANDO MAPPER PARA TRANSFORMAR DATOS =====
// El mapper permite transformar la respuesta antes de devolverla
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.WORKSHEET,
  useFirstRowAsHeaders: true,
  // Mapper recibe el payload y devuelve la transformación
  mapper: (data) => {
    // Transformar datos según necesidades
    const transformed = {
      totalSheets: data.totalSheets,
      sheets: data.sheets.map(sheet => ({
        name: sheet.name,
        // Convertir filas a objetos con datos transformados
        rows: sheet.rows.map(row => {
          if (row.data) {
            // Transformar cada campo
            return {
              ...row.data,
              // Agregar campos calculados
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

// Ejemplo con formato FLAT y mapper
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.FLAT,
  useFirstRowAsHeaders: true,
  mapper: (data) => {
    // Si es formato flat de una sola hoja
    if ('data' in data && Array.isArray(data.data)) {
      return data.data.map((row: any) => ({
        ...row,
        // Agregar validaciones o transformaciones
        isValid: Object.values(row).every(val => val !== null && val !== undefined)
      }));
    }
    return data;
  }
});

// Ejemplo con formato DETAILED y mapper
const result = await ExcelReader.fromFile(file, {
  outputFormat: OutputFormat.DETAILED,
  mapper: (data) => {
    // Agrupar celdas por hoja
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

**Opciones de lectura:**

```typescript
interface IExcelReaderOptions {
  outputFormat?: 'worksheet' | 'detailed' | 'flat' | OutputFormat; // Formato de salida
  mapper?: (data: IJsonWorkbook | IDetailedFormat | IFlatFormat | IFlatFormatMultiSheet) => unknown; // Función para transformar la respuesta
  useFirstRowAsHeaders?: boolean;    // Usar primera fila como headers
  includeEmptyRows?: boolean;        // Incluir filas vacías
  headers?: string[] | Record<number, string>; // Headers personalizados
  sheetName?: string | number;       // Nombre o índice de hoja
  startRow?: number;                 // Fila inicial (1-based)
  endRow?: number;                    // Fila final (1-based)
  startColumn?: number;               // Columna inicial (1-based)
  endColumn?: number;                 // Columna final (1-based)
  includeFormatting?: boolean;        // Incluir información de formato
  includeFormulas?: boolean;          // Incluir fórmulas
  datesAsISO?: boolean;               // Convertir fechas a ISO string
}
```

**Formatos de salida:**

- **`worksheet`** (por defecto): Estructura completa con hojas, filas y celdas
- **`detailed`**: Array de celdas con información de posición (texto, columna, fila, referencia)
- **`flat`**: Solo los datos, sin estructura (arrays u objetos planos)

#### `Worksheet`

Representa una hoja de cálculo individual.

```typescript
const worksheet = builder.addWorksheet('Mi Hoja', {
  tabColor: '#FF0000',
  defaultRowHeight: 20,
  defaultColWidth: 15,
  pageSetup: {
    orientation: 'portrait',
    paperSize: 9
  }
});

// Métodos principales
worksheet.addHeader(header);             // Agregar header principal
worksheet.addSubHeaders(headers);        // Agregar sub-headers
worksheet.addRow(row);                   // Agregar fila de datos
worksheet.addFooter(footer);             // Agregar footer
worksheet.addTable(config);              // Crear nueva tabla
worksheet.finalizeTable();               // Finalizar tabla actual
worksheet.getTable(name);                // Obtener tabla por nombre
worksheet.validate();                    // Validar hoja
```

### Tipos de Datos

#### `CellType`

```typescript
enum CellType {
  STRING = 'string',        // Texto
  NUMBER = 'number',        // Número
  BOOLEAN = 'boolean',      // Verdadero/Falso
  DATE = 'date',            // Fecha
  PERCENTAGE = 'percentage', // Porcentaje
  CURRENCY = 'currency',    // Moneda
  LINK = 'link',           // Hipervínculo
  FORMULA = 'formula'      // Fórmula
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

### Estilos

#### `StyleBuilder`

API fluida para crear estilos de celdas.

```typescript
const style = new StyleBuilder()
  // Fuentes
  .fontName('Arial')
  .fontSize(12)
  .fontBold()
  .fontItalic()
  .fontUnderline()
  .fontColor('#FF0000')
  
  // Fondos y bordes
  .backgroundColor('#FFFF00')
  .border(BorderStyle.THIN, '#000000')
  .borderTop(BorderStyle.MEDIUM, '#000000')
  .borderLeft(BorderStyle.THIN, '#000000')
  .borderBottom(BorderStyle.THIN, '#000000')
  .borderRight(BorderStyle.THIN, '#000000')
  
  // Alineación
  .centerAlign()
  .leftAlign()
  .rightAlign()
  .horizontalAlign(HorizontalAlignment.CENTER)
  .verticalAlign(VerticalAlignment.MIDDLE)
  .wrapText()
  
  // Formatos
  .numberFormat('$#,##0.00')
  .striped()
  
  // Formato condicional
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

// Método estático alternativo
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

## 🎯 Ejemplos Avanzados

### Múltiples Tablas en una Hoja

```typescript
import { ExcelBuilder, CellType, StyleBuilder, BorderStyle } from 'han-excel-builder';

const builder = new ExcelBuilder();
const worksheet = builder.addWorksheet('Reporte Completo');

// ===== PRIMERA TABLA =====
worksheet.addTable({
  name: 'Ventas',
  showBorders: true,
  showStripes: true,
  style: 'TableStyleLight1'
});

worksheet.addHeader({
  key: 'header-ventas',
  type: CellType.STRING,
  value: 'RESUMEN DE VENTAS',
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
  { key: 'producto', type: CellType.STRING, value: 'Producto' },
  { key: 'ventas', type: CellType.CURRENCY, value: 'Ventas' }
]);

worksheet.addRow([
  { key: 'p1', type: CellType.STRING, value: 'Producto A', header: 'Producto' },
  { key: 'v1', type: CellType.CURRENCY, value: 1500, header: 'Ventas' }
]);

worksheet.finalizeTable();

// ===== SEGUNDA TABLA =====
worksheet.addTable({
  name: 'Empleados',
  showBorders: true,
  showStripes: true,
  style: 'TableStyleMedium1'
});

worksheet.addHeader({
  key: 'header-empleados',
  type: CellType.STRING,
  value: 'TOP EMPLEADOS',
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
  { key: 'nombre', type: CellType.STRING, value: 'Nombre' },
  { key: 'ventas', type: CellType.CURRENCY, value: 'Ventas' }
]);

worksheet.addRow([
  { key: 'e1', type: CellType.STRING, value: 'Juan Pérez', header: 'Nombre' },
  { key: 've1', type: CellType.CURRENCY, value: 150000, header: 'Ventas' }
]);

worksheet.finalizeTable();

await builder.generateAndDownload('multiple-tables.xlsx');
```

### Headers Anidados

```typescript
worksheet.addSubHeaders([
  {
    key: 'ventas',
    value: 'Ventas',
    type: CellType.STRING,
    children: [
      {
        key: 'ventas-q1',
        value: 'Q1',
        type: CellType.STRING
      },
      {
        key: 'ventas-q2',
        value: 'Q2',
        type: CellType.STRING
      }
    ]
  },
  {
    key: 'gastos',
    value: 'Gastos',
    type: CellType.STRING,
    children: [
      {
        key: 'gastos-q1',
        value: 'Q1',
        type: CellType.STRING
      },
      {
        key: 'gastos-q2',
        value: 'Q2',
        type: CellType.STRING
      }
    ]
  }
]);
```

### Hipervínculos

```typescript
worksheet.addRow([
  {
    key: 'link-1',
    type: CellType.LINK,
    value: 'Visitar sitio',
    link: 'https://example.com',
    mask: 'Haz clic aquí', // Texto visible
    header: 'Enlace'
  }
]);
```

### Datos con Children (Estructura Jerárquica)

```typescript
worksheet.addRow([
  {
    key: 'row-1',
    type: CellType.STRING,
    value: 'Categoría Principal',
    header: 'Categoría',
    children: [
      {
        key: 'child-1',
        type: CellType.STRING,
        value: 'Subcategoría 1',
        header: 'Subcategoría'
      },
      {
        key: 'child-2',
        type: CellType.NUMBER,
        value: 100,
        header: 'Valor'
      }
    ]
  }
]);
```

### Formato Condicional

```typescript
worksheet.addRow([
  {
    key: 'ventas-1',
    type: CellType.NUMBER,
    value: 1500,
    header: 'Ventas',
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

### Múltiples Hojas de Cálculo

```typescript
const builder = new ExcelBuilder();

// Hoja 1: Resumen
const summarySheet = builder.addWorksheet('Resumen');
summarySheet.addHeader({
  key: 'title',
  value: 'Resumen Ejecutivo',
  type: CellType.STRING,
  mergeCell: true
});

// Hoja 2: Detalles
const detailsSheet = builder.addWorksheet('Detalles');
detailsSheet.addSubHeaders([
  { key: 'fecha', value: 'Fecha', type: CellType.DATE },
  { key: 'monto', value: 'Monto', type: CellType.CURRENCY }
]);

await builder.generateAndDownload('multi-sheet-report.xlsx');
```

### Exportación en Diferentes Formatos

```typescript
// Descarga directa (navegador)
await builder.generateAndDownload('reporte.xlsx');

// Obtener como Buffer
const bufferResult = await builder.toBuffer();
if (bufferResult.success) {
  const buffer = bufferResult.data;
  // Usar buffer...
}

// Obtener como Blob
const blobResult = await builder.toBlob();
if (blobResult.success) {
  const blob = blobResult.data;
  // Usar blob...
}
```

### Sistema de Eventos

```typescript
builder.on('build:started', (event) => {
  console.log('Construcción iniciada');
});

builder.on('build:completed', (event) => {
  console.log('Construcción completada', event.data);
});

builder.on('build:error', (event) => {
  console.error('Error en construcción', event.data.error);
});

// Remover listener
const listenerId = builder.on('build:started', handler);
builder.off('build:started', listenerId);
```

### Leer Excel y Convertir a JSON

```typescript
import { ExcelReader } from 'han-excel-builder';

// Leer desde un archivo (navegador)
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
    
    // Procesar cada hoja
    workbook.sheets.forEach(sheet => {
      console.log(`Procesando hoja: ${sheet.name}`);
      
      // Convertir a array de objetos (si usamos headers)
      const data = sheet.rows.map(row => row.data || {});
      console.log('Datos:', data);
    });
  }
});

// Leer desde ArrayBuffer (desde API)
async function readExcelFromAPI() {
  const response = await fetch('/api/excel-file');
  const buffer = await response.arrayBuffer();
  
  const result = await ExcelReader.fromBuffer(buffer, {
    useFirstRowAsHeaders: true,
    sheetName: 'Ventas' // Leer solo la hoja 'Ventas'
  });

  if (result.success) {
    const sheet = result.data.sheets[0];
    const ventas = sheet.rows.map(row => row.data);
    return ventas;
  }
}

// Leer desde ruta (Node.js)
async function readExcelFromPath() {
  const result = await ExcelReader.fromPath('./reporte.xlsx', {
    useFirstRowAsHeaders: true,
    startRow: 2, // Saltar header
    includeFormulas: true
  });

  if (result.success) {
    console.log(`Tiempo de procesamiento: ${result.processingTime}ms`);
    return result.data;
  }
}
```

## 🧪 Testing

```bash
# Ejecutar tests
npm test

# Ejecutar tests con cobertura
npm run test:coverage

# Ejecutar tests en modo watch
npm run test:watch
```

## 🛠️ Desarrollo

```bash
# Instalar dependencias
npm install

# Iniciar servidor de desarrollo
npm run dev

# Construir para producción
npm run build

# Ejecutar linting
npm run lint

# Formatear código
npm run format

# Verificar tipos
npm run type-check

# Generar documentación
npm run docs
```

## 📋 Migración desde legacy-excel

Si estás migrando desde la versión legacy, aquí hay una comparación rápida:

```typescript
// Forma legacy
const worksheets: IWorksheets[] = [{
  name: "Reporte",
  tables: [{
    headers: [...],
    subHeaders: [...],
    body: [...],
    footers: [...]
  }]
}];
await fileBuilder(worksheets, "reporte");

// Nueva forma
const builder = new ExcelBuilder();
const worksheet = builder.addWorksheet('Reporte');
worksheet.addHeader({...});
worksheet.addSubHeaders([...]);
worksheet.addRow([...]);
worksheet.addFooter([...]);
await builder.generateAndDownload('reporte');
```

## 📚 Recursos Adicionales

- 📖 [Guía de Múltiples Tablas](./MULTIPLE-TABLES-GUIDE.md)
- 📖 [Mejoras Implementadas](./IMPROVEMENTS.md)
- 📖 [Resultados de Pruebas](./TEST-RESULTS.md)

## 🤝 Contribuir

1. Fork el repositorio
2. Crea una rama de feature (`git checkout -b feature/mi-caracteristica`)
3. Commit tus cambios (`git commit -m 'Agregar mi característica'`)
4. Push a la rama (`git push origin feature/mi-caracteristica`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está licenciado bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.

## 🆘 Soporte

- 📖 [Documentación](https://github.com/hannndler/-han-excel)
- 🐛 [Issues](https://github.com/hannndler/-han-excel/issues)
- 💬 [Discussions](https://github.com/hannndler/-han-excel/discussions)

---

Hecho con ❤️ por el equipo de Han Excel
