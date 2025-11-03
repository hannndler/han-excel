# Implementación de Múltiples Tablas en Han Excel

## Descripción

Se ha implementado la funcionalidad para crear múltiples tablas en una sola hoja de cálculo. Cada tabla puede contener su propio header, body, footer y estilos independientes.

## Nuevas Funcionalidades

### Métodos Agregados al Worksheet

#### `addTable(tableConfig?: Partial<ITable>): this`
Crea una nueva tabla y la agrega al worksheet.

**Parámetros:**
- `tableConfig` (opcional): Configuración de la tabla

**Ejemplo:**
```typescript
worksheet.addTable({
  name: 'MiTabla',
  showBorders: true,
  showStripes: true,
  style: 'TableStyleLight1'
});
```

#### `finalizeTable(): this`
Finaliza la tabla actual agregando todos los elementos temporales (headers, subheaders, body, footers) a la última tabla creada.

**Ejemplo:**
```typescript
// Agregar contenido a la tabla
worksheet.addHeader({...});
worksheet.addSubHeaders([...]);
worksheet.addRow([...]);
worksheet.addFooter([...]);

// Finalizar la tabla
worksheet.finalizeTable();
```

#### `getTable(name: string): ITable | undefined`
Obtiene una tabla por su nombre.

**Ejemplo:**
```typescript
const tabla = worksheet.getTable('MiTabla');
```

## Uso Básico

### Crear Múltiples Tablas

```typescript
import { ExcelBuilder, CellType, StyleBuilder } from '../index';

const builder = new ExcelBuilder();
const worksheet = builder.addWorksheet('Mi Hoja');

// ===== PRIMERA TABLA =====
worksheet.addTable({
  name: 'Ventas',
  showBorders: true,
  showStripes: true
});

worksheet.addHeader({
  key: 'header1',
  type: CellType.STRING,
  value: 'TABLA DE VENTAS',
  mergeCell: true,
  styles: new StyleBuilder()
    .fontBold()
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

// Agregar datos
worksheet.addRow([
  {
    key: 'laptop',
    type: CellType.STRING,
    value: 'Laptop',
    header: 'Producto'
  },
  {
    key: 'precio-laptop',
    type: CellType.NUMBER,
    value: 1000,
    header: 'Precio'
  }
]);

// Finalizar la primera tabla
worksheet.finalizeTable();

// ===== SEGUNDA TABLA =====
worksheet.addTable({
  name: 'Empleados',
  showBorders: true,
  showStripes: true
});

worksheet.addHeader({
  key: 'header2',
  type: CellType.STRING,
  value: 'TABLA DE EMPLEADOS',
  mergeCell: true,
  styles: new StyleBuilder()
    .fontBold()
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

// Agregar datos
worksheet.addRow([
  {
    key: 'juan',
    type: CellType.STRING,
    value: 'Juan Pérez',
    header: 'Nombre'
  },
  {
    key: 'edad-juan',
    type: CellType.NUMBER,
    value: 30,
    header: 'Edad'
  }
]);

// Finalizar la segunda tabla
worksheet.finalizeTable();

// Generar el archivo
await builder.generateAndDownload('multiple-tables.xlsx');
```

## Características

### Separación Automática
- Las tablas se separan automáticamente con 2 filas de espacio entre ellas
- Cada tabla mantiene su estructura independiente

### Estilos Independientes
- Cada tabla puede tener sus propios estilos de header, body y footer
- Soporte para bordes y rayas alternadas por tabla
- Diferentes colores y estilos para cada tabla

### Compatibilidad Hacia Atrás
- El código existente sigue funcionando sin cambios
- Si no se usan tablas, el comportamiento es el mismo que antes

## Configuración de Tabla

### Propiedades de ITable

```typescript
interface ITable {
  name?: string;                    // Nombre de la tabla
  headers?: IHeaderCell[];          // Headers principales
  subHeaders?: IHeaderCell[];       // Subheaders
  body?: IDataCell[];              // Datos del cuerpo
  footers?: IFooterCell[];        // Footers
  range?: ICellRange;              // Rango de la tabla
  showBorders?: boolean;           // Mostrar bordes
  showStripes?: boolean;           // Mostrar rayas alternadas
  style?: TableStyle;              // Estilo de tabla
}
```

### Estilos de Tabla Disponibles
- `TableStyleLight1`
- `TableStyleLight2`
- `TableStyleMedium1`
- `TableStyleMedium2`
- `TableStyleDark1`
- `TableStyleDark2`

## Ejemplos

Ver los archivos de ejemplo:
- `src/examples/multiple-tables.ts` - Ejemplo completo con 3 tablas
- `src/examples/simple-multiple-tables.ts` - Ejemplo básico con 2 tablas

## Notas Importantes

1. **Siempre llamar `finalizeTable()`** después de agregar contenido a una tabla
2. **Usar `addTable()`** antes de agregar contenido para cada nueva tabla
3. **Los estilos se aplican automáticamente** si `showBorders` o `showStripes` están habilitados
4. **El espaciado entre tablas** se maneja automáticamente
5. **Compatibilidad total** con el código existente

