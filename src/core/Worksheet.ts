import ExcelJS from 'exceljs';
import {
  IWorksheet,
  IWorksheetConfig,
  ITable
} from '../types/worksheet.types';
import {
  IDataCell,
  IHeaderCell,
  IFooterCell
} from '../types/cell.types';
import { IBuildOptions } from '../types/builder.types';
import { Result, ErrorType, CellType } from '../types/core.types';

/**
 * Worksheet - Representa una hoja de cálculo dentro del builder
 *
2 * Soporta headers, subheaders anidados, rows, footers, children y estilos por celda.
 */
export class Worksheet implements IWorksheet {
  public config: IWorksheetConfig;
  public tables: ITable[] = [];
  public currentRow = 1;
  public currentCol = 1;
  public headerPointers: Map<string, any> = new Map();
  public isBuilt = false;

  // Estructuras temporales para la tabla actual
  private headers: IHeaderCell[] = [];
  private subHeaders: IHeaderCell[] = [];
  private body: IDataCell[] = [];
  private footers: IFooterCell[] = [];

  constructor(config: IWorksheetConfig) {
    this.config = config;
  }

  /**
   * Agrega un header principal
   */
  addHeader(header: IHeaderCell): this {
    this.headers.push(header);
    return this;
  }

  /**
   * Agrega subheaders (ahora soporta anidación)
   */
  addSubHeaders(subHeaders: IHeaderCell[]): this {
    this.subHeaders.push(...subHeaders);
    return this;
  }

  /**
   * Agrega una fila de datos (puede ser jerárquica con childrens)
   */
  addRow(row: IDataCell[] | IDataCell): this {
    if (Array.isArray(row)) {
      this.body.push(...row);
    } else {
      this.body.push(row);
    }
    return this;
  }

  /**
   * Agrega un footer o varios
   */
  addFooter(footer: IFooterCell[] | IFooterCell): this {
    if (Array.isArray(footer)) {
      this.footers.push(...footer);
    } else {
      this.footers.push(footer);
    }
    return this;
  }

  /**
   * Crea una nueva tabla y la agrega al worksheet
   */
  addTable(tableConfig: Partial<ITable> = {}): this {
    const table: ITable = {
      name: tableConfig.name || `Table_${this.tables.length + 1}`,
      headers: tableConfig.headers || [],
      subHeaders: tableConfig.subHeaders || [],
      body: tableConfig.body || [],
      footers: tableConfig.footers || [],
      showBorders: tableConfig.showBorders !== false,
      showStripes: tableConfig.showStripes !== false,
      style: tableConfig.style || 'TableStyleLight1',
      ...tableConfig
    };
    
    this.tables.push(table);
    return this;
  }

  /**
   * Finaliza la tabla actual agregando todos los elementos temporales a la última tabla
   */
  finalizeTable(): this {
    if (this.tables.length === 0) {
      // Si no hay tablas, crear una nueva con los datos temporales
      this.addTable();
    }
    
    const currentTable = this.tables[this.tables.length - 1];
    if (!currentTable) {
      throw new Error('No se pudo obtener la tabla actual');
    }
    
    // Agregar headers, subheaders, body y footers a la tabla actual
    if (this.headers.length > 0) {
      currentTable.headers = [...(currentTable.headers || []), ...this.headers];
    }
    
    if (this.subHeaders.length > 0) {
      currentTable.subHeaders = [...(currentTable.subHeaders || []), ...this.subHeaders];
    }
    
    if (this.body.length > 0) {
      currentTable.body = [...(currentTable.body || []), ...this.body];
    }
    
    if (this.footers.length > 0) {
      currentTable.footers = [...(currentTable.footers || []), ...this.footers];
    }
    
    // Limpiar las estructuras temporales
    this.headers = [];
    this.subHeaders = [];
    this.body = [];
    this.footers = [];
    
    return this;
  }

  /**
   * Obtiene una tabla por nombre
   */
  getTable(name: string): ITable | undefined {
    return this.tables.find(table => table.name === name);
  }

  /**
   * Construye la hoja en el workbook de ExcelJS
   */
  async build(workbook: ExcelJS.Workbook, _options: IBuildOptions = {}): Promise<void> {
    const ws = workbook.addWorksheet(this.config.name, {
      properties: {
        defaultRowHeight: this.config.defaultRowHeight || 20,
        tabColor: this.config.tabColor as any
      },
      pageSetup: this.config.pageSetup as any
    });

    let rowPointer = 1;
    
    // Si hay tablas definidas, construir cada tabla
    if (this.tables.length > 0) {
      for (let i = 0; i < this.tables.length; i++) {
        const table = this.tables[i];
        if (table) {
          rowPointer = await this.buildTable(ws, table, rowPointer, i > 0);
        }
      }
    } else {
      // Construcción tradicional para compatibilidad hacia atrás
      rowPointer = await this.buildLegacyContent(ws, rowPointer);
    }
    
    this.isBuilt = true;
  }

  /**
   * Construye una tabla individual en el worksheet
   */
  private async buildTable(ws: ExcelJS.Worksheet, table: ITable, startRow: number, addSpacing: boolean = false): Promise<number> {
    let rowPointer = startRow;
    
    // Agregar espacio entre tablas si no es la primera
    if (addSpacing) {
      rowPointer += 2; // 2 filas de espacio
    }
    
    // Headers principales de la tabla
    if (table.headers && table.headers.length > 0) {
      for (const header of table.headers) {
        ws.addRow([this.processCellValue(header)]);
        if (header.mergeCell) {
          const maxCols = this.calculateTableMaxColumns(table);
          ws.mergeCells(rowPointer, 1, rowPointer, maxCols);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell: any) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
        // Aplicar dimensiones de celda
        this.applyCellDimensions(ws, rowPointer, 1, header);
        rowPointer++;
      }
    }
    
    // SubHeaders con soporte para anidación
    if (table.subHeaders && table.subHeaders.length > 0) {
      rowPointer = this.buildNestedHeaders(ws, rowPointer, table.subHeaders);
    }
    
    // Body (soporta children)
    if (table.body && table.body.length > 0) {
      for (const row of table.body) {
        rowPointer = this.addDataRowRecursive(ws, rowPointer, row);
      }
    }
    
    // Footers
    if (table.footers && table.footers.length > 0) {
      for (const footer of table.footers) {
        rowPointer = this.addFooterRow(ws, rowPointer, footer);
      }
    }
    
    // Aplicar estilo de tabla si está configurado
    if (table.showBorders || table.showStripes) {
      this.applyTableStyle(ws, table, startRow, rowPointer - 1);
    }
    
    return rowPointer;
  }

  /**
   * Construcción tradicional para compatibilidad hacia atrás
   */
  private async buildLegacyContent(ws: ExcelJS.Worksheet, startRow: number): Promise<number> {
    let rowPointer = startRow;
    
    // Headers principales
    if (this.headers.length > 0) {
      this.headers.forEach(header => {
        ws.addRow([this.processCellValue(header)]);
        if (header.mergeCell) {
          ws.mergeCells(rowPointer, 1, rowPointer, (this.getMaxColumns() || 1));
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell: any) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
        // Aplicar dimensiones de celda
        this.applyCellDimensions(ws, rowPointer, 1, header);
        rowPointer++;
      });
    }
    
    // SubHeaders con soporte para anidación
    if (this.subHeaders.length > 0) {
      rowPointer = this.buildNestedHeaders(ws, rowPointer, this.subHeaders);
    }
    
    // Body (soporta children)
    for (const row of this.body) {
      rowPointer = this.addDataRowRecursive(ws, rowPointer, row);
    }
    
    // Footers
    if (this.footers.length > 0) {
      for (const footer of this.footers) {
        rowPointer = this.addFooterRow(ws, rowPointer, footer);
      }
    }
    
    return rowPointer;
  }

  /**
   * Calcula el número máximo de columnas para una tabla
   */
  private calculateTableMaxColumns(table: ITable): number {
    let maxCols = 0;
    
    if (table.subHeaders && table.subHeaders.length > 0) {
      for (const header of table.subHeaders) {
        maxCols += this.calculateHeaderColSpan(header);
      }
    }
    
    return maxCols || 1;
  }

  /**
   * Aplica el estilo de tabla a un rango específico
   */
  private applyTableStyle(ws: ExcelJS.Worksheet, table: ITable, startRow: number, endRow: number): void {
    const maxCols = this.calculateTableMaxColumns(table);
    
    // Aplicar bordes si está configurado
    if (table.showBorders) {
      for (let row = startRow; row <= endRow; row++) {
        for (let col = 1; col <= maxCols; col++) {
          const cell = ws.getRow(row).getCell(col);
          if (!cell.style) cell.style = {};
          if (!cell.style.border) {
            cell.style.border = {
              top: { style: 'thin', color: { argb: 'FF8EAADB' } },
              left: { style: 'thin', color: { argb: 'FF8EAADB' } },
              bottom: { style: 'thin', color: { argb: 'FF8EAADB' } },
              right: { style: 'thin', color: { argb: 'FF8EAADB' } }
            };
          }
        }
      }
    }
    
    // Aplicar rayas alternadas si está configurado
    if (table.showStripes) {
      for (let row = startRow; row <= endRow; row++) {
        if ((row - startRow) % 2 === 1) { // Filas impares (empezando desde 0)
          for (let col = 1; col <= maxCols; col++) {
            const cell = ws.getRow(row).getCell(col);
            if (!cell.style) cell.style = {};
            if (!cell.style.fill) {
              cell.style.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF2F2F2' }
              };
            }
          }
        }
      }
    }
  }

  /**
   * Construye headers anidados recursivamente
   * @param ws - Worksheet de ExcelJS
   * @param startRow - Fila inicial
   * @param headers - Array de headers a procesar
   * @returns La siguiente fila disponible
   */
    private buildNestedHeaders(ws: ExcelJS.Worksheet, startRow: number, headers: IHeaderCell[]): number {
    let currentRow = startRow;
    const maxDepth = this.getMaxHeaderDepth(headers);
    
    // Crear filas para cada nivel de profundidad
    for (let depth = 0; depth < maxDepth; depth++) {
      // Crear la fila primero
      const row = ws.getRow(currentRow);
      
      // Procesar cada header en este nivel
      let colIndex = 1;
      for (const header of headers) {
        if (depth === 0) {
          // Nivel principal del header
          const headerInfo = this.getHeaderAtDepth(header, depth, colIndex);
          const cell = row.getCell(colIndex);
          cell.value = this.processCellValue(header);
          if (headerInfo.style) {
            cell.style = this.convertStyle(headerInfo.style);
          }
          // Aplicar dimensiones de celda
          this.applyCellDimensions(ws, currentRow, colIndex, header);
          colIndex += headerInfo.colSpan;
        } else {
          // Nivel de children - procesar todos los children directos
          if (header.children && header.children.length > 0) {
            for (const child of header.children) {
              const cell = row.getCell(colIndex);
              cell.value = this.processCellValue(child);
              if (child.styles || header.styles) {
                cell.style = this.convertStyle(child.styles || header.styles);
              }
              // Aplicar dimensiones de celda para children
              this.applyCellDimensions(ws, currentRow, colIndex, child);
              colIndex += this.calculateHeaderColSpan(child);
            }
          } else {
            // Si no tiene children, agregar celda vacía
            const cell = row.getCell(colIndex);
            cell.value = null;
            colIndex += 1;
          }
        }
      }
      
      currentRow++;
    }
    
    // Aplicar todos los merges después de crear todas las filas
    this.applyAllMerges(ws, startRow, currentRow - 1, headers);
    
    return currentRow;
  }

  /**
   * Obtiene información del header en una profundidad específica
   */
  private getHeaderAtDepth(header: IHeaderCell, depth: number, startCol: number): {
    value: string | null;
    style: any;
    colSpan: number;
    mergeRange?: { start: number; end: number } | null;
  } {
    const colSpan = this.calculateHeaderColSpan(header);
    if (depth === 0) {
      // Nivel principal del header
      const mergeRange = colSpan > 1 ? { start: startCol, end: startCol + colSpan - 1 } : null;
      return {
        value: typeof header.value === 'string' ? header.value : String(header.value || ''),
        style: header.styles,
        colSpan,
        mergeRange: mergeRange
      };
    } else if (header.children && header.children.length > 0) {
      // Nivel de children
      const child = header.children[depth];
      if (child) {
        const childColSpan = this.calculateHeaderColSpan(child);
        // Los children también pueden hacer merge si tienen múltiples childrens
        const mergeRange = childColSpan > 1 ? { start: startCol, end: startCol + childColSpan - 1 } : null;
        
        return {
          value: typeof child.value === 'string' ? child.value : String(child.value || ''),
          style: child.styles || header.styles,
          colSpan: childColSpan,
          mergeRange: mergeRange
        };
      }
    }
    
    // Celda vacía para mantener alineación
    return {
      value: null,
      style: null,
      colSpan: 1
    };
  }




  /**
   * Aplica todos los merges (horizontales y verticales) después de crear todas las filas
   */
  private applyAllMerges(ws: ExcelJS.Worksheet, startRow: number, endRow: number, headers: IHeaderCell[]): void {
    const maxDepth = this.getMaxHeaderDepth(headers);
    
    // Solo aplicar merges si hay más de una fila de headers
    if (maxDepth <= 1) return;
    
    // Aplicar merges inteligentes basados en la estructura
    this.applySmartMerges(ws, startRow, endRow, headers);
  }

  /**
   * Aplica merges inteligentes basados en la estructura de headers
   */
  private applySmartMerges(ws: ExcelJS.Worksheet, startRow: number, endRow: number, headers: IHeaderCell[]): void {
    const maxDepth = this.getMaxHeaderDepth(headers);
    
    // Solo aplicar merges si hay más de una fila de headers
    if (maxDepth <= 1) return;
    
    // Aplicar merges para cada header
    let colIndex = 1;
    for (const header of headers) {
      this.applySmartMergesForHeader(ws, startRow, endRow, header, colIndex);
      colIndex += this.calculateHeaderColSpan(header);
    }
  }

  /**
   * Aplica merges inteligentes para un header específico
   */
  private applySmartMergesForHeader(ws: ExcelJS.Worksheet, startRow: number, endRow: number, header: IHeaderCell, startCol: number): void {
    const headerColSpan = this.calculateHeaderColSpan(header);
    
    if (!header.children || header.children.length === 0) {
      // Si no tiene children, hacer merge vertical desde la primera fila hasta la última
      ws.mergeCells(startRow, startCol, endRow, startCol + headerColSpan - 1);
    } else {
      // Si tiene children, aplicar merge horizontal en la primera fila
      if (headerColSpan > 1) {
        ws.mergeCells(startRow, startCol, startRow, startCol + headerColSpan - 1);
      }
      
      // Procesar children recursivamente
      let childColIndex = startCol;
      for (const child of header.children) {
        this.applySmartMergesForHeader(ws, startRow + 1, endRow, child, childColIndex);
        childColIndex += this.calculateHeaderColSpan(child);
      }
    }
  }


 
  /**
   * Calcula el span de columnas para un header
   */
  private calculateHeaderColSpan(header: IHeaderCell): number {
    if (!header.children || header.children.length === 0) {
      return 1;
    }
    
    return header.children.reduce((total, child) => {
      return total + this.calculateHeaderColSpan(child);
    }, 0);
  }

  /**
   * Obtiene la profundidad máxima de headers anidados
   */
  private getMaxHeaderDepth(headers: IHeaderCell[]): number {
    let maxDepth = 1;
    
    for (const header of headers) {
      if (header.children && header.children.length > 0) {
        const childDepth = this.getMaxHeaderDepth(header.children);
        maxDepth = Math.max(maxDepth, childDepth + 1);
      }
    }
    
    return maxDepth;
  }

  /**
   * Obtiene el número máximo de columnas
   */
  private getMaxColumns(): number {
    let maxCols = 0;
    
    for (const header of this.subHeaders) {
      maxCols += this.calculateHeaderColSpan(header);
    }
    
    return maxCols;
  }

  /**
   * Valida la hoja
   */
  validate(): Result<boolean> {
    if (!this.headers.length && !this.body.length) {
      return {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: 'La hoja no tiene datos',
        }
      };
    }
    return { success: true, data: true };
  }

  /**
   * Calcula las posiciones de columnas para los datos basándose en la estructura de subheaders
   */
  private calculateDataColumnPositions(): { [key: string]: number } {
    const positions: { [key: string]: number } = {};
    let currentCol = 1;
    
    for (const header of this.subHeaders) {
      if (header.children && header.children.length > 0) {
        // Si el header tiene children, cada child ocupa una columna
        for (const child of header.children) {
          if (child.key) {
            positions[child.key] = currentCol;
          }
          if (child.value) {
            positions[String(child.value)] = currentCol;
          }
          currentCol++;
        }
      } else {
        // Si el header no tiene children, ocupa una columna
        if (header.key) {
          positions[header.key] = currentCol;
        }
        if (header.value) {
          positions[String(header.value)] = currentCol;
        }
        currentCol++;
      }
    }
    
    return positions;
  }

  /**
   * Agrega una fila de footer
   * @returns el siguiente rowPointer disponible
   */
  private addFooterRow(ws: ExcelJS.Worksheet, rowPointer: number, footer: IFooterCell): number {
    // Calcular las columnas basándose en la estructura de subheaders
    const columnPositions = this.calculateDataColumnPositions();
    
    // Buscar la columna correcta para el footer
    let footerColPosition: number | undefined;
    
    // Intentar encontrar por key primero
    if (footer.key && columnPositions[footer.key]) {
      footerColPosition = columnPositions[footer.key];
    }
    // Si no se encuentra por key, intentar por header
    else if (footer.header && columnPositions[footer.header]) {
      footerColPosition = columnPositions[footer.header];
    }
    
    // Si no se encuentra la posición, usar columna 1 por defecto
    if (footerColPosition === undefined) {
      footerColPosition = 1;
    }
    
    // Escribir el footer en la columna correcta
    const excelRow = ws.getRow(rowPointer);
    const footerCell = excelRow.getCell(footerColPosition);
    footerCell.value = this.processCellValue(footer);
    if (footer.styles) {
      footerCell.style = this.convertStyle(footer.styles);
    }
    if (footer.numberFormat) {
      footerCell.numFmt = footer.numberFormat;
    }
    
    // Aplicar dimensiones de celda
    this.applyCellDimensions(ws, rowPointer, footerColPosition, footer);
    
    // Aplicar merge si está configurado
    if (footer.mergeCell && footer.mergeTo) {
      ws.mergeCells(rowPointer, footerColPosition, rowPointer, footer.mergeTo);
    }
    
    // Si hay children, escribirlos en las columnas correspondientes
    if (footer.children && footer.children.length > 0) {
      for (const child of footer.children) {
        if (child) {
          // Buscar la columna correcta basándose en el header del child
          let colPosition: number | undefined;
          
          // Intentar encontrar por key primero
          if (child.key && columnPositions[child.key]) {
            colPosition = columnPositions[child.key];
          }
          // Si no se encuentra por key, intentar por header
          else if (child.header && columnPositions[child.header]) {
            colPosition = columnPositions[child.header];
          }
          
          if (colPosition !== undefined) {
            const childCell = excelRow.getCell(colPosition);
            childCell.value = this.processCellValue(child);
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
            
            // Aplicar dimensiones de celda para children
            this.applyCellDimensions(ws, rowPointer, colPosition, child);
          }
        }
      }
    }
    
    // Incrementar rowPointer solo si el footer tiene la propiedad jump
    if (footer.jump) {
      return rowPointer + 1;
    }
    
    return rowPointer;
  }

  /**
   * Aplica width y height a una celda/fila
   */
  private applyCellDimensions(ws: ExcelJS.Worksheet, row: number, col: number, cell: IDataCell | IHeaderCell | IFooterCell): void {
    // Aplicar rowHeight si está definido
    if (cell.rowHeight !== undefined) {
      const excelRow = ws.getRow(row);
      excelRow.height = cell.rowHeight;
    }
    
    // Aplicar colWidth si está definido
    if (cell.colWidth !== undefined) {
      const excelCol = ws.getColumn(col);
      excelCol.width = cell.colWidth;
    }
  }

  /**
   * Procesa el valor de una celda considerando links y máscaras
   * Si el tipo es LINK o hay un link, crea un hipervínculo en Excel
   */
  private processCellValue(cell: IDataCell | IHeaderCell | IFooterCell): ExcelJS.CellValue {
    // Si hay un link o el tipo es LINK, crear hipervínculo
    if (cell.link || cell.type === CellType.LINK) {
      const linkUrl = cell.link || (typeof cell.value === 'string' ? cell.value : '');
      
      // Si no hay URL válida, retornar el valor normal
      if (!linkUrl || linkUrl.trim() === '') {
        return cell.value;
      }
      
      // Determinar el texto visible: usar máscara si existe, sino usar value, sino usar la URL
      const displayText = cell.mask || cell.value || linkUrl;
      
      // Crear objeto de hipervínculo para ExcelJS
      return {
        text: String(displayText),
        hyperlink: linkUrl
      } as any;
    }
    
    // Si no hay link, retornar el valor normal
    return cell.value;
  }

  /**
   * Agrega una fila de datos y sus children recursivamente
   * @returns el siguiente rowPointer disponible
   */
  private addDataRowRecursive(ws: ExcelJS.Worksheet, rowPointer: number, row: IDataCell): number {
    // Calcular las columnas basándose en la estructura de subheaders
    const columnPositions = this.calculateDataColumnPositions();

    // Buscar la columna correcta para el dato principal
    let mainColPosition: number | undefined;
    
    // Intentar encontrar por key primero
    if (row.key && columnPositions[row.key]) {
      mainColPosition = columnPositions[row.key];
    }
    // Si no se encuentra por key, intentar por header
    else if (row.header && columnPositions[row.header]) {
      mainColPosition = columnPositions[row.header];
    }
    
    // Si no se encuentra la posición, usar columna 1 por defecto
    if (mainColPosition === undefined) {
      mainColPosition = 1;
    }
    
    // Escribir el dato principal en la columna correcta
    const excelRow = ws.getRow(rowPointer);
    const mainCell = excelRow.getCell(mainColPosition);
    mainCell.value = this.processCellValue(row);
    if (row.styles) {
      mainCell.style = this.convertStyle(row.styles);
    }
    if (row.numberFormat) {
      mainCell.numFmt = row.numberFormat;
    }
    
    // Aplicar dimensiones de celda
    this.applyCellDimensions(ws, rowPointer, mainColPosition, row);
    
    // Si hay children, escribirlos en las columnas correspondientes
    if (row.children && row.children.length > 0) {
      for (const child of row.children) {
        if (child) {
          // Buscar la columna correcta basándose en el header del child
          let colPosition: number | undefined;
          
          // Intentar encontrar por key primero
          if (child.key && columnPositions[child.key]) {
            colPosition = columnPositions[child.key];
          }
          // Si no se encuentra por key, intentar por header
          else if (child.header && columnPositions[child.header]) {
            colPosition = columnPositions[child.header];
          }
          
          if (colPosition !== undefined) {
            const childCell = excelRow.getCell(colPosition);
            childCell.value = this.processCellValue(child);
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
            
            // Aplicar dimensiones de celda para children
            this.applyCellDimensions(ws, rowPointer, colPosition, child);
          }
        }
      }
    }
    
    // Incrementar rowPointer solo si la celda tiene la propiedad jump
    if (row.jump) {
      return rowPointer + 1;
    }
    
    return rowPointer;
  }

  /**
   * Convierte un color a formato ExcelJS (ARGB)
   */
  private convertColor(color: any): any {
    if (!color) return undefined;
    
    // Si ya es un objeto con argb, retornarlo
    if (typeof color === 'object' && color.argb) {
      return color;
    }
    
    // Si es un objeto con r, g, b
    if (typeof color === 'object' && 'r' in color && 'g' in color && 'b' in color) {
      const r = color.r.toString(16).padStart(2, '0');
      const g = color.g.toString(16).padStart(2, '0');
      const b = color.b.toString(16).padStart(2, '0');
      return { argb: `FF${r}${g}${b}`.toUpperCase() };
    }
    
    // Si es un string (hex)
    if (typeof color === 'string') {
      // Remover # si existe
      let hex = color.replace('#', '');
      
      // Si es formato corto (RGB), expandirlo
      if (hex.length === 3) {
        hex = hex.split('').map(c => c + c).join('');
      }
      
      // Asegurar que tenga alpha (FF = completamente opaco)
      if (hex.length === 6) {
        hex = 'FF' + hex.toUpperCase();
      }
      
      return { argb: hex };
    }
    
    // Si es un objeto theme
    if (typeof color === 'object' && 'theme' in color) {
      return color;
    }
    
    return undefined;
  }

  /**
   * Convierte el estilo personalizado a formato compatible con ExcelJS
   */
  private convertStyle(style: any): Partial<ExcelJS.Style> {
    if (!style) return {};
    
    const converted: Partial<ExcelJS.Style> = {};
    
    if (style.font) {
      converted.font = {
        name: style.font.family || style.font.name,
        size: style.font.size,
        bold: style.font.bold,
        italic: style.font.italic,
        underline: style.font.underline,
        color: this.convertColor(style.font.color)
      };
    }
    
    if (style.fill) {
      // En ExcelJS, para patrón sólido, el color de fondo debe ir en fgColor
      // backgroundColor es el color que queremos mostrar como fondo de la celda
      const pattern = style.fill.pattern || 'solid';
      
      // Para patrón sólido: backgroundColor va en fgColor (es el color visible)
      // Para otros patrones: foregroundColor es el color del patrón, backgroundColor es el fondo
      const fgColor = pattern === 'solid' 
        ? (style.fill.backgroundColor || style.fill.foregroundColor)
        : (style.fill.foregroundColor || style.fill.backgroundColor);
      
      // bgColor solo es relevante para patrones no sólidos
      const bgColor = pattern !== 'solid' ? style.fill.backgroundColor : undefined;
      
      converted.fill = {
        type: style.fill.type || 'pattern',
        pattern: pattern,
        fgColor: this.convertColor(fgColor),
        bgColor: bgColor ? this.convertColor(bgColor) : undefined
      };
      
      // Limpiar bgColor si es undefined para evitar problemas
      if (!converted.fill.bgColor) {
        delete converted.fill.bgColor;
      }
    }
    
    if (style.border) {
      converted.border = {};
      
      if (style.border.top) {
        converted.border.top = {
          style: style.border.top.style,
          color: this.convertColor(style.border.top.color)
        };
      }
      
      if (style.border.left) {
        converted.border.left = {
          style: style.border.left.style,
          color: this.convertColor(style.border.left.color)
        };
      }
      
      if (style.border.bottom) {
        converted.border.bottom = {
          style: style.border.bottom.style,
          color: this.convertColor(style.border.bottom.color)
        };
      }
      
      if (style.border.right) {
        converted.border.right = {
          style: style.border.right.style,
          color: this.convertColor(style.border.right.color)
        };
      }
    }
    
    if (style.alignment) {
      converted.alignment = {};
      
      // Horizontal alignment - validar valores permitidos
      if (style.alignment.horizontal !== undefined) {
        const validHorizontal = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
        if (validHorizontal.includes(style.alignment.horizontal)) {
          converted.alignment.horizontal = style.alignment.horizontal as any;
        }
      }
      
      // Vertical alignment - validar valores permitidos
      if (style.alignment.vertical !== undefined) {
        const validVertical = ['top', 'middle', 'bottom', 'distributed', 'justify'];
        if (validVertical.includes(style.alignment.vertical)) {
          converted.alignment.vertical = style.alignment.vertical as any;
        }
      }
      
      // Wrap text
      if (style.alignment.wrapText !== undefined) {
        converted.alignment.wrapText = Boolean(style.alignment.wrapText);
      }
      
      // Shrink to fit
      if (style.alignment.shrinkToFit !== undefined) {
        converted.alignment.shrinkToFit = Boolean(style.alignment.shrinkToFit);
      }
      
      // Indent
      if (style.alignment.indent !== undefined && typeof style.alignment.indent === 'number') {
        converted.alignment.indent = style.alignment.indent;
      }
      
      // Text rotation (0-180 grados)
      if (style.alignment.textRotation !== undefined && typeof style.alignment.textRotation === 'number') {
        converted.alignment.textRotation = style.alignment.textRotation;
      }
      
      // Reading order
      if (style.alignment.readingOrder !== undefined) {
        const validReadingOrder = ['left-to-right', 'right-to-left', 'context'];
        if (validReadingOrder.includes(style.alignment.readingOrder)) {
          converted.alignment.readingOrder = style.alignment.readingOrder as any;
        }
      }
      
      // Solo agregar alignment si tiene al menos una propiedad
      if (Object.keys(converted.alignment).length === 0) {
        delete converted.alignment;
      }
    }
    
    if (style.numFmt) {
      converted.numFmt = style.numFmt;
    }
    
    return converted;
  }
}