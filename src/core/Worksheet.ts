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
import { Result, ErrorType } from '../types/core.types';

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
        ws.addRow([header.value]);
        if (header.mergeCell) {
          const maxCols = this.calculateTableMaxColumns(table);
          ws.mergeCells(rowPointer, 1, rowPointer, maxCols);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell: any) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
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
        ws.addRow([header.value]);
        if (header.mergeCell) {
          ws.mergeCells(rowPointer, 1, rowPointer, (this.getMaxColumns() || 1));
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell: any) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
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
          cell.value = headerInfo.value;
          if (headerInfo.style) {
            cell.style = this.convertStyle(headerInfo.style);
          }
          colIndex += headerInfo.colSpan;
        } else {
          // Nivel de children - procesar todos los children directos
          if (header.children && header.children.length > 0) {
            for (const child of header.children) {
              const cell = row.getCell(colIndex);
              cell.value = typeof child.value === 'string' ? child.value : String(child.value || '');
              if (child.styles || header.styles) {
                cell.style = this.convertStyle(child.styles || header.styles);
              }
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
    footerCell.value = footer.value;
    if (footer.styles) {
      footerCell.style = this.convertStyle(footer.styles);
    }
    if (footer.numberFormat) {
      footerCell.numFmt = footer.numberFormat;
    }
    
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
            childCell.value = child.value;
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
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
    mainCell.value = row.value;
    if (row.styles) {
      mainCell.style = this.convertStyle(row.styles);
    }
    if (row.numberFormat) {
      mainCell.numFmt = row.numberFormat;
    }
    
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
            childCell.value = child.value;
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
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
   * Convierte el estilo personalizado a formato compatible con ExcelJS
   */
  private convertStyle(style: any): Partial<ExcelJS.Style> {
    if (!style) return {};
    
    const converted: Partial<ExcelJS.Style> = {};
    
    if (style.font) {
      converted.font = {
        name: style.font.family,
        size: style.font.size,
        bold: style.font.bold,
        italic: style.font.italic,
        underline: style.font.underline,
        color: style.font.color
      };
    }
    
    if (style.fill) {
      converted.fill = {
        type: style.fill.type,
        pattern: style.fill.pattern,
        fgColor: style.fill.foregroundColor,
        bgColor: style.fill.backgroundColor
      };
    }
    
    if (style.border) {
      converted.border = {
        top: style.border.top,
        left: style.border.left,
        bottom: style.border.bottom,
        right: style.border.right
      };
    }
    
    if (style.alignment) {
      converted.alignment = {
        horizontal: style.alignment.horizontal,
        vertical: style.alignment.vertical,
        wrapText: style.alignment.wrapText,
        indent: style.alignment.indent
      };
    }
    
    if (style.numFmt) {
      converted.numFmt = style.numFmt;
    }
    
    return converted;
  }
}