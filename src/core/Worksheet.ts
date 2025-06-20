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
 * Soporta headers, subheaders, rows, footers, children y estilos por celda.
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
   * Agrega subheaders
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
    // Headers
    if (this.headers.length > 0) {
      this.headers.forEach(header => {
        ws.addRow([header.value]);
        if (header.mergeCell) {
          ws.mergeCells(rowPointer, 1, rowPointer, (this.subHeaders.length || 1));
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell(cell => {
            cell.style = this.convertStyle(header.styles);
          });
        }
        rowPointer++;
      });
    }
    // SubHeaders
    if (this.subHeaders.length > 0) {
      const subHeaderValues = this.subHeaders.map(sh => sh.value);
      ws.addRow(subHeaderValues);
      this.subHeaders.forEach((sh, idx) => {
        if (sh.styles) {
          ws.getRow(rowPointer).getCell(idx + 1).style = this.convertStyle(sh.styles);
        }
      });
      rowPointer++;
    }
    // Body (soporta children)
    for (const row of this.body) {
      rowPointer = this.addDataRowRecursive(ws, rowPointer, row);
    }
    // Footers
    if (this.footers.length > 0) {
      for (const footer of this.footers) {
        ws.addRow([footer.value]);
        if (footer.mergeCell && footer.mergeTo) {
          ws.mergeCells(rowPointer, 1, rowPointer, footer.mergeTo);
        }
        if (footer.styles) {
          ws.getRow(rowPointer).eachCell(cell => {
            cell.style = this.convertStyle(footer.styles);
          });
        }
        rowPointer++;
      }
    }
    this.isBuilt = true;
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
   * Agrega una fila de datos y sus children recursivamente
   * @returns el siguiente rowPointer disponible
   */
  private addDataRowRecursive(ws: ExcelJS.Worksheet, rowPointer: number, row: IDataCell, colPointer = 1): number {
    // Asegura que la fila exista
    const excelRow = ws.getRow(rowPointer);
    const cell = excelRow.getCell(colPointer);
    cell.value = row.value;
    if (row.styles) {
      cell.style = this.convertStyle(row.styles);
    }
    if (row.numberFormat) {
      cell.numFmt = row.numberFormat;
    }
    let maxRowPointer = rowPointer;
    // Si hay children, agregarlos en filas siguientes
    if (row.children && row.children.length > 0) {
      let childRowPointer = rowPointer;
      for (const child of row.children) {
        childRowPointer++;
        const usedRow = this.addDataRowRecursive(ws, childRowPointer, child, colPointer + 1);
        if (usedRow > maxRowPointer) maxRowPointer = usedRow;
      }
    }
    return maxRowPointer;
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
        fgColor: style.fill.fgColor,
        bgColor: style.fill.bgColor
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