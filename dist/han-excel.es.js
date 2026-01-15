import ExcelJS from "exceljs";
import saveAs from "file-saver";
class EventEmitter {
  listeners = /* @__PURE__ */ new Map();
  /**
   * Add an event listener
   */
  on(type, listener, options = {}) {
    if (!this.listeners.has(type)) {
      this.listeners.set(type, []);
    }
    const registration = {
      type,
      listener,
      options: {
        once: false,
        async: false,
        priority: 0,
        stopPropagation: false,
        ...options
      },
      id: this.generateId(),
      active: true,
      timestamp: /* @__PURE__ */ new Date()
    };
    this.listeners.get(type).push(registration);
    this.listeners.get(type).sort((a, b) => (b.options.priority || 0) - (a.options.priority || 0));
    return registration.id;
  }
  /**
   * Add a one-time event listener
   */
  once(type, listener, options = {}) {
    return this.on(type, listener, { ...options, once: true });
  }
  /**
   * Remove an event listener
   */
  off(type, listenerId) {
    const listeners = this.listeners.get(type);
    if (!listeners) {
      return false;
    }
    const index = listeners.findIndex((reg) => reg.id === listenerId);
    if (index === -1) {
      return false;
    }
    listeners.splice(index, 1);
    return true;
  }
  /**
   * Remove all listeners for an event type
   */
  offAll(type) {
    const listeners = this.listeners.get(type);
    if (!listeners) {
      return 0;
    }
    const count = listeners.length;
    this.listeners.delete(type);
    return count;
  }
  /**
   * Emit an event
   */
  async emit(event) {
    const type = event.type || "default";
    const listeners = this.listeners.get(type);
    if (!listeners || listeners.length === 0) {
      return;
    }
    const activeListeners = listeners.filter((reg) => reg.active);
    for (const registration of activeListeners) {
      try {
        if (registration.options.once) {
          registration.active = false;
        }
        if (registration.options.async) {
          await registration.listener(event);
        } else {
          registration.listener(event);
        }
        if (registration.options.stopPropagation) {
          break;
        }
      } catch (error) {
        console.error(`Error in event listener for ${type}:`, error);
      }
    }
    this.cleanupInactiveListeners(type);
  }
  /**
   * Emit an event synchronously
   */
  emitSync(event) {
    const type = event.type || "default";
    const listeners = this.listeners.get(type);
    if (!listeners || listeners.length === 0) {
      return;
    }
    const activeListeners = listeners.filter((reg) => reg.active);
    for (const registration of activeListeners) {
      try {
        if (registration.options.once) {
          registration.active = false;
        }
        registration.listener(event);
        if (registration.options.stopPropagation) {
          break;
        }
      } catch (error) {
        console.error(`Error in event listener for ${type}:`, error);
      }
    }
    this.cleanupInactiveListeners(type);
  }
  /**
   * Clear all listeners
   */
  clear() {
    this.listeners.clear();
  }
  /**
   * Get listeners for an event type
   */
  getListeners(type) {
    return this.listeners.get(type) || [];
  }
  /**
   * Get listener count for an event type
   */
  getListenerCount(type) {
    return this.listeners.get(type)?.length || 0;
  }
  /**
   * Get all registered event types
   */
  getEventTypes() {
    return Array.from(this.listeners.keys());
  }
  // Private methods
  generateId() {
    return Math.random().toString(36).substr(2, 9);
  }
  cleanupInactiveListeners(type) {
    const listeners = this.listeners.get(type);
    if (listeners) {
      const activeListeners = listeners.filter((reg) => reg.active);
      if (activeListeners.length !== listeners.length) {
        this.listeners.set(type, activeListeners);
      }
    }
  }
}
var CellType = /* @__PURE__ */ ((CellType2) => {
  CellType2["STRING"] = "string";
  CellType2["NUMBER"] = "number";
  CellType2["BOOLEAN"] = "boolean";
  CellType2["DATE"] = "date";
  CellType2["PERCENTAGE"] = "percentage";
  CellType2["CURRENCY"] = "currency";
  CellType2["LINK"] = "link";
  CellType2["FORMULA"] = "formula";
  return CellType2;
})(CellType || {});
var NumberFormat = /* @__PURE__ */ ((NumberFormat2) => {
  NumberFormat2["GENERAL"] = "General";
  NumberFormat2["NUMBER"] = "#,##0";
  NumberFormat2["NUMBER_DECIMALS"] = "#,##0.00";
  NumberFormat2["CURRENCY"] = "$#,##0.00";
  NumberFormat2["CURRENCY_INTEGER"] = "$#,##0";
  NumberFormat2["PERCENTAGE"] = "0%";
  NumberFormat2["PERCENTAGE_DECIMALS"] = "0.00%";
  NumberFormat2["DATE"] = "dd/mm/yyyy";
  NumberFormat2["DATE_TIME"] = "dd/mm/yyyy hh:mm";
  NumberFormat2["TIME"] = "hh:mm:ss";
  NumberFormat2["CUSTOM"] = "custom";
  return NumberFormat2;
})(NumberFormat || {});
var HorizontalAlignment = /* @__PURE__ */ ((HorizontalAlignment2) => {
  HorizontalAlignment2["LEFT"] = "left";
  HorizontalAlignment2["CENTER"] = "center";
  HorizontalAlignment2["RIGHT"] = "right";
  HorizontalAlignment2["FILL"] = "fill";
  HorizontalAlignment2["JUSTIFY"] = "justify";
  HorizontalAlignment2["CENTER_CONTINUOUS"] = "centerContinuous";
  HorizontalAlignment2["DISTRIBUTED"] = "distributed";
  return HorizontalAlignment2;
})(HorizontalAlignment || {});
var VerticalAlignment = /* @__PURE__ */ ((VerticalAlignment2) => {
  VerticalAlignment2["TOP"] = "top";
  VerticalAlignment2["MIDDLE"] = "middle";
  VerticalAlignment2["BOTTOM"] = "bottom";
  VerticalAlignment2["DISTRIBUTED"] = "distributed";
  VerticalAlignment2["JUSTIFY"] = "justify";
  return VerticalAlignment2;
})(VerticalAlignment || {});
var BorderStyle = /* @__PURE__ */ ((BorderStyle2) => {
  BorderStyle2["THIN"] = "thin";
  BorderStyle2["MEDIUM"] = "medium";
  BorderStyle2["THICK"] = "thick";
  BorderStyle2["DOTTED"] = "dotted";
  BorderStyle2["DASHED"] = "dashed";
  BorderStyle2["DOUBLE"] = "double";
  BorderStyle2["HAIR"] = "hair";
  BorderStyle2["MEDIUM_DASHED"] = "mediumDashed";
  BorderStyle2["DASH_DOT"] = "dashDot";
  BorderStyle2["MEDIUM_DASH_DOT"] = "mediumDashDot";
  BorderStyle2["DASH_DOT_DOT"] = "dashDotDot";
  BorderStyle2["MEDIUM_DASH_DOT_DOT"] = "mediumDashDotDot";
  BorderStyle2["SLANT_DASH_DOT"] = "slantDashDot";
  return BorderStyle2;
})(BorderStyle || {});
var FontStyle = /* @__PURE__ */ ((FontStyle2) => {
  FontStyle2["NORMAL"] = "normal";
  FontStyle2["BOLD"] = "bold";
  FontStyle2["ITALIC"] = "italic";
  FontStyle2["BOLD_ITALIC"] = "bold italic";
  return FontStyle2;
})(FontStyle || {});
var ErrorType = /* @__PURE__ */ ((ErrorType2) => {
  ErrorType2["VALIDATION_ERROR"] = "VALIDATION_ERROR";
  ErrorType2["BUILD_ERROR"] = "BUILD_ERROR";
  ErrorType2["STYLE_ERROR"] = "STYLE_ERROR";
  ErrorType2["WORKSHEET_ERROR"] = "WORKSHEET_ERROR";
  ErrorType2["CELL_ERROR"] = "CELL_ERROR";
  return ErrorType2;
})(ErrorType || {});
class Worksheet {
  config;
  tables = [];
  currentRow = 1;
  currentCol = 1;
  headerPointers = /* @__PURE__ */ new Map();
  isBuilt = false;
  // Estructuras temporales para la tabla actual
  headers = [];
  subHeaders = [];
  body = [];
  footers = [];
  // Features adicionales
  images = [];
  rowGroups = [];
  columnGroups = [];
  namedRanges = [];
  excelTables = [];
  hiddenRows = /* @__PURE__ */ new Set();
  hiddenColumns = /* @__PURE__ */ new Set();
  pivotTables = [];
  slicers = [];
  watermarks = [];
  dataConnections = [];
  // Estilos y tema del workbook (no se guardan en el objeto de ExcelJS)
  customStyles;
  theme;
  constructor(config) {
    this.config = config;
  }
  /**
   * Agrega un header principal
   */
  addHeader(header) {
    this.headers.push(header);
    return this;
  }
  /**
   * Agrega subheaders (ahora soporta anidación)
   */
  addSubHeaders(subHeaders) {
    this.subHeaders.push(...subHeaders);
    return this;
  }
  /**
   * Agrega una fila de datos (puede ser jerárquica con childrens)
   */
  addRow(row) {
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
  addFooter(footer) {
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
  addTable(tableConfig = {}) {
    const table = {
      name: tableConfig.name || `Table_${this.tables.length + 1}`,
      headers: tableConfig.headers || [],
      subHeaders: tableConfig.subHeaders || [],
      body: tableConfig.body || [],
      footers: tableConfig.footers || [],
      showBorders: tableConfig.showBorders !== false,
      showStripes: tableConfig.showStripes !== false,
      style: tableConfig.style || "TableStyleLight1",
      ...tableConfig
    };
    this.tables.push(table);
    return this;
  }
  /**
   * Finaliza la tabla actual agregando todos los elementos temporales a la última tabla
   */
  finalizeTable() {
    if (this.tables.length === 0) {
      this.addTable();
    }
    const currentTable = this.tables[this.tables.length - 1];
    if (!currentTable) {
      throw new Error("No se pudo obtener la tabla actual");
    }
    if (this.headers.length > 0) {
      currentTable.headers = [...currentTable.headers || [], ...this.headers];
    }
    if (this.subHeaders.length > 0) {
      currentTable.subHeaders = [...currentTable.subHeaders || [], ...this.subHeaders];
    }
    if (this.body.length > 0) {
      currentTable.body = [...currentTable.body || [], ...this.body];
    }
    if (this.footers.length > 0) {
      currentTable.footers = [...currentTable.footers || [], ...this.footers];
    }
    this.headers = [];
    this.subHeaders = [];
    this.body = [];
    this.footers = [];
    return this;
  }
  /**
   * Obtiene una tabla por nombre
   */
  getTable(name) {
    return this.tables.find((table) => table.name === name);
  }
  /**
   * Agrega una imagen al worksheet
   */
  addImage(image) {
    this.images.push(image);
    return this;
  }
  /**
   * Agrupa filas (crea esquema colapsable)
   */
  groupRows(startRow, endRow, collapsed = false) {
    this.rowGroups.push({ start: startRow, end: endRow, collapsed });
    return this;
  }
  /**
   * Agrupa columnas (crea esquema colapsable)
   */
  groupColumns(startCol, endCol, collapsed = false) {
    this.columnGroups.push({ start: startCol, end: endCol, collapsed });
    return this;
  }
  /**
   * Agrega un rango con nombre
   */
  addNamedRange(name, range, scope) {
    let rangeString;
    if (typeof range === "string") {
      rangeString = range;
    } else {
      const startRef = range.start.reference || `${this.numberToColumnLetter(range.start.col)}${range.start.row}`;
      const endRef = range.end.reference || `${this.numberToColumnLetter(range.end.col)}${range.end.row}`;
      rangeString = `${startRef}:${endRef}`;
    }
    const namedRange = { name, range: rangeString };
    if (scope !== void 0) {
      namedRange.scope = scope;
    }
    this.namedRanges.push(namedRange);
    return this;
  }
  /**
   * Agrega una tabla estructurada de Excel
   */
  addExcelTable(table) {
    this.excelTables.push(table);
    return this;
  }
  /**
   * Oculta filas
   */
  hideRows(rows) {
    const rowsArray = Array.isArray(rows) ? rows : [rows];
    rowsArray.forEach((row) => this.hiddenRows.add(row));
    return this;
  }
  /**
   * Muestra filas
   */
  showRows(rows) {
    const rowsArray = Array.isArray(rows) ? rows : [rows];
    rowsArray.forEach((row) => this.hiddenRows.delete(row));
    return this;
  }
  /**
   * Oculta columnas
   */
  hideColumns(columns) {
    const columnsArray = Array.isArray(columns) ? columns : [columns];
    columnsArray.forEach((col) => {
      const colNum = typeof col === "string" ? this.columnLetterToNumber(col) : col;
      this.hiddenColumns.add(colNum);
    });
    return this;
  }
  /**
   * Muestra columnas
   */
  showColumns(columns) {
    const columnsArray = Array.isArray(columns) ? columns : [columns];
    columnsArray.forEach((col) => {
      const colNum = typeof col === "string" ? this.columnLetterToNumber(col) : col;
      this.hiddenColumns.delete(colNum);
    });
    return this;
  }
  /**
   * Agrega una tabla dinámica (pivot table)
   */
  addPivotTable(pivotTable) {
    this.pivotTables.push(pivotTable);
    return this;
  }
  /**
   * Agrega un slicer a una tabla o tabla dinámica
   */
  addSlicer(slicer) {
    this.slicers.push(slicer);
    return this;
  }
  /**
   * Agrega una marca de agua al worksheet
   */
  addWatermark(watermark) {
    this.watermarks.push(watermark);
    return this;
  }
  /**
   * Agrega una conexión de datos
   */
  addDataConnection(connection) {
    this.dataConnections.push(connection);
    return this;
  }
  /**
   * Construye la hoja en el workbook de ExcelJS
   */
  async build(workbook, _options = {}) {
    const ws = workbook.addWorksheet(this.config.name, {
      properties: {
        defaultRowHeight: this.config.defaultRowHeight || 20,
        tabColor: this.config.tabColor
      },
      pageSetup: this.config.pageSetup
    });
    this.customStyles = workbook.__customStyles;
    this.theme = workbook.__theme;
    let rowPointer = 1;
    if (this.tables.length > 0) {
      let tableStartRow = rowPointer;
      for (let i = 0; i < this.tables.length; i++) {
        const table = this.tables[i];
        if (table) {
          tableStartRow = rowPointer;
          rowPointer = await this.buildTable(ws, table, rowPointer, i > 0);
          if (table.autoFilter && rowPointer > tableStartRow) {
            this.applyAutoFilter(ws, table, tableStartRow, rowPointer - 1);
          }
        }
      }
    } else {
      rowPointer = await this.buildLegacyContent(ws, rowPointer);
    }
    if (this.config.autoFilter?.enabled) {
      this.applyWorksheetAutoFilter(ws, rowPointer);
    }
    this.applyViews(ws);
    if (this.config.protected) {
      ws.protect(this.config.protectionPassword || "", {
        selectLockedCells: false,
        selectUnlockedCells: true,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertRows: false,
        insertHyperlinks: false,
        deleteColumns: false,
        deleteRows: false,
        sort: false,
        autoFilter: false,
        pivotTables: false
      });
    }
    for (const image of this.images) {
      await this.applyImage(ws, image);
    }
    for (const group of this.rowGroups) {
      this.applyRowGrouping(ws, group.start, group.end, group.collapsed);
    }
    for (const group of this.columnGroups) {
      this.applyColumnGrouping(ws, group.start, group.end, group.collapsed);
    }
    for (const namedRange of this.namedRanges) {
      workbook.definedNames.add(namedRange.name, namedRange.range);
    }
    for (const excelTable of this.excelTables) {
      this.applyExcelTable(ws, excelTable);
    }
    this.applyAdvancedPrintSettings(ws);
    this.applyHiddenRowsColumns(ws);
    for (const pivotTable of this.pivotTables) {
      await this.applyPivotTable(ws, pivotTable);
    }
    for (const slicer of this.slicers) {
      await this.applySlicer(ws, slicer);
    }
    for (const watermark of this.watermarks) {
      await this.applyWatermark(ws, watermark);
    }
    for (const connection of this.dataConnections) {
      await this.applyDataConnection(workbook, connection);
    }
    this.isBuilt = true;
  }
  /**
   * Construye una tabla individual en el worksheet
   */
  async buildTable(ws, table, startRow, addSpacing = false) {
    let rowPointer = startRow;
    if (addSpacing) {
      rowPointer += 2;
    }
    if (table.headers && table.headers.length > 0) {
      for (const header of table.headers) {
        const cell = ws.getRow(rowPointer).getCell(1);
        if (header.richText && header.richText.length > 0) {
          cell.value = {
            richText: header.richText.map((run) => ({
              text: run.text,
              font: run.font ? { name: run.font } : void 0,
              size: run.size,
              color: run.color ? this.convertColorToExcelJS(run.color) : void 0,
              bold: run.bold,
              italic: run.italic,
              underline: run.underline,
              strike: run.strikethrough
            })).filter((run) => run.text !== void 0)
          };
        } else {
          cell.value = this.processCellValue(header);
        }
        if (header.mergeCell) {
          const maxCols = this.calculateTableMaxColumns(table);
          ws.mergeCells(rowPointer, 1, rowPointer, maxCols);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell2) => {
            cell2.style = this.convertStyle(header.styles);
          });
        }
        if (header.cellProtection) {
          cell.protection = {
            locked: header.cellProtection.locked ?? true,
            hidden: header.cellProtection.hidden ?? false
          };
        } else if (header.protected !== void 0) {
          cell.protection = {
            locked: header.protected,
            hidden: false
          };
        }
        this.applyCellDimensions(ws, rowPointer, 1, header);
        if (header.comment) {
          this.applyCellComment(ws, rowPointer, 1, header.comment);
        }
        if (header.validation) {
          this.applyDataValidation(ws, rowPointer, 1, header.validation);
        }
        if (header.styles?.conditionalFormats) {
          this.applyConditionalFormatting(ws, rowPointer, 1, header.styles.conditionalFormats);
        }
        rowPointer++;
      }
    }
    if (table.subHeaders && table.subHeaders.length > 0) {
      rowPointer = this.buildNestedHeaders(ws, rowPointer, table.subHeaders);
    }
    if (table.body && table.body.length > 0) {
      for (const row of table.body) {
        rowPointer = this.addDataRowRecursive(ws, rowPointer, row);
      }
    }
    if (table.footers && table.footers.length > 0) {
      for (const footer of table.footers) {
        rowPointer = this.addFooterRow(ws, rowPointer, footer);
      }
    }
    if (table.showBorders || table.showStripes) {
      this.applyTableStyle(ws, table, startRow, rowPointer - 1);
    }
    return rowPointer;
  }
  /**
   * Construcción tradicional para compatibilidad hacia atrás
   */
  async buildLegacyContent(ws, startRow) {
    let rowPointer = startRow;
    if (this.headers.length > 0) {
      this.headers.forEach((header) => {
        ws.addRow([this.processCellValue(header)]);
        if (header.mergeCell) {
          ws.mergeCells(rowPointer, 1, rowPointer, this.getMaxColumns() || 1);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
        this.applyCellDimensions(ws, rowPointer, 1, header);
        if (header.comment) {
          this.applyCellComment(ws, rowPointer, 1, header.comment);
        }
        if (header.validation) {
          this.applyDataValidation(ws, rowPointer, 1, header.validation);
        }
        if (header.styles?.conditionalFormats) {
          this.applyConditionalFormatting(ws, rowPointer, 1, header.styles.conditionalFormats);
        }
        rowPointer++;
      });
    }
    if (this.subHeaders.length > 0) {
      rowPointer = this.buildNestedHeaders(ws, rowPointer, this.subHeaders);
    }
    for (const row of this.body) {
      rowPointer = this.addDataRowRecursive(ws, rowPointer, row);
    }
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
  calculateTableMaxColumns(table) {
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
  applyTableStyle(ws, table, startRow, endRow) {
    const maxCols = this.calculateTableMaxColumns(table);
    if (table.showBorders) {
      for (let row = startRow; row <= endRow; row++) {
        for (let col = 1; col <= maxCols; col++) {
          const cell = ws.getRow(row).getCell(col);
          if (!cell.style)
            cell.style = {};
          if (!cell.style.border) {
            cell.style.border = {
              top: { style: "thin", color: { argb: "FF8EAADB" } },
              left: { style: "thin", color: { argb: "FF8EAADB" } },
              bottom: { style: "thin", color: { argb: "FF8EAADB" } },
              right: { style: "thin", color: { argb: "FF8EAADB" } }
            };
          }
        }
      }
    }
    if (table.showStripes) {
      for (let row = startRow; row <= endRow; row++) {
        if ((row - startRow) % 2 === 1) {
          for (let col = 1; col <= maxCols; col++) {
            const cell = ws.getRow(row).getCell(col);
            if (!cell.style)
              cell.style = {};
            if (!cell.style.fill) {
              cell.style.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FFF2F2F2" }
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
  buildNestedHeaders(ws, startRow, headers) {
    let currentRow = startRow;
    const maxDepth = this.getMaxHeaderDepth(headers);
    for (let depth = 0; depth < maxDepth; depth++) {
      const row = ws.getRow(currentRow);
      let colIndex = 1;
      for (const header of headers) {
        if (depth === 0) {
          const headerInfo = this.getHeaderAtDepth(header, depth, colIndex);
          const cell = row.getCell(colIndex);
          cell.value = this.processCellValue(header);
          if (headerInfo.style) {
            cell.style = this.convertStyle(headerInfo.style);
          }
          this.applyCellDimensions(ws, currentRow, colIndex, header);
          if (header.comment) {
            this.applyCellComment(ws, currentRow, colIndex, header.comment);
          }
          if (header.validation) {
            this.applyDataValidation(ws, currentRow, colIndex, header.validation);
          }
          if (header.styles?.conditionalFormats) {
            this.applyConditionalFormatting(ws, currentRow, colIndex, header.styles.conditionalFormats);
          }
          colIndex += headerInfo.colSpan;
        } else {
          if (header.children && header.children.length > 0) {
            for (const child of header.children) {
              const cell = row.getCell(colIndex);
              cell.value = this.processCellValue(child);
              if (child.styles || header.styles) {
                cell.style = this.convertStyle(child.styles || header.styles);
              }
              this.applyCellDimensions(ws, currentRow, colIndex, child);
              if (child.comment) {
                this.applyCellComment(ws, currentRow, colIndex, child.comment);
              }
              if (child.validation) {
                this.applyDataValidation(ws, currentRow, colIndex, child.validation);
              }
              if (child.styles?.conditionalFormats) {
                this.applyConditionalFormatting(ws, currentRow, colIndex, child.styles.conditionalFormats);
              }
              colIndex += this.calculateHeaderColSpan(child);
            }
          } else {
            const cell = row.getCell(colIndex);
            cell.value = null;
            colIndex += 1;
          }
        }
      }
      currentRow++;
    }
    this.applyAllMerges(ws, startRow, currentRow - 1, headers);
    return currentRow;
  }
  /**
   * Obtiene información del header en una profundidad específica
   */
  getHeaderAtDepth(header, depth, startCol) {
    const colSpan = this.calculateHeaderColSpan(header);
    if (depth === 0) {
      const mergeRange = colSpan > 1 ? { start: startCol, end: startCol + colSpan - 1 } : null;
      return {
        value: typeof header.value === "string" ? header.value : String(header.value || ""),
        style: header.styles,
        colSpan,
        mergeRange
      };
    } else if (header.children && header.children.length > 0) {
      const child = header.children[depth];
      if (child) {
        const childColSpan = this.calculateHeaderColSpan(child);
        const mergeRange = childColSpan > 1 ? { start: startCol, end: startCol + childColSpan - 1 } : null;
        return {
          value: typeof child.value === "string" ? child.value : String(child.value || ""),
          style: child.styles || header.styles,
          colSpan: childColSpan,
          mergeRange
        };
      }
    }
    return {
      value: null,
      style: null,
      colSpan: 1
    };
  }
  /**
   * Aplica todos los merges (horizontales y verticales) después de crear todas las filas
   */
  applyAllMerges(ws, startRow, endRow, headers) {
    const maxDepth = this.getMaxHeaderDepth(headers);
    if (maxDepth <= 1)
      return;
    this.applySmartMerges(ws, startRow, endRow, headers);
  }
  /**
   * Aplica merges inteligentes basados en la estructura de headers
   */
  applySmartMerges(ws, startRow, endRow, headers) {
    const maxDepth = this.getMaxHeaderDepth(headers);
    if (maxDepth <= 1)
      return;
    let colIndex = 1;
    for (const header of headers) {
      this.applySmartMergesForHeader(ws, startRow, endRow, header, colIndex);
      colIndex += this.calculateHeaderColSpan(header);
    }
  }
  /**
   * Aplica merges inteligentes para un header específico
   */
  applySmartMergesForHeader(ws, startRow, endRow, header, startCol) {
    const headerColSpan = this.calculateHeaderColSpan(header);
    if (!header.children || header.children.length === 0) {
      ws.mergeCells(startRow, startCol, endRow, startCol + headerColSpan - 1);
    } else {
      if (headerColSpan > 1) {
        ws.mergeCells(startRow, startCol, startRow, startCol + headerColSpan - 1);
      }
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
  calculateHeaderColSpan(header) {
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
  getMaxHeaderDepth(headers) {
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
  getMaxColumns() {
    let maxCols = 0;
    for (const header of this.subHeaders) {
      maxCols += this.calculateHeaderColSpan(header);
    }
    return maxCols;
  }
  /**
   * Valida la hoja
   */
  validate() {
    if (!this.headers.length && !this.body.length) {
      return {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: "La hoja no tiene datos"
        }
      };
    }
    return { success: true, data: true };
  }
  /**
   * Calcula las posiciones de columnas para los datos basándose en la estructura de subheaders
   */
  calculateDataColumnPositions() {
    const positions = {};
    let currentCol = 1;
    for (const header of this.subHeaders) {
      if (header.children && header.children.length > 0) {
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
  addFooterRow(ws, rowPointer, footer) {
    const columnPositions = this.calculateDataColumnPositions();
    let footerColPosition;
    if (footer.key && columnPositions[footer.key]) {
      footerColPosition = columnPositions[footer.key];
    } else if (footer.header && columnPositions[footer.header]) {
      footerColPosition = columnPositions[footer.header];
    }
    if (footerColPosition === void 0) {
      footerColPosition = 1;
    }
    const excelRow = ws.getRow(rowPointer);
    const footerCell = excelRow.getCell(footerColPosition);
    footerCell.value = this.processCellValue(footer);
    if (footer.styles) {
      footerCell.style = this.convertStyle(footer.styles);
    } else if (footer.styleName) {
      const style = this.getPredefinedStyle(footer.styleName);
      if (style) {
        footerCell.style = this.convertStyle(style);
      }
    } else {
      const themeStyle = this.getThemeStyle("footer");
      if (themeStyle) {
        footerCell.style = this.convertStyle(themeStyle);
      }
    }
    if (footer.numberFormat) {
      footerCell.numFmt = footer.numberFormat;
    }
    this.applyCellDimensions(ws, rowPointer, footerColPosition, footer);
    if (footer.comment) {
      this.applyCellComment(ws, rowPointer, footerColPosition, footer.comment);
    }
    if (footer.validation) {
      this.applyDataValidation(ws, rowPointer, footerColPosition, footer.validation);
    }
    if (footer.styles?.conditionalFormats) {
      this.applyConditionalFormatting(ws, rowPointer, footerColPosition, footer.styles.conditionalFormats);
    }
    if (footer.mergeCell && footer.mergeTo) {
      ws.mergeCells(rowPointer, footerColPosition, rowPointer, footer.mergeTo);
    }
    if (footer.children && footer.children.length > 0) {
      for (const child of footer.children) {
        if (child) {
          let colPosition;
          if (child.key && columnPositions[child.key]) {
            colPosition = columnPositions[child.key];
          } else if (child.header && columnPositions[child.header]) {
            colPosition = columnPositions[child.header];
          }
          if (colPosition !== void 0) {
            const childCell = excelRow.getCell(colPosition);
            childCell.value = this.processCellValue(child);
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
            this.applyCellDimensions(ws, rowPointer, colPosition, child);
            if (child.comment) {
              this.applyCellComment(ws, rowPointer, colPosition, child.comment);
            }
            if (child.validation) {
              this.applyDataValidation(ws, rowPointer, colPosition, child.validation);
            }
            if (child.styles?.conditionalFormats) {
              this.applyConditionalFormatting(ws, rowPointer, colPosition, child.styles.conditionalFormats);
            }
          }
        }
      }
    }
    if (footer.jump) {
      return rowPointer + 1;
    }
    return rowPointer;
  }
  /**
   * Aplica width y height a una celda/fila
   */
  applyCellDimensions(ws, row, col, cell) {
    if (cell.rowHeight !== void 0) {
      const excelRow = ws.getRow(row);
      excelRow.height = cell.rowHeight;
    }
    if (cell.colWidth !== void 0) {
      const excelCol = ws.getColumn(col);
      excelCol.width = cell.colWidth;
    }
  }
  /**
   * Aplica comentario a una celda
   */
  applyCellComment(ws, row, col, comment) {
    if (!comment || comment.trim() === "") {
      return;
    }
    const cell = ws.getRow(row).getCell(col);
    if (typeof comment === "string") {
      cell.note = comment;
    }
  }
  /**
   * Aplica validación de datos a una celda
   */
  applyDataValidation(ws, row, col, validation) {
    if (!validation) {
      return;
    }
    const cell = ws.getRow(row).getCell(col);
    const validationType = validation.type === "time" ? "date" : validation.type;
    const dataValidation = {
      type: validationType,
      allowBlank: validation.allowBlank ?? true,
      formulae: []
      // Inicializar como array vacío, se llenará si hay fórmulas
    };
    if (validation.operator) {
      dataValidation.operator = validation.operator;
    }
    if (validation.formula1 !== void 0) {
      if (typeof validation.formula1 === "string") {
        dataValidation.formulae = [validation.formula1];
      } else if (validation.formula1 instanceof Date) {
        dataValidation.formulae = [validation.formula1.toISOString()];
      } else {
        dataValidation.formulae = [validation.formula1];
      }
    }
    if (validation.formula2 !== void 0) {
      if (!dataValidation.formulae) {
        dataValidation.formulae = [];
      }
      if (typeof validation.formula2 === "string") {
        dataValidation.formulae.push(validation.formula2);
      } else if (validation.formula2 instanceof Date) {
        dataValidation.formulae.push(validation.formula2.toISOString());
      } else {
        dataValidation.formulae.push(validation.formula2);
      }
    }
    if (validation.showErrorMessage) {
      dataValidation.showErrorMessage = true;
      if (validation.errorMessage) {
        dataValidation.error = validation.errorMessage;
      }
    }
    if (validation.showInputMessage) {
      dataValidation.showInputMessage = true;
      if (validation.inputMessage) {
        dataValidation.prompt = validation.inputMessage;
      }
    }
    cell.dataValidation = dataValidation;
  }
  /**
   * Aplica formato condicional a una celda
   */
  applyConditionalFormatting(ws, row, col, conditionalFormats) {
    if (!conditionalFormats || conditionalFormats.length === 0) {
      return;
    }
    const cell = ws.getRow(row).getCell(col);
    const cellAddress = cell.address;
    conditionalFormats.forEach((format, index) => {
      const rule = {
        type: format.type,
        priority: format.priority ?? index + 1,
        stopIfTrue: format.stopIfTrue ?? false
      };
      if (format.operator) {
        rule.operator = format.operator;
      }
      if (format.formula) {
        rule.formulae = [format.formula];
      } else if (format.values && format.values.length > 0) {
        rule.formulae = format.values.map((v) => {
          if (typeof v === "string") {
            return v;
          } else if (v instanceof Date) {
            return v.toISOString();
          } else {
            return String(v);
          }
        });
      }
      if (format.style) {
        const style = this.convertStyle(format.style);
        rule.style = style;
      }
      ws.addConditionalFormatting({
        ref: cellAddress,
        rules: [rule]
      });
    });
  }
  /**
   * Aplica filtro automático a una tabla
   */
  applyAutoFilter(ws, table, startRow, endRow) {
    if (!table.autoFilter) {
      return;
    }
    const maxCols = this.calculateTableMaxColumns(table);
    const headerRow = startRow;
    const dataEndRow = endRow;
    if (maxCols > 0 && dataEndRow >= headerRow) {
      ws.autoFilter = {
        from: {
          row: headerRow,
          column: 1
        },
        to: {
          row: dataEndRow,
          column: maxCols
        }
      };
    }
  }
  /**
   * Aplica filtro automático a nivel de worksheet
   */
  applyWorksheetAutoFilter(ws, lastRow) {
    const autoFilterConfig = this.config.autoFilter;
    if (!autoFilterConfig || !autoFilterConfig.enabled) {
      return;
    }
    if (autoFilterConfig.range) {
      ws.autoFilter = {
        from: {
          row: autoFilterConfig.range.start?.row || 1,
          column: autoFilterConfig.range.start?.col || 1
        },
        to: {
          row: autoFilterConfig.range.end?.row || lastRow,
          column: autoFilterConfig.range.end?.col || ws.columnCount || 1
        }
      };
      return;
    }
    if (autoFilterConfig.startRow !== void 0 || autoFilterConfig.endRow !== void 0 || autoFilterConfig.startColumn !== void 0 || autoFilterConfig.endColumn !== void 0) {
      ws.autoFilter = {
        from: {
          row: autoFilterConfig.startRow || 1,
          column: autoFilterConfig.startColumn || 1
        },
        to: {
          row: autoFilterConfig.endRow || lastRow,
          column: autoFilterConfig.endColumn || ws.columnCount || 1
        }
      };
      return;
    }
    const startRow = this.headers.length > 0 ? this.headers.length + (this.subHeaders.length > 0 ? this.getMaxHeaderDepth(this.subHeaders) : 0) : 1;
    const maxCols = this.getMaxColumns() || ws.columnCount || 1;
    if (lastRow >= startRow && maxCols > 0) {
      ws.autoFilter = {
        from: {
          row: startRow,
          column: 1
        },
        to: {
          row: lastRow,
          column: maxCols
        }
      };
    }
  }
  /**
   * Procesa el valor de una celda considerando links y máscaras
   * Si el tipo es LINK o hay un link, crea un hipervínculo en Excel
   */
  processCellValue(cell) {
    if (cell.link || cell.type === CellType.LINK) {
      const linkUrl = cell.link || (typeof cell.value === "string" ? cell.value : "");
      if (!linkUrl || linkUrl.trim() === "") {
        return cell.value;
      }
      const displayText = cell.mask || cell.value || linkUrl;
      return {
        text: String(displayText),
        hyperlink: linkUrl
      };
    }
    return cell.value;
  }
  /**
   * Agrega una fila de datos y sus children recursivamente
   * @returns el siguiente rowPointer disponible
   */
  addDataRowRecursive(ws, rowPointer, row) {
    const columnPositions = this.calculateDataColumnPositions();
    let mainColPosition;
    if (row.key && columnPositions[row.key]) {
      mainColPosition = columnPositions[row.key];
    } else if (row.header && columnPositions[row.header]) {
      mainColPosition = columnPositions[row.header];
    }
    if (mainColPosition === void 0) {
      mainColPosition = 1;
    }
    const excelRow = ws.getRow(rowPointer);
    const mainCell = excelRow.getCell(mainColPosition);
    if (row.richText && row.richText.length > 0) {
      mainCell.value = {
        richText: row.richText.map((run) => ({
          text: run.text,
          font: run.font ? { name: run.font } : void 0,
          size: run.size,
          color: run.color ? this.convertColorToExcelJS(run.color) : void 0,
          bold: run.bold,
          italic: run.italic,
          underline: run.underline,
          strike: run.strikethrough
        })).filter((run) => run.text !== void 0)
      };
    } else {
      mainCell.value = this.processCellValue(row);
    }
    if (row.styles) {
      mainCell.style = this.convertStyle(row.styles);
    } else if (row.styleName) {
      const style = this.getPredefinedStyle(row.styleName);
      if (style) {
        mainCell.style = this.convertStyle(style);
      }
    } else {
      const themeStyle = this.getThemeStyle("body", rowPointer);
      if (themeStyle) {
        mainCell.style = this.convertStyle(themeStyle);
      }
    }
    if (row.numberFormat) {
      mainCell.numFmt = row.numberFormat;
    }
    if (row.cellProtection) {
      mainCell.protection = {
        locked: row.cellProtection.locked ?? true,
        hidden: row.cellProtection.hidden ?? false
      };
    } else if (row.protected !== void 0) {
      mainCell.protection = {
        locked: row.protected,
        hidden: false
      };
    }
    this.applyCellDimensions(ws, rowPointer, mainColPosition, row);
    if (row.comment) {
      this.applyCellComment(ws, rowPointer, mainColPosition, row.comment);
    }
    if (row.validation) {
      this.applyDataValidation(ws, rowPointer, mainColPosition, row.validation);
    }
    if (row.styles?.conditionalFormats) {
      this.applyConditionalFormatting(ws, rowPointer, mainColPosition, row.styles.conditionalFormats);
    }
    if (row.children && row.children.length > 0) {
      for (const child of row.children) {
        if (child) {
          let colPosition;
          if (child.key && columnPositions[child.key]) {
            colPosition = columnPositions[child.key];
          } else if (child.header && columnPositions[child.header]) {
            colPosition = columnPositions[child.header];
          }
          if (colPosition !== void 0) {
            const childCell = excelRow.getCell(colPosition);
            childCell.value = this.processCellValue(child);
            if (child.styles) {
              childCell.style = this.convertStyle(child.styles);
            }
            if (child.numberFormat) {
              childCell.numFmt = child.numberFormat;
            }
            this.applyCellDimensions(ws, rowPointer, colPosition, child);
            if (child.comment) {
              this.applyCellComment(ws, rowPointer, colPosition, child.comment);
            }
            if (child.validation) {
              this.applyDataValidation(ws, rowPointer, colPosition, child.validation);
            }
            if (child.styles?.conditionalFormats) {
              this.applyConditionalFormatting(ws, rowPointer, colPosition, child.styles.conditionalFormats);
            }
          }
        }
      }
    }
    if (row.jump) {
      return rowPointer + 1;
    }
    return rowPointer;
  }
  /**
   * Convierte un color a formato ExcelJS (ARGB)
   */
  convertColor(color) {
    if (!color)
      return void 0;
    if (typeof color === "object" && color.argb) {
      return color;
    }
    if (typeof color === "object" && "r" in color && "g" in color && "b" in color) {
      const r = color.r.toString(16).padStart(2, "0");
      const g = color.g.toString(16).padStart(2, "0");
      const b = color.b.toString(16).padStart(2, "0");
      return { argb: `FF${r}${g}${b}`.toUpperCase() };
    }
    if (typeof color === "string") {
      let hex = color.replace("#", "");
      if (hex.length === 3) {
        hex = hex.split("").map((c) => c + c).join("");
      }
      if (hex.length === 6) {
        hex = "FF" + hex.toUpperCase();
      }
      return { argb: hex };
    }
    if (typeof color === "object" && "theme" in color) {
      return color;
    }
    return void 0;
  }
  /**
   * Convierte el estilo personalizado a formato compatible con ExcelJS
   */
  convertStyle(style) {
    if (!style)
      return {};
    const converted = {};
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
      const pattern = style.fill.pattern || "solid";
      const fgColor = pattern === "solid" ? style.fill.backgroundColor || style.fill.foregroundColor : style.fill.foregroundColor || style.fill.backgroundColor;
      const bgColor = pattern !== "solid" ? style.fill.backgroundColor : void 0;
      converted.fill = {
        type: style.fill.type || "pattern",
        pattern,
        fgColor: this.convertColor(fgColor),
        bgColor: bgColor ? this.convertColor(bgColor) : void 0
      };
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
      if (style.alignment.horizontal !== void 0) {
        const validHorizontal = ["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
        if (validHorizontal.includes(style.alignment.horizontal)) {
          converted.alignment.horizontal = style.alignment.horizontal;
        }
      }
      if (style.alignment.vertical !== void 0) {
        const validVertical = ["top", "middle", "bottom", "distributed", "justify"];
        if (validVertical.includes(style.alignment.vertical)) {
          converted.alignment.vertical = style.alignment.vertical;
        }
      }
      if (style.alignment.wrapText !== void 0) {
        converted.alignment.wrapText = Boolean(style.alignment.wrapText);
      }
      if (style.alignment.shrinkToFit !== void 0) {
        converted.alignment.shrinkToFit = Boolean(style.alignment.shrinkToFit);
      }
      if (style.alignment.indent !== void 0 && typeof style.alignment.indent === "number") {
        converted.alignment.indent = style.alignment.indent;
      }
      if (style.alignment.textRotation !== void 0 && typeof style.alignment.textRotation === "number") {
        converted.alignment.textRotation = style.alignment.textRotation;
      }
      if (style.alignment.readingOrder !== void 0) {
        const validReadingOrder = ["left-to-right", "right-to-left", "context"];
        if (validReadingOrder.includes(style.alignment.readingOrder)) {
          converted.alignment.readingOrder = style.alignment.readingOrder;
        }
      }
      if (Object.keys(converted.alignment).length === 0) {
        delete converted.alignment;
      }
    }
    if (style.numFmt) {
      converted.numFmt = style.numFmt;
    }
    return converted;
  }
  /**
   * Convierte un número de columna a letra (1 = A, 2 = B, etc.)
   */
  numberToColumnLetter(columnNumber) {
    let result = "";
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + columnNumber % 26) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }
  /**
   * Convierte letra de columna a número (A = 1, B = 2, etc.)
   */
  columnLetterToNumber(columnLetter) {
    let result = 0;
    for (let i = 0; i < columnLetter.length; i++) {
      result = result * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return result;
  }
  /**
   * Aplica una imagen al worksheet
   */
  async applyImage(ws, image) {
    try {
      let row;
      let col;
      if (typeof image.position.row === "string") {
        const match = image.position.row.match(/([A-Z]+)(\d+)/);
        if (match && match[1] && match[2]) {
          col = this.columnLetterToNumber(match[1]);
          row = parseInt(match[2], 10);
        } else {
          row = parseInt(image.position.row, 10) || 1;
          col = typeof image.position.col === "string" ? this.columnLetterToNumber(image.position.col) : typeof image.position.col === "number" ? image.position.col : 1;
        }
      } else {
        row = image.position.row;
        col = typeof image.position.col === "string" ? this.columnLetterToNumber(image.position.col) : typeof image.position.col === "number" ? image.position.col : 1;
      }
      let imageBuffer;
      if (typeof image.buffer === "string") {
        let base64Data;
        if (image.buffer.startsWith("data:")) {
          const parts = image.buffer.split(",");
          base64Data = parts[1] || image.buffer;
        } else {
          base64Data = image.buffer;
        }
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        imageBuffer = bytes;
      } else if (image.buffer instanceof ArrayBuffer) {
        imageBuffer = new Uint8Array(image.buffer);
      } else {
        imageBuffer = image.buffer;
      }
      const imagePosition = {
        tl: { col: col - 1, row: row - 1 }
      };
      if (image.size) {
        if (image.size.width && image.size.height) {
          imagePosition.ext = {
            width: image.size.width,
            height: image.size.height
          };
        } else if (image.size.scaleX && image.size.scaleY) {
          imagePosition.ext = {
            width: 100 * (image.size.scaleX || 1),
            height: 100 * (image.size.scaleY || 1)
          };
        }
      }
      ws.addImage({
        buffer: imageBuffer,
        extension: image.extension
      }, imagePosition);
      if (image.hyperlink) {
        const cell = ws.getRow(row).getCell(col);
        cell.value = {
          text: image.description || "",
          hyperlink: image.hyperlink
        };
      }
    } catch (error) {
      console.warn("Error adding image to worksheet:", error);
    }
  }
  /**
   * Aplica agrupación de filas
   */
  applyRowGrouping(ws, startRow, endRow, collapsed = false) {
    for (let row = startRow; row <= endRow; row++) {
      const excelRow = ws.getRow(row);
      if (!excelRow.outlineLevel) {
        excelRow.outlineLevel = 1;
      }
      if (collapsed && row === startRow) {
        try {
          excelRow.collapsed = true;
        } catch {
        }
      }
    }
  }
  /**
   * Aplica agrupación de columnas
   */
  applyColumnGrouping(ws, startCol, endCol, collapsed = false) {
    for (let col = startCol; col <= endCol; col++) {
      const excelCol = ws.getColumn(col);
      if (!excelCol.outlineLevel) {
        excelCol.outlineLevel = 1;
      }
      if (collapsed && col === startCol) {
        try {
          excelCol.collapsed = true;
        } catch {
        }
      }
    }
  }
  /**
   * Aplica una tabla estructurada de Excel
   */
  applyExcelTable(ws, table) {
    try {
      const range = `${table.range.start}:${table.range.end}`;
      const tableConfig = {
        name: table.name,
        ref: range,
        headerRow: table.headerRow !== false,
        totalsRow: table.totalRow === true
      };
      if (table.style) {
        tableConfig.style = {
          theme: table.style,
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false
        };
      }
      if (table.columns && table.columns.length > 0) {
        tableConfig.columns = table.columns.map((col) => ({
          name: col.name,
          filterButton: col.filterButton !== false,
          totalsRowFunction: col.totalsRowFunction || "none",
          totalsRowFormula: col.totalsRowFormula
        }));
      }
      ws.addTable(tableConfig);
    } catch (error) {
      console.warn("Error adding Excel table:", error);
    }
  }
  /**
   * Aplica configuración avanzada de impresión
   */
  applyAdvancedPrintSettings(ws) {
    if (this.config.printHeadersFooters) {
      const headerFooter = {};
      if (this.config.printHeadersFooters.header) {
        const left = this.config.printHeadersFooters.header.left || "";
        const center = this.config.printHeadersFooters.header.center || "";
        const right = this.config.printHeadersFooters.header.right || "";
        headerFooter.oddHeader = `${left}&C${center}&R${right}`;
      }
      if (this.config.printHeadersFooters.footer) {
        const left = this.config.printHeadersFooters.footer.left || "";
        const center = this.config.printHeadersFooters.footer.center || "";
        const right = this.config.printHeadersFooters.footer.right || "";
        headerFooter.oddFooter = `${left}&C${center}&R${right}`;
      }
      if (Object.keys(headerFooter).length > 0) {
        ws.headerFooter = headerFooter;
      }
    }
    if (this.config.printRepeat) {
      if (this.config.printRepeat.rows) {
        if (Array.isArray(this.config.printRepeat.rows)) {
          const rowsStr = this.config.printRepeat.rows.map((r) => r.toString()).join(":");
          ws.pageSetup.printTitlesRow = `$${rowsStr}`;
        } else {
          ws.pageSetup.printTitlesRow = `$${this.config.printRepeat.rows}`;
        }
      }
      if (this.config.printRepeat.columns) {
        if (Array.isArray(this.config.printRepeat.columns)) {
          const colsStr = this.config.printRepeat.columns.map((c) => typeof c === "number" ? this.numberToColumnLetter(c) : c).join(":");
          ws.pageSetup.printTitlesColumn = `$${colsStr}`;
        } else {
          ws.pageSetup.printTitlesColumn = `$${this.config.printRepeat.columns}`;
        }
      }
    }
  }
  /**
   * Aplica filas y columnas ocultas
   */
  applyHiddenRowsColumns(ws) {
    for (const rowNum of this.hiddenRows) {
      const row = ws.getRow(rowNum);
      row.hidden = true;
    }
    for (const colNum of this.hiddenColumns) {
      const column = ws.getColumn(colNum);
      column.hidden = true;
    }
  }
  /**
   * Aplica una tabla dinámica (pivot table)
   */
  async applyPivotTable(ws, pivotTable) {
    try {
      if (pivotTable.sourceSheet) {
        const workbook = ws.workbook;
        const sourceSheet = workbook.getWorksheet(pivotTable.sourceSheet);
        if (!sourceSheet) {
          console.warn(`Source sheet "${pivotTable.sourceSheet}" not found for pivot table "${pivotTable.name}"`);
          return;
        }
      }
      const pivotConfig = {
        name: pivotTable.name,
        ref: pivotTable.ref,
        sourceRange: pivotTable.sourceRange,
        fields: {}
      };
      if (pivotTable.fields.rows && pivotTable.fields.rows.length > 0) {
        pivotConfig.fields.rows = pivotTable.fields.rows;
      }
      if (pivotTable.fields.columns && pivotTable.fields.columns.length > 0) {
        pivotConfig.fields.columns = pivotTable.fields.columns;
      }
      if (pivotTable.fields.values && pivotTable.fields.values.length > 0) {
        pivotConfig.fields.values = pivotTable.fields.values.map((v) => ({
          name: v.name,
          stat: v.stat
        }));
      }
      if (pivotTable.fields.filters && pivotTable.fields.filters.length > 0) {
        pivotConfig.fields.filters = pivotTable.fields.filters;
      }
      if (pivotTable.options) {
        pivotConfig.options = pivotTable.options;
      }
      if (ws.addPivotTable) {
        ws.addPivotTable(pivotConfig);
      } else {
        console.warn("Pivot tables require ExcelJS 4.5.0+. Feature may not be fully supported.");
      }
    } catch (error) {
      console.warn("Error adding pivot table:", error);
    }
  }
  /**
   * Convierte un color a formato ExcelJS
   */
  convertColorToExcelJS(color) {
    if (typeof color === "string") {
      if (color.startsWith("#")) {
        const hex = color.substring(1);
        return { argb: `FF${hex.toUpperCase()}` };
      }
      return { argb: "FF000000" };
    } else if ("r" in color && "g" in color && "b" in color) {
      const hex = [color.r, color.g, color.b].map((x) => {
        const hex2 = x.toString(16);
        return hex2.length === 1 ? "0" + hex2 : hex2;
      }).join("").toUpperCase();
      return { argb: `FF${hex}` };
    } else if ("theme" in color) {
      return { theme: color.theme };
    }
    return { argb: "FF000000" };
  }
  /**
   * Aplica views (freeze panes, split panes, sheet views)
   */
  applyViews(ws) {
    const views = [];
    if (this.config.freezePanes) {
      const freezeView = {
        state: "frozen",
        xSplit: this.config.freezePanes.col - 1,
        ySplit: this.config.freezePanes.row - 1,
        topLeftCell: this.config.freezePanes.reference || this.numberToColumnLetter(this.config.freezePanes.col) + String(this.config.freezePanes.row),
        activeCell: this.config.freezePanes.reference || this.numberToColumnLetter(this.config.freezePanes.col) + String(this.config.freezePanes.row)
      };
      views.push(freezeView);
    } else if (this.config.splitPanes) {
      const splitConfig = this.config.splitPanes;
      const splitView = {
        state: "split",
        xSplit: splitConfig.xSplit || 0,
        ySplit: splitConfig.ySplit || 0
      };
      if (splitConfig.topLeftCell) {
        splitView.topLeftCell = splitConfig.topLeftCell;
      }
      if (splitConfig.activePane) {
        const paneMap = {
          "topLeft": "topLeft",
          "topRight": "topRight",
          "bottomLeft": "bottomLeft",
          "bottomRight": "bottomRight"
        };
        splitView.activePane = paneMap[splitConfig.activePane] || "topLeft";
      }
      views.push(splitView);
    } else if (this.config.views) {
      const viewConfig = this.config.views;
      const view = {
        state: viewConfig.state === "pageBreakPreview" || viewConfig.state === "pageLayout" ? "normal" : viewConfig.state || "normal"
      };
      if (viewConfig.zoomScale !== void 0) {
        view.zoomScale = viewConfig.zoomScale;
      }
      if (viewConfig.zoomScaleNormal !== void 0) {
        view.zoomScaleNormal = viewConfig.zoomScaleNormal;
      }
      if (viewConfig.showGridLines !== void 0) {
        view.showGridLines = viewConfig.showGridLines;
      }
      if (viewConfig.showRowColHeaders !== void 0) {
        view.showRowColHeaders = viewConfig.showRowColHeaders;
      }
      if (viewConfig.showRuler !== void 0) {
        view.showRuler = viewConfig.showRuler;
      }
      if (viewConfig.rightToLeft !== void 0) {
        view.rightToLeft = viewConfig.rightToLeft;
      }
      views.push(view);
    } else if (this.config.zoom) {
      views.push({
        state: "normal",
        zoomScale: this.config.zoom
      });
    }
    if (views.length > 0) {
      ws.views = views;
    }
  }
  /**
   * Obtiene un estilo predefinido del workbook
   */
  getPredefinedStyle(styleName) {
    if (this.customStyles && this.customStyles[styleName]) {
      return this.customStyles[styleName];
    }
    return void 0;
  }
  /**
   * Obtiene un estilo del tema para una sección específica
   */
  getThemeStyle(section, rowIndex) {
    if (!this.theme || this.theme.autoApplySectionStyles === false) {
      return void 0;
    }
    if (!this.customStyles) {
      return void 0;
    }
    let styleName = "";
    if (section === "header") {
      styleName = "__theme_header";
    } else if (section === "subHeader") {
      styleName = "__theme_subHeader";
    } else if (section === "body") {
      if (rowIndex !== void 0 && rowIndex % 2 === 1 && this.customStyles["__theme_body_alt"]) {
        styleName = "__theme_body_alt";
      } else {
        styleName = "__theme_body";
      }
    } else if (section === "footer") {
      styleName = "__theme_footer";
    }
    return this.customStyles[styleName];
  }
  /**
   * Aplica un slicer a una tabla o tabla dinámica
   */
  async applySlicer(ws, slicer) {
    try {
      console.warn("Slicers require advanced ExcelJS XML manipulation. Feature documented but not fully implemented.");
      const colNum = typeof slicer.position.col === "string" ? this.columnLetterToNumber(slicer.position.col) : slicer.position.col;
      const cell = ws.getRow(slicer.position.row).getCell(colNum);
      cell.note = `Slicer: ${slicer.name} for table "${slicer.targetTable}" on column "${slicer.column}"`;
    } catch (error) {
      console.warn("Error adding slicer:", error);
    }
  }
  /**
   * Aplica una marca de agua al worksheet
   */
  async applyWatermark(ws, watermark) {
    try {
      if (watermark.image) {
        const imageConfig = {
          ...watermark.image,
          position: watermark.position ? {
            row: watermark.position.vertical === "top" ? 1 : watermark.position.vertical === "bottom" ? 1e3 : 500,
            col: watermark.position.horizontal === "left" ? 1 : watermark.position.horizontal === "right" ? 20 : 10
          } : { row: 500, col: 10 },
          size: watermark.image.size || {
            width: 400,
            height: 300,
            scaleX: watermark.opacity || 0.3,
            scaleY: watermark.opacity || 0.3
          }
        };
        await this.applyImage(ws, imageConfig);
      } else if (watermark.text) {
        const centerRow = Math.floor((ws.rowCount || 100) / 2);
        const centerCol = Math.floor((ws.columnCount || 20) / 2);
        const cell = ws.getRow(centerRow).getCell(centerCol);
        cell.value = watermark.text;
        cell.style = {
          font: {
            size: watermark.fontSize || 72,
            color: { argb: this.convertColorToExcelJS(watermark.fontColor || "#CCCCCC").argb },
            italic: true
          },
          alignment: {
            horizontal: "center",
            vertical: "middle"
          }
        };
      }
    } catch (error) {
      console.warn("Error adding watermark:", error);
    }
  }
  /**
   * Aplica una conexión de datos
   */
  async applyDataConnection(workbook, connection) {
    try {
      console.warn("Data connections require advanced ExcelJS XML manipulation. Feature documented but not fully implemented.");
      if (!workbook.model) {
        workbook.model = {};
      }
      if (!workbook.model.dataConnections) {
        workbook.model.dataConnections = [];
      }
      workbook.model.dataConnections.push({
        name: connection.name,
        type: connection.type,
        connectionString: connection.connectionString,
        commandText: connection.commandText,
        refresh: connection.refresh,
        credentials: connection.credentials ? {
          username: connection.credentials.username,
          integratedSecurity: connection.credentials.integratedSecurity
          // No guardar password por seguridad
        } : void 0
      });
    } catch (error) {
      console.warn("Error adding data connection:", error);
    }
  }
}
var BuilderEventType = /* @__PURE__ */ ((BuilderEventType2) => {
  BuilderEventType2["WORKSHEET_ADDED"] = "worksheetAdded";
  BuilderEventType2["WORKSHEET_REMOVED"] = "worksheetRemoved";
  BuilderEventType2["WORKSHEET_UPDATED"] = "worksheetUpdated";
  BuilderEventType2["BUILD_STARTED"] = "buildStarted";
  BuilderEventType2["BUILD_PROGRESS"] = "buildProgress";
  BuilderEventType2["BUILD_COMPLETED"] = "buildCompleted";
  BuilderEventType2["BUILD_ERROR"] = "buildError";
  BuilderEventType2["DOWNLOAD_STARTED"] = "downloadStarted";
  BuilderEventType2["DOWNLOAD_PROGRESS"] = "downloadProgress";
  BuilderEventType2["DOWNLOAD_COMPLETED"] = "downloadCompleted";
  BuilderEventType2["DOWNLOAD_ERROR"] = "downloadError";
  return BuilderEventType2;
})(BuilderEventType || {});
class ExcelBuilder {
  config;
  worksheets = /* @__PURE__ */ new Map();
  currentWorksheet;
  isBuilding = false;
  stats;
  eventEmitter;
  cellStyles = /* @__PURE__ */ new Map();
  theme;
  /**
   * Creates a new ExcelBuilder instance
   * 
   * @param {IExcelBuilderConfig} config - Configuration options for the builder
   * @param {IWorkbookMetadata} [config.metadata] - Workbook metadata (title, author, description, etc.)
   * @param {Partial<IWorksheetConfig>} [config.defaultWorksheetConfig] - Default configuration for all worksheets
   * @param {boolean} [config.enableValidation=true] - Enable data validation
   * @param {boolean} [config.enableEvents=true] - Enable event system
   * @param {boolean} [config.enablePerformanceMonitoring=false] - Enable performance monitoring
   * @param {number} [config.maxWorksheets=255] - Maximum number of worksheets allowed
   * @param {number} [config.maxRowsPerWorksheet=1048576] - Maximum rows per worksheet
   * @param {number} [config.maxColumnsPerWorksheet=16384] - Maximum columns per worksheet
   * @param {number} [config.memoryLimit=104857600] - Memory limit in bytes (100MB default)
   * 
   * @example
   * ```typescript
   * const builder = new ExcelBuilder({
   *   metadata: {
   *     title: 'Annual Report',
   *     author: 'John Doe',
   *     description: 'Company annual report for 2024'
   *   },
   *   enableValidation: true,
   *   enableEvents: true
   * });
   * ```
   */
  constructor(config = {}) {
    this.config = {
      enableValidation: true,
      enableEvents: true,
      enablePerformanceMonitoring: false,
      maxWorksheets: 255,
      maxRowsPerWorksheet: 1048576,
      maxColumnsPerWorksheet: 16384,
      memoryLimit: 100 * 1024 * 1024,
      // 100MB
      ...config
    };
    this.stats = this.initializeStats();
    this.eventEmitter = new EventEmitter();
  }
  /**
   * Add a new worksheet to the workbook
   * 
   * Creates a new worksheet with the specified name and configuration. The worksheet
   * becomes the current worksheet automatically. If a worksheet with the same name
   * already exists, an error is thrown.
   * 
   * @param {string} name - Unique name for the worksheet (required, must be unique)
   * @param {Partial<IWorksheetConfig>} [worksheetConfig={}] - Configuration for the worksheet
   * @param {string} [worksheetConfig.tabColor] - Tab color (hex format, e.g., '#FF0000')
   * @param {number} [worksheetConfig.defaultRowHeight=20] - Default row height in points
   * @param {number} [worksheetConfig.defaultColWidth=10] - Default column width in characters
   * @param {boolean} [worksheetConfig.hidden=false] - Whether the worksheet is hidden
   * @param {boolean} [worksheetConfig.protected=false] - Whether the worksheet is protected
   * @param {string} [worksheetConfig.protectionPassword] - Password for worksheet protection
   * @param {boolean} [worksheetConfig.showGridLines=true] - Show grid lines
   * @param {boolean} [worksheetConfig.showRowColHeaders=true] - Show row and column headers
   * @param {number} [worksheetConfig.zoom] - Zoom level (10-400)
   * 
   * @returns {IWorksheet} The newly created worksheet instance
   * 
   * @throws {Error} If a worksheet with the same name already exists
   * 
   * @example
   * ```typescript
   * // Simple worksheet
   * const sheet1 = builder.addWorksheet('Sales');
   * 
   * // Worksheet with configuration
   * const sheet2 = builder.addWorksheet('Summary', {
   *   tabColor: '#4472C4',
   *   defaultRowHeight: 25,
   *   defaultColWidth: 15,
   *   protected: true,
   *   protectionPassword: 'mypassword'
   * });
   * ```
   */
  addWorksheet(name, worksheetConfig = {}) {
    if (this.worksheets.has(name)) {
      throw new Error(`Worksheet "${name}" already exists`);
    }
    const config = {
      name,
      defaultRowHeight: 20,
      defaultColWidth: 10,
      ...this.config.defaultWorksheetConfig,
      ...worksheetConfig
    };
    const worksheet = new Worksheet(config);
    this.worksheets.set(name, worksheet);
    this.currentWorksheet = worksheet;
    this.emitEvent(BuilderEventType.WORKSHEET_ADDED, { worksheetName: name });
    return worksheet;
  }
  /**
   * Get a worksheet by name
   * 
   * Retrieves an existing worksheet from the workbook by its name.
   * Returns undefined if the worksheet doesn't exist.
   * 
   * @param {string} name - Name of the worksheet to retrieve
   * @returns {IWorksheet | undefined} The worksheet if found, undefined otherwise
   * 
   * @example
   * ```typescript
   * const worksheet = builder.getWorksheet('Sales');
   * if (worksheet) {
   *   worksheet.addRow([...]);
   * }
   * ```
   */
  getWorksheet(name) {
    return this.worksheets.get(name);
  }
  /**
   * Remove a worksheet by name
   * 
   * Removes a worksheet from the workbook. If the removed worksheet was the current
   * worksheet, the current worksheet is cleared.
   * 
   * @param {string} name - Name of the worksheet to remove
   * @returns {boolean} True if the worksheet was found and removed, false otherwise
   * 
   * @example
   * ```typescript
   * const removed = builder.removeWorksheet('OldSheet');
   * if (removed) {
   *   console.log('Worksheet removed successfully');
   * }
   * ```
   */
  removeWorksheet(name) {
    const worksheet = this.worksheets.get(name);
    if (!worksheet) {
      return false;
    }
    this.worksheets.delete(name);
    if (this.currentWorksheet === worksheet) {
      this.currentWorksheet = void 0;
    }
    this.emitEvent(BuilderEventType.WORKSHEET_REMOVED, { worksheetName: name });
    return true;
  }
  /**
   * Set the current worksheet
   * 
   * Sets the active worksheet. Operations like addRow() will be performed on the
   * current worksheet. When you add a new worksheet, it automatically becomes the current one.
   * 
   * @param {string} name - Name of the worksheet to set as current
   * @returns {boolean} True if the worksheet was found and set, false otherwise
   * 
   * @example
   * ```typescript
   * builder.addWorksheet('Sheet1');
   * builder.addWorksheet('Sheet2');
   * 
   * // Switch back to Sheet1
   * builder.setCurrentWorksheet('Sheet1');
   * ```
   */
  setCurrentWorksheet(name) {
    const worksheet = this.worksheets.get(name);
    if (!worksheet) {
      return false;
    }
    this.currentWorksheet = worksheet;
    return true;
  }
  /**
   * Build the workbook and return as ArrayBuffer
   * 
   * Compiles all worksheets, applies themes and styles, and generates the Excel file
   * as an ArrayBuffer. This is the core method that all export methods use internally.
   * 
   * The build process:
   * 1. Creates a new ExcelJS workbook
   * 2. Applies workbook metadata
   * 3. Applies theme (if set)
   * 4. Adds predefined cell styles
   * 5. Builds each worksheet
   * 6. Writes to buffer with compression
   * 
   * @param {IBuildOptions} [options={}] - Build options
   * @param {'xlsx' | 'xls' | 'csv'} [options.format='xlsx'] - Output format
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * @param {number} [options.compressionLevel=6] - Compression level (0-9, higher = more compression)
   * @param {boolean} [options.optimizeForSpeed=false] - Optimize for speed over file size
   * 
   * @returns {Promise<Result<ArrayBuffer>>} Result containing the Excel file as ArrayBuffer
   * 
   * @throws {Error} If build is already in progress (prevents concurrent builds)
   * 
   * @example
   * ```typescript
   * // Basic build
   * const result = await builder.build();
   * if (result.success) {
   *   const buffer = result.data;
   *   // Use buffer...
   * }
   * 
   * // Build with options
   * const result = await builder.build({
   *   compressionLevel: 9, // Maximum compression
   *   optimizeForSpeed: true
   * });
   * ```
   */
  async build(options = {}) {
    if (this.isBuilding) {
      return {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: "Build already in progress",
          stack: new Error().stack || ""
        }
      };
    }
    this.isBuilding = true;
    const startTime = Date.now();
    try {
      this.emitEvent(BuilderEventType.BUILD_STARTED);
      const workbook = new ExcelJS.Workbook();
      if (this.config.metadata) {
        workbook.creator = this.config.metadata.author || "Han Excel Builder";
        workbook.lastModifiedBy = this.config.metadata.author || "Han Excel Builder";
        workbook.created = this.config.metadata.created || /* @__PURE__ */ new Date();
        workbook.modified = this.config.metadata.modified || /* @__PURE__ */ new Date();
        if (this.config.metadata.title)
          workbook.title = this.config.metadata.title;
        if (this.config.metadata.subject)
          workbook.subject = this.config.metadata.subject;
        if (this.config.metadata.keywords)
          workbook.keywords = this.config.metadata.keywords;
        if (this.config.metadata.category)
          workbook.category = this.config.metadata.category;
        if (this.config.metadata.description)
          workbook.description = this.config.metadata.description;
      }
      if (this.theme) {
        this.applyTheme(workbook, this.theme);
      }
      for (const [name, style] of this.cellStyles.entries()) {
        this.addStyleToWorkbook(workbook, name, style);
      }
      for (const worksheet of this.worksheets.values()) {
        await worksheet.build(workbook, options);
      }
      const buffer = await workbook.xlsx.writeBuffer({
        compression: options.compressionLevel || 6
      });
      const endTime = Date.now();
      this.stats.buildTime = endTime - startTime;
      this.stats.fileSize = buffer.byteLength;
      const successResult = {
        success: true,
        data: buffer
      };
      this.emitEvent(BuilderEventType.BUILD_COMPLETED, {
        buildTime: this.stats.buildTime,
        fileSize: this.stats.fileSize
      });
      return successResult;
    } catch (error) {
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : "Unknown build error",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      this.emitEvent(BuilderEventType.BUILD_ERROR, { error: errorResult.error });
      return errorResult;
    } finally {
      this.isBuilding = false;
    }
  }
  /**
   * Generate and download the file (Browser only)
   * 
   * Builds the Excel file and automatically triggers a download in the user's browser.
   * This is the simplest method for browser environments - just one method call!
   * 
   * **Note**: This method is designed for browser environments. For Node.js, use `saveToFile()` instead.
   * 
   * @param {string} fileName - Name of the file to download (e.g., 'report.xlsx')
   * @param {IDownloadOptions} [options={}] - Download options
   * @param {string} [options.mimeType] - MIME type (default: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
   * @param {number} [options.compressionLevel=6] - Compression level (0-9)
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * 
   * @returns {Promise<Result<void>>} Result indicating success or failure
   * 
   * @example
   * ```typescript
   * // Simple download
   * const result = await builder.generateAndDownload('sales-report.xlsx');
   * 
   * if (result.success) {
   *   console.log('File downloaded successfully!');
   * } else {
   *   console.error('Download failed:', result.error);
   * }
   * 
   * // With options
   * await builder.generateAndDownload('report.xlsx', {
   *   compressionLevel: 9,
   *   mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
   * });
   * ```
   */
  async generateAndDownload(fileName, options = {}) {
    const buildResult = await this.build(options);
    if (!buildResult.success) {
      return buildResult;
    }
    try {
      this.emitEvent(BuilderEventType.DOWNLOAD_STARTED, { fileName });
      const blob = new Blob([buildResult.data], {
        type: options.mimeType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      saveAs(blob, fileName);
      this.emitEvent(BuilderEventType.DOWNLOAD_COMPLETED, { fileName });
      return { success: true, data: void 0 };
    } catch (error) {
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : "Download failed",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      this.emitEvent(BuilderEventType.DOWNLOAD_ERROR, { error: errorResult.error });
      return errorResult;
    }
  }
  /**
   * Save file to disk (Node.js only)
   * 
   * Builds the Excel file and saves it directly to the file system. This is the Node.js
   * equivalent of `generateAndDownload()` - just as simple! Automatically creates parent
   * directories if they don't exist.
   * 
   * **Note**: This method only works in Node.js environments. For browsers, use `generateAndDownload()`.
   * 
   * @param {string} filePath - Full path where to save the file (e.g., './output/report.xlsx')
   * @param {ISaveFileOptions} [options={}] - Save options
   * @param {boolean} [options.createDir=true] - Create parent directories if they don't exist
   * @param {string} [options.encoding='binary'] - File encoding ('binary', 'base64', etc.)
   * @param {number} [options.compressionLevel=6] - Compression level (0-9)
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * 
   * @returns {Promise<Result<void>>} Result indicating success or failure
   * 
   * @throws {Error} If called in browser environment (use `generateAndDownload()` instead)
   * @throws {Error} If Node.js modules (fs, path, buffer) are not available
   * 
   * @example
   * ```typescript
   * // Simple save - creates directories automatically
   * const result = await builder.saveToFile('./output/report.xlsx');
   * 
   * if (result.success) {
   *   console.log('File saved successfully!');
   * }
   * 
   * // With options
   * await builder.saveToFile('./reports/sales.xlsx', {
   *   createDir: true,  // Create ./reports/ if it doesn't exist
   *   encoding: 'binary',
   *   compressionLevel: 9
   * });
   * ```
   */
  async saveToFile(filePath, options = {}) {
    const buildResult = await this.build(options);
    if (!buildResult.success) {
      return buildResult;
    }
    try {
      if (typeof window !== "undefined") {
        const errorResult = {
          success: false,
          error: {
            type: ErrorType.BUILD_ERROR,
            message: "saveToFile() is only available in Node.js. Use generateAndDownload() in the browser.",
            stack: ""
          }
        };
        return errorResult;
      }
      this.emitEvent(BuilderEventType.DOWNLOAD_STARTED, { fileName: filePath });
      const nodeModules = await (async () => {
        try {
          const fs = await import("./__vite-browser-external-d06ac358.js");
          const path = await import("./__vite-browser-external-d06ac358.js");
          const buffer2 = await import("./index-8081eac4.js").then((n) => n.i);
          return { fs, path, Buffer: buffer2.Buffer };
        } catch {
          throw new Error("Node.js modules not available. saveToFile() requires Node.js environment.");
        }
      })();
      if (options.createDir !== false) {
        const dir = nodeModules.path.dirname(filePath);
        try {
          await nodeModules.fs.mkdir(dir, { recursive: true });
        } catch (error) {
          if (error?.code !== "EEXIST") {
            throw error;
          }
        }
      }
      const buffer = nodeModules.Buffer.from(buildResult.data);
      await nodeModules.fs.writeFile(filePath, buffer, { encoding: options.encoding || "binary" });
      this.emitEvent(BuilderEventType.DOWNLOAD_COMPLETED, { fileName: filePath });
      return { success: true, data: void 0 };
    } catch (error) {
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : "Failed to save file",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      this.emitEvent(BuilderEventType.DOWNLOAD_ERROR, { error: errorResult.error });
      return errorResult;
    }
  }
  /**
   * Save to stream (Node.js only) - For large files
   * 
   * Builds the Excel file and writes it directly to a writable stream. This is ideal
   * for very large files or when you need to stream the data (e.g., HTTP responses,
   * file uploads, etc.).
   * 
   * **Note**: This method only works in Node.js environments.
   * 
   * @param {NodeJS.WritableStream} writeStream - Writable stream to write the file to
   * @param {IBuildOptions} [options={}] - Build options
   * @param {number} [options.compressionLevel=6] - Compression level (0-9)
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * 
   * @returns {Promise<Result<void>>} Result indicating success or failure
   * 
   * @throws {Error} If called in browser environment
   * 
   * @example
   * ```typescript
   * import fs from 'fs';
   * 
   * // Save to file stream
   * const writeStream = fs.createWriteStream('./output/report.xlsx');
   * const result = await builder.saveToStream(writeStream);
   * 
   * if (result.success) {
   *   writeStream.end();
   *   console.log('File streamed successfully!');
   * }
   * 
   * // HTTP response stream
   * app.get('/download', async (req, res) => {
   *   res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
   *   await builder.saveToStream(res);
   * });
   * ```
   */
  async saveToStream(writeStream, options = {}) {
    const buildResult = await this.build(options);
    if (!buildResult.success) {
      return buildResult;
    }
    try {
      if (typeof window !== "undefined") {
        const errorResult = {
          success: false,
          error: {
            type: ErrorType.BUILD_ERROR,
            message: "saveToStream() is only available in Node.js.",
            stack: ""
          }
        };
        return errorResult;
      }
      this.emitEvent(BuilderEventType.DOWNLOAD_STARTED, { fileName: "stream" });
      const bufferModule = await import("./index-8081eac4.js").then((n) => n.i);
      const buffer = bufferModule.Buffer.from(buildResult.data);
      return new Promise((resolve) => {
        writeStream.write(buffer, (error) => {
          if (error) {
            const errorResult = {
              success: false,
              error: {
                type: ErrorType.BUILD_ERROR,
                message: error.message || "Failed to write to stream",
                stack: error.stack || ""
              }
            };
            this.emitEvent(BuilderEventType.DOWNLOAD_ERROR, { error: errorResult.error });
            resolve(errorResult);
          } else {
            this.emitEvent(BuilderEventType.DOWNLOAD_COMPLETED, { fileName: "stream" });
            resolve({ success: true, data: void 0 });
          }
        });
      });
    } catch (error) {
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : "Failed to save to stream",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      this.emitEvent(BuilderEventType.DOWNLOAD_ERROR, { error: errorResult.error });
      return errorResult;
    }
  }
  /**
   * Get workbook as ArrayBuffer
   * 
   * Builds the Excel file and returns it as an ArrayBuffer. This is useful when you need
   * the raw binary data for custom handling (e.g., sending via WebSocket, processing with
   * other libraries, manual file operations, etc.).
   * 
   * Works in both browser and Node.js environments.
   * 
   * @param {IBuildOptions} [options={}] - Build options
   * @param {number} [options.compressionLevel=6] - Compression level (0-9)
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * 
   * @returns {Promise<Result<ArrayBuffer>>} Result containing the Excel file as ArrayBuffer
   * 
   * @example
   * ```typescript
   * // Browser: Get buffer for custom handling
   * const result = await builder.toBuffer();
   * if (result.success) {
   *   const buffer = result.data;
   *   // Upload to server, send via WebSocket, etc.
   * }
   * 
   * // Node.js: Get buffer for manual file write
   * const result = await builder.toBuffer();
   * if (result.success) {
   *   const fs = await import('fs/promises');
   *   await fs.writeFile('./report.xlsx', Buffer.from(result.data));
   * }
   * ```
   */
  async toBuffer(options = {}) {
    return this.build(options);
  }
  /**
   * Get workbook as Blob
   * 
   * Builds the Excel file and returns it as a Blob object. This is useful in browser
   * environments when you need to upload to a server, create object URLs for preview,
   * or handle the file programmatically without triggering an automatic download.
   * 
   * **Note**: Blob is a browser API. In Node.js, use `toBuffer()` instead.
   * 
   * @param {IBuildOptions} [options={}] - Build options
   * @param {number} [options.compressionLevel=6] - Compression level (0-9)
   * @param {boolean} [options.includeStyles=true] - Include cell styles
   * 
   * @returns {Promise<Result<Blob>>} Result containing the Excel file as Blob
   * 
   * @example
   * ```typescript
   * // Get as Blob for upload
   * const result = await builder.toBlob();
   * if (result.success) {
   *   const blob = result.data;
   *   
   *   // Upload to server
   *   const formData = new FormData();
   *   formData.append('file', blob, 'report.xlsx');
   *   await fetch('/api/upload', { method: 'POST', body: formData });
   *   
   *   // Or create preview URL
   *   const url = URL.createObjectURL(blob);
   *   window.open(url);
   * }
   * ```
   */
  async toBlob(options = {}) {
    const buildResult = await this.build(options);
    if (!buildResult.success) {
      return buildResult;
    }
    const blob = new Blob([buildResult.data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    return { success: true, data: blob };
  }
  /**
   * Validate the workbook
   * 
   * Performs validation checks on the workbook to ensure it's ready for building.
   * Validates that worksheets exist and each worksheet is valid.
   * 
   * @returns {Result<boolean>} Result indicating if the workbook is valid
   * - `success: true` - Workbook is valid and ready to build
   * - `success: false` - Validation errors found (check `error.message` for details)
   * 
   * @example
   * ```typescript
   * const validation = builder.validate();
   * if (!validation.success) {
   *   console.error('Validation errors:', validation.error?.message);
   *   return;
   * }
   * 
   * // Safe to build
   * await builder.build();
   * ```
   */
  validate() {
    const errors = [];
    if (this.worksheets.size === 0) {
      errors.push("No worksheets found");
    }
    for (const [name, worksheet] of this.worksheets.entries()) {
      const worksheetValidation = worksheet.validate();
      if (!worksheetValidation.success) {
        errors.push(`Worksheet "${name}": ${worksheetValidation.error?.message}`);
      }
    }
    if (errors.length > 0) {
      return {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: errors.join("; "),
          stack: new Error().stack || ""
        }
      };
    }
    return { success: true, data: true };
  }
  /**
   * Clear all worksheets and reset the builder
   * 
   * Removes all worksheets, clears predefined cell styles, resets the theme,
   * and clears the current worksheet. This effectively resets the builder to
   * its initial state.
   * 
   * @returns {void}
   * 
   * @example
   * ```typescript
   * // Clear everything and start fresh
   * builder.clear();
   * 
   * // Now add new worksheets
   * builder.addWorksheet('NewSheet');
   * ```
   */
  clear() {
    this.worksheets.clear();
    this.currentWorksheet = void 0;
    this.cellStyles.clear();
    this.theme = void 0;
  }
  /**
   * Get workbook statistics
   * 
   * Returns build statistics including build time, file size, number of worksheets,
   * cells, styles used, and performance metrics. Statistics are updated after each build.
   * 
   * @returns {IBuildStats} Statistics object containing:
   * - `totalWorksheets` - Number of worksheets
   * - `totalCells` - Total number of cells
   * - `memoryUsage` - Memory usage in bytes
   * - `buildTime` - Last build time in milliseconds
   * - `fileSize` - Last build file size in bytes
   * - `stylesUsed` - Number of unique styles used
   * - `formulasUsed` - Number of formulas
   * - `conditionalFormatsUsed` - Number of conditional formats
   * - `performance` - Performance breakdown by operation
   * 
   * @example
   * ```typescript
   * await builder.build();
   * const stats = builder.getStats();
   * 
   * console.log(`Build time: ${stats.buildTime}ms`);
   * console.log(`File size: ${stats.fileSize} bytes`);
   * console.log(`Worksheets: ${stats.totalWorksheets}`);
   * ```
   */
  getStats() {
    return { ...this.stats };
  }
  /**
   * Add a predefined cell style
   * 
   * Defines a reusable cell style that can be referenced by name in cells using
   * the `styleName` property. This is useful for maintaining consistent styling
   * across the workbook and reducing code duplication.
   * 
   * Styles are stored at the workbook level and can be used in any worksheet.
   * 
   * @param {string} name - Unique name for the style (used to reference it later)
   * @param {IStyle} style - Style object created with StyleBuilder
   * @returns {this} Returns the builder instance for method chaining
   * 
   * @example
   * ```typescript
   * // Define reusable styles
   * builder.addCellStyle('headerStyle', StyleBuilder.create()
   *   .font({ name: 'Arial', size: 14, bold: true })
   *   .fill({ backgroundColor: '#4472C4' })
   *   .fontColor('#FFFFFF')
   *   .build()
   * );
   * 
   * // Use in cells
   * worksheet.addHeader({
   *   key: 'title',
   *   value: 'Report',
   *   type: CellType.STRING,
   *   styleName: 'headerStyle' // Reference the predefined style
   * });
   * ```
   */
  addCellStyle(name, style) {
    this.cellStyles.set(name, style);
    return this;
  }
  /**
   * Get a predefined cell style by name
   * 
   * Retrieves a previously defined cell style by its name. Returns undefined
   * if the style doesn't exist.
   * 
   * @param {string} name - Name of the style to retrieve
   * @returns {IStyle | undefined} The style if found, undefined otherwise
   * 
   * @example
   * ```typescript
   * const style = builder.getCellStyle('headerStyle');
   * if (style) {
   *   console.log('Style found:', style);
   * }
   * ```
   */
  getCellStyle(name) {
    return this.cellStyles.get(name);
  }
  /**
   * Set workbook theme
   * 
   * Applies a color and font theme to the entire workbook. Themes affect all
   * worksheets and can automatically apply styles to table sections (header, body, footer)
   * if `autoApplySectionStyles` is enabled.
   * 
   * Themes include:
   * - Color palette (dark1, light1, dark2, light2, accent1-6, hyperlink colors)
   * - Font families (major and minor fonts for latin, eastAsian, complexScript)
   * - Optional section styles for automatic styling
   * 
   * @param {IWorkbookTheme} theme - Theme configuration object
   * @param {string} [theme.name] - Theme name
   * @param {object} [theme.colors] - Color palette
   * @param {object} [theme.fonts] - Font configuration
   * @param {object} [theme.sectionStyles] - Styles for table sections
   * @param {boolean} [theme.autoApplySectionStyles=true] - Auto-apply section styles
   * 
   * @returns {this} Returns the builder instance for method chaining
   * 
   * @example
   * ```typescript
   * builder.setTheme({
   *   name: 'Corporate Theme',
   *   colors: {
   *     dark1: '#000000',
   *     light1: '#FFFFFF',
   *     accent1: '#4472C4',
   *     accent2: '#ED7D31'
   *   },
   *   fonts: {
   *     major: { latin: 'Calibri' },
   *     minor: { latin: 'Calibri' }
   *   },
   *   autoApplySectionStyles: true
   * });
   * ```
   */
  setTheme(theme) {
    this.theme = theme;
    return this;
  }
  /**
   * Get current workbook theme
   * 
   * Retrieves the currently active theme, if one has been set.
   * 
   * @returns {IWorkbookTheme | undefined} The current theme, or undefined if no theme is set
   * 
   * @example
   * ```typescript
   * const theme = builder.getTheme();
   * if (theme) {
   *   console.log('Active theme:', theme.name);
   * }
   * ```
   */
  getTheme() {
    return this.theme;
  }
  /**
   * Register an event listener
   * 
   * Subscribes to builder events to monitor the build process. Returns a listener ID
   * that can be used to remove the listener later.
   * 
   * Available events:
   * - `build:started` - Build process started
   * - `build:completed` - Build completed successfully
   * - `build:error` - Build failed with error
   * - `download:started` - File download/save started
   * - `download:completed` - File download/save completed
   * - `download:error` - File download/save failed
   * - `worksheet:added` - New worksheet added
   * - `worksheet:removed` - Worksheet removed
   * 
   * @param {BuilderEventType} eventType - Type of event to listen for
   * @param {(event: IBuilderEvent) => void} listener - Callback function to execute when event fires
   * @returns {string} Listener ID (use with `off()` to remove the listener)
   * 
   * @example
   * ```typescript
   * const listenerId = builder.on('build:started', (event) => {
   *   console.log('Build started at', event.timestamp);
   * });
   * 
   * builder.on('build:completed', (event) => {
   *   console.log('Build completed:', event.data);
   * });
   * 
   * builder.on('build:error', (event) => {
   *   console.error('Build error:', event.data.error);
   * });
   * ```
   */
  on(eventType, listener) {
    return this.eventEmitter.on(eventType, listener);
  }
  /**
   * Remove an event listener
   * 
   * Unsubscribes from a specific event by removing the listener with the given ID.
   * 
   * @param {BuilderEventType} eventType - Type of event
   * @param {string} listenerId - Listener ID returned from `on()`
   * @returns {boolean} True if the listener was found and removed, false otherwise
   * 
   * @example
   * ```typescript
   * const listenerId = builder.on('build:started', handler);
   * 
   * // Later, remove the listener
   * builder.off('build:started', listenerId);
   * ```
   */
  off(eventType, listenerId) {
    return this.eventEmitter.off(eventType, listenerId);
  }
  /**
   * Remove all event listeners
   * 
   * Removes all listeners for a specific event type, or all listeners for all events
   * if no event type is specified.
   * 
   * @param {BuilderEventType} [eventType] - Event type to clear listeners for. If omitted, clears all listeners
   * @returns {void}
   * 
   * @example
   * ```typescript
   * // Remove all listeners for 'build:started' event
   * builder.removeAllListeners('build:started');
   * 
   * // Remove all listeners for all events
   * builder.removeAllListeners();
   * ```
   */
  removeAllListeners(eventType) {
    if (eventType) {
      this.eventEmitter.offAll(eventType);
    } else {
      this.eventEmitter.clear();
    }
  }
  /**
   * Private methods
   */
  /**
   * Emit an event to all registered listeners
   * @private
   */
  emitEvent(type, data) {
    const event = {
      type,
      data: data || {},
      timestamp: /* @__PURE__ */ new Date()
    };
    this.eventEmitter.emitSync(event);
  }
  /**
   * Initialize build statistics
   * @private
   */
  initializeStats() {
    return {
      totalWorksheets: 0,
      totalCells: 0,
      memoryUsage: 0,
      buildTime: 0,
      fileSize: 0,
      stylesUsed: 0,
      formulasUsed: 0,
      conditionalFormatsUsed: 0,
      performance: {
        headersTime: 0,
        dataTime: 0,
        stylesTime: 0,
        writeTime: 0
      }
    };
  }
  /**
   * Apply theme to workbook
   * 
   * Internal method that applies the theme configuration to the ExcelJS workbook.
   * Converts theme colors and fonts to ExcelJS format.
   * 
   * @private
   * @param {ExcelJS.Workbook} workbook - ExcelJS workbook instance
   * @param {IWorkbookTheme} theme - Theme configuration
   */
  applyTheme(workbook, theme) {
    if (!workbook.model) {
      return;
    }
    const excelTheme = {
      name: theme.name || "Custom Theme"
    };
    if (theme.colors) {
      excelTheme.colors = {};
      if (theme.colors.dark1)
        excelTheme.colors.dark1 = this.convertColorToTheme(theme.colors.dark1);
      if (theme.colors.light1)
        excelTheme.colors.light1 = this.convertColorToTheme(theme.colors.light1);
      if (theme.colors.dark2)
        excelTheme.colors.dark2 = this.convertColorToTheme(theme.colors.dark2);
      if (theme.colors.light2)
        excelTheme.colors.light2 = this.convertColorToTheme(theme.colors.light2);
      if (theme.colors.accent1)
        excelTheme.colors.accent1 = this.convertColorToTheme(theme.colors.accent1);
      if (theme.colors.accent2)
        excelTheme.colors.accent2 = this.convertColorToTheme(theme.colors.accent2);
      if (theme.colors.accent3)
        excelTheme.colors.accent3 = this.convertColorToTheme(theme.colors.accent3);
      if (theme.colors.accent4)
        excelTheme.colors.accent4 = this.convertColorToTheme(theme.colors.accent4);
      if (theme.colors.accent5)
        excelTheme.colors.accent5 = this.convertColorToTheme(theme.colors.accent5);
      if (theme.colors.accent6)
        excelTheme.colors.accent6 = this.convertColorToTheme(theme.colors.accent6);
      if (theme.colors.hyperlink)
        excelTheme.colors.hyperlink = this.convertColorToTheme(theme.colors.hyperlink);
      if (theme.colors.followedHyperlink)
        excelTheme.colors.followedHyperlink = this.convertColorToTheme(theme.colors.followedHyperlink);
    }
    if (theme.fonts) {
      excelTheme.fonts = {};
      if (theme.fonts.major) {
        excelTheme.fonts.major = {
          latin: theme.fonts.major.latin || "Calibri",
          eastAsian: theme.fonts.major.eastAsian || theme.fonts.major.latin || "Calibri",
          complexScript: theme.fonts.major.complexScript || theme.fonts.major.latin || "Calibri"
        };
      }
      if (theme.fonts.minor) {
        excelTheme.fonts.minor = {
          latin: theme.fonts.minor.latin || "Calibri",
          eastAsian: theme.fonts.minor.eastAsian || theme.fonts.minor.latin || "Calibri",
          complexScript: theme.fonts.minor.complexScript || theme.fonts.minor.latin || "Calibri"
        };
      }
    }
    workbook.model = workbook.model || {};
    workbook.model.theme = excelTheme;
  }
  /**
   * Convert color to theme format
   * 
   * Converts a Color value (hex string, RGB object, or theme color) to the format
   * expected by ExcelJS themes (hex string without #).
   * 
   * @private
   * @param {Color} color - Color to convert
   * @returns {string} Hex color string without # prefix
   */
  convertColorToTheme(color) {
    if (typeof color === "string") {
      return color.startsWith("#") ? color.substring(1) : color;
    }
    if ("r" in color && "g" in color && "b" in color) {
      return `${color.r.toString(16).padStart(2, "0")}${color.g.toString(16).padStart(2, "0")}${color.b.toString(16).padStart(2, "0")}`;
    }
    return "000000";
  }
  /**
   * Add style to workbook
   * 
   * Stores a predefined style in the workbook so it can be accessed during worksheet
   * building. ExcelJS doesn't support named styles directly, so we store them in a custom
   * property that worksheets can access when building cells.
   * 
   * @private
   * @param {ExcelJS.Workbook} workbook - ExcelJS workbook instance
   * @param {string} name - Style name
   * @param {IStyle} style - Style object
   * 
   * @remarks
   * ExcelJS applies styles per cell, not as named styles. This method stores styles
   * in a way that worksheets can retrieve them when building cells that reference
   * the style by name.
   */
  addStyleToWorkbook(workbook, name, style) {
    workbook.__customStyles = workbook.__customStyles || {};
    workbook.__customStyles[name] = style;
  }
}
var OutputFormat = /* @__PURE__ */ ((OutputFormat2) => {
  OutputFormat2["WORKSHEET"] = "worksheet";
  OutputFormat2["DETAILED"] = "detailed";
  OutputFormat2["FLAT"] = "flat";
  return OutputFormat2;
})(OutputFormat || {});
class ExcelReader {
  /**
   * Read Excel file from ArrayBuffer
   */
  static async fromBuffer(buffer, options = {}) {
    const startTime = Date.now();
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const outputFormat = options.outputFormat || OutputFormat.WORKSHEET;
      const processingTime = Date.now() - startTime;
      let result;
      switch (outputFormat) {
        case OutputFormat.DETAILED:
          result = this.convertToDetailedFormat(workbook, options);
          break;
        case OutputFormat.FLAT:
          result = this.convertToFlatFormat(workbook, options);
          break;
        case OutputFormat.WORKSHEET:
        default:
          result = this.convertWorkbookToJson(workbook, options);
          break;
      }
      if (options.mapper) {
        try {
          switch (outputFormat) {
            case OutputFormat.DETAILED:
              result = options.mapper(result);
              break;
            case OutputFormat.FLAT:
              result = options.mapper(result);
              break;
            case OutputFormat.WORKSHEET:
            default:
              result = options.mapper(result);
              break;
          }
        } catch (mapperError) {
          const errorResult = {
            success: false,
            error: {
              type: ErrorType.VALIDATION_ERROR,
              message: mapperError instanceof Error ? `Mapper function error: ${mapperError.message}` : "Error in mapper function",
              stack: mapperError instanceof Error ? mapperError.stack || "" : ""
            }
          };
          return {
            ...errorResult,
            processingTime: Date.now() - startTime
          };
        }
      }
      const successResult = {
        success: true,
        data: result,
        processingTime
      };
      return successResult;
    } catch (error) {
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: error instanceof Error ? error.message : "Error reading Excel file",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      const errorResponse = {
        success: false,
        error: errorResult.error,
        processingTime: Date.now() - startTime
      };
      return errorResponse;
    }
  }
  /**
   * Read Excel file from Blob
   */
  static async fromBlob(blob, options = {}) {
    const arrayBuffer = await blob.arrayBuffer();
    return this.fromBuffer(arrayBuffer, options);
  }
  /**
   * Read Excel file from File (browser)
   */
  static async fromFile(file, options = {}) {
    return this.fromBlob(file, options);
  }
  /**
   * Read Excel file from path (Node.js)
   * Note: This method only works in Node.js environment
   */
  /**
   * Read Excel file from path (Node.js only)
   * Note: This method only works in Node.js environment
   */
  static async fromPath(filePath, options = {}) {
    try {
      const fs = await import("./__vite-browser-external-d06ac358.js");
      const buffer = await fs.readFile(filePath);
      const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
      return this.fromBuffer(arrayBuffer, options);
    } catch (error) {
      const isBrowserError = error instanceof Error && (error.message.includes("Cannot find module") || error.message.includes("fs") || typeof window !== "undefined");
      const errorResult = {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: isBrowserError ? "fromPath() method requires Node.js environment. Use fromFile() or fromBlob() in browser." : error instanceof Error ? error.message : "Error reading file from path",
          stack: error instanceof Error ? error.stack || "" : ""
        }
      };
      const errorResponse = {
        ...errorResult,
        processingTime: 0
      };
      return errorResponse;
    }
  }
  /**
   * Convert ExcelJS Workbook to JSON
   */
  static convertWorkbookToJson(workbook, options) {
    const {
      includeEmptyRows = false,
      useFirstRowAsHeaders = false,
      headers,
      sheetName,
      startRow = 1,
      endRow,
      startColumn = 1,
      endColumn,
      includeFormatting = false,
      includeFormulas = false,
      datesAsISO = true
    } = options;
    const metadata = {
      title: workbook.title,
      author: workbook.creator,
      company: workbook.company,
      created: workbook.created,
      modified: workbook.modified,
      description: workbook.description
    };
    let sheetsToProcess = [];
    if (sheetName !== void 0) {
      if (typeof sheetName === "number") {
        const sheet = workbook.worksheets[sheetName];
        if (sheet)
          sheetsToProcess.push(sheet);
      } else {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet)
          sheetsToProcess.push(sheet);
      }
    } else {
      sheetsToProcess = workbook.worksheets;
    }
    const sheets = sheetsToProcess.map((worksheet) => {
      const sheetOptions = {
        includeEmptyRows: includeEmptyRows ?? false,
        useFirstRowAsHeaders: useFirstRowAsHeaders ?? false,
        startRow: startRow ?? 1,
        startColumn: startColumn ?? 1,
        includeFormatting: includeFormatting ?? false,
        includeFormulas: includeFormulas ?? false,
        datesAsISO: datesAsISO ?? true
      };
      if (headers !== void 0) {
        sheetOptions.headers = headers;
      }
      if (endRow !== void 0) {
        sheetOptions.endRow = endRow;
      }
      if (endColumn !== void 0) {
        sheetOptions.endColumn = endColumn;
      }
      return this.convertSheetToJson(worksheet, sheetOptions);
    });
    const workbookResult = {
      sheets,
      totalSheets: sheets.length
    };
    const hasMetadata = Object.values(metadata).some((val) => val !== void 0 && val !== null);
    if (hasMetadata) {
      workbookResult.metadata = metadata;
    }
    return workbookResult;
  }
  /**
   * Convert ExcelJS Worksheet to JSON
   */
  static convertSheetToJson(worksheet, options) {
    const {
      includeEmptyRows,
      useFirstRowAsHeaders,
      headers,
      startRow,
      endRow,
      startColumn,
      endColumn,
      includeFormatting,
      includeFormulas,
      datesAsISO
    } = options;
    const rows = [];
    let headerRow;
    let maxColumns = 0;
    const actualStartRow = Math.max(startRow, 1);
    const actualEndRow = endRow || worksheet.rowCount || worksheet.lastRow?.number || 1;
    const actualStartCol = Math.max(startColumn, 1);
    const actualEndCol = endColumn || worksheet.columnCount || worksheet.lastColumn?.number || 1;
    for (let rowNum = actualStartRow; rowNum <= actualEndRow; rowNum++) {
      const excelRow = worksheet.getRow(rowNum);
      const cells = [];
      let hasData = false;
      for (let colNum = actualStartCol; colNum <= actualEndCol; colNum++) {
        const cell = excelRow.getCell(colNum);
        if (!cell.value && !includeEmptyRows) {
          continue;
        }
        const jsonCell = this.convertCellToJson(cell, {
          includeFormatting,
          includeFormulas,
          datesAsISO
        });
        cells.push(jsonCell);
        hasData = true;
      }
      if (cells.length > maxColumns) {
        maxColumns = cells.length;
      }
      if (!hasData && !includeEmptyRows) {
        continue;
      }
      if (useFirstRowAsHeaders && rowNum === actualStartRow) {
        headerRow = cells.map((cell) => {
          if (headers && Array.isArray(headers)) {
            return headers[cells.indexOf(cell)] || String(cell.value || "");
          } else if (headers && typeof headers === "object") {
            return headers[actualStartCol + cells.indexOf(cell)] || String(cell.value || "");
          }
          return String(cell.value || "");
        });
        continue;
      }
      let rowData;
      if (useFirstRowAsHeaders && headerRow) {
        rowData = {};
        cells.forEach((cell, index) => {
          const header = headerRow[index] || `column_${index + 1}`;
          rowData[header] = cell.value;
        });
      }
      const jsonRow = {
        rowNumber: rowNum,
        cells
      };
      if (rowData) {
        jsonRow.data = rowData;
      }
      rows.push(jsonRow);
    }
    const sheet = {
      name: worksheet.name,
      index: worksheet.id || 0,
      rows,
      totalRows: rows.length,
      totalColumns: maxColumns
    };
    if (headerRow) {
      sheet.headers = headerRow;
    }
    return sheet;
  }
  /**
   * Convert ExcelJS Cell to JSON
   */
  static convertCellToJson(cell, options) {
    const { includeFormatting, includeFormulas, datesAsISO } = options;
    let value = cell.value;
    let type;
    if (cell.type === ExcelJS.ValueType.Null || cell.value === null || cell.value === void 0) {
      value = null;
      type = "null";
    } else if (cell.type === ExcelJS.ValueType.Number) {
      value = cell.value;
      type = "number";
    } else if (cell.type === ExcelJS.ValueType.String) {
      value = cell.value;
      type = "string";
    } else if (cell.type === ExcelJS.ValueType.Date) {
      const dateValue = cell.value;
      value = datesAsISO ? dateValue.toISOString() : dateValue;
      type = "date";
    } else if (cell.type === ExcelJS.ValueType.Boolean) {
      value = cell.value;
      type = "boolean";
    } else if (cell.type === ExcelJS.ValueType.Formula) {
      if (includeFormulas && cell.formula) {
        value = cell.result || cell.value;
        type = "formula";
      } else {
        value = cell.result || cell.value;
        type = typeof cell.result === "number" ? "number" : typeof cell.result === "string" ? "string" : "unknown";
      }
    } else if (cell.type === ExcelJS.ValueType.Hyperlink) {
      const hyperlinkValue = cell.value;
      if (typeof hyperlinkValue === "object" && hyperlinkValue !== null) {
        value = hyperlinkValue.text || hyperlinkValue.hyperlink || cell.value;
      } else {
        value = hyperlinkValue;
      }
      type = "hyperlink";
    } else {
      value = cell.value;
      type = "unknown";
    }
    const jsonCell = {
      value,
      type,
      reference: cell.address
    };
    if (includeFormatting && cell.numFmt) {
      jsonCell.formattedValue = String(value);
    }
    if (includeFormulas && cell.formula) {
      jsonCell.formula = cell.formula;
    }
    if (cell.note) {
      const note = cell.note;
      if (typeof note === "string") {
        jsonCell.comment = note;
      } else if (note && typeof note === "object" && "texts" in note) {
        const texts = note.texts;
        if (Array.isArray(texts) && texts.length > 0) {
          jsonCell.comment = texts.map((t) => t.text || "").join("");
        }
      } else if (note && typeof note === "object" && "text" in note) {
        jsonCell.comment = String(note.text);
      }
    }
    return jsonCell;
  }
  /**
   * Convert workbook to detailed format (with position information)
   */
  static convertToDetailedFormat(workbook, options) {
    const {
      includeEmptyRows = false,
      includeFormatting = false,
      includeFormulas = false,
      datesAsISO = true,
      sheetName,
      startRow = 1,
      endRow,
      startColumn = 1,
      endColumn
    } = options;
    const cells = [];
    const metadata = {
      title: workbook.title,
      author: workbook.creator,
      company: workbook.company,
      created: workbook.created,
      modified: workbook.modified,
      description: workbook.description
    };
    let sheetsToProcess = [];
    if (sheetName !== void 0) {
      if (typeof sheetName === "number") {
        const sheet = workbook.worksheets[sheetName];
        if (sheet)
          sheetsToProcess.push(sheet);
      } else {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet)
          sheetsToProcess.push(sheet);
      }
    } else {
      sheetsToProcess = workbook.worksheets;
    }
    for (const worksheet of sheetsToProcess) {
      const actualStartRow = Math.max(startRow, 1);
      const actualEndRow = endRow || worksheet.rowCount || worksheet.lastRow?.number || 1;
      const actualStartCol = Math.max(startColumn, 1);
      const actualEndCol = endColumn || worksheet.columnCount || worksheet.lastColumn?.number || 1;
      for (let rowNum = actualStartRow; rowNum <= actualEndRow; rowNum++) {
        const excelRow = worksheet.getRow(rowNum);
        for (let colNum = actualStartCol; colNum <= actualEndCol; colNum++) {
          const cell = excelRow.getCell(colNum);
          if (!cell.value && !includeEmptyRows) {
            continue;
          }
          const columnLetter = this.numberToColumnLetter(colNum);
          const cellValue = this.getCellValue(cell, { includeFormatting, includeFormulas, datesAsISO });
          const detailedCell = {
            value: cellValue.value,
            text: String(cellValue.value ?? ""),
            column: colNum,
            columnLetter,
            row: rowNum,
            reference: cell.address || `${columnLetter}${rowNum}`,
            sheet: worksheet.name
          };
          if (cellValue.type) {
            detailedCell.type = cellValue.type;
          }
          if (cellValue.formattedValue) {
            detailedCell.formattedValue = cellValue.formattedValue;
          }
          if (cellValue.formula) {
            detailedCell.formula = cellValue.formula;
          }
          if (cell.note) {
            const note = cell.note;
            if (typeof note === "string") {
              detailedCell.comment = note;
            } else if (note && typeof note === "object" && "texts" in note) {
              const texts = note.texts;
              if (Array.isArray(texts) && texts.length > 0) {
                detailedCell.comment = texts.map((t) => t.text || "").join("");
              }
            } else if (note && typeof note === "object" && "text" in note) {
              detailedCell.comment = String(note.text);
            }
          }
          cells.push(detailedCell);
        }
      }
    }
    const result = {
      cells,
      totalCells: cells.length
    };
    const hasMetadata = Object.values(metadata).some((val) => val !== void 0 && val !== null);
    if (hasMetadata) {
      result.metadata = metadata;
    }
    return result;
  }
  /**
   * Convert workbook to flat format (just data)
   */
  static convertToFlatFormat(workbook, options) {
    const {
      useFirstRowAsHeaders = false,
      includeEmptyRows = false,
      sheetName,
      startRow = 1,
      endRow,
      startColumn = 1,
      endColumn
    } = options;
    const metadata = {
      title: workbook.title,
      author: workbook.creator,
      company: workbook.company,
      created: workbook.created,
      modified: workbook.modified,
      description: workbook.description
    };
    let sheetsToProcess = [];
    if (sheetName !== void 0) {
      if (typeof sheetName === "number") {
        const sheet = workbook.worksheets[sheetName];
        if (sheet)
          sheetsToProcess.push(sheet);
      } else {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet)
          sheetsToProcess.push(sheet);
      }
    } else {
      sheetsToProcess = workbook.worksheets;
    }
    if (sheetsToProcess.length === 1) {
      const worksheet = sheetsToProcess[0];
      const flatOptions = {
        useFirstRowAsHeaders,
        includeEmptyRows,
        startRow
      };
      if (endRow !== void 0) {
        flatOptions.endRow = endRow;
      }
      if (startColumn !== void 0) {
        flatOptions.startColumn = startColumn;
      }
      if (endColumn !== void 0) {
        flatOptions.endColumn = endColumn;
      }
      const flatData = this.convertSheetToFlat(worksheet, flatOptions);
      return flatData;
    }
    const sheets = {};
    for (const worksheet of sheetsToProcess) {
      const flatOptions = {
        useFirstRowAsHeaders,
        includeEmptyRows,
        startRow
      };
      if (endRow !== void 0) {
        flatOptions.endRow = endRow;
      }
      if (startColumn !== void 0) {
        flatOptions.startColumn = startColumn;
      }
      if (endColumn !== void 0) {
        flatOptions.endColumn = endColumn;
      }
      const flatData = this.convertSheetToFlat(worksheet, flatOptions);
      sheets[worksheet.name] = flatData;
    }
    const result = {
      sheets,
      totalSheets: Object.keys(sheets).length
    };
    const hasMetadata = Object.values(metadata).some((val) => val !== void 0 && val !== null);
    if (hasMetadata) {
      result.metadata = metadata;
    }
    return result;
  }
  /**
   * Convert a single sheet to flat format
   */
  static convertSheetToFlat(worksheet, options) {
    const {
      useFirstRowAsHeaders,
      includeEmptyRows,
      startRow,
      endRow,
      startColumn,
      endColumn
    } = options;
    const actualStartRow = Math.max(startRow, 1);
    const actualEndRow = endRow || worksheet.rowCount || worksheet.lastRow?.number || 1;
    const actualStartCol = Math.max(startColumn || 1, 1);
    const actualEndCol = endColumn || worksheet.columnCount || worksheet.lastColumn?.number || 1;
    const data = [];
    let headers;
    if (useFirstRowAsHeaders) {
      const headerRow = worksheet.getRow(actualStartRow);
      headers = [];
      for (let colNum = actualStartCol; colNum <= actualEndCol; colNum++) {
        const cell = headerRow.getCell(colNum);
        headers.push(String(cell.value || `Column${colNum}`));
      }
    }
    const dataStartRow = useFirstRowAsHeaders ? actualStartRow + 1 : actualStartRow;
    for (let rowNum = dataStartRow; rowNum <= actualEndRow; rowNum++) {
      const excelRow = worksheet.getRow(rowNum);
      const rowValues = [];
      let hasData = false;
      for (let colNum = actualStartCol; colNum <= actualEndCol; colNum++) {
        const cell = excelRow.getCell(colNum);
        const cellValue = this.getCellValue(cell, { includeFormatting: false, includeFormulas: false, datesAsISO: true });
        rowValues.push(cellValue.value);
        if (cellValue.value !== null && cellValue.value !== void 0 && cellValue.value !== "") {
          hasData = true;
        }
      }
      if (!hasData && !includeEmptyRows) {
        continue;
      }
      if (useFirstRowAsHeaders && headers) {
        const rowObject = {};
        headers.forEach((header, index) => {
          rowObject[header] = rowValues[index];
        });
        data.push(rowObject);
      } else {
        data.push(rowValues);
      }
    }
    const result = {
      data,
      totalRows: data.length,
      sheet: worksheet.name
    };
    if (headers) {
      result.headers = headers;
    }
    return result;
  }
  /**
   * Get cell value with type information
   */
  static getCellValue(cell, options) {
    const { includeFormatting, includeFormulas, datesAsISO } = options;
    let value = cell.value;
    let type;
    let formattedValue;
    let formula;
    if (cell.type === ExcelJS.ValueType.Null || cell.value === null || cell.value === void 0) {
      value = null;
      type = "null";
    } else if (cell.type === ExcelJS.ValueType.Number) {
      value = cell.value;
      type = "number";
    } else if (cell.type === ExcelJS.ValueType.String) {
      value = cell.value;
      type = "string";
    } else if (cell.type === ExcelJS.ValueType.Date) {
      const dateValue = cell.value;
      value = datesAsISO ? dateValue.toISOString() : dateValue;
      type = "date";
    } else if (cell.type === ExcelJS.ValueType.Boolean) {
      value = cell.value;
      type = "boolean";
    } else if (cell.type === ExcelJS.ValueType.Formula) {
      if (includeFormulas && cell.formula) {
        formula = cell.formula;
        value = cell.result || cell.value;
        type = "formula";
      } else {
        value = cell.result || cell.value;
        type = typeof cell.result === "number" ? "number" : typeof cell.result === "string" ? "string" : "unknown";
      }
    } else if (cell.type === ExcelJS.ValueType.Hyperlink) {
      const hyperlinkValue = cell.value;
      if (typeof hyperlinkValue === "object" && hyperlinkValue !== null) {
        value = hyperlinkValue.text || hyperlinkValue.hyperlink || cell.value;
      } else {
        value = hyperlinkValue;
      }
      type = "hyperlink";
    } else {
      value = cell.value;
      type = "unknown";
    }
    if (includeFormatting && cell.numFmt) {
      formattedValue = String(value);
    }
    return {
      value,
      type,
      ...formattedValue && { formattedValue },
      ...formula && { formula }
    };
  }
  /**
   * Convert column number to letter (1 = A, 2 = B, 27 = AA, etc.)
   */
  static numberToColumnLetter(columnNumber) {
    let result = "";
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + columnNumber % 26) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }
}
class StyleBuilder {
  style = {};
  constructor() {
    this.style.alignment = {
      horizontal: HorizontalAlignment.CENTER,
      vertical: VerticalAlignment.MIDDLE,
      wrapText: true,
      shrinkToFit: true
    };
  }
  /**
   * Create a new StyleBuilder instance
   */
  static create() {
    return new StyleBuilder();
  }
  /**
   * Set font name
   */
  fontName(name) {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.name = name;
    return this;
  }
  /**
   * Set font size
   */
  fontSize(size) {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.size = size;
    return this;
  }
  /**
   * Set font style
   */
  fontStyle(style) {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.style = style;
    return this;
  }
  /**
   * Set font color
   */
  fontColor(color) {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.color = color;
    return this;
  }
  /**
   * Make font bold
   */
  fontBold() {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.bold = true;
    return this;
  }
  /**
   * Make font italic
   */
  fontItalic() {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.italic = true;
    return this;
  }
  /**
   * Make font underlined
   */
  fontUnderline() {
    if (!this.style.font) {
      this.style.font = {};
    }
    this.style.font.underline = true;
    return this;
  }
  /**
   * Set border on all sides
   */
  border(style, color) {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border = { style };
    if (color !== void 0) {
      border.color = color;
    }
    this.style.border.top = border;
    this.style.border.left = border;
    this.style.border.bottom = border;
    this.style.border.right = border;
    return this;
  }
  /**
   * Set top border
   */
  borderTop(style, color) {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border = { style };
    if (color !== void 0) {
      border.color = color;
    }
    this.style.border.top = border;
    return this;
  }
  /**
   * Set left border
   */
  borderLeft(style, color) {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border = { style };
    if (color !== void 0) {
      border.color = color;
    }
    this.style.border.left = border;
    return this;
  }
  /**
   * Set bottom border
   */
  borderBottom(style, color) {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border = { style };
    if (color !== void 0) {
      border.color = color;
    }
    this.style.border.bottom = border;
    return this;
  }
  /**
   * Set right border
   */
  borderRight(style, color) {
    if (!this.style.border) {
      this.style.border = {};
    }
    const border = { style };
    if (color !== void 0) {
      border.color = color;
    }
    this.style.border.right = border;
    return this;
  }
  /**
   * Set background color
   */
  backgroundColor(color) {
    if (!this.style.fill) {
      this.style.fill = { type: "pattern" };
    }
    this.style.fill.backgroundColor = color;
    this.style.fill.pattern = "solid";
    return this;
  }
  /**
   * Set horizontal alignment
   */
  horizontalAlign(alignment) {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = alignment;
    return this;
  }
  /**
   * Set vertical alignment
   */
  verticalAlign(alignment) {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.vertical = alignment;
    return this;
  }
  /**
   * Center align text
   */
  centerAlign() {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.CENTER;
    this.style.alignment.vertical = VerticalAlignment.MIDDLE;
    return this;
  }
  /**
   * Left align text
   */
  leftAlign() {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.LEFT;
    return this;
  }
  /**
   * Right align text
   */
  rightAlign() {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.horizontal = HorizontalAlignment.RIGHT;
    return this;
  }
  /**
   * Wrap text
   */
  wrapText() {
    if (!this.style.alignment) {
      this.style.alignment = {};
    }
    this.style.alignment.wrapText = true;
    return this;
  }
  /**
   * Set number format
   */
  numberFormat(format) {
    this.style.numberFormat = format;
    return this;
  }
  /**
   * Set striped rows
   */
  striped() {
    this.style.striped = true;
    return this;
  }
  /**
   * Add conditional formatting
   */
  conditionalFormat(format) {
    if (!this.style.conditionalFormats) {
      this.style.conditionalFormats = [];
    }
    this.style.conditionalFormats.push(format);
    return this;
  }
  /**
   * Build the final style
   */
  build() {
    return this.style;
  }
  /**
   * Reset the builder
   */
  reset() {
    this.style = {};
    this.style.alignment = {
      horizontal: HorizontalAlignment.CENTER,
      vertical: VerticalAlignment.MIDDLE,
      wrapText: true
    };
    return this;
  }
  /**
   * Clone the current style
   */
  clone() {
    const cloned = new StyleBuilder();
    cloned.style = JSON.parse(JSON.stringify(this.style));
    return cloned;
  }
}
var CellEventType = /* @__PURE__ */ ((CellEventType2) => {
  CellEventType2["CREATED"] = "created";
  CellEventType2["UPDATED"] = "updated";
  CellEventType2["DELETED"] = "deleted";
  CellEventType2["STYLED"] = "styled";
  CellEventType2["VALIDATED"] = "validated";
  return CellEventType2;
})(CellEventType || {});
var WorksheetEventType = /* @__PURE__ */ ((WorksheetEventType2) => {
  WorksheetEventType2["CREATED"] = "created";
  WorksheetEventType2["UPDATED"] = "updated";
  WorksheetEventType2["DELETED"] = "deleted";
  WorksheetEventType2["TABLE_ADDED"] = "tableAdded";
  WorksheetEventType2["TABLE_REMOVED"] = "tableRemoved";
  WorksheetEventType2["CELL_ADDED"] = "cellAdded";
  WorksheetEventType2["CELL_UPDATED"] = "cellUpdated";
  WorksheetEventType2["CELL_DELETED"] = "cellDeleted";
  return WorksheetEventType2;
})(WorksheetEventType || {});
var StylePreset = /* @__PURE__ */ ((StylePreset2) => {
  StylePreset2["HEADER"] = "header";
  StylePreset2["SUBHEADER"] = "subheader";
  StylePreset2["DATA"] = "data";
  StylePreset2["FOOTER"] = "footer";
  StylePreset2["TOTAL"] = "total";
  StylePreset2["HIGHLIGHT"] = "highlight";
  StylePreset2["WARNING"] = "warning";
  StylePreset2["ERROR"] = "error";
  StylePreset2["SUCCESS"] = "success";
  StylePreset2["INFO"] = "info";
  return StylePreset2;
})(StylePreset || {});
export {
  BorderStyle,
  BuilderEventType,
  CellEventType,
  CellType,
  ErrorType,
  EventEmitter,
  ExcelBuilder,
  ExcelReader,
  FontStyle,
  HorizontalAlignment,
  NumberFormat,
  StyleBuilder,
  StylePreset,
  VerticalAlignment,
  Worksheet,
  WorksheetEventType,
  ExcelBuilder as default
};
//# sourceMappingURL=han-excel.es.js.map
