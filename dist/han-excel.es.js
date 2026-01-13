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
    let rowPointer = 1;
    if (this.tables.length > 0) {
      for (let i = 0; i < this.tables.length; i++) {
        const table = this.tables[i];
        if (table) {
          rowPointer = await this.buildTable(ws, table, rowPointer, i > 0);
        }
      }
    } else {
      rowPointer = await this.buildLegacyContent(ws, rowPointer);
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
        ws.addRow([this.processCellValue(header)]);
        if (header.mergeCell) {
          const maxCols = this.calculateTableMaxColumns(table);
          ws.mergeCells(rowPointer, 1, rowPointer, maxCols);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell) => {
            cell.style = this.convertStyle(header.styles);
          });
        }
        this.applyCellDimensions(ws, rowPointer, 1, header);
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
    }
    if (footer.numberFormat) {
      footerCell.numFmt = footer.numberFormat;
    }
    this.applyCellDimensions(ws, rowPointer, footerColPosition, footer);
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
    mainCell.value = this.processCellValue(row);
    if (row.styles) {
      mainCell.style = this.convertStyle(row.styles);
    }
    if (row.numberFormat) {
      mainCell.numFmt = row.numberFormat;
    }
    this.applyCellDimensions(ws, rowPointer, mainColPosition, row);
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
   */
  getWorksheet(name) {
    return this.worksheets.get(name);
  }
  /**
   * Remove a worksheet by name
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
   * Generate and download the file
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
   * Get workbook as buffer
   */
  async toBuffer(options = {}) {
    return this.build(options);
  }
  /**
   * Get workbook as blob
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
   * Clear all worksheets
   */
  clear() {
    this.worksheets.clear();
    this.currentWorksheet = void 0;
  }
  /**
   * Get workbook statistics
   */
  getStats() {
    return { ...this.stats };
  }
  /**
   * Event handling methods
   */
  on(eventType, listener) {
    return this.eventEmitter.on(eventType, listener);
  }
  off(eventType, listenerId) {
    return this.eventEmitter.off(eventType, listenerId);
  }
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
  emitEvent(type, data) {
    const event = {
      type,
      data: data || {},
      timestamp: /* @__PURE__ */ new Date()
    };
    this.eventEmitter.emitSync(event);
  }
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
