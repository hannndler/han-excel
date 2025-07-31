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
    if (this.headers.length > 0) {
      this.headers.forEach((header) => {
        ws.addRow([header.value]);
        if (header.mergeCell) {
          ws.mergeCells(rowPointer, 1, rowPointer, this.getMaxColumns() || 1);
        }
        if (header.styles) {
          ws.getRow(rowPointer).eachCell((cell) => {
            cell.style = this.convertStyle(header.styles);
          });
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
    this.isBuilt = true;
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
          cell.value = headerInfo.value;
          if (headerInfo.style) {
            cell.style = this.convertStyle(headerInfo.style);
          }
          colIndex += headerInfo.colSpan;
        } else {
          if (header.children && header.children.length > 0) {
            for (const child of header.children) {
              const cell = row.getCell(colIndex);
              cell.value = typeof child.value === "string" ? child.value : String(child.value || "");
              if (child.styles || header.styles) {
                cell.style = this.convertStyle(child.styles || header.styles);
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
      console.log(" paso dios mio header", header.value);
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
    footerCell.value = footer.value;
    if (footer.styles) {
      footerCell.style = this.convertStyle(footer.styles);
    }
    if (footer.numberFormat) {
      footerCell.numFmt = footer.numberFormat;
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
    if (footer.jump) {
      return rowPointer + 1;
    }
    return rowPointer;
  }
  /**
   * Agrega una fila de datos y sus children recursivamente
   * @returns el siguiente rowPointer disponible
   */
  addDataRowRecursive(ws, rowPointer, row) {
    const columnPositions = this.calculateDataColumnPositions();
    console.log("columnPositions", columnPositions);
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
    mainCell.value = row.value;
    if (row.styles) {
      mainCell.style = this.convertStyle(row.styles);
    }
    if (row.numberFormat) {
      mainCell.numFmt = row.numberFormat;
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
    if (row.jump) {
      return rowPointer + 1;
    }
    return rowPointer;
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
