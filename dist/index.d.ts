import { CellValue } from 'exceljs';
import { default as default_2 } from 'exceljs';

/**
 * Border style options
 */
export declare enum BorderStyle {
    THIN = "thin",
    MEDIUM = "medium",
    THICK = "thick",
    DOTTED = "dotted",
    DASHED = "dashed",
    DOUBLE = "double",
    HAIR = "hair",
    MEDIUM_DASHED = "mediumDashed",
    DASH_DOT = "dashDot",
    MEDIUM_DASH_DOT = "mediumDashDot",
    DASH_DOT_DOT = "dashDotDot",
    MEDIUM_DASH_DOT_DOT = "mediumDashDotDot",
    SLANT_DASH_DOT = "slantDashDot"
}

/**
 * Builder event types
 */
export declare enum BuilderEventType {
    WORKSHEET_ADDED = "worksheetAdded",
    WORKSHEET_REMOVED = "worksheetRemoved",
    WORKSHEET_UPDATED = "worksheetUpdated",
    BUILD_STARTED = "buildStarted",
    BUILD_PROGRESS = "buildProgress",
    BUILD_COMPLETED = "buildCompleted",
    BUILD_ERROR = "buildError",
    DOWNLOAD_STARTED = "downloadStarted",
    DOWNLOAD_PROGRESS = "downloadProgress",
    DOWNLOAD_COMPLETED = "downloadCompleted",
    DOWNLOAD_ERROR = "downloadError"
}

/**
 * Cell event types
 */
export declare enum CellEventType {
    CREATED = "created",
    UPDATED = "updated",
    DELETED = "deleted",
    STYLED = "styled",
    VALIDATED = "validated"
}

/**
 * Supported cell data types
 */
export declare enum CellType {
    STRING = "string",
    NUMBER = "number",
    BOOLEAN = "boolean",
    DATE = "date",
    PERCENTAGE = "percentage",
    CURRENCY = "currency",
    LINK = "link",
    FORMULA = "formula"
}

/**
 * Color type - can be hex string, RGB object, or theme color
 */
export declare type Color = string | {
    r: number;
    g: number;
    b: number;
} | {
    theme: number;
};

declare type DetailedMapper = (data: IDetailedFormat) => unknown;

/**
 * Error types for validation
 */
export declare enum ErrorType {
    VALIDATION_ERROR = "VALIDATION_ERROR",
    BUILD_ERROR = "BUILD_ERROR",
    STYLE_ERROR = "STYLE_ERROR",
    WORKSHEET_ERROR = "WORKSHEET_ERROR",
    CELL_ERROR = "CELL_ERROR"
}

/**
 * EventEmitter class for handling events
 */
export declare class EventEmitter {
    private listeners;
    /**
     * Add an event listener
     */
    on<T = any>(type: string, listener: EventListener_3<T>, options?: EventListenerOptions_2): string;
    /**
     * Add a one-time event listener
     */
    once<T = any>(type: string, listener: EventListener_3<T>, options?: EventListenerOptions_2): string;
    /**
     * Remove an event listener
     */
    off(type: string, listenerId: string): boolean;
    /**
     * Remove all listeners for an event type
     */
    offAll(type: string): number;
    /**
     * Emit an event
     */
    emit<T = any>(event: T): Promise<void>;
    /**
     * Emit an event synchronously
     */
    emitSync<T = any>(event: T): void;
    /**
     * Clear all listeners
     */
    clear(): void;
    /**
     * Get listeners for an event type
     */
    getListeners(type: string): EventListenerRegistration[];
    /**
     * Get listener count for an event type
     */
    getListenerCount(type: string): number;
    /**
     * Get all registered event types
     */
    getEventTypes(): string[];
    private generateId;
    private cleanupInactiveListeners;
}

/**
 * Event listener function type
 */
declare type EventListener_2 = (event: IBuilderEvent) => void;
export { EventListener_2 as EventListener }

/**
 * Simple EventEmitter implementation
 */
/**
 * Event listener function type
 */
declare type EventListener_3<T = any> = (event: T) => void | Promise<void>;

/**
 * Event listener options
 */
declare interface EventListenerOptions_2 {
    /** Whether to execute the listener only once */
    once?: boolean;
    /** Whether to execute the listener asynchronously */
    async?: boolean;
    /** Priority of the listener (higher = executed first) */
    priority?: number;
    /** Whether to stop event propagation */
    stopPropagation?: boolean;
}

/**
 * Event listener registration
 */
declare interface EventListenerRegistration {
    /** Event type */
    type: string;
    /** Listener function */
    listener: EventListener_3;
    /** Listener options */
    options: EventListenerOptions_2;
    /** Registration ID */
    id: string;
    /** Whether the listener is active */
    active: boolean;
    /** Registration timestamp */
    timestamp: Date;
}

/**
 * ExcelBuilder class for creating Excel workbooks
 *
 * Main entry point for creating Excel files. Supports multiple worksheets, themes,
 * predefined styles, and comprehensive Excel features.
 *
 * @class ExcelBuilder
 * @implements {IExcelBuilder}
 */
declare class ExcelBuilder implements IExcelBuilder {
    config: IExcelBuilderConfig;
    worksheets: Map<string, IWorksheet>;
    currentWorksheet: IWorksheet | undefined;
    isBuilding: boolean;
    stats: IBuildStats;
    private eventEmitter;
    private cellStyles;
    private theme;
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
    constructor(config?: IExcelBuilderConfig);
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
    addWorksheet(name: string, worksheetConfig?: Partial<IWorksheetConfig>): IWorksheet;
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
    getWorksheet(name: string): IWorksheet | undefined;
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
    removeWorksheet(name: string): boolean;
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
    setCurrentWorksheet(name: string): boolean;
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
    build(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
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
    generateAndDownload(fileName: string, options?: IDownloadOptions): Promise<Result<void>>;
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
    saveToFile(filePath: string, options?: ISaveFileOptions): Promise<Result<void>>;
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
    saveToStream(writeStream: {
        write: (chunk: any, callback?: (error?: Error | null) => void) => boolean;
    }, options?: IBuildOptions): Promise<Result<void>>;
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
    toBuffer(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
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
    toBlob(options?: IBuildOptions): Promise<Result<Blob>>;
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
    validate(): Result<boolean>;
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
    clear(): void;
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
    getStats(): IBuildStats;
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
    addCellStyle(name: string, style: IStyle): this;
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
    getCellStyle(name: string): IStyle | undefined;
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
    setTheme(theme: IWorkbookTheme): this;
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
    getTheme(): IWorkbookTheme | undefined;
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
    on(eventType: BuilderEventType, listener: (event: IBuilderEvent) => void): string;
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
    off(eventType: BuilderEventType, listenerId: string): boolean;
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
    removeAllListeners(eventType?: BuilderEventType): void;
    /**
     * Private methods
     */
    /**
     * Emit an event to all registered listeners
     * @private
     */
    private emitEvent;
    /**
     * Initialize build statistics
     * @private
     */
    private initializeStats;
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
    private applyTheme;
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
    private convertColorToTheme;
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
    private addStyleToWorkbook;
}
export { ExcelBuilder }
export default ExcelBuilder;

/**
 * ExcelReader class for reading Excel files and converting to JSON
 */
export declare class ExcelReader {
    /**
     * Read Excel file from ArrayBuffer
     */
    static fromBuffer<T extends OutputFormat = OutputFormat.WORKSHEET>(buffer: ArrayBuffer, options?: IExcelReaderOptions): Promise<ExcelReaderResult<T>>;
    /**
     * Read Excel file from Blob
     */
    static fromBlob<T extends OutputFormat = OutputFormat.WORKSHEET>(blob: Blob, options?: IExcelReaderOptions): Promise<ExcelReaderResult<T>>;
    /**
     * Read Excel file from File (browser)
     */
    static fromFile<T extends OutputFormat = OutputFormat.WORKSHEET>(file: File, options?: IExcelReaderOptions): Promise<ExcelReaderResult<T>>;
    /**
     * Read Excel file from path (Node.js)
     * Note: This method only works in Node.js environment
     */
    /**
     * Read Excel file from path (Node.js only)
     * Note: This method only works in Node.js environment
     */
    static fromPath<T extends OutputFormat = OutputFormat.WORKSHEET>(filePath: string, options?: IExcelReaderOptions): Promise<ExcelReaderResult<T>>;
    /**
     * Convert ExcelJS Workbook to JSON
     */
    private static convertWorkbookToJson;
    /**
     * Convert ExcelJS Worksheet to JSON
     */
    private static convertSheetToJson;
    /**
     * Convert ExcelJS Cell to JSON
     */
    private static convertCellToJson;
    /**
     * Convert workbook to detailed format (with position information)
     */
    private static convertToDetailedFormat;
    /**
     * Convert workbook to flat format (just data)
     */
    private static convertToFlatFormat;
    /**
     * Convert a single sheet to flat format
     */
    private static convertSheetToFlat;
    /**
     * Get cell value with type information
     */
    private static getCellValue;
    /**
     * Convert column number to letter (1 = A, 2 = B, 27 = AA, etc.)
     */
    private static numberToColumnLetter;
}

/**
 * Reader result - generic type based on output format
 */
declare type ExcelReaderResult<T extends OutputFormat = OutputFormat.WORKSHEET> = T extends OutputFormat.DETAILED ? Result<IDetailedFormat> & {
    processingTime?: number;
} : T extends OutputFormat.FLAT ? Result<IFlatFormat | IFlatFormatMultiSheet> & {
    processingTime?: number;
} : Result<IJsonWorkbook> & {
    processingTime?: number;
};

declare type FlatMapper = (data: IFlatFormat | IFlatFormatMultiSheet) => unknown;

/**
 * Font style options
 */
export declare enum FontStyle {
    NORMAL = "normal",
    BOLD = "bold",
    ITALIC = "italic",
    BOLD_ITALIC = "bold italic"
}

/**
 * Horizontal alignment options
 */
export declare enum HorizontalAlignment {
    LEFT = "left",
    CENTER = "center",
    RIGHT = "right",
    FILL = "fill",
    JUSTIFY = "justify",
    CENTER_CONTINUOUS = "centerContinuous",
    DISTRIBUTED = "distributed"
}

/**
 * Alignment configuration interface
 */
export declare interface IAlignment {
    /** Horizontal alignment */
    horizontal?: HorizontalAlignment;
    /** Vertical alignment */
    vertical?: VerticalAlignment;
    /** Text rotation (0-180 degrees) */
    textRotation?: number;
    /** Whether to wrap text */
    wrapText?: boolean;
    /** Whether to shrink text to fit */
    shrinkToFit?: boolean;
    /** Indent level */
    indent?: number;
    /** Whether to merge cells */
    mergeCell?: boolean;
    /** Reading order */
    readingOrder?: 'left-to-right' | 'right-to-left';
}

/**
 * Base cell properties interface
 */
export declare interface IBaseCell {
    /** Unique identifier for the cell */
    key: string;
    /** Cell data type */
    type: CellType;
    /** Cell value */
    value: CellValue;
    /** Optional cell reference (e.g., A1, B2) */
    reference?: string;
    /** Whether to merge this cell with others */
    mergeCell?: boolean;
    /** Number of columns to merge (if mergeCell is true) */
    mergeTo?: number;
    /** Row height for this cell */
    rowHeight?: number;
    /** Column width for this cell */
    colWidth?: number;
    /** Whether to move to next row after this cell */
    jump?: boolean;
    /** Hyperlink URL */
    link?: string;
    /** Text mask for hyperlink (displayed text when link is present) */
    mask?: string;
    /** Excel formula */
    formula?: string;
    /** Number format for numeric cells */
    numberFormat?: NumberFormat | string;
    /** Custom number format string */
    customNumberFormat?: string;
    /** Whether the cell is protected */
    protected?: boolean;
    /** Whether the cell is hidden */
    hidden?: boolean;
    /** Cell comment */
    comment?: string;
    /** Data validation rules */
    validation?: IDataValidation;
    /** Optional styles for the cell */
    styles?: IStyle;
    /** Predefined style name (references a style added via addCellStyle) */
    styleName?: string;
    /** Legacy children cells */
    childrens?: IBaseCell[];
    /** Modern children cells */
    children?: IBaseCell[];
}

/**
 * Border configuration interface
 */
export declare interface IBorder {
    /** Border style */
    style?: BorderStyle;
    /** Border color */
    color?: Color;
    /** Border width */
    width?: number;
}

/**
 * Border sides interface
 */
export declare interface IBorderSides {
    /** Top border */
    top?: IBorder;
    /** Left border */
    left?: IBorder;
    /** Bottom border */
    bottom?: IBorder;
    /** Right border */
    right?: IBorder;
    /** Diagonal border */
    diagonal?: IBorder;
    /** Diagonal direction */
    diagonalDirection?: 'up' | 'down' | 'both';
}

/**
 * Builder event interface
 */
export declare interface IBuilderEvent {
    type: BuilderEventType;
    data?: Record<string, unknown>;
    timestamp: Date;
}

/**
 * Builder event listener interface
 */
export declare interface IBuilderEventListener {
    (event: IBuilderEvent): void;
}

/**
 * Builder validation result interface
 */
export declare interface IBuilderValidationResult {
    /** Whether the builder is valid */
    isValid: boolean;
    /** Validation errors */
    errors: string[];
    /** Validation warnings */
    warnings: string[];
    /** Worksheet validation results */
    worksheetResults: Map<string, boolean>;
}

/**
 * Build options interface
 */
export declare interface IBuildOptions {
    /** Output format */
    format?: 'xlsx' | 'xls' | 'csv';
    /** Whether to include styles */
    includeStyles?: boolean;
    /** Whether to include formulas */
    includeFormulas?: boolean;
    /** Whether to include comments */
    includeComments?: boolean;
    /** Whether to include data validation */
    includeValidation?: boolean;
    /** Whether to include conditional formatting */
    includeConditionalFormatting?: boolean;
    /** Compression level (0-9) */
    compressionLevel?: number;
    /** Whether to optimize for size */
    optimizeForSize?: boolean;
    /** Whether to optimize for speed */
    optimizeForSpeed?: boolean;
}

/**
 * Build statistics interface
 */
export declare interface IBuildStats {
    /** Total number of worksheets */
    totalWorksheets: number;
    /** Total number of cells */
    totalCells: number;
    /** Total memory usage in bytes */
    memoryUsage: number;
    /** Build time in milliseconds */
    buildTime: number;
    /** File size in bytes */
    fileSize: number;
    /** Number of styles used */
    stylesUsed: number;
    /** Number of formulas used */
    formulasUsed: number;
    /** Number of conditional formats used */
    conditionalFormatsUsed: number;
    /** Performance metrics */
    performance: {
        /** Time spent building headers */
        headersTime: number;
        /** Time spent building data */
        dataTime: number;
        /** Time spent applying styles */
        stylesTime: number;
        /** Time spent writing to buffer */
        writeTime: number;
    };
}

/**
 * Cell data for different types
 */
export declare interface ICellData {
    /** String cell data */
    string?: {
        value: string;
        maxLength?: number;
        trim?: boolean;
    };
    /** Number cell data */
    number?: {
        value: number;
        min?: number;
        max?: number;
        precision?: number;
        allowNegative?: boolean;
    };
    /** Date cell data */
    date?: {
        value: Date;
        min?: Date;
        max?: Date;
        format?: string;
    };
    /** Boolean cell data */
    boolean?: {
        value: boolean;
        trueText?: string;
        falseText?: string;
    };
    /** Percentage cell data */
    percentage?: {
        value: number;
        min?: number;
        max?: number;
        precision?: number;
        showSymbol?: boolean;
    };
    /** Currency cell data */
    currency?: {
        value: number;
        currency?: string;
        precision?: number;
        showSymbol?: boolean;
    };
    /** Link cell data */
    link?: {
        value: string;
        text?: string;
        tooltip?: string;
    };
    /** Formula cell data */
    formula?: {
        value: string;
        result?: CellValue;
    };
}

/**
 * Cell event interface
 */
export declare interface ICellEvent {
    type: CellEventType;
    cell: IDataCell | IHeaderCell | IFooterCell;
    position: ICellPosition;
    timestamp: Date;
    data?: Record<string, unknown>;
}

/**
 * Cell position interface
 */
export declare interface ICellPosition {
    /** Row index (1-based) */
    row: number;
    /** Column index (1-based) */
    col: number;
    /** Cell reference (e.g., A1) */
    reference: string;
}

/**
 * Cell range interface
 */
export declare interface ICellRange {
    /** Start position */
    start: ICellPosition;
    /** End position */
    end: ICellPosition;
    /** Range reference (e.g., A1:B10) */
    reference: string;
}

/**
 * Cell type validation interface
 */
export declare interface ICellTypeValidation {
    /** Expected cell type */
    expectedType: CellType;
    /** Whether to allow null/undefined values */
    allowNull?: boolean;
    /** Whether to allow empty strings */
    allowEmpty?: boolean;
    /** Type conversion options */
    conversion?: {
        /** Whether to attempt type conversion */
        enabled: boolean;
        /** Whether to be strict about conversion */
        strict: boolean;
    };
}

/**
 * Cell validation result
 */
export declare interface ICellValidationResult {
    /** Whether the cell is valid */
    isValid: boolean;
    /** Validation errors */
    errors: string[];
    /** Validation warnings */
    warnings: string[];
}

/**
 * Conditional formatting interface
 */
export declare interface IConditionalFormat {
    /** Condition type */
    type: 'cellIs' | 'containsText' | 'beginsWith' | 'endsWith' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors' | 'timePeriod' | 'top' | 'bottom' | 'aboveAverage' | 'belowAverage' | 'duplicateValues' | 'uniqueValues' | 'expression' | 'colorScale' | 'dataBar' | 'iconSet';
    /** Condition operator */
    operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
    /** Condition values */
    values?: Array<string | number | Date>;
    /** Condition formula */
    formula?: string;
    /** Style to apply when condition is met */
    style?: IStyle;
    /** Priority of the condition */
    priority?: number;
    /** Whether to stop if true */
    stopIfTrue?: boolean;
}

/**
 * Data cell interface
 */
export declare interface IDataCell extends IBaseCell {
    /** Reference to header key */
    header: string;
    /** Reference to main header key */
    mainHeaderKey?: string;
    /** Child data cells */
    children?: IDataCell[];
    /** Whether this cell has alternating row color */
    striped?: boolean;
    /** Row index */
    rowIndex?: number;
    /** Column index */
    colIndex?: number;
}

/**
 * Data connection configuration
 */
export declare interface IDataConnection {
    /** Connection name */
    name: string;
    /** Connection type */
    type: 'odbc' | 'oledb' | 'web' | 'text' | 'xml';
    /** Connection string or URL */
    connectionString: string;
    /** Command text (SQL query, etc.) */
    commandText?: string;
    /** Refresh settings */
    refresh?: {
        /** Auto refresh on open */
        refreshOnOpen?: boolean;
        /** Refresh interval in minutes */
        refreshInterval?: number;
    };
    /** Credentials */
    credentials?: {
        /** Username */
        username?: string;
        /** Password */
        password?: string;
        /** Integrated security */
        integratedSecurity?: boolean;
    };
}

/**
 * Data validation interface
 */
export declare interface IDataValidation {
    /** Validation type */
    type: 'list' | 'whole' | 'decimal' | 'textLength' | 'date' | 'time' | 'custom';
    /** Validation operator */
    operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
    /** Validation formula or values */
    formula1?: string | number | Date;
    /** Second validation formula or value (for between/notBetween) */
    formula2?: string | number | Date;
    /** Whether to show error message */
    showErrorMessage?: boolean;
    /** Error message text */
    errorMessage?: string;
    /** Whether to show input message */
    showInputMessage?: boolean;
    /** Input message text */
    inputMessage?: string;
    /** Whether to allow blank values */
    allowBlank?: boolean;
}

/**
 * Detailed cell format - includes position information
 */
declare interface IDetailedCell {
    /** Cell value */
    value: unknown;
    /** Cell text (string representation) */
    text: string;
    /** Column number (1-based) */
    column: number;
    /** Column letter (e.g., A, B, C) */
    columnLetter: string;
    /** Row number (1-based) */
    row: number;
    /** Cell reference (e.g., A1) */
    reference: string;
    /** Sheet name */
    sheet: string;
    /** Cell type */
    type?: string;
    /** Formatted value (if includeFormatting is true) */
    formattedValue?: string;
    /** Formula (if includeFormulas is true) */
    formula?: string;
    /** Cell comment */
    comment?: string;
}

/**
 * Detailed format result - array of cells with position
 */
declare interface IDetailedFormat {
    /** Array of all cells with detailed information */
    cells: IDetailedCell[];
    /** Total number of cells */
    totalCells: number;
    /** Workbook metadata */
    metadata?: {
        title?: string;
        author?: string;
        company?: string;
        created?: Date | string;
        modified?: Date | string;
        description?: string;
    };
}

/**
 * Download options interface
 */
export declare interface IDownloadOptions extends IBuildOptions {
    /** File name */
    fileName?: string;
    /** Whether to show download progress */
    showProgress?: boolean;
    /** Progress callback */
    onProgress?: (progress: number) => void;
    /** Whether to auto-download */
    autoDownload?: boolean;
    /** MIME type */
    mimeType?: string;
}

/**
 * Error interface
 */
export declare interface IError {
    type: ErrorType;
    message: string;
    code?: string;
    details?: Record<string, unknown>;
    stack?: string;
}

/**
 * Error result interface
 */
export declare interface IErrorResult {
    success: false;
    error: IError;
}

/**
 * Event emitter interface
 */
export declare interface IEventEmitter {
    on(event: BuilderEventType, listener: EventListener_2): void;
    off(event: BuilderEventType, listener: EventListener_2): void;
    emit(event: BuilderEventType, data?: Record<string, unknown>): void;
    removeAllListeners(event?: BuilderEventType): void;
}

/**
 * Excel builder interface
 */
export declare interface IExcelBuilder {
    /** Builder configuration */
    config: IExcelBuilderConfig;
    /** Worksheets in the workbook */
    worksheets: Map<string, IWorksheet>;
    /** Current worksheet */
    currentWorksheet: IWorksheet | undefined;
    /** Whether the builder is building */
    isBuilding: boolean;
    /** Build statistics */
    stats: IBuildStats;
    /** Add a new worksheet */
    addWorksheet(name: string, config?: Partial<IWorksheetConfig>): IWorksheet;
    /** Get a worksheet by name */
    getWorksheet(name: string): IWorksheet | undefined;
    /** Remove a worksheet */
    removeWorksheet(name: string): boolean;
    /** Set the current worksheet */
    setCurrentWorksheet(name: string): boolean;
    /** Build the workbook */
    build(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
    /** Generate and download the file */
    generateAndDownload(fileName: string, options?: IDownloadOptions): Promise<Result<void>>;
    /** Save file to disk (Node.js only) - Similar to generateAndDownload but for Node.js */
    saveToFile(filePath: string, options?: ISaveFileOptions): Promise<Result<void>>;
    /** Save to stream (Node.js only) - For large files */
    saveToStream(writeStream: {
        write: (chunk: any, callback?: (error?: Error | null) => void) => boolean;
    }, options?: IBuildOptions): Promise<Result<void>>;
    /** Get workbook as buffer */
    toBuffer(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
    /** Get workbook as blob */
    toBlob(options?: IBuildOptions): Promise<Result<Blob>>;
    /** Validate the workbook */
    validate(): Result<boolean>;
    /** Clear all worksheets */
    clear(): void;
    /** Get workbook statistics */
    getStats(): IBuildStats;
    /** Add a predefined cell style */
    addCellStyle(name: string, style: IStyle): this;
    /** Get a predefined cell style by name */
    getCellStyle(name: string): IStyle | undefined;
    /** Set workbook theme */
    setTheme(theme: IWorkbookTheme): this;
    /** Get current workbook theme */
    getTheme(): IWorkbookTheme | undefined;
}

/**
 * Excel builder configuration interface
 */
export declare interface IExcelBuilderConfig {
    /** Workbook metadata */
    metadata?: IWorkbookMetadata;
    /** Default worksheet configuration */
    defaultWorksheetConfig?: Partial<IWorksheetConfig>;
    /** Default styles */
    defaultStyles?: {
        header?: IStyle;
        subheader?: IStyle;
        data?: IStyle;
        footer?: IStyle;
        total?: IStyle;
    };
    /** Whether to enable validation */
    enableValidation?: boolean;
    /** Whether to enable events */
    enableEvents?: boolean;
    /** Whether to enable performance monitoring */
    enablePerformanceMonitoring?: boolean;
    /** Maximum number of worksheets */
    maxWorksheets?: number;
    /** Maximum number of rows per worksheet */
    maxRowsPerWorksheet?: number;
    /** Maximum number of columns per worksheet */
    maxColumnsPerWorksheet?: number;
    /** Memory limit in bytes */
    memoryLimit?: number;
}

/**
 * Options for reading Excel files
 */
declare interface IExcelReaderOptions {
    /** Output format (default: 'worksheet') */
    outputFormat?: OutputFormat | 'worksheet' | 'detailed' | 'flat';
    /** Mapper function to transform the response data */
    mapper?: WorksheetMapper | DetailedMapper | FlatMapper;
    /** Whether to include empty rows */
    includeEmptyRows?: boolean;
    /** Whether to use first row as headers */
    useFirstRowAsHeaders?: boolean;
    /** Custom headers mapping (column index -> header name) */
    headers?: string[] | Record<number, string>;
    /** Sheet name or index to read (if not specified, reads all sheets) */
    sheetName?: string | number;
    /** Starting row (1-based, default: 1) */
    startRow?: number;
    /** Ending row (1-based, if not specified, reads until end) */
    endRow?: number;
    /** Starting column (1-based, default: 1) */
    startColumn?: number;
    /** Ending column (1-based, if not specified, reads until end) */
    endColumn?: number;
    /** Whether to include cell formatting information */
    includeFormatting?: boolean;
    /** Whether to include formulas */
    includeFormulas?: boolean;
    /** Date format for date cells */
    dateFormat?: string;
    /** Whether to convert dates to ISO strings */
    datesAsISO?: boolean;
}

/**
 * Excel structured table configuration
 */
export declare interface IExcelTable {
    /** Table name */
    name: string;
    /** Table range (start and end cells) */
    range: {
        /** Start cell reference (e.g., 'A1') */
        start: string;
        /** End cell reference (e.g., 'D10') */
        end: string;
    };
    /** Table style */
    style?: 'TableStyleLight1' | 'TableStyleLight2' | 'TableStyleLight3' | 'TableStyleLight4' | 'TableStyleLight5' | 'TableStyleLight6' | 'TableStyleLight7' | 'TableStyleLight8' | 'TableStyleLight9' | 'TableStyleLight10' | 'TableStyleLight11' | 'TableStyleLight12' | 'TableStyleLight13' | 'TableStyleLight14' | 'TableStyleLight15' | 'TableStyleLight16' | 'TableStyleLight17' | 'TableStyleLight18' | 'TableStyleLight19' | 'TableStyleLight20' | 'TableStyleLight21' | 'TableStyleMedium1' | 'TableStyleMedium2' | 'TableStyleMedium3' | 'TableStyleMedium4' | 'TableStyleMedium5' | 'TableStyleMedium6' | 'TableStyleMedium7' | 'TableStyleMedium8' | 'TableStyleMedium9' | 'TableStyleMedium10' | 'TableStyleMedium11' | 'TableStyleMedium12' | 'TableStyleMedium13' | 'TableStyleMedium14' | 'TableStyleMedium15' | 'TableStyleMedium16' | 'TableStyleMedium17' | 'TableStyleMedium18' | 'TableStyleMedium19' | 'TableStyleMedium20' | 'TableStyleMedium21' | 'TableStyleMedium22' | 'TableStyleMedium23' | 'TableStyleMedium24' | 'TableStyleMedium25' | 'TableStyleMedium26' | 'TableStyleMedium27' | 'TableStyleMedium28' | 'TableStyleDark1' | 'TableStyleDark2' | 'TableStyleDark3' | 'TableStyleDark4' | 'TableStyleDark5' | 'TableStyleDark6' | 'TableStyleDark7' | 'TableStyleDark8' | 'TableStyleDark9' | 'TableStyleDark10' | 'TableStyleDark11';
    /** Whether to show header row */
    headerRow?: boolean;
    /** Whether to show total row */
    totalRow?: boolean;
    /** Column definitions */
    columns?: Array<{
        /** Column name */
        name: string;
        /** Column filter button */
        filterButton?: boolean;
        /** Totals row function */
        totalsRowFunction?: 'none' | 'sum' | 'min' | 'max' | 'average' | 'count' | 'countNums' | 'stdDev' | 'var' | 'custom';
        /** Totals row formula */
        totalsRowFormula?: string;
    }>;
}

/**
 * Fill pattern interface
 */
export declare interface IFill {
    /** Fill type */
    type: 'pattern' | 'gradient';
    /** Pattern type (for pattern fills) */
    pattern?: 'none' | 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray125' | 'gray0625' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis';
    /** Background color */
    backgroundColor?: Color;
    /** Foreground color */
    foregroundColor?: Color;
    /** Gradient type (for gradient fills) */
    gradient?: 'linear' | 'path';
    /** Gradient stops */
    stops?: Array<{
        position: number;
        color: Color;
    }>;
    /** Gradient angle (for linear gradients) */
    angle?: number;
}

/**
 * Flat format result - just the data values
 */
declare interface IFlatFormat {
    /** Array of row data (as objects if useFirstRowAsHeaders is true, or as arrays) */
    data: Array<Record<string, unknown> | unknown[]>;
    /** Headers (if useFirstRowAsHeaders is true) */
    headers?: string[];
    /** Sheet name */
    sheet?: string;
    /** Total number of rows */
    totalRows: number;
}

/**
 * Flat format result for multiple sheets
 */
declare interface IFlatFormatMultiSheet {
    /** Data organized by sheet name */
    sheets: Record<string, IFlatFormat>;
    /** Total number of sheets */
    totalSheets: number;
    /** Workbook metadata */
    metadata?: {
        title?: string;
        author?: string;
        company?: string;
        created?: Date | string;
        modified?: Date | string;
        description?: string;
    };
}

/**
 * Font configuration interface
 */
export declare interface IFont {
    /** Font name */
    name?: string;
    /** Font size */
    size?: number;
    /** Font style */
    style?: FontStyle;
    /** Font color */
    color?: Color;
    /** Whether the font is bold */
    bold?: boolean;
    /** Whether the font is italic */
    italic?: boolean;
    /** Whether the font is underlined */
    underline?: boolean;
    /** Whether the font is strikethrough */
    strikethrough?: boolean;
    /** Font family */
    family?: string;
    /** Font scheme */
    scheme?: 'major' | 'minor' | 'none';
}

/**
 * Footer cell interface
 */
export declare interface IFooterCell extends IBaseCell {
    /** Reference to header key */
    header: string;
    /** Child footer cells */
    children?: IDataCell[];
    /** Whether this is a total row */
    isTotal?: boolean;
    /** Footer type */
    footerType?: 'total' | 'subtotal' | 'average' | 'count' | 'custom';
}

/**
 * Format validation interface
 */
export declare interface IFormatValidation {
    /** Format pattern (regex or format string) */
    pattern: string | RegExp;
    /** Format type */
    type: 'regex' | 'date' | 'email' | 'url' | 'phone' | 'custom';
    /** Custom format function */
    formatFunction?: (value: unknown) => boolean;
}

/**
 * Header cell interface
 */
export declare interface IHeaderCell extends IBaseCell {
    /** Reference to parent header key */
    mainHeaderKey?: string;
    /** Child headers */
    children?: IHeaderCell[];
    /** Whether this is a main header */
    isMainHeader?: boolean;
    /** Header level (1 = main, 2 = sub, etc.) */
    level?: number;
}

/**
 * Cell data in JSON format
 */
declare interface IJsonCell {
    /** Cell value */
    value: unknown;
    /** Cell type */
    type?: string;
    /** Cell reference (e.g., A1) */
    reference?: string;
    /** Formatted value (if includeFormatting is true) */
    formattedValue?: string;
    /** Formula (if includeFormulas is true) */
    formula?: string;
    /** Cell comment */
    comment?: string;
}

/**
 * Row data in JSON format
 */
declare interface IJsonRow {
    /** Row number (1-based) */
    rowNumber: number;
    /** Cells in the row */
    cells: IJsonCell[];
    /** Row as object (if useFirstRowAsHeaders is true) */
    data?: Record<string, unknown>;
}

/**
 * Sheet data in JSON format
 */
declare interface IJsonSheet {
    /** Sheet name */
    name: string;
    /** Sheet index */
    index: number;
    /** Rows in the sheet */
    rows: IJsonRow[];
    /** Headers (if useFirstRowAsHeaders is true) */
    headers?: string[];
    /** Total number of rows */
    totalRows: number;
    /** Total number of columns */
    totalColumns: number;
}

/**
 * Workbook data in JSON format
 */
declare interface IJsonWorkbook {
    /** Workbook metadata */
    metadata?: {
        title?: string;
        author?: string;
        company?: string;
        created?: Date | string;
        modified?: Date | string;
        description?: string;
    };
    /** Sheets in the workbook */
    sheets: IJsonSheet[];
    /** Total number of sheets */
    totalSheets: number;
}

/**
 * Length validation interface
 */
export declare interface ILengthValidation {
    /** Minimum length */
    min?: number;
    /** Maximum length */
    max?: number;
    /** Exact length */
    exact?: number;
    /** Whether to trim whitespace before validation */
    trim?: boolean;
}

/**
 * Pivot table configuration
 */
export declare interface IPivotTable {
    /** Pivot table name */
    name: string;
    /** Reference cell where pivot table starts (e.g., 'A1') */
    ref: string;
    /** Source data range (e.g., 'A1:D100') */
    sourceRange: string;
    /** Source worksheet name (if different from current) */
    sourceSheet?: string;
    /** Pivot table fields configuration */
    fields: {
        /** Rows fields */
        rows?: string[];
        /** Columns fields */
        columns?: string[];
        /** Values fields with aggregation function */
        values?: Array<{
            name: string;
            stat: 'sum' | 'count' | 'average' | 'min' | 'max' | 'stdDev' | 'var';
        }>;
        /** Filters fields */
        filters?: string[];
    };
    /** Pivot table options */
    options?: {
        /** Show grand totals for rows */
        showRowGrandTotals?: boolean;
        /** Show grand totals for columns */
        showColGrandTotals?: boolean;
        /** Show headers */
        showHeaders?: boolean;
    };
}

/**
 * Protection configuration interface
 */
export declare interface IProtection {
    /** Whether the cell is locked */
    locked?: boolean;
    /** Whether the cell is hidden */
    hidden?: boolean;
}

/**
 * Range validation interface
 */
export declare interface IRangeValidation {
    /** Minimum value */
    min?: number | Date | string;
    /** Maximum value */
    max?: number | Date | string;
    /** Whether the range is inclusive */
    inclusive?: boolean;
    /** Custom range function */
    rangeFunction?: (value: unknown) => boolean;
}

/**
 * Reference validation interface
 */
export declare interface IReferenceValidation {
    /** Reference type */
    type: 'formula' | 'hyperlink' | 'comment' | 'validation';
    /** Reference target */
    target: string;
    /** Whether the reference is required */
    required?: boolean;
    /** Reference validation function */
    validateReference?: (reference: string) => boolean;
}

/**
 * Rich text run interface (for formatted text within a cell)
 */
export declare interface IRichTextRun {
    /** Text content */
    text: string;
    /** Font name */
    font?: string;
    /** Font size */
    size?: number;
    /** Font color */
    color?: string | {
        r: number;
        g: number;
        b: number;
    } | {
        theme: number;
    };
    /** Bold */
    bold?: boolean;
    /** Italic */
    italic?: boolean;
    /** Underline */
    underline?: boolean;
    /** Strikethrough */
    strikethrough?: boolean;
}

/**
 * Save file options interface (for Node.js)
 */
export declare interface ISaveFileOptions extends IBuildOptions {
    /** Whether to create parent directories if they don't exist (default: true) */
    createDir?: boolean;
    /** File encoding (default: 'binary') */
    encoding?: 'ascii' | 'utf8' | 'utf-8' | 'utf16le' | 'ucs2' | 'ucs-2' | 'base64' | 'latin1' | 'binary' | 'hex';
}

/**
 * Slicer configuration for tables and pivot tables
 */
export declare interface ISlicer {
    /** Slicer name */
    name: string;
    /** Target table or pivot table name */
    targetTable: string;
    /** Column/field to create slicer for */
    column: string;
    /** Position where slicer should be placed */
    position: {
        /** Row number (1-based) */
        row: number;
        /** Column number or letter (1-based or 'A', 'B', etc.) */
        col: number | string;
    };
    /** Slicer size */
    size?: {
        /** Width in pixels */
        width?: number;
        /** Height in pixels */
        height?: number;
    };
    /** Slicer style */
    style?: {
        /** Caption style */
        caption?: string;
        /** Item style */
        itemStyle?: string;
    };
}

/**
 * Main style interface
 */
export declare interface IStyle {
    /** Font configuration */
    font?: IFont;
    /** Border configuration */
    border?: IBorderSides;
    /** Fill configuration */
    fill?: IFill;
    /** Alignment configuration */
    alignment?: IAlignment;
    /** Protection configuration */
    protection?: IProtection;
    /** Conditional formatting */
    conditionalFormats?: IConditionalFormat[];
    /** Number format */
    numberFormat?: string;
    /** Whether to apply alternating row colors */
    striped?: boolean;
    /** Custom CSS-like properties */
    custom?: Record<string, unknown>;
}

/**
 * Style builder interface
 */
export declare interface IStyleBuilder {
    /** Set font name */
    fontName(name: string): IStyleBuilder;
    /** Set font size */
    fontSize(size: number): IStyleBuilder;
    /** Set font style */
    fontStyle(style: FontStyle): IStyleBuilder;
    /** Set font color */
    fontColor(color: Color): IStyleBuilder;
    /** Make font bold */
    fontBold(): IStyleBuilder;
    /** Make font italic */
    fontItalic(): IStyleBuilder;
    /** Make font underlined */
    fontUnderline(): IStyleBuilder;
    /** Set border */
    border(style: BorderStyle, color?: Color): IStyleBuilder;
    /** Set specific border */
    borderTop(style: BorderStyle, color?: Color): IStyleBuilder;
    borderLeft(style: BorderStyle, color?: Color): IStyleBuilder;
    borderBottom(style: BorderStyle, color?: Color): IStyleBuilder;
    borderRight(style: BorderStyle, color?: Color): IStyleBuilder;
    /** Set background color */
    backgroundColor(color: Color): IStyleBuilder;
    /** Set horizontal alignment */
    horizontalAlign(alignment: HorizontalAlignment): IStyleBuilder;
    /** Set vertical alignment */
    verticalAlign(alignment: VerticalAlignment): IStyleBuilder;
    /** Center align text */
    centerAlign(): IStyleBuilder;
    /** Left align text */
    leftAlign(): IStyleBuilder;
    /** Right align text */
    rightAlign(): IStyleBuilder;
    /** Wrap text */
    wrapText(): IStyleBuilder;
    /** Set number format */
    numberFormat(format: string): IStyleBuilder;
    /** Set striped rows */
    striped(): IStyleBuilder;
    /** Add conditional formatting */
    conditionalFormat(format: IConditionalFormat): IStyleBuilder;
    /** Build the final style */
    build(): IStyle;
}

/**
 * Style theme interface
 */
export declare interface IStyleTheme {
    /** Theme name */
    name: string;
    /** Theme description */
    description?: string;
    /** Color palette */
    colors: {
        primary: Color;
        secondary: Color;
        accent: Color;
        background: Color;
        text: Color;
        border: Color;
        success: Color;
        warning: Color;
        error: Color;
        info: Color;
    };
    /** Font family */
    fontFamily: string;
    /** Base font size */
    fontSize: number;
    /** Style presets */
    presets: Record<StylePreset, IStyle>;
}

/**
 * Success result interface
 */
export declare interface ISuccessResult<T = unknown> {
    success: true;
    data: T;
    message?: string;
}

/**
 * Table structure interface
 */
export declare interface ITable {
    /** Table name */
    name?: string;
    /** Table headers */
    headers?: IHeaderCell[];
    /** Table sub-headers */
    subHeaders?: IHeaderCell[];
    /** Table data rows */
    body?: IDataCell[];
    /** Table footers */
    footers?: IFooterCell[];
    /** Table range */
    range?: ICellRange;
    /** Whether to show table borders */
    showBorders?: boolean;
    /** Whether to show alternating row colors */
    showStripes?: boolean;
    /** Table style */
    style?: 'TableStyleLight1' | 'TableStyleLight2' | 'TableStyleMedium1' | 'TableStyleMedium2' | 'TableStyleDark1' | 'TableStyleDark2';
    /** Auto filter for this table */
    autoFilter?: boolean;
}

/**
 * Unique validation interface
 */
export declare interface IUniqueValidation {
    /** Scope for uniqueness check */
    scope: 'worksheet' | 'table' | 'column' | 'row' | 'custom';
    /** Custom scope function */
    scopeFunction?: (cell: IDataCell | IHeaderCell | IFooterCell) => string;
    /** Whether to ignore case */
    ignoreCase?: boolean;
    /** Whether to ignore whitespace */
    ignoreWhitespace?: boolean;
}

/**
 * Validation context interface
 */
export declare interface IValidationContext {
    /** Cell being validated */
    cell: IDataCell | IHeaderCell | IFooterCell;
    /** Cell position */
    position?: {
        row: number;
        col: number;
    };
    /** Worksheet name */
    worksheetName?: string;
    /** Table name */
    tableName?: string;
    /** Validation rules */
    rules?: IValidationRule[];
    /** Additional context data */
    data?: Record<string, unknown>;
}

/**
 * Validation engine interface
 */
export declare interface IValidationEngine {
    /** Validation schemas */
    schemas: Map<string, IValidationSchema>;
    /** Active schema */
    activeSchema?: IValidationSchema;
    /** Whether validation is enabled */
    enabled: boolean;
    /** Validation cache */
    cache: Map<string, IValidationResult>;
    /** Add a validation schema */
    addSchema(schema: IValidationSchema): void;
    /** Remove a validation schema */
    removeSchema(name: string): boolean;
    /** Set the active schema */
    setActiveSchema(name: string): boolean;
    /** Validate a cell */
    validateCell(cell: IDataCell | IHeaderCell | IFooterCell, context?: IValidationContext): IValidationResult;
    /** Validate a worksheet */
    validateWorksheet(worksheet: unknown): IValidationResult[];
    /** Clear validation cache */
    clearCache(): void;
    /** Get validation statistics */
    getStats(): IValidationStats;
}

/**
 * Validation result interface
 */
export declare interface IValidationResult {
    /** Whether the validation passed */
    isValid: boolean;
    /** Validation errors */
    errors: string[];
    /** Validation warnings */
    warnings: string[];
    /** Validation info messages */
    info: string[];
    /** Suggested fixes */
    suggestions: string[];
    /** Validation metadata */
    metadata?: Record<string, unknown>;
}

/**
 * Validation rule interface
 */
export declare interface IValidationRule {
    /** Rule name */
    name: string;
    /** Rule description */
    description?: string;
    /** Rule type */
    type: 'required' | 'type' | 'range' | 'length' | 'format' | 'custom' | 'unique' | 'reference';
    /** Rule severity */
    severity: 'error' | 'warning' | 'info';
    /** Whether the rule is enabled */
    enabled: boolean;
    /** Rule parameters */
    params?: Record<string, unknown>;
    /** Custom validation function */
    validator?: (value: unknown, context?: IValidationContext) => IValidationResult;
}

/**
 * Validation schema interface
 */
export declare interface IValidationSchema {
    /** Schema name */
    name: string;
    /** Schema description */
    description?: string;
    /** Schema version */
    version?: string;
    /** Default rules */
    defaultRules: IValidationRule[];
    /** Cell type rules */
    cellTypeRules: Map<CellType, IValidationRule[]>;
    /** Custom rules */
    customRules: Map<string, IValidationRule>;
    /** Whether the schema is enabled */
    enabled: boolean;
}

/**
 * Validation statistics interface
 */
export declare interface IValidationStats {
    /** Total validations performed */
    totalValidations: number;
    /** Number of passed validations */
    passedValidations: number;
    /** Number of failed validations */
    failedValidations: number;
    /** Number of warnings */
    warnings: number;
    /** Average validation time in milliseconds */
    averageValidationTime: number;
    /** Cache hit rate */
    cacheHitRate: number;
    /** Most common validation errors */
    commonErrors: Array<{
        error: string;
        count: number;
    }>;
}

/**
 * Watermark configuration
 */
export declare interface IWatermark {
    /** Watermark text */
    text?: string;
    /** Watermark image (alternative to text) */
    image?: IWorksheetImage;
    /** Position */
    position?: {
        /** Horizontal position: 'left' | 'center' | 'right' */
        horizontal?: 'left' | 'center' | 'right';
        /** Vertical position: 'top' | 'middle' | 'bottom' */
        vertical?: 'top' | 'middle' | 'bottom';
    };
    /** Opacity (0-1) */
    opacity?: number;
    /** Rotation in degrees */
    rotation?: number;
    /** Font size (if using text) */
    fontSize?: number;
    /** Font color (if using text) */
    fontColor?: string;
}

/**
 * Workbook metadata interface
 */
export declare interface IWorkbookMetadata {
    /** Workbook author */
    author?: string;
    /** Workbook title */
    title?: string;
    /** Workbook subject */
    subject?: string;
    /** Workbook keywords */
    keywords?: string;
    /** Workbook category */
    category?: string;
    /** Workbook description */
    description?: string;
    /** Workbook company */
    company?: string;
    /** Workbook manager */
    manager?: string;
    /** Creation date */
    created?: Date;
    /** Last modified date */
    modified?: Date;
    /** Application name */
    application?: string;
    /** Application version */
    appVersion?: string;
    /** Hyperlink base */
    hyperlinkBase?: string;
}

/**
 * Workbook theme configuration
 */
export declare interface IWorkbookTheme {
    /** Theme name */
    name?: string;
    /** Color scheme */
    colors?: {
        /** Dark 1 color */
        dark1?: Color;
        /** Light 1 color */
        light1?: Color;
        /** Dark 2 color */
        dark2?: Color;
        /** Light 2 color */
        light2?: Color;
        /** Accent 1 color */
        accent1?: Color;
        /** Accent 2 color */
        accent2?: Color;
        /** Accent 3 color */
        accent3?: Color;
        /** Accent 4 color */
        accent4?: Color;
        /** Accent 5 color */
        accent5?: Color;
        /** Accent 6 color */
        accent6?: Color;
        /** Hyperlink color */
        hyperlink?: Color;
        /** Followed hyperlink color */
        followedHyperlink?: Color;
    };
    /** Font scheme */
    fonts?: {
        /** Major font (headings) */
        major?: {
            latin?: string;
            eastAsian?: string;
            complexScript?: string;
        };
        /** Minor font (body) */
        minor?: {
            latin?: string;
            eastAsian?: string;
            complexScript?: string;
        };
    };
    /** Section styles - automatically applied to headers, footers, body, etc. */
    sectionStyles?: {
        /** Style for main headers */
        header?: {
            backgroundColor?: Color;
            fontColor?: Color;
            fontSize?: number;
            fontBold?: boolean;
            borderColor?: Color;
        };
        /** Style for subheaders */
        subHeader?: {
            backgroundColor?: Color;
            fontColor?: Color;
            fontSize?: number;
            fontBold?: boolean;
            borderColor?: Color;
        };
        /** Style for body/data rows */
        body?: {
            backgroundColor?: Color;
            fontColor?: Color;
            fontSize?: number;
            alternatingRowColor?: Color;
            borderColor?: Color;
        };
        /** Style for footers */
        footer?: {
            backgroundColor?: Color;
            fontColor?: Color;
            fontSize?: number;
            fontBold?: boolean;
            borderColor?: Color;
        };
    };
    /** Whether to automatically apply section styles (default: true) */
    autoApplySectionStyles?: boolean;
}

/**
 * Worksheet interface
 */
export declare interface IWorksheet {
    /** Worksheet configuration */
    config: IWorksheetConfig;
    /** Tables in the worksheet */
    tables: ITable[];
    /** Current row pointer */
    currentRow: number;
    /** Current column pointer */
    currentCol: number;
    /** Header pointers for navigation */
    headerPointers: Map<string, ICellPosition>;
    /** Whether the worksheet has been built */
    isBuilt: boolean;
    /** Add a header */
    addHeader(header: IHeaderCell): this;
    /** Add subheaders */
    addSubHeaders(subHeaders: IHeaderCell[]): this;
    /** Add a row or rows */
    addRow(row: IDataCell[] | IDataCell): this;
    /** Add a footer or footers */
    addFooter(footer: IFooterCell[] | IFooterCell): this;
    /** Add a new table to the worksheet */
    addTable(tableConfig?: Partial<ITable>): this;
    /** Finalize the current table with temporary data */
    finalizeTable(): this;
    /** Get a table by name */
    getTable(name: string): ITable | undefined;
    /** Add an image to the worksheet */
    addImage(image: IWorksheetImage): this;
    /** Group rows (create collapsible outline) */
    groupRows(startRow: number, endRow: number, collapsed?: boolean): this;
    /** Group columns (create collapsible outline) */
    groupColumns(startCol: number, endCol: number, collapsed?: boolean): this;
    /** Add a named range */
    addNamedRange(name: string, range: string | ICellRange, scope?: string): this;
    /** Add an Excel structured table */
    addExcelTable(table: IExcelTable): this;
    /** Hide rows */
    hideRows(rows: number | number[]): this;
    /** Show rows */
    showRows(rows: number | number[]): this;
    /** Hide columns */
    hideColumns(columns: number | string | (number | string)[]): this;
    /** Show columns */
    showColumns(columns: number | string | (number | string)[]): this;
    /** Add a pivot table */
    addPivotTable(pivotTable: IPivotTable): this;
    /** Add a slicer to a table or pivot table */
    addSlicer(slicer: ISlicer): this;
    /** Add a watermark to the worksheet */
    addWatermark(watermark: IWatermark): this;
    /** Add a data connection */
    addDataConnection(connection: IDataConnection): this;
    /** Build the worksheet */
    build(workbook: any, options?: any): Promise<void>;
    /** Validate the worksheet */
    validate(): Result<boolean>;
}

/**
 * Worksheet configuration interface
 */
export declare interface IWorksheetConfig {
    /** Worksheet name */
    name: string;
    /** Tab color */
    tabColor?: Color;
    /** Default row height */
    defaultRowHeight?: number;
    /** Default column width */
    defaultColWidth?: number;
    /** Whether the worksheet is hidden */
    hidden?: boolean;
    /** Whether the worksheet is protected */
    protected?: boolean;
    /** Protection password */
    protectionPassword?: string;
    /** Whether to show grid lines */
    showGridLines?: boolean;
    /** Whether to show row and column headers */
    showRowColHeaders?: boolean;
    /** Zoom level (1-400) */
    zoom?: number;
    /** Freeze panes position */
    freezePanes?: ICellPosition;
    /** Print area */
    printArea?: ICellRange;
    /** Fit to page settings */
    fitToPage?: {
        fitToWidth?: number;
        fitToHeight?: number;
    };
    /** Page setup */
    pageSetup?: {
        orientation?: 'portrait' | 'landscape';
        paperSize?: number;
        fitToPage?: boolean;
        fitToWidth?: number;
        fitToHeight?: number;
        scale?: number;
        horizontalCentered?: boolean;
        verticalCentered?: boolean;
        margins?: {
            top?: number;
            left?: number;
            bottom?: number;
            right?: number;
            header?: number;
            footer?: number;
        };
    };
    /** Auto filter configuration */
    autoFilter?: {
        /** Enable auto filter for the worksheet */
        enabled?: boolean;
        /** Auto filter range (if not specified, applies to all data) */
        range?: ICellRange;
        /** Start row for auto filter (1-based, default: first data row) */
        startRow?: number;
        /** End row for auto filter (1-based, default: last data row) */
        endRow?: number;
        /** Start column for auto filter (1-based, default: 1) */
        startColumn?: number;
        /** End column for auto filter (1-based, default: last column) */
        endColumn?: number;
    };
    /** Print headers/footers configuration */
    printHeadersFooters?: {
        /** Header text (left, center, right) */
        header?: {
            left?: string;
            center?: string;
            right?: string;
        };
        /** Footer text (left, center, right) */
        footer?: {
            left?: string;
            center?: string;
            right?: string;
        };
    };
    /** Repeat rows/columns on each printed page */
    printRepeat?: {
        /** Rows to repeat (e.g., "1:2" or [1, 2]) */
        rows?: string | number[];
        /** Columns to repeat (e.g., "A:B" or [1, 2]) */
        columns?: string | number[];
    };
    /** Split panes configuration (divides window into panes) */
    splitPanes?: {
        /** Horizontal split position (column number, 0 = no split) */
        xSplit?: number;
        /** Vertical split position (row number, 0 = no split) */
        ySplit?: number;
        /** Top-left cell in bottom-right pane */
        topLeftCell?: string;
        /** Active pane (topLeft, topRight, bottomLeft, bottomRight) */
        activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
    };
    /** Sheet views configuration */
    views?: {
        /** View state (normal, pageBreakPreview, pageLayout) */
        state?: 'normal' | 'pageBreakPreview' | 'pageLayout';
        /** Zoom level (10-400) */
        zoomScale?: number;
        /** Normal zoom level */
        zoomScaleNormal?: number;
        /** Show grid lines */
        showGridLines?: boolean;
        /** Show row and column headers */
        showRowColHeaders?: boolean;
        /** Show ruler (page layout view) */
        showRuler?: boolean;
        /** Right-to-left */
        rightToLeft?: boolean;
    };
}

/**
 * Worksheet event interface
 */
export declare interface IWorksheetEvent {
    type: WorksheetEventType;
    worksheet: IWorksheet;
    data?: Record<string, unknown>;
    timestamp: Date;
}

/**
 * Image configuration for worksheet
 */
export declare interface IWorksheetImage {
    /** Image buffer (ArrayBuffer, Uint8Array, or base64 string) */
    buffer: ArrayBuffer | Uint8Array | string;
    /** Image name/ID */
    name?: string;
    /** Image extension (png, jpeg, gif, etc.) */
    extension: 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp' | 'webp';
    /** Position - can be cell reference or absolute position */
    position: {
        /** Cell reference (e.g., 'A1') or row number (1-based) */
        row: string | number;
        /** Column letter (e.g., 'A') or column number (1-based) */
        col: string | number;
    };
    /** Image size */
    size?: {
        /** Width in pixels or Excel units */
        width?: number;
        /** Height in pixels or Excel units */
        height?: number;
        /** Scale factor (0-1) */
        scaleX?: number;
        /** Scale factor (0-1) */
        scaleY?: number;
    };
    /** Hyperlink for image (optional) */
    hyperlink?: string;
    /** Image description/alt text */
    description?: string;
}

/**
 * Worksheet statistics
 */
export declare interface IWorksheetStats {
    /** Total number of cells */
    totalCells: number;
    /** Number of header cells */
    headerCells: number;
    /** Number of data cells */
    dataCells: number;
    /** Number of footer cells */
    footerCells: number;
    /** Number of tables */
    tables: number;
    /** Used range */
    usedRange: ICellRange;
    /** Memory usage in bytes */
    memoryUsage: number;
}

/**
 * Worksheet validation result
 */
export declare interface IWorksheetValidationResult {
    /** Whether the worksheet is valid */
    isValid: boolean;
    /** Validation errors */
    errors: string[];
    /** Validation warnings */
    warnings: string[];
    /** Cell validation results */
    cellResults: Map<string, boolean>;
}

/**
 * Number format options
 */
export declare enum NumberFormat {
    GENERAL = "General",
    NUMBER = "#,##0",
    NUMBER_DECIMALS = "#,##0.00",
    CURRENCY = "$#,##0.00",
    CURRENCY_INTEGER = "$#,##0",
    PERCENTAGE = "0%",
    PERCENTAGE_DECIMALS = "0.00%",
    DATE = "dd/mm/yyyy",
    DATE_TIME = "dd/mm/yyyy hh:mm",
    TIME = "hh:mm:ss",
    CUSTOM = "custom"
}

/**
 * Output format types
 */
declare enum OutputFormat {
    /** Format by worksheet (structured with sheets, rows, cells) */
    WORKSHEET = "worksheet",
    /** Detailed format with text, column, row information */
    DETAILED = "detailed",
    /** Flat format - just the data without structure */
    FLAT = "flat"
}

/**
 * Result union type
 */
export declare type Result<T = unknown> = ISuccessResult<T> | IErrorResult;

/**
 * StyleBuilder class providing a fluent API for creating Excel styles
 */
export declare class StyleBuilder implements IStyleBuilder {
    private style;
    constructor();
    /**
     * Create a new StyleBuilder instance
     */
    static create(): StyleBuilder;
    /**
     * Set font name
     */
    fontName(name: string): StyleBuilder;
    /**
     * Set font size
     */
    fontSize(size: number): StyleBuilder;
    /**
     * Set font style
     */
    fontStyle(style: FontStyle): StyleBuilder;
    /**
     * Set font color
     */
    fontColor(color: Color): StyleBuilder;
    /**
     * Make font bold
     */
    fontBold(): StyleBuilder;
    /**
     * Make font italic
     */
    fontItalic(): StyleBuilder;
    /**
     * Make font underlined
     */
    fontUnderline(): StyleBuilder;
    /**
     * Set border on all sides
     */
    border(style: BorderStyle, color?: Color): StyleBuilder;
    /**
     * Set top border
     */
    borderTop(style: BorderStyle, color?: Color): StyleBuilder;
    /**
     * Set left border
     */
    borderLeft(style: BorderStyle, color?: Color): StyleBuilder;
    /**
     * Set bottom border
     */
    borderBottom(style: BorderStyle, color?: Color): StyleBuilder;
    /**
     * Set right border
     */
    borderRight(style: BorderStyle, color?: Color): StyleBuilder;
    /**
     * Set background color
     */
    backgroundColor(color: Color): StyleBuilder;
    /**
     * Set horizontal alignment
     */
    horizontalAlign(alignment: HorizontalAlignment): StyleBuilder;
    /**
     * Set vertical alignment
     */
    verticalAlign(alignment: VerticalAlignment): StyleBuilder;
    /**
     * Center align text
     */
    centerAlign(): StyleBuilder;
    /**
     * Left align text
     */
    leftAlign(): StyleBuilder;
    /**
     * Right align text
     */
    rightAlign(): StyleBuilder;
    /**
     * Wrap text
     */
    wrapText(): StyleBuilder;
    /**
     * Set number format
     */
    numberFormat(format: string): StyleBuilder;
    /**
     * Set striped rows
     */
    striped(): StyleBuilder;
    /**
     * Add conditional formatting
     */
    conditionalFormat(format: IConditionalFormat): StyleBuilder;
    /**
     * Build the final style
     */
    build(): IStyle;
    /**
     * Reset the builder
     */
    reset(): StyleBuilder;
    /**
     * Clone the current style
     */
    clone(): StyleBuilder;
}

/**
 * Style preset types
 */
export declare enum StylePreset {
    HEADER = "header",
    SUBHEADER = "subheader",
    DATA = "data",
    FOOTER = "footer",
    TOTAL = "total",
    HIGHLIGHT = "highlight",
    WARNING = "warning",
    ERROR = "error",
    SUCCESS = "success",
    INFO = "info"
}

/**
 * Vertical alignment options
 */
export declare enum VerticalAlignment {
    TOP = "top",
    MIDDLE = "middle",
    BOTTOM = "bottom",
    DISTRIBUTED = "distributed",
    JUSTIFY = "justify"
}

/**
 * Worksheet - Representa una hoja de cálculo dentro del builder
 *
 2 * Soporta headers, subheaders anidados, rows, footers, children y estilos por celda.
 */
export declare class Worksheet implements IWorksheet {
    config: IWorksheetConfig;
    tables: ITable[];
    currentRow: number;
    currentCol: number;
    headerPointers: Map<string, any>;
    isBuilt: boolean;
    private headers;
    private subHeaders;
    private body;
    private footers;
    private images;
    private rowGroups;
    private columnGroups;
    private namedRanges;
    private excelTables;
    private hiddenRows;
    private hiddenColumns;
    private pivotTables;
    private slicers;
    private watermarks;
    private dataConnections;
    private customStyles?;
    private theme?;
    constructor(config: IWorksheetConfig);
    /**
     * Agrega un header principal
     */
    addHeader(header: IHeaderCell): this;
    /**
     * Agrega subheaders (ahora soporta anidación)
     */
    addSubHeaders(subHeaders: IHeaderCell[]): this;
    /**
     * Agrega una fila de datos (puede ser jerárquica con childrens)
     */
    addRow(row: IDataCell[] | IDataCell): this;
    /**
     * Agrega un footer o varios
     */
    addFooter(footer: IFooterCell[] | IFooterCell): this;
    /**
     * Crea una nueva tabla y la agrega al worksheet
     */
    addTable(tableConfig?: Partial<ITable>): this;
    /**
     * Finaliza la tabla actual agregando todos los elementos temporales a la última tabla
     */
    finalizeTable(): this;
    /**
     * Obtiene una tabla por nombre
     */
    getTable(name: string): ITable | undefined;
    /**
     * Agrega una imagen al worksheet
     */
    addImage(image: IWorksheetImage): this;
    /**
     * Agrupa filas (crea esquema colapsable)
     */
    groupRows(startRow: number, endRow: number, collapsed?: boolean): this;
    /**
     * Agrupa columnas (crea esquema colapsable)
     */
    groupColumns(startCol: number, endCol: number, collapsed?: boolean): this;
    /**
     * Agrega un rango con nombre
     */
    addNamedRange(name: string, range: string | ICellRange, scope?: string): this;
    /**
     * Agrega una tabla estructurada de Excel
     */
    addExcelTable(table: IExcelTable): this;
    /**
     * Oculta filas
     */
    hideRows(rows: number | number[]): this;
    /**
     * Muestra filas
     */
    showRows(rows: number | number[]): this;
    /**
     * Oculta columnas
     */
    hideColumns(columns: number | string | (number | string)[]): this;
    /**
     * Muestra columnas
     */
    showColumns(columns: number | string | (number | string)[]): this;
    /**
     * Agrega una tabla dinámica (pivot table)
     */
    addPivotTable(pivotTable: IPivotTable): this;
    /**
     * Agrega un slicer a una tabla o tabla dinámica
     */
    addSlicer(slicer: ISlicer): this;
    /**
     * Agrega una marca de agua al worksheet
     */
    addWatermark(watermark: IWatermark): this;
    /**
     * Agrega una conexión de datos
     */
    addDataConnection(connection: IDataConnection): this;
    /**
     * Construye la hoja en el workbook de ExcelJS
     */
    build(workbook: default_2.Workbook, _options?: IBuildOptions): Promise<void>;
    /**
     * Construye una tabla individual en el worksheet
     */
    private buildTable;
    /**
     * Construcción tradicional para compatibilidad hacia atrás
     */
    private buildLegacyContent;
    /**
     * Calcula el número máximo de columnas para una tabla
     */
    private calculateTableMaxColumns;
    /**
     * Aplica el estilo de tabla a un rango específico
     */
    private applyTableStyle;
    /**
     * Construye headers anidados recursivamente
     * @param ws - Worksheet de ExcelJS
     * @param startRow - Fila inicial
     * @param headers - Array de headers a procesar
     * @returns La siguiente fila disponible
     */
    private buildNestedHeaders;
    /**
     * Obtiene información del header en una profundidad específica
     */
    private getHeaderAtDepth;
    /**
     * Aplica todos los merges (horizontales y verticales) después de crear todas las filas
     */
    private applyAllMerges;
    /**
     * Aplica merges inteligentes basados en la estructura de headers
     */
    private applySmartMerges;
    /**
     * Aplica merges inteligentes para un header específico
     */
    private applySmartMergesForHeader;
    /**
     * Calcula el span de columnas para un header
     */
    private calculateHeaderColSpan;
    /**
     * Obtiene la profundidad máxima de headers anidados
     */
    private getMaxHeaderDepth;
    /**
     * Obtiene el número máximo de columnas
     */
    private getMaxColumns;
    /**
     * Valida la hoja
     */
    validate(): Result<boolean>;
    /**
     * Calcula las posiciones de columnas para los datos basándose en la estructura de subheaders
     */
    private calculateDataColumnPositions;
    /**
     * Agrega una fila de footer
     * @returns el siguiente rowPointer disponible
     */
    private addFooterRow;
    /**
     * Aplica width y height a una celda/fila
     */
    private applyCellDimensions;
    /**
     * Aplica comentario a una celda
     */
    private applyCellComment;
    /**
     * Aplica validación de datos a una celda
     */
    private applyDataValidation;
    /**
     * Aplica formato condicional a una celda
     */
    private applyConditionalFormatting;
    /**
     * Aplica filtro automático a una tabla
     */
    private applyAutoFilter;
    /**
     * Aplica filtro automático a nivel de worksheet
     */
    private applyWorksheetAutoFilter;
    /**
     * Procesa el valor de una celda considerando links y máscaras
     * Si el tipo es LINK o hay un link, crea un hipervínculo en Excel
     */
    private processCellValue;
    /**
     * Agrega una fila de datos y sus children recursivamente
     * @returns el siguiente rowPointer disponible
     */
    private addDataRowRecursive;
    /**
     * Convierte un color a formato ExcelJS (ARGB)
     */
    private convertColor;
    /**
     * Convierte el estilo personalizado a formato compatible con ExcelJS
     */
    private convertStyle;
    /**
     * Convierte un número de columna a letra (1 = A, 2 = B, etc.)
     */
    private numberToColumnLetter;
    /**
     * Convierte letra de columna a número (A = 1, B = 2, etc.)
     */
    private columnLetterToNumber;
    /**
     * Aplica una imagen al worksheet
     */
    private applyImage;
    /**
     * Aplica agrupación de filas
     */
    private applyRowGrouping;
    /**
     * Aplica agrupación de columnas
     */
    private applyColumnGrouping;
    /**
     * Aplica una tabla estructurada de Excel
     */
    private applyExcelTable;
    /**
     * Aplica configuración avanzada de impresión
     */
    private applyAdvancedPrintSettings;
    /**
     * Aplica filas y columnas ocultas
     */
    private applyHiddenRowsColumns;
    /**
     * Aplica una tabla dinámica (pivot table)
     */
    private applyPivotTable;
    /**
     * Convierte un color a formato ExcelJS
     */
    private convertColorToExcelJS;
    /**
     * Aplica views (freeze panes, split panes, sheet views)
     */
    private applyViews;
    /**
     * Obtiene un estilo predefinido del workbook
     */
    private getPredefinedStyle;
    /**
     * Obtiene un estilo del tema para una sección específica
     */
    private getThemeStyle;
    /**
     * Aplica un slicer a una tabla o tabla dinámica
     */
    private applySlicer;
    /**
     * Aplica una marca de agua al worksheet
     */
    private applyWatermark;
    /**
     * Aplica una conexión de datos
     */
    private applyDataConnection;
}

/**
 * Worksheet event types
 */
export declare enum WorksheetEventType {
    CREATED = "created",
    UPDATED = "updated",
    DELETED = "deleted",
    TABLE_ADDED = "tableAdded",
    TABLE_REMOVED = "tableRemoved",
    CELL_ADDED = "cellAdded",
    CELL_UPDATED = "cellUpdated",
    CELL_DELETED = "cellDeleted"
}

/**
 * Mapper function types for different output formats
 */
declare type WorksheetMapper = (data: IJsonWorkbook) => unknown;

export { }
