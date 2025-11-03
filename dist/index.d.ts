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
 */
declare class ExcelBuilder implements IExcelBuilder {
    config: IExcelBuilderConfig;
    worksheets: Map<string, IWorksheet>;
    currentWorksheet: IWorksheet | undefined;
    isBuilding: boolean;
    stats: IBuildStats;
    private eventEmitter;
    constructor(config?: IExcelBuilderConfig);
    /**
     * Add a new worksheet to the workbook
     */
    addWorksheet(name: string, worksheetConfig?: Partial<IWorksheetConfig>): IWorksheet;
    /**
     * Get a worksheet by name
     */
    getWorksheet(name: string): IWorksheet | undefined;
    /**
     * Remove a worksheet by name
     */
    removeWorksheet(name: string): boolean;
    /**
     * Set the current worksheet
     */
    setCurrentWorksheet(name: string): boolean;
    /**
     * Build the workbook and return as ArrayBuffer
     */
    build(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
    /**
     * Generate and download the file
     */
    generateAndDownload(fileName: string, options?: IDownloadOptions): Promise<Result<void>>;
    /**
     * Get workbook as buffer
     */
    toBuffer(options?: IBuildOptions): Promise<Result<ArrayBuffer>>;
    /**
     * Get workbook as blob
     */
    toBlob(options?: IBuildOptions): Promise<Result<Blob>>;
    /**
     * Validate the workbook
     */
    validate(): Result<boolean>;
    /**
     * Clear all worksheets
     */
    clear(): void;
    /**
     * Get workbook statistics
     */
    getStats(): IBuildStats;
    /**
     * Event handling methods
     */
    on(eventType: BuilderEventType, listener: (event: IBuilderEvent) => void): string;
    off(eventType: BuilderEventType, listenerId: string): boolean;
    removeAllListeners(eventType?: BuilderEventType): void;
    /**
     * Private methods
     */
    private emitEvent;
    private initializeStats;
}
export { ExcelBuilder }
export default ExcelBuilder;

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
     * Agrega una fila de datos y sus children recursivamente
     * @returns el siguiente rowPointer disponible
     */
    private addDataRowRecursive;
    /**
     * Convierte el estilo personalizado a formato compatible con ExcelJS
     */
    private convertStyle;
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

export { }
