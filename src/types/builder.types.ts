/**
 * Builder-specific type definitions
 */

import { IWorkbookMetadata, Result } from './core.types';
import { IWorksheet, IWorksheetConfig } from './worksheet.types';
import { IStyle } from './style.types';

// Re-export ErrorType for convenience
export { ErrorType } from './core.types';

/**
 * Excel builder configuration interface
 */
export interface IExcelBuilderConfig {
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
 * Build options interface
 */
export interface IBuildOptions {
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
 * Download options interface
 */
export interface IDownloadOptions extends IBuildOptions {
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
 * Excel builder interface
 */
export interface IExcelBuilder {
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
 * Build statistics interface
 */
export interface IBuildStats {
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
 * Builder event types
 */
export enum BuilderEventType {
  WORKSHEET_ADDED = 'worksheetAdded',
  WORKSHEET_REMOVED = 'worksheetRemoved',
  WORKSHEET_UPDATED = 'worksheetUpdated',
  BUILD_STARTED = 'buildStarted',
  BUILD_PROGRESS = 'buildProgress',
  BUILD_COMPLETED = 'buildCompleted',
  BUILD_ERROR = 'buildError',
  DOWNLOAD_STARTED = 'downloadStarted',
  DOWNLOAD_PROGRESS = 'downloadProgress',
  DOWNLOAD_COMPLETED = 'downloadCompleted',
  DOWNLOAD_ERROR = 'downloadError'
}

/**
 * Builder event interface
 */
export interface IBuilderEvent {
  type: BuilderEventType;
  data?: Record<string, unknown>;
  timestamp: Date;
}

/**
 * Builder event listener interface
 */
export interface IBuilderEventListener {
  (event: IBuilderEvent): void;
}

/**
 * Builder validation result interface
 */
export interface IBuilderValidationResult {
  /** Whether the builder is valid */
  isValid: boolean;
  /** Validation errors */
  errors: string[];
  /** Validation warnings */
  warnings: string[];
  /** Worksheet validation results */
  worksheetResults: Map<string, boolean>;
} 