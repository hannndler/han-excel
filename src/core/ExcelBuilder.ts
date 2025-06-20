/**
 * ExcelBuilder - Main class for creating Excel workbooks
 */

import ExcelJS from 'exceljs';
import saveAs from 'file-saver';
import { EventEmitter } from '../utils/EventEmitter';
import { Worksheet } from './Worksheet';
import {
  IExcelBuilder,
  IExcelBuilderConfig,
  IBuildOptions,
  IDownloadOptions,
  IBuildStats,
  BuilderEventType,
  IBuilderEvent,
  ErrorType
} from '../types/builder.types';
import {
  IWorksheet,
  IWorksheetConfig
} from '../types/worksheet.types';
import { 
  Result,
  ISuccessResult,
  IErrorResult
} from '../types/core.types';

/**
 * ExcelBuilder class for creating Excel workbooks
 */
export class ExcelBuilder implements IExcelBuilder {
  public config: IExcelBuilderConfig;
  public worksheets: Map<string, IWorksheet> = new Map();
  public currentWorksheet: IWorksheet | undefined;
  public isBuilding = false;
  public stats: IBuildStats;

  private eventEmitter: EventEmitter;

  constructor(config: IExcelBuilderConfig = {}) {
    this.config = {
      enableValidation: true,
      enableEvents: true,
      enablePerformanceMonitoring: false,
      maxWorksheets: 255,
      maxRowsPerWorksheet: 1048576,
      maxColumnsPerWorksheet: 16384,
      memoryLimit: 100 * 1024 * 1024, // 100MB
      ...config
    };

    this.stats = this.initializeStats();
    this.eventEmitter = new EventEmitter();
  }

  /**
   * Add a new worksheet to the workbook
   */
  addWorksheet(name: string, worksheetConfig: Partial<IWorksheetConfig> = {}): IWorksheet {
    if (this.worksheets.has(name)) {
      throw new Error(`Worksheet "${name}" already exists`);
    }

    const config: IWorksheetConfig = {
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
  getWorksheet(name: string): IWorksheet | undefined {
    return this.worksheets.get(name);
  }

  /**
   * Remove a worksheet by name
   */
  removeWorksheet(name: string): boolean {
    const worksheet = this.worksheets.get(name);
    if (!worksheet) {
      return false;
    }

    this.worksheets.delete(name);
    
    // If this was the current worksheet, clear it
    if (this.currentWorksheet === worksheet) {
      this.currentWorksheet = undefined;
    }
    
    this.emitEvent(BuilderEventType.WORKSHEET_REMOVED, { worksheetName: name });
    
    return true;
  }

  /**
   * Set the current worksheet
   */
  setCurrentWorksheet(name: string): boolean {
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
  async build(options: IBuildOptions = {}): Promise<Result<ArrayBuffer>> {
    if (this.isBuilding) {
      return {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: 'Build already in progress',
          stack: new Error().stack || ''
        }
      };
    }

    this.isBuilding = true;
    const startTime = Date.now();
    
    try {
      this.emitEvent(BuilderEventType.BUILD_STARTED);
      
      const workbook = new ExcelJS.Workbook();
      
      // Add metadata
      if (this.config.metadata) {
        workbook.creator = this.config.metadata.author || 'Han Excel Builder';
        workbook.lastModifiedBy = this.config.metadata.author || 'Han Excel Builder';
        workbook.created = this.config.metadata.created || new Date();
        workbook.modified = this.config.metadata.modified || new Date();
        if (this.config.metadata.title) workbook.title = this.config.metadata.title;
        if (this.config.metadata.subject) workbook.subject = this.config.metadata.subject;
        if (this.config.metadata.keywords) workbook.keywords = this.config.metadata.keywords;
        if (this.config.metadata.category) workbook.category = this.config.metadata.category;
        if (this.config.metadata.description) workbook.description = this.config.metadata.description;
      }

      // Build each worksheet
      for (const worksheet of this.worksheets.values()) {
        await (worksheet as Worksheet).build(workbook, options);
      }

      // Write to buffer
      const buffer = await workbook.xlsx.writeBuffer({
        compression: options.compressionLevel || 6
      } as any);

      const endTime = Date.now();
      this.stats.buildTime = endTime - startTime;
      this.stats.fileSize = buffer.byteLength;
      
      const successResult: ISuccessResult<ArrayBuffer> = {
        success: true,
        data: buffer
      };

      this.emitEvent(BuilderEventType.BUILD_COMPLETED, {
        buildTime: this.stats.buildTime,
        fileSize: this.stats.fileSize
      });

      return successResult;

    } catch (error) {
      const errorResult: IErrorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : 'Unknown build error',
          stack: error instanceof Error ? error.stack || '' : ''
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
  async generateAndDownload(fileName: string, options: IDownloadOptions = {}): Promise<Result<void>> {
    const buildResult = await this.build(options);
    
    if (!buildResult.success) {
      return buildResult;
    }

    try {
      this.emitEvent(BuilderEventType.DOWNLOAD_STARTED, { fileName });
      
      const blob = new Blob([buildResult.data], { 
        type: options.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      saveAs(blob, fileName);
      
      this.emitEvent(BuilderEventType.DOWNLOAD_COMPLETED, { fileName });
      
      return { success: true, data: undefined };

    } catch (error) {
      const errorResult: IErrorResult = {
        success: false,
        error: {
          type: ErrorType.BUILD_ERROR,
          message: error instanceof Error ? error.message : 'Download failed',
          stack: error instanceof Error ? error.stack || '' : ''
        }
      };

      this.emitEvent(BuilderEventType.DOWNLOAD_ERROR, { error: errorResult.error });
      return errorResult;
    }
  }

  /**
   * Get workbook as buffer
   */
  async toBuffer(options: IBuildOptions = {}): Promise<Result<ArrayBuffer>> {
    return this.build(options);
  }

  /**
   * Get workbook as blob
   */
  async toBlob(options: IBuildOptions = {}): Promise<Result<Blob>> {
    const buildResult = await this.build(options);
    
    if (!buildResult.success) {
      return buildResult;
    }

    const blob = new Blob([buildResult.data], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    return { success: true, data: blob };
  }

  /**
   * Validate the workbook
   */
  validate(): Result<boolean> {
    const errors: string[] = [];
    
    if (this.worksheets.size === 0) {
      errors.push('No worksheets found');
    }

    // Validate each worksheet
    for (const [name, worksheet] of this.worksheets.entries()) {
      const worksheetValidation = (worksheet as Worksheet).validate();
      if (!worksheetValidation.success) {
        errors.push(`Worksheet "${name}": ${worksheetValidation.error?.message}`);
      }
    }

    if (errors.length > 0) {
      return {
        success: false,
        error: {
          type: ErrorType.VALIDATION_ERROR,
          message: errors.join('; '),
          stack: new Error().stack || ''
        }
      };
    }

    return { success: true, data: true };
  }

  /**
   * Clear all worksheets
   */
  clear(): void {
    this.worksheets.clear();
    this.currentWorksheet = undefined;
  }

  /**
   * Get workbook statistics
   */
  getStats(): IBuildStats {
    return { ...this.stats };
  }

  /**
   * Event handling methods
   */
  on(eventType: BuilderEventType, listener: (event: IBuilderEvent) => void): string {
    return this.eventEmitter.on(eventType, listener);
  }

  off(eventType: BuilderEventType, listenerId: string): boolean {
    return this.eventEmitter.off(eventType, listenerId);
  }

  removeAllListeners(eventType?: BuilderEventType): void {
    if (eventType) {
      this.eventEmitter.offAll(eventType);
    } else {
      this.eventEmitter.clear();
    }
  }

  /**
   * Private methods
   */
  private emitEvent(type: BuilderEventType, data?: Record<string, unknown>): void {
    const event: IBuilderEvent = {
      type,
      data: data || {},
      timestamp: new Date()
    };
    this.eventEmitter.emitSync(event);
  }

  private initializeStats(): IBuildStats {
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