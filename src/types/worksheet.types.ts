/**
 * Worksheet-specific type definitions
 */

import { IHeaderCell, IDataCell, IFooterCell, ICellPosition, ICellRange } from './cell.types';
import { Color, Result } from './core.types';

/**
 * Worksheet configuration interface
 */
export interface IWorksheetConfig {
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
 * Table structure interface
 */
export interface ITable {
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
 * Worksheet interface
 */
export interface IWorksheet {
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
  /** Build the worksheet */
  build(workbook: any, options?: any): Promise<void>;
  /** Validate the worksheet */
  validate(): Result<boolean>;
}

/**
 * Worksheet event types
 */
export enum WorksheetEventType {
  CREATED = 'created',
  UPDATED = 'updated',
  DELETED = 'deleted',
  TABLE_ADDED = 'tableAdded',
  TABLE_REMOVED = 'tableRemoved',
  CELL_ADDED = 'cellAdded',
  CELL_UPDATED = 'cellUpdated',
  CELL_DELETED = 'cellDeleted'
}

/**
 * Worksheet event interface
 */
export interface IWorksheetEvent {
  type: WorksheetEventType;
  worksheet: IWorksheet;
  data?: Record<string, unknown>;
  timestamp: Date;
}

/**
 * Worksheet validation result
 */
export interface IWorksheetValidationResult {
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
 * Worksheet statistics
 */
export interface IWorksheetStats {
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