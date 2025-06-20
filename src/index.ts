/**
 * Han Excel Builder - Main entry point
 * 
 * Advanced Excel file generator with TypeScript support, comprehensive styling, and optimized performance
 */

// Core exports
export { ExcelBuilder } from './core/ExcelBuilder';
export { Worksheet } from './core/Worksheet';
export { StyleBuilder } from './styles/StyleBuilder';

// Type exports
export * from './types/core.types';
export * from './types/cell.types';
export * from './types/worksheet.types';
export * from './types/style.types';
export * from './types/builder.types';
export * from './types/validation.types';
export * from './types/events.types';

// Constants and enums
export {
  CellType,
  NumberFormat,
  HorizontalAlignment,
  VerticalAlignment,
  BorderStyle,
  FontStyle,
  ErrorType
} from './types/core.types';

export {
  BuilderEventType,
  StylePreset
} from './types';

// Utility exports
export { EventEmitter } from './utils/EventEmitter';

// Default export
export { ExcelBuilder as default } from './core/ExcelBuilder'; 