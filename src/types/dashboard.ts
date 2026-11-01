export interface LegacyMacro {
  id: string;
  name: string;
  moduleName: string;
  moduleType: 'standard' | 'document' | 'class' | 'userform';
  procedureType: 'Sub' | 'Function' | 'Property' | 'Event';
  description: string;
  triggerEvent: string;
  originalVba: string;
  code: string; // Active editable VBA code
  modernizedTs?: string;
  isAutoRun?: boolean;
  isEnhanced?: boolean;
  enhancementTags?: string[];
  lastModified?: string;
  lineCount?: number;
}

export interface VbaModule {
  name: string;
  type: 'standard' | 'document' | 'class' | 'userform';
  code: string;
  procedures: LegacyMacro[];
}

export interface SheetProtectionMeta {
  wasProtected: boolean;
  algorithm?: string;
  hashValue?: string;
  unlockedAt?: string;
}

export interface SheetData {
  name: string;
  rowCount: number;
  columnCount: number;
  columns: string[];
  rows: Record<string, any>[];
  formulas: Record<string, string>; // e.g., "E2": "=C2*D2"
  isProtected?: boolean;
  protectionMeta?: SheetProtectionMeta;
}

export interface ColumnMetadata {
  type: 'currency' | 'percentage' | 'number' | 'date' | 'text';
  sampleValue?: any;
}

export interface Dataset {
  id: string;
  name: string;
  filename: string;
  fileSize: string;
  uploadedAt: string;
  sheets: SheetData[];
  activeSheetName: string;
  legacyMacros: LegacyMacro[];
  modules?: VbaModule[];
  columnMetadata: Record<string, ColumnMetadata>;
  notes?: string;
  hasVbaProject?: boolean;
  detectedModules?: string[];
  rawZip?: any; // JSZip instance to repack upgraded workbook
  unprotectedBlob?: Blob;
  unprotectionLog?: string[];
}

export type WidgetType = 'metric' | 'line' | 'area' | 'bar' | 'donut' | 'table';
export type AggregationType = 'sum' | 'avg' | 'min' | 'max' | 'count';
export type WidgetWidth = '1/3' | '1/2' | '2/3' | 'full';

export interface ReportingWidget {
  id: string;
  title: string;
  type: WidgetType;
  sheetName: string;
  xAxis: string; // Dimension or Category
  yAxis: string; // Numeric Metric
  aggregation: AggregationType;
  width: WidgetWidth;
  colorTheme: 'emerald' | 'indigo' | 'amber' | 'rose' | 'cyan' | 'slate';
  filterField?: string;
  filterValue?: string;
}

