import { CellRange } from './spreadsheet';

export type AggregationType = 'sum' | 'count' | 'average' | 'min' | 'max';

export interface PivotFieldConfig {
  columnIndex: number;
  label: string;
}

export interface PivotValueFieldConfig extends PivotFieldConfig {
  aggregation: AggregationType;
}

export interface PivotConfig {
  sourceRange: CellRange;
  rowFields: PivotFieldConfig[];
  columnFields: PivotFieldConfig[];
  valueFields: PivotValueFieldConfig[];
}

export interface PivotResult {
  headers: string[];
  rows: (string | number)[][];
}
