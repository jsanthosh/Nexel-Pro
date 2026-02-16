import { CellRange } from './spreadsheet';

export type ChartType = 'bar' | 'line' | 'pie';

export interface ChartInstance {
  id: string;
  chartType: ChartType;
  dataRange: CellRange | null;
  position: { x: number; y: number };
  size: { width: number; height: number };
}
