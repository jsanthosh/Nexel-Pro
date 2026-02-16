export type HAlign = 'left' | 'center' | 'right';
export type VAlign = 'top' | 'middle' | 'bottom';

export type NumberFormatType = 'general' | 'number' | 'currency' | 'accounting' | 'percentage' | 'date' | 'time' | 'text' | 'custom';

export interface CellFormatting {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  backgroundColor: string | null;
  textColor: string | null;
  hAlign: HAlign;
  vAlign: VAlign;
  wrapText: boolean;
  indent: number;
  textRotation: number;
  fontFamily: string;
  numberFormat: NumberFormatType;
  decimalPlaces: number;
  currencyCode: string;
  dateFormat: string;
  customFormat: string;
}

export interface CellData {
  value: string;
  displayValue: string;
  formatting: CellFormatting;
}

export type SheetData = Map<string, CellData>;

export interface CellAddress {
  row: number;
  col: number;
}

export interface CellRange {
  start: CellAddress;
  end: CellAddress;
}

export const DEFAULT_FORMATTING: CellFormatting = {
  bold: false,
  italic: false,
  underline: false,
  backgroundColor: null,
  textColor: null,
  hAlign: 'left',
  vAlign: 'middle',
  wrapText: false,
  indent: 0,
  textRotation: 0,
  fontFamily: 'Default',
  numberFormat: 'general',
  decimalPlaces: 2,
  currencyCode: 'USD',
  dateFormat: 'mm/dd/yyyy',
  customFormat: '',
};

export const DEFAULT_CELL: CellData = {
  value: '',
  displayValue: '',
  formatting: { ...DEFAULT_FORMATTING },
};

export type ConditionType =
  | 'greaterThan'
  | 'lessThan'
  | 'equalTo'
  | 'between'
  | 'textContains'
  | 'isEmpty'
  | 'isNotEmpty'
  | 'duplicateValues';

export interface ConditionalFormatRule {
  id: string;
  range: CellRange;
  condition: ConditionType;
  values: string[];
  formatting: Partial<CellFormatting>;
  priority: number;
}

export type SortDirection = 'asc' | 'desc' | null;

export interface TableColumnFilter {
  checkedValues: Set<string> | null;
}

export interface TableStyleDef {
  id: string;
  name: string;
  headerBg: string;
  headerText: string;
  oddRowBg: string;
  evenRowBg: string;
  borderColor: string;
}

export const TABLE_STYLES: TableStyleDef[] = [
  { id: 'blue',    name: 'Ocean Blue',  headerBg: '#4472C4', headerText: '#fff', oddRowBg: '#D9E2F3', evenRowBg: '#ffffff', borderColor: '#4472C4' },
  { id: 'teal',    name: 'Teal',        headerBg: '#2B9E8F', headerText: '#fff', oddRowBg: '#D4EFED', evenRowBg: '#ffffff', borderColor: '#2B9E8F' },
  { id: 'indigo',  name: 'Indigo',      headerBg: '#4F46E5', headerText: '#fff', oddRowBg: '#E0E7FF', evenRowBg: '#ffffff', borderColor: '#4F46E5' },
  { id: 'purple',  name: 'Purple',      headerBg: '#7C3AED', headerText: '#fff', oddRowBg: '#EDE9FE', evenRowBg: '#ffffff', borderColor: '#7C3AED' },
  { id: 'rose',    name: 'Rose',        headerBg: '#E11D48', headerText: '#fff', oddRowBg: '#FFE4E6', evenRowBg: '#ffffff', borderColor: '#E11D48' },
  { id: 'orange',  name: 'Sunset',      headerBg: '#EA580C', headerText: '#fff', oddRowBg: '#FED7AA', evenRowBg: '#ffffff', borderColor: '#EA580C' },
  { id: 'green',   name: 'Forest',      headerBg: '#16A34A', headerText: '#fff', oddRowBg: '#DCFCE7', evenRowBg: '#ffffff', borderColor: '#16A34A' },
  { id: 'slate',   name: 'Slate',       headerBg: '#475569', headerText: '#fff', oddRowBg: '#E2E8F0', evenRowBg: '#ffffff', borderColor: '#475569' },
  { id: 'amber',   name: 'Amber',       headerBg: '#D97706', headerText: '#fff', oddRowBg: '#FEF3C7', evenRowBg: '#ffffff', borderColor: '#D97706' },
  { id: 'cyan',    name: 'Cyan',        headerBg: '#0891B2', headerText: '#fff', oddRowBg: '#CFFAFE', evenRowBg: '#ffffff', borderColor: '#0891B2' },
  { id: 'pink',    name: 'Blossom',     headerBg: '#DB2777', headerText: '#fff', oddRowBg: '#FCE7F3', evenRowBg: '#ffffff', borderColor: '#DB2777' },
  { id: 'neutral', name: 'Minimal',     headerBg: '#F3F4F6', headerText: '#111', oddRowBg: '#F9FAFB', evenRowBg: '#ffffff', borderColor: '#D1D5DB' },
];

export interface TableDefinition {
  id: string;
  range: CellRange;
  styleId: string;
  sortColumn: number | null;
  sortDirection: SortDirection;
  filters: Map<number, TableColumnFilter>;
}

export interface SheetTabConfig {
  id: string;
  name: string;
  color: string | null;
}

export type SpreadsheetAction =
  | { type: 'FORMAT'; range: string; formatting: Partial<CellFormatting> }
  | { type: 'SET_COLOR'; range: string; backgroundColor?: string; textColor?: string }
  | { type: 'SET_VALUE'; range: string; value: string }
  | { type: 'CREATE_SHEET'; name?: string }
  | { type: 'CREATE_CHART'; range: string; chartType?: 'bar' | 'line' | 'pie' }
  | { type: 'ADD_CF_RULE'; range: string; condition: ConditionType; values: string[]; formatting: Partial<CellFormatting> }
  | { type: 'DELETE_CF_RULES'; range?: string }
  | { type: 'CREATE_TABLE'; range: string; styleId?: string }
  | { type: 'DELETE_TABLE'; range: string }
  | { type: 'INSERT_ROW'; at: number; count?: number }
  | { type: 'DELETE_ROW'; at: number; count?: number }
  | { type: 'INSERT_COL'; at: number; count?: number }
  | { type: 'DELETE_COL'; at: number; count?: number }
  | { type: 'FREEZE_ROWS'; count: number }
  | { type: 'CLEAR_RANGE'; range: string }
  | { type: 'SET_COL_WIDTH'; column: string; width: number }
  | { type: 'QUERY_RESULT'; message: string; value?: number | string }
  | { type: 'ERROR'; message: string };

export interface ClaudeResponse {
  actions: SpreadsheetAction[];
  explanation: string;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  isLoading?: boolean;
}
