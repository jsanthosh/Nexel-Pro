import {
  SheetData, CellFormatting, DEFAULT_FORMATTING,
  TableDefinition, TableColumnFilter, ConditionalFormatRule,
  SheetTabConfig, CellRange, SortDirection,
} from '../types/spreadsheet';
import type { SpreadsheetState, HistoryState, WorkbookState } from '../hooks/useSpreadsheet';

export type { WorkbookState };

// ── Serialized types (JSON-safe, no Maps/Sets) ──

export interface SerializedDocument {
  version: number;
  sheets: Record<string, SerializedSheetState>;
  tabs: SheetTabConfig[];
  activeSheetId: string;
}

interface SerializedSheetState {
  cells: Record<string, SerializedCellData>;
  rowCount: number;
  colCount: number;
  colWidths: Record<string, number>;
  rowHeights?: Record<string, number>;
  freezeRows: number;
  conditionalFormatRules: ConditionalFormatRule[];
  tables: SerializedTableDefinition[];
}

interface SerializedCellData {
  v: string;
  d: string;
  f?: Partial<CellFormatting>;
}

interface SerializedTableDefinition {
  id: string;
  range: CellRange;
  styleId?: string;
  sortColumn: number | null;
  sortDirection: SortDirection;
  filters: Record<string, { checkedValues: string[] | null }>;
}

// ── Serialize (client state → JSON-safe) ──

export function serializeWorkbook(workbook: WorkbookState): SerializedDocument {
  const sheets: Record<string, SerializedSheetState> = {};
  for (const [id, historyState] of Object.entries(workbook.sheets)) {
    sheets[id] = serializeSheet(historyState.present);
  }
  return { version: 1, sheets, tabs: workbook.tabs, activeSheetId: workbook.activeSheetId };
}

function serializeSheet(state: SpreadsheetState): SerializedSheetState {
  const cells: Record<string, SerializedCellData> = {};
  state.cells.forEach((cell, key) => {
    if (cell.value === '' && isDefaultFormatting(cell.formatting)) return;
    const entry: SerializedCellData = { v: cell.value, d: cell.displayValue };
    const sparseFmt = sparseFormatting(cell.formatting);
    if (sparseFmt) entry.f = sparseFmt;
    cells[key] = entry;
  });

  const colWidths: Record<string, number> = {};
  state.colWidths.forEach((w, col) => { colWidths[String(col)] = w; });

  const rowHeights: Record<string, number> = {};
  state.rowHeights.forEach((h, row) => { rowHeights[String(row)] = h; });

  return {
    cells,
    rowCount: state.rowCount,
    colCount: state.colCount,
    colWidths,
    rowHeights: Object.keys(rowHeights).length > 0 ? rowHeights : undefined,
    freezeRows: state.freezeRows,
    conditionalFormatRules: state.conditionalFormatRules,
    tables: state.tables.map(serializeTable),
  };
}

function serializeTable(table: TableDefinition): SerializedTableDefinition {
  const filters: Record<string, { checkedValues: string[] | null }> = {};
  table.filters.forEach((f, col) => {
    filters[String(col)] = { checkedValues: f.checkedValues ? Array.from(f.checkedValues) : null };
  });
  return {
    id: table.id,
    range: table.range,
    styleId: table.styleId,
    sortColumn: table.sortColumn,
    sortDirection: table.sortDirection,
    filters,
  };
}

// ── Deserialize (JSON-safe → client state) ──

export function deserializeWorkbook(doc: SerializedDocument): WorkbookState {
  const sheets: Record<string, HistoryState> = {};
  for (const [id, serialized] of Object.entries(doc.sheets)) {
    sheets[id] = { past: [], present: deserializeSheet(serialized), future: [] };
  }
  return { sheets, tabs: doc.tabs, activeSheetId: doc.activeSheetId };
}

function deserializeSheet(s: SerializedSheetState): SpreadsheetState {
  const cells: SheetData = new Map();
  for (const [key, entry] of Object.entries(s.cells)) {
    cells.set(key, {
      value: entry.v,
      displayValue: entry.d,
      formatting: { ...DEFAULT_FORMATTING, ...(entry.f || {}) },
    });
  }

  const colWidths = new Map<number, number>();
  for (const [col, w] of Object.entries(s.colWidths)) {
    colWidths.set(Number(col), w);
  }

  const rowHeights = new Map<number, number>();
  if (s.rowHeights) {
    for (const [row, h] of Object.entries(s.rowHeights)) {
      rowHeights.set(Number(row), h);
    }
  }

  return {
    cells,
    rowCount: s.rowCount,
    colCount: s.colCount,
    selectedRange: null,
    activeCell: null,
    colWidths,
    rowHeights,
    freezeRows: s.freezeRows,
    conditionalFormatRules: s.conditionalFormatRules ?? [],
    tables: (s.tables ?? []).map(deserializeTable),
  };
}

function deserializeTable(t: SerializedTableDefinition): TableDefinition {
  const filters = new Map<number, TableColumnFilter>();
  for (const [col, f] of Object.entries(t.filters)) {
    filters.set(Number(col), { checkedValues: f.checkedValues ? new Set(f.checkedValues) : null });
  }
  return { id: t.id, range: t.range, styleId: t.styleId ?? 'blue', sortColumn: t.sortColumn, sortDirection: t.sortDirection, filters };
}

// ── Helpers ──

function isDefaultFormatting(f: CellFormatting): boolean {
  return (
    f.bold === false && f.italic === false && f.underline === false &&
    f.backgroundColor === null && f.textColor === null &&
    f.hAlign === 'left' && f.vAlign === 'middle' &&
    (f.wrapText ?? false) === false && (f.indent ?? 0) === 0 && (f.textRotation ?? 0) === 0 &&
    (f.numberFormat ?? 'general') === 'general' && (f.decimalPlaces ?? 2) === 2 &&
    (f.currencyCode ?? 'USD') === 'USD' && (f.dateFormat ?? 'mm/dd/yyyy') === 'mm/dd/yyyy' &&
    (f.customFormat ?? '') === ''
  );
}

function sparseFormatting(f: CellFormatting): Partial<CellFormatting> | null {
  const sparse: Partial<CellFormatting> = {};
  let hasAny = false;
  for (const key of Object.keys(DEFAULT_FORMATTING) as (keyof CellFormatting)[]) {
    if (f[key] !== DEFAULT_FORMATTING[key]) {
      (sparse as any)[key] = f[key];
      hasAny = true;
    }
  }
  return hasAny ? sparse : null;
}
