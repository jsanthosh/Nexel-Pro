import { useReducer, useCallback } from 'react';
import {
  SheetData, CellRange, CellAddress, CellFormatting,
  SpreadsheetAction, ConditionalFormatRule, SheetTabConfig, TableDefinition,
  TableColumnFilter, SortDirection, DEFAULT_CELL, DEFAULT_FORMATTING,
} from '../types/spreadsheet';
import { colIndexToLetter, iterateRange } from '../utils/cellRange';
import { rangesOverlap, detectDataRegion } from '../utils/tableUtils';
import { evaluateCell, recomputeFormulas, parseNum } from '../utils/formulaEngine';
import { applyActions as applyActionsUtil } from '../utils/actionApplier';
import { exportToCSV, downloadCSV, importFromCSV, importFromXLSX, importAllSheetsFromXLSX } from '../utils/csvHandler';
import { generateSeries } from '../utils/fillSeries';

export interface SpreadsheetState {
  cells: SheetData;
  rowCount: number;
  colCount: number;
  selectedRange: CellRange | null;
  activeCell: CellAddress | null;
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  freezeRows: number;
  conditionalFormatRules: ConditionalFormatRule[];
  tables: TableDefinition[];
}

// ── History wrapper ────────────────────────────────────────────────────────────
export interface HistoryState {
  past: SpreadsheetState[];
  present: SpreadsheetState;
  future: SpreadsheetState[];
}

const MAX_HISTORY = 50;

// ── Inner action union (all undoable + non-undoable) ──────────────────────────
type Action =
  | { type: 'SET_CELL'; row: number; col: number; value: string }
  | { type: 'APPLY_FORMAT'; range: CellRange; formatting: Partial<CellFormatting> }
  | { type: 'APPLY_COLOR'; range: CellRange; backgroundColor?: string; textColor?: string }
  | { type: 'APPLY_ACTIONS'; actions: SpreadsheetAction[] }
  | { type: 'SET_SELECTION'; range: CellRange | null; activeCell: CellAddress | null }
  | { type: 'IMPORT_CSV'; cells: SheetData; rowCount: number; colCount: number }
  | { type: 'FILL_SERIES'; range: CellRange }
  | { type: 'FILL_DOWN';   range: CellRange }
  | { type: 'FILL_RIGHT';  range: CellRange }
  | { type: 'INSERT_ROW';  at: number }
  | { type: 'DELETE_ROW';  at: number }
  | { type: 'INSERT_COL';  at: number }
  | { type: 'DELETE_COL';  at: number }
  | { type: 'SET_COL_WIDTH'; col: number; width: number }
  | { type: 'SET_ROW_HEIGHT'; row: number; height: number }
  | { type: 'SET_FREEZE_ROWS'; rows: number }
  | { type: 'ADD_CF_RULE'; rule: ConditionalFormatRule }
  | { type: 'DELETE_CF_RULE'; id: string }
  | { type: 'CREATE_TABLE'; range: CellRange; styleId?: string }
  | { type: 'SET_TABLE_STYLE'; id: string; styleId: string }
  | { type: 'DELETE_TABLE'; id: string }
  | { type: 'SORT_TABLE'; id: string; column: number; direction: SortDirection }
  | { type: 'FILTER_TABLE'; id: string; column: number; checkedValues: Set<string> | null }
  | { type: 'UNDO' }
  | { type: 'REDO' }
  | { type: 'ADD_SHEET'; id: string; name: string }
  | { type: 'DELETE_SHEET'; id: string }
  | { type: 'RENAME_SHEET'; id: string; name: string }
  | { type: 'SET_SHEET_COLOR'; id: string; color: string | null }
  | { type: 'REORDER_SHEET'; id: string; newIndex: number }
  | { type: 'SWITCH_SHEET'; id: string }
  | { type: 'LOAD_WORKBOOK'; workbook: WorkbookState };

function createEmptyHistory(): HistoryState {
  return {
    past: [],
    present: {
      cells: new Map(),
      rowCount: 100,
      colCount: 26,
      selectedRange: null,
      activeCell: null,
      colWidths: new Map(),
      rowHeights: new Map(),
      freezeRows: 0,
      conditionalFormatRules: [],
      tables: [],
    },
    future: [],
  };
}

export interface WorkbookState {
  sheets: Record<string, HistoryState>;
  tabs: SheetTabConfig[];
  activeSheetId: string;
}

const DEFAULT_SHEET_ID = 'sheet-default';

const initialWorkbook: WorkbookState = {
  sheets: { [DEFAULT_SHEET_ID]: createEmptyHistory() },
  tabs: [{ id: DEFAULT_SHEET_ID, name: 'Sheet1', color: null }],
  activeSheetId: DEFAULT_SHEET_ID,
};

// ── Pure cell-state reducer (no history concerns) ─────────────────────────────
function cellReducer(state: SpreadsheetState, action: Action): SpreadsheetState {
  switch (action.type) {
    case 'SET_CELL': {
      const key = `${action.row}:${action.col}`;
      const existing = state.cells.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
      const updated = new Map(state.cells);
      const displayValue = evaluateCell(action.value, updated);
      updated.set(key, { ...existing, value: action.value, displayValue });
      return { ...state, cells: recomputeFormulas(updated) };
    }

    case 'APPLY_FORMAT': {
      const updated = new Map(state.cells);
      for (const key of iterateRange(action.range)) {
        const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
        updated.set(key, { ...existing, formatting: { ...existing.formatting, ...action.formatting } });
      }
      return { ...state, cells: updated };
    }

    case 'APPLY_COLOR': {
      const updated = new Map(state.cells);
      for (const key of iterateRange(action.range)) {
        const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
        updated.set(key, {
          ...existing,
          formatting: {
            ...existing.formatting,
            ...(action.backgroundColor !== undefined && { backgroundColor: action.backgroundColor }),
            ...(action.textColor !== undefined && { textColor: action.textColor }),
          },
        });
      }
      return { ...state, cells: updated };
    }

    case 'APPLY_ACTIONS':
      return { ...state, cells: applyActionsUtil(state.cells, action.actions) };

    case 'SET_SELECTION':
      return { ...state, selectedRange: action.range, activeCell: action.activeCell };

    case 'FILL_SERIES': {
      const { range } = action;
      const rowSpan = range.end.row - range.start.row + 1;
      const colSpan = range.end.col - range.start.col + 1;
      const goDown = rowSpan >= colSpan;
      const updated = new Map(state.cells);

      if (goDown) {
        for (let c = range.start.col; c <= range.end.col; c++) {
          const seeds: string[] = [];
          const seedCount = Math.min(2, rowSpan - 1) || 1;
          for (let s = 0; s < seedCount; s++) {
            seeds.push(updated.get(`${range.start.row + s}:${c}`)?.value ?? '');
          }
          if (seeds.every(v => v === '')) continue;
          const series = generateSeries(seeds, rowSpan);
          for (let r = 0; r < rowSpan; r++) {
            const key = `${range.start.row + r}:${c}`;
            const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
            updated.set(key, { ...existing, value: series[r], displayValue: series[r] });
          }
        }
      } else {
        for (let r = range.start.row; r <= range.end.row; r++) {
          const seeds: string[] = [];
          const seedCount = Math.min(2, colSpan - 1) || 1;
          for (let s = 0; s < seedCount; s++) {
            seeds.push(updated.get(`${r}:${range.start.col + s}`)?.value ?? '');
          }
          if (seeds.every(v => v === '')) continue;
          const series = generateSeries(seeds, colSpan);
          for (let c = 0; c < colSpan; c++) {
            const key = `${r}:${range.start.col + c}`;
            const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
            updated.set(key, { ...existing, value: series[c], displayValue: series[c] });
          }
        }
      }
      return { ...state, cells: recomputeFormulas(updated) };
    }

    case 'FILL_DOWN': {
      const { range } = action;
      if (range.end.row === range.start.row) return state;
      const updated = new Map(state.cells);
      for (let c = range.start.col; c <= range.end.col; c++) {
        const src = updated.get(`${range.start.row}:${c}`) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
        for (let r = range.start.row + 1; r <= range.end.row; r++) {
          const key = `${r}:${c}`;
          const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
          updated.set(key, { ...existing, value: src.value, displayValue: src.displayValue });
        }
      }
      return { ...state, cells: recomputeFormulas(updated) };
    }

    case 'FILL_RIGHT': {
      const { range } = action;
      if (range.end.col === range.start.col) return state;
      const updated = new Map(state.cells);
      for (let r = range.start.row; r <= range.end.row; r++) {
        const src = updated.get(`${r}:${range.start.col}`) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
        for (let c = range.start.col + 1; c <= range.end.col; c++) {
          const key = `${r}:${c}`;
          const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
          updated.set(key, { ...existing, value: src.value, displayValue: src.displayValue });
        }
      }
      return { ...state, cells: recomputeFormulas(updated) };
    }

    case 'IMPORT_CSV':
      return { ...state, cells: action.cells, rowCount: action.rowCount, colCount: action.colCount };

    case 'INSERT_ROW': {
      const updated = new Map<string, typeof DEFAULT_CELL>();
      state.cells.forEach((cell, key) => {
        const [r, c] = key.split(':').map(Number);
        const newKey = r >= action.at ? `${r + 1}:${c}` : key;
        updated.set(newKey, cell);
      });
      const shiftedTables = state.tables.map(t => {
        const { range } = t;
        if (action.at <= range.start.row) {
          return { ...t, range: { start: { ...range.start, row: range.start.row + 1 }, end: { ...range.end, row: range.end.row + 1 } } };
        } else if (action.at <= range.end.row) {
          return { ...t, range: { ...range, end: { ...range.end, row: range.end.row + 1 } } };
        }
        return t;
      });
      const newRowHeightsIns = new Map(state.rowHeights);
      const rowEntries = Array.from(newRowHeightsIns.entries()).sort((a, b) => b[0] - a[0]);
      for (const [row, h] of rowEntries) {
        if (row >= action.at) { newRowHeightsIns.delete(row); newRowHeightsIns.set(row + 1, h); }
      }
      return { ...state, cells: recomputeFormulas(updated), rowCount: state.rowCount + 1, tables: shiftedTables, rowHeights: newRowHeightsIns };
    }

    case 'DELETE_ROW': {
      const updated = new Map<string, typeof DEFAULT_CELL>();
      state.cells.forEach((cell, key) => {
        const [r, c] = key.split(':').map(Number);
        if (r === action.at) return;
        const newKey = r > action.at ? `${r - 1}:${c}` : key;
        updated.set(newKey, cell);
      });
      const shiftedTables = state.tables
        .filter(t => action.at !== t.range.start.row) // delete table if header row is deleted
        .map(t => {
          const { range } = t;
          if (action.at < range.start.row) {
            return { ...t, range: { start: { ...range.start, row: range.start.row - 1 }, end: { ...range.end, row: range.end.row - 1 } } };
          } else if (action.at <= range.end.row) {
            return { ...t, range: { ...range, end: { ...range.end, row: range.end.row - 1 } } };
          }
          return t;
        });
      const newRowHeightsDel = new Map(state.rowHeights);
      newRowHeightsDel.delete(action.at);
      const rowEntriesDel = Array.from(newRowHeightsDel.entries()).sort((a, b) => a[0] - b[0]);
      for (const [row, h] of rowEntriesDel) {
        if (row > action.at) { newRowHeightsDel.delete(row); newRowHeightsDel.set(row - 1, h); }
      }
      return { ...state, cells: recomputeFormulas(updated), rowCount: Math.max(state.rowCount - 1, 1), tables: shiftedTables, rowHeights: newRowHeightsDel };
    }

    case 'INSERT_COL': {
      const updated = new Map<string, typeof DEFAULT_CELL>();
      state.cells.forEach((cell, key) => {
        const [r, c] = key.split(':').map(Number);
        const newKey = c >= action.at ? `${r}:${c + 1}` : key;
        updated.set(newKey, cell);
      });
      const newColWidths = new Map(state.colWidths);
      // Shift column widths
      const entries = Array.from(newColWidths.entries()).sort((a, b) => b[0] - a[0]);
      for (const [col, w] of entries) {
        if (col >= action.at) { newColWidths.delete(col); newColWidths.set(col + 1, w); }
      }
      return { ...state, cells: recomputeFormulas(updated), colCount: state.colCount + 1, colWidths: newColWidths };
    }

    case 'DELETE_COL': {
      const updated = new Map<string, typeof DEFAULT_CELL>();
      state.cells.forEach((cell, key) => {
        const [r, c] = key.split(':').map(Number);
        if (c === action.at) return;
        const newKey = c > action.at ? `${r}:${c - 1}` : key;
        updated.set(newKey, cell);
      });
      const newColWidths = new Map(state.colWidths);
      newColWidths.delete(action.at);
      const entries = Array.from(newColWidths.entries()).sort((a, b) => a[0] - b[0]);
      for (const [col, w] of entries) {
        if (col > action.at) { newColWidths.delete(col); newColWidths.set(col - 1, w); }
      }
      return { ...state, cells: recomputeFormulas(updated), colCount: Math.max(state.colCount - 1, 1), colWidths: newColWidths };
    }

    case 'SET_COL_WIDTH': {
      const newColWidths = new Map(state.colWidths);
      newColWidths.set(action.col, action.width);
      return { ...state, colWidths: newColWidths };
    }

    case 'SET_ROW_HEIGHT': {
      const newRowHeights = new Map(state.rowHeights);
      newRowHeights.set(action.row, action.height);
      return { ...state, rowHeights: newRowHeights };
    }

    case 'SET_FREEZE_ROWS':
      return { ...state, freezeRows: action.rows };

    case 'ADD_CF_RULE':
      return { ...state, conditionalFormatRules: [...state.conditionalFormatRules, action.rule] };

    case 'DELETE_CF_RULE':
      return { ...state, conditionalFormatRules: state.conditionalFormatRules.filter(r => r.id !== action.id) };

    case 'CREATE_TABLE': {
      // Auto-detect data region when a single cell is selected
      let range = action.range;
      if (range.start.row === range.end.row && range.start.col === range.end.col) {
        range = detectDataRegion(range.start, state.cells, state.rowCount, state.colCount);
      }
      if (range.end.row - range.start.row < 1) return state; // need at least header + 1 data row
      for (const t of state.tables) {
        if (rangesOverlap(t.range, range)) return state;
      }
      const id = `table-${Date.now()}`;
      const filters = new Map<number, TableColumnFilter>();
      const updated = new Map(state.cells);
      for (let c = range.start.col; c <= range.end.col; c++) {
        filters.set(c, { checkedValues: null });
        const headerKey = `${range.start.row}:${c}`;
        const existing = updated.get(headerKey);
        if (!existing || existing.value.trim() === '') {
          const label = `Column ${colIndexToLetter(c)}`;
          updated.set(headerKey, { value: label, displayValue: label, formatting: { ...DEFAULT_FORMATTING } });
        }
      }
      const table: TableDefinition = { id, range, styleId: action.styleId ?? 'blue', sortColumn: null, sortDirection: null, filters };
      return { ...state, cells: updated, tables: [...state.tables, table] };
    }

    case 'SET_TABLE_STYLE': {
      return {
        ...state,
        tables: state.tables.map(t => t.id === action.id ? { ...t, styleId: action.styleId } : t),
      };
    }

    case 'DELETE_TABLE':
      return { ...state, tables: state.tables.filter(t => t.id !== action.id) };

    case 'SORT_TABLE': {
      const table = state.tables.find(t => t.id === action.id);
      if (!table) return state;
      if (action.direction === null) {
        // Clear sort state only (data stays in current order)
        const updatedTables = state.tables.map(t =>
          t.id === action.id ? { ...t, sortColumn: null, sortDirection: null } : t
        );
        return { ...state, tables: updatedTables };
      }
      const { range } = table;
      const dataStartRow = range.start.row + 1;
      const rows: Map<string, typeof DEFAULT_CELL>[] = [];
      for (let r = dataStartRow; r <= range.end.row; r++) {
        const rowCells = new Map<string, typeof DEFAULT_CELL>();
        for (let c = range.start.col; c <= range.end.col; c++) {
          const cell = state.cells.get(`${r}:${c}`);
          if (cell) rowCells.set(String(c), cell);
        }
        rows.push(rowCells);
      }
      rows.sort((a, b) => {
        const aVal = a.get(String(action.column))?.displayValue ?? '';
        const bVal = b.get(String(action.column))?.displayValue ?? '';
        const aNum = parseNum(aVal);
        const bNum = parseNum(bVal);
        let cmp: number;
        if (!isNaN(aNum) && !isNaN(bNum) && aVal !== '' && bVal !== '') {
          cmp = aNum - bNum;
        } else {
          cmp = aVal.localeCompare(bVal);
        }
        return action.direction === 'desc' ? -cmp : cmp;
      });
      const updated = new Map(state.cells);
      rows.forEach((rowCells, i) => {
        const targetRow = dataStartRow + i;
        for (let c = range.start.col; c <= range.end.col; c++) {
          const key = `${targetRow}:${c}`;
          const cellData = rowCells.get(String(c));
          if (cellData) updated.set(key, cellData);
          else updated.delete(key);
        }
      });
      const updatedTables = state.tables.map(t =>
        t.id === action.id ? { ...t, sortColumn: action.column, sortDirection: action.direction } : t
      );
      return { ...state, cells: recomputeFormulas(updated), tables: updatedTables };
    }

    case 'FILTER_TABLE': {
      const updatedTables = state.tables.map(t => {
        if (t.id !== action.id) return t;
        const newFilters = new Map(t.filters);
        newFilters.set(action.column, { checkedValues: action.checkedValues });
        return { ...t, filters: newFilters };
      });
      return { ...state, tables: updatedTables };
    }

    default:
      return state;
  }
}

// ── History reducer ───────────────────────────────────────────────────────────
// SET_SELECTION never pushes to history (it's just cursor movement).
// UNDO / REDO walk the history stack.
// Everything else saves current present to past before applying the change.
function historyReducer(state: HistoryState, action: Action): HistoryState {
  if (action.type === 'UNDO') {
    if (state.past.length === 0) return state;
    const previous = state.past[state.past.length - 1];
    // Preserve the current selection so the cursor doesn't jump on undo
    const restored = { ...previous, selectedRange: state.present.selectedRange, activeCell: state.present.activeCell };
    return {
      past: state.past.slice(0, -1),
      present: restored,
      future: [state.present, ...state.future].slice(0, MAX_HISTORY),
    };
  }

  if (action.type === 'REDO') {
    if (state.future.length === 0) return state;
    const next = state.future[0];
    const restored = { ...next, selectedRange: state.present.selectedRange, activeCell: state.present.activeCell };
    return {
      past: [...state.past, state.present].slice(-MAX_HISTORY),
      present: restored,
      future: state.future.slice(1),
    };
  }

  // Selection and resize changes don't touch history
  if (action.type === 'SET_SELECTION' || action.type === 'SET_COL_WIDTH' || action.type === 'SET_ROW_HEIGHT') {
    return { ...state, present: cellReducer(state.present, action) };
  }

  // All other actions: save current present, apply change, clear redo stack
  const newPresent = cellReducer(state.present, action);
  if (newPresent === state.present) return state; // no-op (e.g. FILL_DOWN on single row)
  return {
    past: [...state.past, state.present].slice(-MAX_HISTORY),
    present: newPresent,
    future: [],
  };
}

// ── Workbook reducer (routes to per-sheet history) ────────────────────────────
function workbookReducer(state: WorkbookState, action: Action): WorkbookState {
  switch (action.type) {
    case 'LOAD_WORKBOOK':
      return action.workbook;

    case 'ADD_SHEET':
      return {
        ...state,
        sheets: { ...state.sheets, [action.id]: createEmptyHistory() },
        tabs: [...state.tabs, { id: action.id, name: action.name, color: null }],
        activeSheetId: action.id,
      };

    case 'DELETE_SHEET': {
      if (state.tabs.length <= 1) return state;
      const newTabs = state.tabs.filter(t => t.id !== action.id);
      const newSheets = { ...state.sheets };
      delete newSheets[action.id];
      let newActive = state.activeSheetId;
      if (newActive === action.id) {
        const deletedIdx = state.tabs.findIndex(t => t.id === action.id);
        newActive = newTabs[Math.max(0, deletedIdx - 1)].id;
      }
      return { ...state, sheets: newSheets, tabs: newTabs, activeSheetId: newActive };
    }

    case 'RENAME_SHEET':
      return {
        ...state,
        tabs: state.tabs.map(t => t.id === action.id ? { ...t, name: action.name } : t),
      };

    case 'SET_SHEET_COLOR':
      return {
        ...state,
        tabs: state.tabs.map(t => t.id === action.id ? { ...t, color: action.color } : t),
      };

    case 'REORDER_SHEET': {
      const oldIdx = state.tabs.findIndex(t => t.id === action.id);
      if (oldIdx === -1 || oldIdx === action.newIndex) return state;
      const newTabs = [...state.tabs];
      const [removed] = newTabs.splice(oldIdx, 1);
      newTabs.splice(action.newIndex, 0, removed);
      return { ...state, tabs: newTabs };
    }

    case 'SWITCH_SHEET': {
      if (action.id === state.activeSheetId || !state.sheets[action.id]) return state;
      return { ...state, activeSheetId: action.id };
    }

    default: {
      const activeHistory = state.sheets[state.activeSheetId];
      const newHistory = historyReducer(activeHistory, action);
      if (newHistory === activeHistory) return state;
      return {
        ...state,
        sheets: { ...state.sheets, [state.activeSheetId]: newHistory },
      };
    }
  }
}

// ── Hook ──────────────────────────────────────────────────────────────────────
export function useSpreadsheet(overrideInitial?: WorkbookState) {
  const [workbook, dispatch] = useReducer(workbookReducer, overrideInitial ?? initialWorkbook);
  const activeHistory = workbook.sheets[workbook.activeSheetId];
  const state = activeHistory.present;

  const setCellValue = useCallback((row: number, col: number, value: string) => {
    dispatch({ type: 'SET_CELL', row, col, value });
  }, []);

  const applyFormat = useCallback((range: CellRange, formatting: Partial<CellFormatting>) => {
    dispatch({ type: 'APPLY_FORMAT', range, formatting });
  }, []);

  const applyColor = useCallback((range: CellRange, backgroundColor?: string, textColor?: string) => {
    dispatch({ type: 'APPLY_COLOR', range, backgroundColor, textColor });
  }, []);

  const applyActions = useCallback((actions: SpreadsheetAction[]) => {
    dispatch({ type: 'APPLY_ACTIONS', actions });
  }, []);

  const setSelection = useCallback((range: CellRange | null, activeCell: CellAddress | null) => {
    dispatch({ type: 'SET_SELECTION', range, activeCell });
  }, []);

  const undo = useCallback(() => dispatch({ type: 'UNDO' }), []);
  const redo = useCallback(() => dispatch({ type: 'REDO' }), []);

  const getActiveCellValue = useCallback(() => {
    if (!state.activeCell) return '';
    return state.cells.get(`${state.activeCell.row}:${state.activeCell.col}`)?.value ?? '';
  }, [state.activeCell, state.cells]);

  const getActiveCellFormatting = useCallback((): CellFormatting => {
    if (!state.activeCell) return { ...DEFAULT_FORMATTING };
    return state.cells.get(`${state.activeCell.row}:${state.activeCell.col}`)?.formatting ?? { ...DEFAULT_FORMATTING };
  }, [state.activeCell, state.cells]);

  const handleExportCSV = useCallback(() => {
    downloadCSV(exportToCSV(state.cells, state.rowCount, state.colCount));
  }, [state.cells, state.rowCount, state.colCount]);

  const handleImportCSV = useCallback((file: File) => {
    const reader = new FileReader();
    const isXLSX = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    if (isXLSX) {
      reader.onload = (e) => {
        const imported = importAllSheetsFromXLSX(e.target?.result as ArrayBuffer);
        // Build a full workbook with all sheets
        const sheets: Record<string, HistoryState> = {};
        for (const [id, sheet] of Object.entries(imported.sheets)) {
          sheets[id] = {
            past: [],
            present: {
              cells: sheet.cells,
              rowCount: sheet.rowCount,
              colCount: sheet.colCount,
              selectedRange: null,
              activeCell: null,
              colWidths: new Map(),
              rowHeights: new Map(),
              freezeRows: 0,
              conditionalFormatRules: [],
              tables: [],
            },
            future: [],
          };
        }
        const tabs: SheetTabConfig[] = imported.sheetNames.map(s => ({ id: s.id, name: s.name, color: null }));
        const activeSheetId = imported.sheetNames[0]?.id ?? 'sheet-default';
        dispatch({ type: 'LOAD_WORKBOOK', workbook: { sheets, tabs, activeSheetId } });
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = (e) => {
        const result = importFromCSV(e.target?.result as string);
        dispatch({ type: 'IMPORT_CSV', ...result });
      };
      reader.readAsText(file);
    }
  }, []);

  const fillSeries = useCallback((range: CellRange) => dispatch({ type: 'FILL_SERIES', range }), []);
  const fillDown   = useCallback((range: CellRange) => dispatch({ type: 'FILL_DOWN',   range }), []);
  const fillRight  = useCallback((range: CellRange) => dispatch({ type: 'FILL_RIGHT',  range }), []);

  const insertRow = useCallback((at: number) => dispatch({ type: 'INSERT_ROW', at }), []);
  const deleteRow = useCallback((at: number) => dispatch({ type: 'DELETE_ROW', at }), []);
  const insertCol = useCallback((at: number) => dispatch({ type: 'INSERT_COL', at }), []);
  const deleteCol = useCallback((at: number) => dispatch({ type: 'DELETE_COL', at }), []);
  const setColWidth = useCallback((col: number, width: number) => dispatch({ type: 'SET_COL_WIDTH', col, width }), []);
  const setRowHeight = useCallback((row: number, height: number) => dispatch({ type: 'SET_ROW_HEIGHT', row, height }), []);
  const setFreezeRows = useCallback((rows: number) => dispatch({ type: 'SET_FREEZE_ROWS', rows }), []);
  const addCFRule = useCallback((rule: ConditionalFormatRule) => dispatch({ type: 'ADD_CF_RULE', rule }), []);
  const deleteCFRule = useCallback((id: string) => dispatch({ type: 'DELETE_CF_RULE', id }), []);

  const createTable = useCallback((range: CellRange, styleId?: string) => dispatch({ type: 'CREATE_TABLE', range, styleId }), []);
  const deleteTable = useCallback((id: string) => dispatch({ type: 'DELETE_TABLE', id }), []);
  const setTableStyle = useCallback((id: string, styleId: string) => dispatch({ type: 'SET_TABLE_STYLE', id, styleId }), []);
  const sortTable = useCallback((id: string, column: number, direction: SortDirection) =>
    dispatch({ type: 'SORT_TABLE', id, column, direction }), []);
  const filterTable = useCallback((id: string, column: number, checkedValues: Set<string> | null) =>
    dispatch({ type: 'FILTER_TABLE', id, column, checkedValues }), []);

  const getSerializedCells = useCallback(() => {
    const result: { address: string; value: string; displayValue: string }[] = [];
    state.cells.forEach((cell, key) => {
      if (cell.value !== '') {
        const [r, c] = key.split(':').map(Number);
        result.push({ address: `${String.fromCharCode(65 + c)}${r + 1}`, value: cell.value, displayValue: cell.displayValue });
      }
    });
    return result;
  }, [state.cells]);

  // Sheet management
  const addSheet = useCallback((customName?: string) => {
    const id = `sheet-${Date.now()}`;
    const usedNames = new Set(workbook.tabs.map(t => t.name));
    let name: string;
    if (customName && !usedNames.has(customName)) {
      name = customName;
    } else {
      let n = workbook.tabs.length + 1;
      while (usedNames.has(`Sheet${n}`)) n++;
      name = `Sheet${n}`;
    }
    dispatch({ type: 'ADD_SHEET', id, name });
  }, [workbook.tabs]);

  const deleteSheet = useCallback((id: string) => {
    dispatch({ type: 'DELETE_SHEET', id });
  }, []);

  const renameSheet = useCallback((id: string, name: string) => {
    dispatch({ type: 'RENAME_SHEET', id, name });
  }, []);

  const setSheetColor = useCallback((id: string, color: string | null) => {
    dispatch({ type: 'SET_SHEET_COLOR', id, color });
  }, []);

  const reorderSheet = useCallback((id: string, newIndex: number) => {
    dispatch({ type: 'REORDER_SHEET', id, newIndex });
  }, []);

  const switchSheet = useCallback((id: string) => {
    dispatch({ type: 'SWITCH_SHEET', id });
  }, []);

  const loadWorkbook = useCallback((wb: WorkbookState) => {
    dispatch({ type: 'LOAD_WORKBOOK', workbook: wb });
  }, []);

  return {
    _workbook: workbook,
    loadWorkbook,
    ...state,
    tables: state.tables ?? [],
    canUndo: activeHistory.past.length > 0,
    canRedo: activeHistory.future.length > 0,
    undo,
    redo,
    setCellValue,
    applyFormat,
    applyColor,
    applyActions,
    setSelection,
    getActiveCellValue,
    getActiveCellFormatting,
    exportCSV: handleExportCSV,
    importCSV: handleImportCSV,
    fillSeries,
    fillDown,
    fillRight,
    insertRow,
    deleteRow,
    insertCol,
    deleteCol,
    setColWidth,
    setRowHeight,
    setFreezeRows,
    addCFRule,
    deleteCFRule,
    getSerializedCells,
    // Table management
    createTable,
    deleteTable,
    setTableStyle,
    sortTable,
    filterTable,
    // Sheet management
    tabs: workbook.tabs,
    activeSheetId: workbook.activeSheetId,
    addSheet,
    deleteSheet,
    renameSheet,
    setSheetColor,
    reorderSheet,
    switchSheet,
  };
}
