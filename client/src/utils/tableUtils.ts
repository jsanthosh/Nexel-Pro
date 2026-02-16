import { CellFormatting, CellRange, CellAddress, SheetData, TableDefinition, TABLE_STYLES } from '../types/spreadsheet';
import { isInRange } from './cellRange';

/**
 * Auto-detect the contiguous data region from a single cell.
 * Expands outward from the seed cell until hitting empty rows/columns.
 * Similar to Excel's "current region" (Ctrl+Shift+*).
 */
export function detectDataRegion(seed: CellAddress, cells: SheetData, maxRow: number, maxCol: number): CellRange {
  const hasData = (r: number, c: number) => {
    const val = cells.get(`${r}:${c}`)?.value ?? '';
    return val.trim() !== '';
  };

  // Expand up
  let top = seed.row;
  while (top > 0) {
    let rowEmpty = true;
    // Check across current width estimate
    for (let c = Math.max(0, seed.col - 50); c <= Math.min(maxCol - 1, seed.col + 50); c++) {
      if (hasData(top - 1, c)) { rowEmpty = false; break; }
    }
    if (rowEmpty) break;
    top--;
  }

  // Expand down
  let bottom = seed.row;
  while (bottom < maxRow - 1) {
    let rowEmpty = true;
    for (let c = Math.max(0, seed.col - 50); c <= Math.min(maxCol - 1, seed.col + 50); c++) {
      if (hasData(bottom + 1, c)) { rowEmpty = false; break; }
    }
    if (rowEmpty) break;
    bottom++;
  }

  // Expand left
  let left = seed.col;
  while (left > 0) {
    let colEmpty = true;
    for (let r = top; r <= bottom; r++) {
      if (hasData(r, left - 1)) { colEmpty = false; break; }
    }
    if (colEmpty) break;
    left--;
  }

  // Expand right
  let right = seed.col;
  while (right < maxCol - 1) {
    let colEmpty = true;
    for (let r = top; r <= bottom; r++) {
      if (hasData(r, right + 1)) { colEmpty = false; break; }
    }
    if (colEmpty) break;
    right++;
  }

  return { start: { row: top, col: left }, end: { row: bottom, col: right } };
}

export function rangesOverlap(a: CellRange, b: CellRange): boolean {
  return !(
    a.end.row < b.start.row || b.end.row < a.start.row ||
    a.end.col < b.start.col || b.end.col < a.start.col
  );
}

export function findTableForCell(
  row: number, col: number, tables: TableDefinition[],
): TableDefinition | undefined {
  if (!Array.isArray(tables)) return undefined;
  return tables.find(t => isInRange(row, col, t.range));
}

export function computeTableFormattingMap(
  tables: TableDefinition[],
): Map<string, Partial<CellFormatting>> {
  const map = new Map<string, Partial<CellFormatting>>();
  if (!Array.isArray(tables)) return map;

  for (const table of tables) {
    const { range } = table;
    const style = TABLE_STYLES.find(s => s.id === (table.styleId ?? 'blue')) ?? TABLE_STYLES[0];
    const headerRow = range.start.row;

    for (let c = range.start.col; c <= range.end.col; c++) {
      map.set(`${headerRow}:${c}`, {
        bold: true,
        backgroundColor: style.headerBg,
        textColor: style.headerText,
      });

      for (let r = range.start.row + 1; r <= range.end.row; r++) {
        const dataRowIndex = r - range.start.row - 1;
        const isOddRow = dataRowIndex % 2 === 1;
        map.set(`${r}:${c}`, {
          backgroundColor: isOddRow ? style.oddRowBg : style.evenRowBg,
        });
      }
    }
  }

  return map;
}

export function computeHiddenRows(
  tables: TableDefinition[],
  cells: SheetData,
): Set<number> {
  const hidden = new Set<number>();
  if (!Array.isArray(tables)) return hidden;

  for (const table of tables) {
    const { range, filters } = table;
    const dataStartRow = range.start.row + 1;

    for (let r = dataStartRow; r <= range.end.row; r++) {
      for (const [col, filter] of filters) {
        if (filter.checkedValues === null) continue;
        const cellVal = cells.get(`${r}:${col}`)?.displayValue ?? '';
        if (!filter.checkedValues.has(cellVal)) {
          hidden.add(r);
          break;
        }
      }
    }
  }

  return hidden;
}

export function getUniqueColumnValues(
  table: TableDefinition,
  col: number,
  cells: SheetData,
): string[] {
  const values = new Set<string>();
  for (let r = table.range.start.row + 1; r <= table.range.end.row; r++) {
    values.add(cells.get(`${r}:${col}`)?.displayValue ?? '');
  }
  return Array.from(values).sort();
}

export function computeTableCellClasses(
  tables: TableDefinition[],
): Map<string, string> {
  const map = new Map<string, string>();
  if (!Array.isArray(tables)) return map;
  for (const table of tables) {
    const { range } = table;
    for (let r = range.start.row; r <= range.end.row; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        const classes: string[] = ['table-cell'];
        if (r === range.start.row) classes.push('table-border-top');
        if (r === range.end.row) classes.push('table-border-bottom');
        if (c === range.start.col) classes.push('table-border-left');
        if (c === range.end.col) classes.push('table-border-right');
        if (r === range.start.row) classes.push('table-header-cell');
        map.set(`${r}:${c}`, classes.join(' '));
      }
    }
  }
  return map;
}
