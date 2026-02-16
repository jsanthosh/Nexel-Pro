import { CellAddress, CellRange } from '../types/spreadsheet';

export function colIndexToLetter(index: number): string {
  let result = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function colLetterToIndex(letters: string): number {
  return letters
    .toUpperCase()
    .split('')
    .reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1;
}

export function parseCellAddress(address: string): CellAddress {
  const match = address.trim().match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return {
    col: colLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

export function parseRange(rangeStr: string): CellRange {
  const parts = rangeStr.split(':');
  const start = parseCellAddress(parts[0].trim());
  const end = parts[1] ? parseCellAddress(parts[1].trim()) : start;
  return {
    start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) },
  };
}

export function iterateRange(range: CellRange): string[] {
  const keys: string[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      keys.push(`${r}:${c}`);
    }
  }
  return keys;
}

export function cellKey(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

export function isInRange(row: number, col: number, range: CellRange | null): boolean {
  if (!range) return false;
  return (
    row >= range.start.row &&
    row <= range.end.row &&
    col >= range.start.col &&
    col <= range.end.col
  );
}

export function rangeToString(range: CellRange): string {
  const start = cellKey(range.start.row, range.start.col);
  const end = cellKey(range.end.row, range.end.col);
  return start === end ? start : `${start}:${end}`;
}
