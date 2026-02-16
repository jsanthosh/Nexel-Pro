import * as XLSX from 'xlsx';
import { SheetData, DEFAULT_CELL, DEFAULT_FORMATTING } from '../types/spreadsheet';

export function exportToCSV(cells: SheetData, rowCount: number, colCount: number): string {
  const rows: string[] = [];
  for (let r = 0; r < rowCount; r++) {
    const cols: string[] = [];
    for (let c = 0; c < colCount; c++) {
      const cell = cells.get(`${r}:${c}`);
      const value = cell?.displayValue ?? '';
      cols.push(
        value.includes(',') || value.includes('"') || value.includes('\n')
          ? `"${value.replace(/"/g, '""')}"`
          : value
      );
    }
    rows.push(cols.join(','));
  }

  // Trim trailing empty rows
  while (rows.length > 0 && rows[rows.length - 1].replace(/,/g, '').trim() === '') {
    rows.pop();
  }

  return rows.join('\n');
}

export function downloadCSV(csvContent: string, filename = 'spreadsheet.csv'): void {
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

export function importFromCSV(
  csvContent: string
): { cells: SheetData; rowCount: number; colCount: number } {
  const cells: SheetData = new Map();
  const lines = parseCSVLines(csvContent);
  let maxCol = 0;

  lines.forEach((cols, rowIdx) => {
    maxCol = Math.max(maxCol, cols.length);
    cols.forEach((value, colIdx) => {
      const trimmed = value.trim();
      if (trimmed !== '') {
        cells.set(`${rowIdx}:${colIdx}`, {
          ...DEFAULT_CELL,
          formatting: { ...DEFAULT_FORMATTING },
          value: trimmed,
          displayValue: trimmed,
        });
      }
    });
  });

  return {
    cells,
    rowCount: Math.max(lines.length + 10, 100),
    colCount: Math.max(maxCol + 5, 26),
  };
}

export function importFromXLSX(
  data: ArrayBuffer
): { cells: SheetData; rowCount: number; colCount: number } {
  const result = importAllSheetsFromXLSX(data);
  const firstSheet = Object.values(result.sheets)[0];
  return firstSheet ?? { cells: new Map(), rowCount: 100, colCount: 26 };
}

export interface ImportedSheet {
  cells: SheetData;
  rowCount: number;
  colCount: number;
}

export interface ImportedWorkbook {
  sheets: Record<string, ImportedSheet>;
  sheetNames: { id: string; name: string }[];
}

export function importAllSheetsFromXLSX(data: ArrayBuffer): ImportedWorkbook {
  const workbook = XLSX.read(data, { type: 'array' });
  const sheets: Record<string, ImportedSheet> = {};
  const sheetNames: { id: string; name: string }[] = [];

  workbook.SheetNames.forEach((name, idx) => {
    const sheet = workbook.Sheets[name];
    const cells: SheetData = new Map();
    const id = `sheet-import-${Date.now()}-${idx}`;

    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    if (range) {
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[addr];
          if (cell) {
            const value = cell.w ?? String(cell.v ?? '');
            if (value !== '') {
              cells.set(`${r}:${c}`, {
                ...DEFAULT_CELL,
                formatting: { ...DEFAULT_FORMATTING },
                value,
                displayValue: value,
              });
            }
          }
        }
      }
    }

    sheets[id] = {
      cells,
      rowCount: range ? Math.max(range.e.r + 11, 100) : 100,
      colCount: range ? Math.max(range.e.c + 6, 26) : 26,
    };
    sheetNames.push({ id, name });
  });

  return { sheets, sheetNames };
}

function parseCSVLines(csv: string): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  let i = 0;

  while (i < csv.length) {
    const char = csv[i];
    if (char === '"') {
      if (inQuotes && csv[i + 1] === '"') { currentField += '"'; i += 2; continue; }
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      currentRow.push(currentField); currentField = '';
    } else if ((char === '\n' || (char === '\r' && csv[i + 1] === '\n')) && !inQuotes) {
      currentRow.push(currentField); rows.push(currentRow);
      currentRow = []; currentField = '';
      if (char === '\r') i++;
    } else {
      currentField += char;
    }
    i++;
  }
  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField); rows.push(currentRow);
  }
  return rows;
}
