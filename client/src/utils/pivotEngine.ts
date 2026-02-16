import { SheetData } from '../types/spreadsheet';
import { PivotConfig, PivotResult, AggregationType } from '../types/pivot';

export function computePivotTable(cells: SheetData, config: PivotConfig): PivotResult {
  const { sourceRange, rowFields, columnFields, valueFields } = config;

  if (rowFields.length === 0 || valueFields.length === 0) {
    return { headers: [], rows: [] };
  }

  // Extract source data
  const headers: string[] = [];
  for (let c = sourceRange.start.col; c <= sourceRange.end.col; c++) {
    const cell = cells.get(`${sourceRange.start.row}:${c}`);
    headers.push(cell?.displayValue ?? cell?.value ?? `Col ${c + 1}`);
  }

  const dataRows: string[][] = [];
  for (let r = sourceRange.start.row + 1; r <= sourceRange.end.row; r++) {
    const row: string[] = [];
    for (let c = sourceRange.start.col; c <= sourceRange.end.col; c++) {
      const cell = cells.get(`${r}:${c}`);
      row.push(cell?.displayValue ?? cell?.value ?? '');
    }
    dataRows.push(row);
  }

  // Group data: rowKey -> colKey -> number[][]
  const groupMap = new Map<string, Map<string, number[][]>>();

  for (const row of dataRows) {
    const rowKey = rowFields.map(f => row[f.columnIndex] ?? '').join('||');
    const colKey = columnFields.length > 0
      ? columnFields.map(f => row[f.columnIndex] ?? '').join('||')
      : '__all__';

    if (!groupMap.has(rowKey)) groupMap.set(rowKey, new Map());
    const colMap = groupMap.get(rowKey)!;
    if (!colMap.has(colKey)) colMap.set(colKey, valueFields.map(() => []));

    const valArrays = colMap.get(colKey)!;
    valueFields.forEach((vf, i) => {
      const num = Number(row[vf.columnIndex]);
      if (!isNaN(num)) valArrays[i].push(num);
    });
  }

  // Unique column keys
  const uniqueColKeys = [...new Set(
    [...groupMap.values()].flatMap(m => [...m.keys()])
  )].sort();

  // Build result headers
  const resultHeaders: string[] = [
    ...rowFields.map(f => f.label),
    ...uniqueColKeys.flatMap(ck =>
      valueFields.map(vf =>
        columnFields.length > 0
          ? `${ck} - ${vf.label} (${vf.aggregation})`
          : `${vf.label} (${vf.aggregation})`
      )
    ),
  ];

  // Build result rows
  const resultRows: (string | number)[][] = [];
  const sortedRowKeys = [...groupMap.keys()].sort();

  for (const rowKey of sortedRowKeys) {
    const rowParts = rowKey.split('||');
    const colMap = groupMap.get(rowKey)!;
    const resultRow: (string | number)[] = [...rowParts];

    for (const colKey of uniqueColKeys) {
      const valArrays = colMap.get(colKey) ?? valueFields.map(() => []);
      for (let i = 0; i < valueFields.length; i++) {
        resultRow.push(aggregate(valArrays[i], valueFields[i].aggregation));
      }
    }

    resultRows.push(resultRow);
  }

  return { headers: resultHeaders, rows: resultRows };
}

function aggregate(values: number[], type: AggregationType): number {
  if (values.length === 0) return 0;
  switch (type) {
    case 'sum': return values.reduce((a, b) => a + b, 0);
    case 'count': return values.length;
    case 'average': return Math.round((values.reduce((a, b) => a + b, 0) / values.length) * 100) / 100;
    case 'min': return Math.min(...values);
    case 'max': return Math.max(...values);
  }
}
