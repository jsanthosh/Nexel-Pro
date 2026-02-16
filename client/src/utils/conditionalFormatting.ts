import { CellData, CellFormatting, ConditionalFormatRule, CellRange } from '../types/spreadsheet';
import { isInRange } from './cellRange';

export function evaluateConditionalFormatting(
  row: number,
  col: number,
  cellData: CellData,
  rules: ConditionalFormatRule[],
  allCells: Map<string, CellData>,
): Partial<CellFormatting> | null {
  const applicableRules = rules
    .filter(rule => isInRange(row, col, rule.range))
    .sort((a, b) => b.priority - a.priority);

  let merged: Partial<CellFormatting> | null = null;

  for (const rule of applicableRules) {
    if (evaluateCondition(cellData, rule, allCells)) {
      merged = { ...(merged ?? {}), ...rule.formatting };
    }
  }

  return merged;
}

function evaluateCondition(
  cellData: CellData,
  rule: ConditionalFormatRule,
  allCells: Map<string, CellData>,
): boolean {
  const displayVal = cellData.displayValue || cellData.value;
  const numVal = Number(displayVal);

  switch (rule.condition) {
    case 'greaterThan':
      return !isNaN(numVal) && numVal > Number(rule.values[0]);
    case 'lessThan':
      return !isNaN(numVal) && numVal < Number(rule.values[0]);
    case 'equalTo':
      return displayVal === rule.values[0] || (!isNaN(numVal) && numVal === Number(rule.values[0]));
    case 'between': {
      const min = Number(rule.values[0]);
      const max = Number(rule.values[1]);
      return !isNaN(numVal) && numVal >= min && numVal <= max;
    }
    case 'textContains':
      return displayVal.toLowerCase().includes((rule.values[0] ?? '').toLowerCase());
    case 'isEmpty':
      return displayVal === '';
    case 'isNotEmpty':
      return displayVal !== '';
    case 'duplicateValues': {
      if (displayVal === '') return false;
      let count = 0;
      const { start, end } = rule.range;
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          const cell = allCells.get(`${r}:${c}`);
          const val = cell?.displayValue || cell?.value || '';
          if (val === displayVal) count++;
          if (count > 1) return true;
        }
      }
      return false;
    }
    default:
      return false;
  }
}
