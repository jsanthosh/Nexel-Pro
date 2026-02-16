import { SheetData, SpreadsheetAction, DEFAULT_CELL, DEFAULT_FORMATTING } from '../types/spreadsheet';
import { parseRange, iterateRange } from './cellRange';
import { evaluateCell, recomputeFormulas } from './formulaEngine';

export function applyActions(cells: SheetData, actions: SpreadsheetAction[]): SheetData {
  let updated = new Map(cells);

  for (const action of actions) {
    switch (action.type) {
      case 'FORMAT': {
        const range = parseRange(action.range);
        for (const key of iterateRange(range)) {
          const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
          updated.set(key, {
            ...existing,
            formatting: { ...existing.formatting, ...action.formatting },
          });
        }
        break;
      }

      case 'SET_COLOR': {
        const range = parseRange(action.range);
        for (const key of iterateRange(range)) {
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
        break;
      }

      case 'SET_VALUE': {
        const range = parseRange(action.range);
        for (const key of iterateRange(range)) {
          const existing = updated.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
          const displayValue = evaluateCell(action.value, updated);
          updated.set(key, { ...existing, value: action.value, displayValue });
        }
        break;
      }

      // QUERY_RESULT and ERROR don't mutate cells
      default:
        break;
    }
  }

  return recomputeFormulas(updated);
}
