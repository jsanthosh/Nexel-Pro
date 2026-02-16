import { useState, useCallback } from 'react';
import { ChatMessage, CellRange, ClaudeResponse, SpreadsheetAction, ConditionalFormatRule, ConditionType } from '../types/spreadsheet';
import { parseRange, rangeToString, colLetterToIndex, iterateRange } from '../utils/cellRange';
import { ChartType } from '../types/chart';
import { rangesOverlap } from '../utils/tableUtils';

interface UseSpreadsheetReturn {
  selectedRange: any;
  rowCount: number;
  colCount: number;
  applyActions: (actions: SpreadsheetAction[]) => void;
  addSheet: (name?: string) => void;
  getSerializedCells: () => { address: string; value: string; displayValue: string }[];
  // Extended methods for chat-driven actions
  addCFRule: (rule: ConditionalFormatRule) => void;
  deleteCFRule: (id: string) => void;
  conditionalFormatRules: ConditionalFormatRule[];
  createTable: (range: CellRange, styleId?: string) => void;
  deleteTable: (id: string) => void;
  tables: { id: string; range: CellRange; styleId: string }[];
  insertRow: (at: number) => void;
  deleteRow: (at: number) => void;
  insertCol: (at: number) => void;
  deleteCol: (at: number) => void;
  setFreezeRows: (rows: number) => void;
  setColWidth: (col: number, width: number) => void;
  setCellValue: (row: number, col: number, value: string) => void;
}

interface AddChartFn {
  (dataRange: CellRange, chartType?: ChartType): void;
}

export function useChat(spreadsheet: UseSpreadsheetReturn, addChart?: AddChartFn) {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: '0',
      role: 'assistant',
      content: 'Hi! I\'m your spreadsheet assistant. I can do everything — just ask! For example:\n\u2022 "Bold A1:B10"\n\u2022 "Highlight values > 500 in red"\n\u2022 "Format as table"\n\u2022 "Insert 3 rows before row 5"\n\u2022 "Create a sales dashboard"\n\u2022 "Freeze the header row"',
      timestamp: new Date(),
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const [model, setModel] = useState<string>('claude-sonnet-4-5-20250929');

  const sendMessage = useCallback(async (userInput: string) => {
    const userMsg: ChatMessage = {
      id: Date.now().toString(),
      role: 'user',
      content: userInput,
      timestamp: new Date(),
    };

    const loadingMsg: ChatMessage = {
      id: `loading-${Date.now()}`,
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      isLoading: true,
    };

    setMessages(prev => [...prev, userMsg, loadingMsg]);
    setIsLoading(true);

    // Build conversation history (last 10 non-loading messages)
    const history = messages
      .filter(m => !m.isLoading)
      .slice(-10)
      .map(m => ({ role: m.role, content: m.content }));

    // Build spreadsheet context
    const context = {
      cells: spreadsheet.getSerializedCells(),
      selectedRange: spreadsheet.selectedRange
        ? rangeToString(spreadsheet.selectedRange)
        : null,
      rowCount: spreadsheet.rowCount,
      colCount: spreadsheet.colCount,
      conditionalFormatRules: spreadsheet.conditionalFormatRules.map(r => ({
        id: r.id,
        range: rangeToString(r.range),
        condition: r.condition,
        values: r.values,
        formatting: r.formatting,
      })),
      tables: spreadsheet.tables.map(t => ({
        id: t.id,
        range: rangeToString(t.range),
        styleId: t.styleId,
      })),
    };

    try {
      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: userInput, context, history, model }),
      });

      if (!response.ok) throw new Error(`Server error: ${response.status}`);

      const data: ClaudeResponse = await response.json();

      // Apply spreadsheet actions
      if (data.actions?.length) {
        const cellActions: SpreadsheetAction[] = []; // FORMAT, SET_COLOR, SET_VALUE batched via applyActions

        for (const action of data.actions) {
          switch (action.type) {
            case 'CREATE_SHEET':
              // Flush any accumulated cell actions to the CURRENT sheet before switching
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              spreadsheet.addSheet(action.name);
              break;

            case 'CREATE_CHART':
              if (addChart) {
                try {
                  const dataRange = parseRange(action.range);
                  addChart(dataRange, action.chartType ?? 'bar');
                } catch {}
              }
              break;

            case 'ADD_CF_RULE': {
              // Flush cell actions first so CF applies to up-to-date data
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              try {
                const range = parseRange(action.range);
                const rule: ConditionalFormatRule = {
                  id: `cf-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
                  range,
                  condition: action.condition as ConditionType,
                  values: action.values ?? [],
                  formatting: action.formatting ?? {},
                  priority: spreadsheet.conditionalFormatRules.length,
                };
                spreadsheet.addCFRule(rule);
              } catch {}
              break;
            }

            case 'DELETE_CF_RULES': {
              if (action.range) {
                try {
                  const range = parseRange(action.range);
                  for (const rule of [...spreadsheet.conditionalFormatRules]) {
                    if (rangesOverlap(rule.range, range)) {
                      spreadsheet.deleteCFRule(rule.id);
                    }
                  }
                } catch {}
              } else {
                for (const rule of [...spreadsheet.conditionalFormatRules]) {
                  spreadsheet.deleteCFRule(rule.id);
                }
              }
              break;
            }

            case 'CREATE_TABLE': {
              // Flush cell actions first so table sees current data
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              try {
                const range = parseRange(action.range);
                spreadsheet.createTable(range, action.styleId);
              } catch {}
              break;
            }

            case 'DELETE_TABLE': {
              try {
                const range = parseRange(action.range);
                const table = spreadsheet.tables.find(t => rangesOverlap(t.range, range));
                if (table) spreadsheet.deleteTable(table.id);
              } catch {}
              break;
            }

            case 'INSERT_ROW': {
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              const count = action.count ?? 1;
              const at = action.at - 1; // convert 1-indexed to 0-indexed
              for (let i = 0; i < count; i++) {
                spreadsheet.insertRow(at);
              }
              break;
            }

            case 'DELETE_ROW': {
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              const count = action.count ?? 1;
              const at = action.at - 1;
              for (let i = 0; i < count; i++) {
                spreadsheet.deleteRow(at);
              }
              break;
            }

            case 'INSERT_COL': {
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              const count = action.count ?? 1;
              const at = action.at - 1;
              for (let i = 0; i < count; i++) {
                spreadsheet.insertCol(at);
              }
              break;
            }

            case 'DELETE_COL': {
              if (cellActions.length) {
                spreadsheet.applyActions([...cellActions]);
                cellActions.length = 0;
              }
              const count = action.count ?? 1;
              const at = action.at - 1;
              for (let i = 0; i < count; i++) {
                spreadsheet.deleteCol(at);
              }
              break;
            }

            case 'FREEZE_ROWS':
              spreadsheet.setFreezeRows(action.count);
              break;

            case 'SET_COL_WIDTH': {
              try {
                const colIdx = colLetterToIndex(action.column);
                spreadsheet.setColWidth(colIdx, action.width);
              } catch {}
              break;
            }

            case 'CLEAR_RANGE': {
              try {
                const range = parseRange(action.range);
                for (const key of iterateRange(range)) {
                  const [r, c] = key.split(':').map(Number);
                  spreadsheet.setCellValue(r, c, '');
                }
              } catch {}
              break;
            }

            // Batch-able cell actions
            case 'FORMAT':
            case 'SET_COLOR':
            case 'SET_VALUE':
              cellActions.push(action);
              break;

            // Informational — no mutation (message extracted below)
            case 'QUERY_RESULT':
            case 'ERROR':
              break;
          }
        }

        // Flush remaining cell actions
        if (cellActions.length) {
          spreadsheet.applyActions(cellActions);
        }
      }

      // Extract detailed messages from QUERY_RESULT and ERROR actions
      const queryMessages: string[] = [];
      for (const a of data.actions ?? []) {
        if ((a.type === 'QUERY_RESULT' || a.type === 'ERROR') && a.message) {
          queryMessages.push(a.message);
        }
      }

      // Show detailed query/error messages if available, fall back to explanation
      const displayContent = queryMessages.length > 0
        ? queryMessages.join('\n\n') + (data.explanation ? `\n\n${data.explanation}` : '')
        : data.explanation || 'Done!';

      const assistantMsg: ChatMessage = {
        id: Date.now().toString(),
        role: 'assistant',
        content: displayContent,
        timestamp: new Date(),
      };

      setMessages(prev => prev.filter(m => !m.isLoading).concat(assistantMsg));
    } catch (err) {
      const errorMsg: ChatMessage = {
        id: Date.now().toString(),
        role: 'assistant',
        content: `Error: ${err instanceof Error ? err.message : 'Failed to connect to assistant'}`,
        timestamp: new Date(),
      };
      setMessages(prev => prev.filter(m => !m.isLoading).concat(errorMsg));
    } finally {
      setIsLoading(false);
    }
  }, [messages, spreadsheet, addChart, model]);

  return { messages, isLoading, sendMessage, model, setModel };
}
