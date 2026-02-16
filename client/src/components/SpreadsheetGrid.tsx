import React, { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import ReactDOM from 'react-dom';
import { SheetData, CellRange, CellAddress, CellData, CellFormatting, ConditionalFormatRule, TableDefinition, SortDirection, DEFAULT_CELL, DEFAULT_FORMATTING } from '../types/spreadsheet';
import { colIndexToLetter, isInRange, cellKey } from '../utils/cellRange';
import { evaluateConditionalFormatting } from '../utils/conditionalFormatting';
import { computeTableFormattingMap, computeHiddenRows, computeTableCellClasses, findTableForCell } from '../utils/tableUtils';
import Cell from './Cell';
import TableFilterDropdown from './TableFilterDropdown';

interface SpreadsheetGridProps {
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
  onCellChange: (row: number, col: number, value: string) => void;
  onSelectionChange: (range: CellRange, activeCell: CellAddress) => void;
  onUndo: () => void;
  onRedo: () => void;
  onFillDown: (range: CellRange) => void;
  onFillRight: (range: CellRange) => void;
  onFillSeries: (range: CellRange) => void;
  onColWidthChange: (col: number, width: number) => void;
  onRowHeightChange: (row: number, height: number) => void;
  onInsertRow: (at: number) => void;
  onDeleteRow: (at: number) => void;
  onInsertCol: (at: number) => void;
  onDeleteCol: (at: number) => void;
  onSetFreezeRows: (rows: number) => void;
  onSortTable: (tableId: string, col: number, direction: SortDirection) => void;
  onFilterTable: (tableId: string, col: number, checkedValues: Set<string> | null) => void;
  onDeleteTable: (tableId: string) => void;
  onFormat?: (range: CellRange, formatting: Partial<CellFormatting>) => void;
}

type ContextMenu = { x: number; y: number; type: 'row' | 'col' | 'cell'; index: number; indices?: number[]; tableId?: string } | null;

const ROW_H = 24;
const ROW_HEADER_W = 40;
const DEFAULT_COL_W = 100;
const OVERSCAN_ROWS = 10;
const OVERSCAN_COLS = 5;

export default function SpreadsheetGrid({
  cells, rowCount, colCount, selectedRange, activeCell,
  colWidths, rowHeights, freezeRows, conditionalFormatRules, tables,
  onCellChange, onSelectionChange, onUndo, onRedo,
  onFillDown, onFillRight, onFillSeries, onColWidthChange, onRowHeightChange,
  onInsertRow, onDeleteRow, onInsertCol, onDeleteCol, onSetFreezeRows,
  onSortTable, onFilterTable, onDeleteTable, onFormat,
}: SpreadsheetGridProps) {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState('');
  const [editOriginalValue, setEditOriginalValue] = useState('');
  const [isSelecting, setIsSelecting] = useState(false);
  const [selectionStart, setSelectionStart] = useState<CellAddress | null>(null);
  const [headerSelectMode, setHeaderSelectMode] = useState<'none' | 'col' | 'row'>('none');
  const [contextMenu, setContextMenu] = useState<ContextMenu>(null);
  const [extraSelections, setExtraSelections] = useState<CellRange[]>([]);
  const [filterDropdown, setFilterDropdown] = useState<{ table: TableDefinition; col: number; position: { x: number; y: number } } | null>(null);
  const [formulaRefStart, setFormulaRefStart] = useState<CellAddress | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const copiedRef = useRef<{ data: CellData[][]; range: CellRange } | null>(null);
  const resizeRef = useRef<{ col: number; startX: number; startWidth: number; otherCols?: number[] } | null>(null);
  const rowResizeRef = useRef<{ row: number; startY: number; startHeight: number; otherRows?: number[] } | null>(null);
  const editValueRef = useRef(editValue);
  editValueRef.current = editValue;

  // ── Fill handle drag state ──
  const [isFillDragging, setIsFillDragging] = useState(false);
  const [fillDragEnd, setFillDragEnd] = useState<CellAddress | null>(null);
  const fillDragBaseRef = useRef<CellRange | null>(null);

  // ── Viewport virtualization state ──
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportW, setViewportW] = useState(1200);
  const [viewportH, setViewportH] = useState(600);

  // Track viewport size via ResizeObserver
  useEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    const ro = new ResizeObserver((entries) => {
      const entry = entries[0];
      if (entry) {
        setViewportW(entry.contentRect.width);
        setViewportH(entry.contentRect.height);
      }
    });
    ro.observe(el);
    // Set initial size
    setViewportW(el.clientWidth);
    setViewportH(el.clientHeight);
    return () => ro.disconnect();
  }, []);

  // Scroll handler
  const handleScroll = useCallback(() => {
    const el = gridRef.current;
    if (!el) return;
    setScrollTop(el.scrollTop);
    setScrollLeft(el.scrollLeft);
  }, []);

  // ── Pre-compute column positions (cumulative X offsets) ──
  const colPositions = useMemo(() => {
    const positions = new Float64Array(colCount + 1);
    positions[0] = 0;
    for (let c = 0; c < colCount; c++) {
      positions[c + 1] = positions[c] + (colWidths.get(c) ?? DEFAULT_COL_W);
    }
    return positions;
  }, [colCount, colWidths]);

  // ── Pre-compute row positions (cumulative Y offsets) ──
  const rowPositions = useMemo(() => {
    const positions = new Float64Array(rowCount + 1);
    positions[0] = 0;
    for (let r = 0; r < rowCount; r++) {
      positions[r + 1] = positions[r] + (rowHeights.get(r) ?? ROW_H);
    }
    return positions;
  }, [rowCount, rowHeights]);

  const totalGridWidth = ROW_HEADER_W + colPositions[colCount];
  const totalGridHeight = ROW_H + rowPositions[rowCount]; // header row + data rows

  // ── Compute visible row/col range ──
  const { visRowStart, visRowEnd, visColStart, visColEnd } = useMemo(() => {
    // Rows: binary search on cumulative positions
    const adjScrollTop = Math.max(0, scrollTop - ROW_H); // subtract header row
    let rs = 0;
    for (let r = 0; r < rowCount; r++) {
      if (rowPositions[r + 1] > adjScrollTop) { rs = r; break; }
    }
    rs = Math.max(0, rs - OVERSCAN_ROWS);

    const bottomEdge = adjScrollTop + viewportH;
    let re = rowCount - 1;
    for (let r = rs; r < rowCount; r++) {
      if (rowPositions[r] > bottomEdge) { re = r; break; }
    }
    re = Math.min(rowCount - 1, re + OVERSCAN_ROWS);

    // Cols: binary search on cumulative positions
    const adjScrollLeft = Math.max(0, scrollLeft - ROW_HEADER_W);
    let cs = 0;
    for (let c = 0; c < colCount; c++) {
      if (colPositions[c + 1] > adjScrollLeft) { cs = c; break; }
    }
    cs = Math.max(0, cs - OVERSCAN_COLS);

    const rightEdge = adjScrollLeft + viewportW;
    let ce = colCount - 1;
    for (let c = cs; c < colCount; c++) {
      if (colPositions[c] > rightEdge) { ce = c; break; }
    }
    ce = Math.min(colCount - 1, ce + OVERSCAN_COLS);

    return { visRowStart: rs, visRowEnd: re, visColStart: cs, visColEnd: ce };
  }, [scrollTop, scrollLeft, viewportW, viewportH, rowCount, colCount, colPositions, rowPositions]);

  // Pre-compute conditional formatting overrides
  const cfMap = useMemo(() => {
    if (conditionalFormatRules.length === 0) return null;
    const map = new Map<string, Partial<CellFormatting>>();
    const allRangeKeys = new Set<string>();
    for (const rule of conditionalFormatRules) {
      const { start, end } = rule.range;
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          allRangeKeys.add(`${r}:${c}`);
        }
      }
    }
    for (const key of allRangeKeys) {
      const [r, c] = key.split(':').map(Number);
      const cellData = cells.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
      const override = evaluateConditionalFormatting(r, c, cellData, conditionalFormatRules, cells);
      if (override) map.set(key, override);
    }
    return map;
  }, [cells, conditionalFormatRules]);

  // Pre-compute table formatting, hidden rows, and cell classes
  const safeTables = tables ?? [];
  const tableFormatMap = useMemo(() => computeTableFormattingMap(safeTables), [safeTables]);
  const hiddenRows = useMemo(() => computeHiddenRows(safeTables, cells), [safeTables, cells]);
  const tableCellClasses = useMemo(() => computeTableCellClasses(safeTables), [safeTables]);

  // Close context menu on outside click
  useEffect(() => {
    const close = () => setContextMenu(null);
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  // Col/Row resize: track mousemove/mouseup globally
  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (resizeRef.current) {
        const { col, startX, startWidth, otherCols } = resizeRef.current;
        const newWidth = Math.max(40, startWidth + e.clientX - startX);
        onColWidthChange(col, newWidth);
        // Apply same width to all other selected columns
        if (otherCols) {
          for (const oc of otherCols) onColWidthChange(oc, newWidth);
        }
      }
      if (rowResizeRef.current) {
        const { row, startY, startHeight, otherRows } = rowResizeRef.current;
        const newHeight = Math.max(20, startHeight + e.clientY - startY);
        onRowHeightChange(row, newHeight);
        if (otherRows) {
          for (const or_ of otherRows) onRowHeightChange(or_, newHeight);
        }
      }
    };
    const onUp = () => { resizeRef.current = null; rowResizeRef.current = null; };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    return () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
  }, [onColWidthChange, onRowHeightChange]);

  // Fill handle drag: start from the fill handle (small square at bottom-right of active cell)
  const handleFillHandleMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!selectedRange) return;
    fillDragBaseRef.current = selectedRange;
    setIsFillDragging(true);
    setFillDragEnd(null);
  }, [selectedRange]);

  // Fill handle drag: track which cell mouse is over
  useEffect(() => {
    if (!isFillDragging) return;
    const onMove = (e: MouseEvent) => {
      const wrapper = gridRef.current;
      if (!wrapper || !fillDragBaseRef.current) return;
      const rect = wrapper.getBoundingClientRect();
      const x = e.clientX - rect.left + wrapper.scrollLeft - ROW_HEADER_W;
      const y = e.clientY - rect.top + wrapper.scrollTop - ROW_H;
      let row = 0;
      for (let r = 0; r < rowCount; r++) {
        if (rowPositions[r + 1] > y) { row = r; break; }
        row = r;
      }
      row = Math.max(0, Math.min(rowCount - 1, row));
      // Find col from cumulative positions
      let col = 0;
      for (let c = 0; c < colCount; c++) {
        if (colPositions[c + 1] > x) { col = c; break; }
        col = c;
      }
      col = Math.max(0, Math.min(colCount - 1, col));
      setFillDragEnd({ row, col });
    };
    const onUp = () => {
      setIsFillDragging(false);
      const base = fillDragBaseRef.current;
      const end = fillDragEnd;
      if (base && end) {
        // Determine fill direction: expand down or right from base
        const baseEndRow = base.end.row;
        const baseEndCol = base.end.col;
        if (end.row > baseEndRow) {
          // Fill down
          const fillRange: CellRange = {
            start: base.start,
            end: { row: end.row, col: base.end.col },
          };
          onFillSeries(fillRange);
        } else if (end.col > baseEndCol) {
          // Fill right
          const fillRange: CellRange = {
            start: base.start,
            end: { row: base.end.row, col: end.col },
          };
          onFillSeries(fillRange);
        } else if (end.row < base.start.row) {
          // Fill up
          const fillRange: CellRange = {
            start: { row: end.row, col: base.start.col },
            end: base.end,
          };
          onFillSeries(fillRange);
        } else if (end.col < base.start.col) {
          // Fill left
          const fillRange: CellRange = {
            start: { row: base.start.row, col: end.col },
            end: base.end,
          };
          onFillSeries(fillRange);
        }
      }
      setFillDragEnd(null);
      fillDragBaseRef.current = null;
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    return () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isFillDragging, fillDragEnd, rowCount, colCount, colPositions, rowPositions, onFillSeries]);

  // Compute the fill drag highlight range
  const fillHighlightRange = useMemo((): CellRange | null => {
    if (!isFillDragging || !fillDragBaseRef.current || !fillDragEnd) return null;
    const base = fillDragBaseRef.current;
    const end = fillDragEnd;
    if (end.row > base.end.row) {
      return { start: { row: base.end.row + 1, col: base.start.col }, end: { row: end.row, col: base.end.col } };
    } else if (end.col > base.end.col) {
      return { start: { row: base.start.row, col: base.end.col + 1 }, end: { row: base.end.row, col: end.col } };
    } else if (end.row < base.start.row) {
      return { start: { row: end.row, col: base.start.col }, end: { row: base.start.row - 1, col: base.end.col } };
    } else if (end.col < base.start.col) {
      return { start: { row: base.start.row, col: end.col }, end: { row: base.end.row, col: base.start.col - 1 } };
    }
    return null;
  }, [isFillDragging, fillDragEnd]);

  const navigate = useCallback((row: number, col: number) => {
    const r = Math.max(0, Math.min(row, rowCount - 1));
    const c = Math.max(0, Math.min(col, colCount - 1));
    const addr = { row: r, col: c };
    setExtraSelections([]);
    onSelectionChange({ start: addr, end: addr }, addr);
  }, [rowCount, colCount, onSelectionChange]);

  // Select all helper
  const selectAll = useCallback(() => {
    const range: CellRange = { start: { row: 0, col: 0 }, end: { row: rowCount - 1, col: colCount - 1 } };
    setExtraSelections([]);
    onSelectionChange(range, { row: 0, col: 0 });
  }, [rowCount, colCount, onSelectionChange]);

  // Enter edit mode
  const enterEditMode = useCallback((replaceWith?: string) => {
    if (!activeCell) return;
    const currentVal = cells.get(`${activeCell.row}:${activeCell.col}`)?.value ?? '';
    const startVal = replaceWith !== undefined ? replaceWith : currentVal;
    setEditOriginalValue(currentVal);
    setEditValue(startVal);
    setIsEditing(true);
  }, [activeCell, cells]);

  // Commit (or discard) the local buffer to global state
  const exitEditMode = useCallback((commit: boolean) => {
    if (commit && activeCell && editValue !== editOriginalValue) {
      onCellChange(activeCell.row, activeCell.col, editValue);
    }
    setIsEditing(false);
    setTimeout(() => gridRef.current?.focus(), 0);
  }, [activeCell, editValue, editOriginalValue, onCellChange]);

  const handleGridKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (!activeCell || isEditing) return;
    const { row, col } = activeCell;

    switch (e.key) {
      case 'ArrowRight': navigate(row, col + 1); e.preventDefault(); break;
      case 'ArrowLeft':  navigate(row, col - 1); e.preventDefault(); break;
      case 'ArrowDown':  navigate(row + 1, col); e.preventDefault(); break;
      case 'ArrowUp':    navigate(row - 1, col); e.preventDefault(); break;
      case 'Tab':        navigate(row, e.shiftKey ? col - 1 : col + 1); e.preventDefault(); break;
      case 'Enter':      enterEditMode(); e.preventDefault(); break;
      case 'F2':         enterEditMode(); e.preventDefault(); break;
      case 'Delete':
      case 'Backspace':  onCellChange(row, col, ''); e.preventDefault(); break;
      default:
        if (e.ctrlKey || e.metaKey) {
          if (e.key === 'z') { e.shiftKey ? onRedo() : onUndo(); e.preventDefault(); }
          else if (e.key === 'y') { onRedo(); e.preventDefault(); }
          else if (e.key === 'a') {
            // Select all
            selectAll();
            e.preventDefault();
          } else if (e.key === 'b') {
            // Toggle bold
            if (onFormat && selectedRange) {
              const curBold = cells.get(`${row}:${col}`)?.formatting?.bold ?? false;
              onFormat(selectedRange, { bold: !curBold });
              e.preventDefault();
            }
          } else if (e.key === 'i') {
            // Toggle italic
            if (onFormat && selectedRange) {
              const curItalic = cells.get(`${row}:${col}`)?.formatting?.italic ?? false;
              onFormat(selectedRange, { italic: !curItalic });
              e.preventDefault();
            }
          } else if (e.key === 'u') {
            // Toggle underline
            if (onFormat && selectedRange) {
              const curUnderline = cells.get(`${row}:${col}`)?.formatting?.underline ?? false;
              onFormat(selectedRange, { underline: !curUnderline });
              e.preventDefault();
            }
          } else if (e.key === 'd') {
            if (selectedRange) { onFillDown(selectedRange); e.preventDefault(); }
          } else if (e.key === 'r') {
            if (selectedRange) { onFillRight(selectedRange); e.preventDefault(); }
          } else if (e.key === 'c') {
            if (selectedRange) {
              const data: CellData[][] = [];
              for (let r = selectedRange.start.row; r <= selectedRange.end.row; r++) {
                const row: CellData[] = [];
                for (let c = selectedRange.start.col; c <= selectedRange.end.col; c++) {
                  row.push(cells.get(`${r}:${c}`) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } });
                }
                data.push(row);
              }
              copiedRef.current = { data, range: selectedRange };
              const text = data.map(r => r.map(c => c.value).join('\t')).join('\n');
              navigator.clipboard.writeText(text).catch(() => {});
              e.preventDefault();
            }
          } else if (e.key === 'v') {
            if (activeCell) {
              if (copiedRef.current) {
                const { data } = copiedRef.current;
                data.forEach((row, ri) => {
                  row.forEach((cell, ci) => {
                    onCellChange(activeCell.row + ri, activeCell.col + ci, cell.value);
                  });
                });
              } else {
                navigator.clipboard.readText().then(text => {
                  text.split('\n').forEach((row, ri) => {
                    row.split('\t').forEach((val, ci) => {
                      onCellChange(activeCell.row + ri, activeCell.col + ci, val);
                    });
                  });
                }).catch(() => {});
              }
              e.preventDefault();
            }
          }
        } else if (e.key.length === 1 && !e.altKey) {
          enterEditMode(e.key);
          e.preventDefault();
        }
    }
  }, [activeCell, isEditing, navigate, enterEditMode, onCellChange, onUndo, onRedo,
      onFillDown, onFillRight, selectedRange, cells, selectAll, onFormat]);

  const handleCellKeyDown = useCallback((e: React.KeyboardEvent, row: number, col: number) => {
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        exitEditMode(true);
        navigate(row + 1, col);
        break;
      case 'Tab':
        e.preventDefault();
        exitEditMode(true);
        navigate(row, e.shiftKey ? col - 1 : col + 1);
        break;
      case 'Escape':
        e.preventDefault();
        exitEditMode(false);
        break;
      case 'ArrowUp':
        e.preventDefault();
        exitEditMode(true);
        navigate(row - 1, col);
        break;
      case 'ArrowDown':
        e.preventDefault();
        exitEditMode(true);
        navigate(row + 1, col);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        exitEditMode(true);
        navigate(row, col - 1);
        break;
      case 'ArrowRight':
        e.preventDefault();
        exitEditMode(true);
        navigate(row, col + 1);
        break;
    }
  }, [exitEditMode, navigate]);

  // Check if cursor is at a position where a cell reference can be inserted in a formula
  const isFormulaRefPosition = useCallback((val: string): boolean => {
    if (!val.startsWith('=')) return false;
    if (val.length === 1) return true;
    const lastChar = val[val.length - 1];
    return /[=+\-*/,(>&<]/.test(lastChar);
  }, []);

  const insertFormulaRef = useCallback((ref: string) => {
    setEditValue(prev => {
      const refPattern = /[A-Z]+\d+(:[A-Z]+\d+)?$/i;
      if (formulaRefStart && refPattern.test(prev)) {
        return prev.replace(refPattern, ref);
      }
      return prev + ref;
    });
  }, [formulaRefStart]);

  const handleMouseDown = useCallback((row: number, col: number, e: React.MouseEvent) => {
    e.preventDefault();

    // Right-click within an already-selected range: preserve selection for context menu
    if (e.button === 2 && selectedRange && isInRange(row, col, selectedRange)) {
      gridRef.current?.focus();
      return;
    }
    if (e.button === 2 && extraSelections.some(s => isInRange(row, col, s))) {
      gridRef.current?.focus();
      return;
    }

    if (isEditing && activeCell && !(row === activeCell.row && col === activeCell.col)) {
      const currentVal = editValueRef.current;
      if (isFormulaRefPosition(currentVal)) {
        const ref = cellKey(row, col);
        setEditValue(currentVal + ref);
        setFormulaRefStart({ row, col });
        setIsSelecting(true);
        setSelectionStart({ row, col });
        return;
      }
      if (currentVal.startsWith('=') && formulaRefStart) {
        const ref = cellKey(row, col);
        insertFormulaRef(ref);
        setFormulaRefStart({ row, col });
        return;
      }
    }

    if (isEditing) exitEditMode(true);
    setFormulaRefStart(null);
    setExtraSelections([]);
    setHeaderSelectMode('none');
    const addr = { row, col };

    if (e.shiftKey && activeCell) {
      onSelectionChange(
        {
          start: { row: Math.min(activeCell.row, row), col: Math.min(activeCell.col, col) },
          end:   { row: Math.max(activeCell.row, row), col: Math.max(activeCell.col, col) },
        },
        activeCell,
      );
    } else {
      setIsSelecting(true);
      setSelectionStart(addr);
      onSelectionChange({ start: addr, end: addr }, addr);
    }
    gridRef.current?.focus();
  }, [isEditing, exitEditMode, onSelectionChange, activeCell, isFormulaRefPosition, insertFormulaRef, formulaRefStart, selectedRange, extraSelections]);

  const handleDoubleClick = useCallback((row: number, col: number) => {
    const val = cells.get(`${row}:${col}`)?.value ?? '';
    setEditOriginalValue(val);
    setEditValue(val);
    setIsEditing(true);
  }, [cells]);

  const handleMouseEnter = useCallback((row: number, col: number) => {
    if (!isSelecting || !selectionStart) return;

    if (isEditing && formulaRefStart) {
      const startR = Math.min(formulaRefStart.row, row);
      const startC = Math.min(formulaRefStart.col, col);
      const endR = Math.max(formulaRefStart.row, row);
      const endC = Math.max(formulaRefStart.col, col);
      const ref = (startR === endR) && (startC === endC)
        ? cellKey(startR, startC)
        : `${cellKey(startR, startC)}:${cellKey(endR, endC)}`;
      insertFormulaRef(ref);
      return;
    }

    onSelectionChange(
      {
        start: { row: Math.min(selectionStart.row, row), col: Math.min(selectionStart.col, col) },
        end:   { row: Math.max(selectionStart.row, row), col: Math.max(selectionStart.col, col) },
      },
      selectionStart
    );
  }, [isSelecting, selectionStart, onSelectionChange, isEditing, formulaRefStart, insertFormulaRef]);

  const handleMouseUp = useCallback(() => {
    setIsSelecting(false);
    setHeaderSelectMode('none');
    if (formulaRefStart) setFormulaRefStart(null);
  }, [formulaRefStart]);

  // ── Build the grid column template (needed for CSS grid) ──
  // Only include visible columns, but use explicit grid-column placement
  const colTemplate = `${ROW_HEADER_W}px ${Array.from({ length: colCount }, (_, c) => `${colWidths.get(c) ?? DEFAULT_COL_W}px`).join(' ')}`;
  const rowTemplate = `${ROW_H}px ${Array.from({ length: rowCount }, (_, r) => `${rowHeights.get(r) ?? ROW_H}px`).join(' ')}`;

  // ── Determine which frozen rows to always render ──
  const frozenRowSet = useMemo(() => {
    const s = new Set<number>();
    for (let r = 0; r < freezeRows; r++) {
      if (!hiddenRows.has(r)) s.add(r);
    }
    return s;
  }, [freezeRows, hiddenRows]);

  // ── Build set of rows to render: frozen + visible range (deduplicated) ──
  const rowsToRender = useMemo(() => {
    const rows: number[] = [];
    // Frozen rows first
    frozenRowSet.forEach(r => rows.push(r));
    // Visible rows (skip if already in frozen set)
    for (let r = visRowStart; r <= visRowEnd; r++) {
      if (!frozenRowSet.has(r) && !hiddenRows.has(r)) {
        rows.push(r);
      }
    }
    return rows;
  }, [frozenRowSet, visRowStart, visRowEnd, hiddenRows]);

  // Auto-fit column width by measuring content
  const autoFitColumn = useCallback((col: number) => {
    const PADDING = 16; // 6px padding each side + extra
    const MIN_WIDTH = 40;
    const HEADER_WIDTH = 30; // approximate width for header letter
    let maxWidth = HEADER_WIDTH;

    // Create a hidden canvas for text measurement
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) { onColWidthChange(col, DEFAULT_COL_W); return; }

    for (let r = 0; r < rowCount; r++) {
      const cell = cells.get(`${r}:${col}`);
      if (!cell || !cell.displayValue) continue;
      const fmt = cell.formatting;
      const fontStyle = fmt.italic ? 'italic' : 'normal';
      const fontWeight = fmt.bold ? 'bold' : 'normal';
      const fontFamily = (fmt.fontFamily && fmt.fontFamily !== 'Default') ? fmt.fontFamily : 'Inter, sans-serif';
      ctx.font = `${fontStyle} ${fontWeight} 13px ${fontFamily}`;
      const width = ctx.measureText(cell.displayValue).width + PADDING;
      if (width > maxWidth) maxWidth = width;
    }

    onColWidthChange(col, Math.max(MIN_WIDTH, Math.ceil(maxWidth)));
  }, [cells, rowCount, onColWidthChange]);

  // Get all selected columns (from selectedRange + extraSelections)
  const getSelectedCols = useCallback((): number[] => {
    const cols = new Set<number>();
    if (selectedRange) {
      const minC = Math.min(selectedRange.start.col, selectedRange.end.col);
      const maxC = Math.max(selectedRange.start.col, selectedRange.end.col);
      for (let c = minC; c <= maxC; c++) cols.add(c);
    }
    for (const s of extraSelections) {
      const minC = Math.min(s.start.col, s.end.col);
      const maxC = Math.max(s.start.col, s.end.col);
      for (let c = minC; c <= maxC; c++) cols.add(c);
    }
    return Array.from(cols);
  }, [selectedRange, extraSelections]);

  // Get all selected rows
  const getSelectedRows = useCallback((): number[] => {
    const rows = new Set<number>();
    if (selectedRange) {
      const minR = Math.min(selectedRange.start.row, selectedRange.end.row);
      const maxR = Math.max(selectedRange.start.row, selectedRange.end.row);
      for (let r = minR; r <= maxR; r++) rows.add(r);
    }
    for (const s of extraSelections) {
      const minR = Math.min(s.start.row, s.end.row);
      const maxR = Math.max(s.start.row, s.end.row);
      for (let r = minR; r <= maxR; r++) rows.add(r);
    }
    return Array.from(rows);
  }, [selectedRange, extraSelections]);

  // Visible columns range
  const colsToRender = useMemo(() => {
    const cols: number[] = [];
    for (let c = visColStart; c <= visColEnd; c++) {
      cols.push(c);
    }
    return cols;
  }, [visColStart, visColEnd]);

  return (
    <div
      className="grid-wrapper"
      ref={gridRef}
      tabIndex={0}
      onKeyDown={handleGridKeyDown}
      onMouseUp={handleMouseUp}
      onMouseLeave={handleMouseUp}
      onScroll={handleScroll}
    >
      <div
        className="grid-container"
        style={{
          gridTemplateColumns: colTemplate,
          gridTemplateRows: rowTemplate,
          width: totalGridWidth,
          height: totalGridHeight,
        }}
      >
        {/* Corner — click to select all */}
        <div
          className="grid-corner"
          style={{ gridRow: 1, gridColumn: 1 }}
          onClick={() => {
            if (isEditing) exitEditMode(true);
            selectAll();
            gridRef.current?.focus();
          }}
        />

        {/* Column headers — only visible cols */}
        {colsToRender.map(c => {
          const isColSelected = (selectedRange && c >= Math.min(selectedRange.start.col, selectedRange.end.col) && c <= Math.max(selectedRange.start.col, selectedRange.end.col))
            || extraSelections.some(s => c >= Math.min(s.start.col, s.end.col) && c <= Math.max(s.start.col, s.end.col));
          return (
          <div
            key={`ch-${c}`}
            className={`grid-col-header${isColSelected ? ' grid-col-header--selected' : ''}`}
            style={{ gridRow: 1, gridColumn: c + 2 }}
            onMouseDown={(e) => {
              e.preventDefault();
              e.stopPropagation();
              // Right-click: preserve selection if this col is already selected
              if (e.button === 2 && isColSelected) {
                gridRef.current?.focus();
                return;
              }
              if (isEditing) exitEditMode(true);
              const colRange: CellRange = { start: { row: 0, col: c }, end: { row: rowCount - 1, col: c } };

              if ((e.ctrlKey || e.metaKey) && selectedRange) {
                setExtraSelections(prev => [...prev, colRange]);
              } else if (e.shiftKey && activeCell) {
                const minCol = Math.min(activeCell.col, c);
                const maxCol = Math.max(activeCell.col, c);
                onSelectionChange(
                  { start: { row: 0, col: minCol }, end: { row: rowCount - 1, col: maxCol } },
                  activeCell,
                );
                setExtraSelections([]);
              } else {
                onSelectionChange(colRange, colRange.start);
                setExtraSelections([]);
                setIsSelecting(true);
                setSelectionStart({ row: 0, col: c });
                setHeaderSelectMode('col');
              }
              gridRef.current?.focus();
            }}
            onMouseEnter={() => {
              if (isSelecting && headerSelectMode === 'col' && selectionStart) {
                const minCol = Math.min(selectionStart.col, c);
                const maxCol = Math.max(selectionStart.col, c);
                onSelectionChange(
                  { start: { row: 0, col: minCol }, end: { row: rowCount - 1, col: maxCol } },
                  selectionStart,
                );
              }
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              const selCols = getSelectedCols();
              const indices = selCols.includes(c) ? selCols : [c];
              setContextMenu({ x: e.clientX, y: e.clientY, type: 'col', index: c, indices });
            }}
          >
            {colIndexToLetter(c)}
            <div
              className="col-resize-handle"
              onMouseDown={(e) => {
                e.preventDefault();
                e.stopPropagation();
                // Find other selected columns to resize together
                const selCols = getSelectedCols();
                const otherCols = selCols.filter(sc => sc !== c);
                resizeRef.current = {
                  col: c,
                  startX: e.clientX,
                  startWidth: colWidths.get(c) ?? DEFAULT_COL_W,
                  otherCols: otherCols.length > 0 ? otherCols : undefined,
                };
              }}
              onDoubleClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                // Auto-fit this column and all selected columns
                const selCols = getSelectedCols();
                if (selCols.includes(c)) {
                  for (const sc of selCols) autoFitColumn(sc);
                } else {
                  autoFitColumn(c);
                }
              }}
            />
          </div>
          );
        })}

        {/* Data rows — only visible + frozen rows */}
        {rowsToRender.map(r => {
          const isFrozen = r < freezeRows;
          const frozenTop = ROW_H + rowPositions[r];
          const gridRow = r + 2; // +1 for header, +1 for 1-based
          const isRowSelected = (selectedRange && r >= Math.min(selectedRange.start.row, selectedRange.end.row) && r <= Math.max(selectedRange.start.row, selectedRange.end.row))
            || extraSelections.some(s => r >= Math.min(s.start.row, s.end.row) && r <= Math.max(s.start.row, s.end.row));

          return (
            <React.Fragment key={r}>
              <div
                className={`grid-row-header${isRowSelected ? ' grid-row-header--selected' : ''}`}
                style={{
                  gridRow, gridColumn: 1,
                  ...(isFrozen ? { position: 'sticky', top: frozenTop, zIndex: 3 } : {}),
                }}
                onMouseDown={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  // Right-click: preserve selection if this row is already selected
                  if (e.button === 2 && isRowSelected) {
                    gridRef.current?.focus();
                    return;
                  }
                  if (isEditing) exitEditMode(true);
                  const rowRange: CellRange = { start: { row: r, col: 0 }, end: { row: r, col: colCount - 1 } };

                  if ((e.ctrlKey || e.metaKey) && selectedRange) {
                    setExtraSelections(prev => [...prev, rowRange]);
                  } else if (e.shiftKey && activeCell) {
                    const minRow = Math.min(activeCell.row, r);
                    const maxRow = Math.max(activeCell.row, r);
                    onSelectionChange(
                      { start: { row: minRow, col: 0 }, end: { row: maxRow, col: colCount - 1 } },
                      activeCell,
                    );
                    setExtraSelections([]);
                  } else {
                    onSelectionChange(rowRange, rowRange.start);
                    setExtraSelections([]);
                    setIsSelecting(true);
                    setSelectionStart({ row: r, col: 0 });
                    setHeaderSelectMode('row');
                  }
                  gridRef.current?.focus();
                }}
                onMouseEnter={() => {
                  if (isSelecting && headerSelectMode === 'row' && selectionStart) {
                    const minRow = Math.min(selectionStart.row, r);
                    const maxRow = Math.max(selectionStart.row, r);
                    onSelectionChange(
                      { start: { row: minRow, col: 0 }, end: { row: maxRow, col: colCount - 1 } },
                      selectionStart,
                    );
                  }
                }}
                onContextMenu={(e) => {
                  e.preventDefault();
                  const selRows = getSelectedRows();
                  const indices = selRows.includes(r) ? selRows : [r];
                  setContextMenu({ x: e.clientX, y: e.clientY, type: 'row', index: r, indices });
                }}
              >
                {r + 1}
                <div
                  className="row-resize-handle"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    const selRows = getSelectedRows();
                    const otherRows = selRows.filter(sr => sr !== r);
                    rowResizeRef.current = {
                      row: r,
                      startY: e.clientY,
                      startHeight: rowHeights.get(r) ?? ROW_H,
                      otherRows: otherRows.length > 0 ? otherRows : undefined,
                    };
                  }}
                />
              </div>
              {colsToRender.map(c => {
                const key = `${r}:${c}`;
                const cellIsActive = activeCell?.row === r && activeCell?.col === c;
                const editing = cellIsActive && isEditing;

                const tableFmt = tableFormatMap.get(key);
                const cfFmt = cfMap?.get(key);
                const mergedFormatting = tableFmt || cfFmt
                  ? { ...tableFmt, ...cfFmt }
                  : null;

                const tableClass = tableCellClasses.get(key);
                let headerControls: React.ReactNode = undefined;
                const cellTable = tableClass?.includes('table-header-cell')
                  ? findTableForCell(r, c, safeTables)
                  : undefined;
                if (cellTable) {
                  const curDir = cellTable.sortColumn === c ? cellTable.sortDirection : null;
                  const hasFilter = cellTable.filters.get(c)?.checkedValues !== null;
                  const hasSort = curDir !== null;
                  headerControls = (
                    <div className="table-header-controls" onMouseDown={e => e.stopPropagation()}>
                      <button
                        className={`table-header-btn${hasFilter ? ' table-header-btn--filtered' : ''}${hasSort ? ' table-header-btn--sorted' : ''}`}
                        title="Sort & Filter"
                        onClick={(e) => {
                          const rect = (e.target as HTMLElement).getBoundingClientRect();
                          setFilterDropdown({ table: cellTable, col: c, position: { x: rect.left, y: rect.bottom + 2 } });
                        }}
                      >
                        {curDir === 'asc' ? '\u25B2' : curDir === 'desc' ? '\u25BC' : '\u25BE'}
                      </button>
                    </div>
                  );
                }

                return (
                  <Cell
                    key={key}
                    row={r}
                    col={c}
                    data={cells.get(key) ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } }}
                    isActive={cellIsActive}
                    isEditing={editing}
                    isSelected={isInRange(r, c, selectedRange) || extraSelections.some(s => isInRange(r, c, s))}
                    editValue={editing ? editValue : undefined}
                    frozenTop={isFrozen ? frozenTop : undefined}
                    conditionalFormatting={mergedFormatting}
                    extraClassName={tableClass}
                    tableHeaderControls={headerControls}
                    gridRow={gridRow}
                    gridColumn={c + 2}
                    onChange={setEditValue}
                    onFocus={() => {
                      const addr = { row: r, col: c };
                      onSelectionChange({ start: addr, end: addr }, addr);
                    }}
                    onKeyDown={(e) => handleCellKeyDown(e, r, c)}
                    onMouseDown={(e) => handleMouseDown(r, c, e)}
                    onDoubleClick={() => handleDoubleClick(r, c)}
                    onMouseEnter={() => handleMouseEnter(r, c)}
                    onContextMenu={(e: React.MouseEvent) => {
                      const t = findTableForCell(r, c, safeTables);
                      if (t) {
                        e.preventDefault();
                        setContextMenu({ x: e.clientX, y: e.clientY, type: 'cell', index: 0, tableId: t.id });
                      }
                    }}
                    onFillHandleMouseDown={cellIsActive ? handleFillHandleMouseDown : undefined}
                    isFillHighlighted={fillHighlightRange ? isInRange(r, c, fillHighlightRange) : false}
                  />
                );
              })}
            </React.Fragment>
          );
        })}
      </div>

      {/* Context menu */}
      {contextMenu && (
        <div
          className="context-menu"
          style={{ top: contextMenu.y, left: contextMenu.x }}
          onMouseDown={(e) => e.stopPropagation()}
        >
          {contextMenu.type === 'row' ? (() => {
            const indices = contextMenu.indices ?? [contextMenu.index];
            const count = indices.length;
            return (
              <>
                <button onClick={() => {
                  // Insert rows above the topmost selected row
                  const minRow = Math.min(...indices);
                  for (let i = 0; i < count; i++) onInsertRow(minRow);
                  setContextMenu(null);
                }}>
                  Insert {count} row{count > 1 ? 's' : ''} above
                </button>
                <button onClick={() => {
                  // Delete selected rows from bottom to top to preserve indices
                  const sorted = [...indices].sort((a, b) => b - a);
                  for (const idx of sorted) onDeleteRow(idx);
                  setContextMenu(null);
                }}>
                  Delete {count} row{count > 1 ? 's' : ''}
                </button>
                <button onClick={() => {
                  onSetFreezeRows(freezeRows === contextMenu.index + 1 ? 0 : contextMenu.index + 1);
                  setContextMenu(null);
                }}>
                  {freezeRows === contextMenu.index + 1 ? 'Unfreeze rows' : `Freeze rows above ${contextMenu.index + 2}`}
                </button>
              </>
            );
          })() : contextMenu.type === 'cell' && contextMenu.tableId ? (
            <button onClick={() => { onDeleteTable(contextMenu.tableId!); setContextMenu(null); }}>
              Delete Table
            </button>
          ) : (() => {
            const indices = contextMenu.indices ?? [contextMenu.index];
            const count = indices.length;
            return (
              <>
                <button onClick={() => {
                  // Insert cols left of the leftmost selected col
                  const minCol = Math.min(...indices);
                  for (let i = 0; i < count; i++) onInsertCol(minCol);
                  setContextMenu(null);
                }}>
                  Insert {count} col{count > 1 ? 's' : ''} left
                </button>
                <button onClick={() => {
                  // Delete selected cols from right to left to preserve indices
                  const sorted = [...indices].sort((a, b) => b - a);
                  for (const idx of sorted) onDeleteCol(idx);
                  setContextMenu(null);
                }}>
                  Delete {count} col{count > 1 ? 's' : ''}
                </button>
              </>
            );
          })()}
        </div>
      )}

      {/* Table filter dropdown */}
      {filterDropdown && ReactDOM.createPortal(
        <TableFilterDropdown
          table={filterDropdown.table}
          col={filterDropdown.col}
          cells={cells}
          position={filterDropdown.position}
          onFilter={onFilterTable}
          onSort={onSortTable}
          onClose={() => setFilterDropdown(null)}
        />,
        document.body,
      )}
    </div>
  );
}
