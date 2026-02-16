import React, { useRef, useEffect, useCallback, useMemo, CSSProperties } from 'react';
import ReactDOM from 'react-dom';
import { CellData, CellFormatting, DEFAULT_CELL, DEFAULT_FORMATTING } from '../types/spreadsheet';
import { isNumericValue } from '../utils/formulaEngine';
import { formatCellValue, applyCustomFormatString } from '../utils/numberFormat';
import FormulaAutocomplete, { getAutocompleteMatch } from './FormulaAutocomplete';

interface CellProps {
  row: number;
  col: number;
  data: CellData;
  isActive: boolean;
  isEditing: boolean;
  isSelected: boolean;
  editValue?: string;
  frozenTop?: number;
  conditionalFormatting?: Partial<CellFormatting> | null;
  extraClassName?: string;
  tableHeaderControls?: React.ReactNode;
  gridRow?: number;
  gridColumn?: number;
  onChange: (value: string) => void;
  onFocus: () => void;
  onKeyDown: (e: React.KeyboardEvent) => void;
  onMouseDown: (e: React.MouseEvent) => void;
  onDoubleClick: () => void;
  onMouseEnter: (e: React.MouseEvent) => void;
  onContextMenu?: (e: React.MouseEvent) => void;
  onFillHandleMouseDown?: (e: React.MouseEvent) => void;
  isFillHighlighted?: boolean;
}

const Cell = React.memo(function Cell({
  data, isActive, isEditing, isSelected, editValue, frozenTop, conditionalFormatting,
  extraClassName, tableHeaderControls, gridRow, gridColumn,
  onChange, onFocus, onKeyDown, onMouseDown, onDoubleClick, onMouseEnter, onContextMenu,
  onFillHandleMouseDown, isFillHighlighted,
}: CellProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const cellDivRef = useRef<HTMLDivElement>(null);
  const cell = data ?? { ...DEFAULT_CELL, formatting: { ...DEFAULT_FORMATTING } };
  const formatting = conditionalFormatting
    ? { ...cell.formatting, ...conditionalFormatting }
    : cell.formatting;

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
      const len = inputRef.current.value.length;
      inputRef.current.setSelectionRange(len, len);
    }
  }, [isEditing]);

  const currentEditValue = isEditing ? (editValue ?? cell.value) : '';
  const autocompleteMatch = isEditing ? getAutocompleteMatch(currentEditValue) : null;

  const handleAutocompleteSelect = useCallback((funcName: string) => {
    if (!autocompleteMatch) return;
    // Replace the partial token with the full function name + opening paren
    const newValue = currentEditValue.slice(0, currentEditValue.length - autocompleteMatch.token.length) + funcName + '(';
    onChange(newValue);
    // Re-focus the input
    setTimeout(() => inputRef.current?.focus(), 0);
  }, [autocompleteMatch, currentEditValue, onChange]);

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    // If autocomplete is showing, Tab accepts the first match
    if (autocompleteMatch && autocompleteMatch.matches.length > 0 && e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      handleAutocompleteSelect(autocompleteMatch.matches[0].name);
      return;
    }
    onKeyDown(e);
  }, [autocompleteMatch, handleAutocompleteSelect, onKeyDown]);

  // Calculate autocomplete position
  let acPosition = { x: 0, y: 0 };
  if (autocompleteMatch && cellDivRef.current) {
    const rect = cellDivRef.current.getBoundingClientRect();
    acPosition = { x: rect.left, y: rect.bottom + 2 };
  }

  const tdStyle: CSSProperties = {
    backgroundColor: isSelected
      ? (formatting.backgroundColor ? formatting.backgroundColor + 'cc' : '#e0e7ff')
      : (formatting.backgroundColor ?? 'transparent'),
    outline: isActive ? '2px solid #4f46e5' : 'none',
    outlineOffset: '-1px',
    ...(frozenTop !== undefined ? { position: 'sticky', top: frozenTop, zIndex: 1 } : {}),
  };

  // Apply number formatting to display value
  const numberFormat = formatting.numberFormat ?? 'general';
  const decimalPlaces = formatting.decimalPlaces ?? 2;
  const currencyCode = formatting.currencyCode ?? 'USD';
  const dateFormat = formatting.dateFormat ?? 'mm/dd/yyyy';
  const customFormat = formatting.customFormat ?? '';
  const formattedDisplayVal = useMemo(() => {
    if (numberFormat === 'custom' && customFormat) {
      return applyCustomFormatString(cell.displayValue, customFormat);
    }
    if (numberFormat !== 'general' && numberFormat !== 'text') {
      return formatCellValue(cell.displayValue, numberFormat, decimalPlaces, currencyCode, dateFormat);
    }
    return cell.displayValue;
  }, [cell.displayValue, numberFormat, decimalPlaces, currencyCode, dateFormat, customFormat]);

  // Auto right-align numbers like Excel (when user hasn't explicitly set alignment)
  const autoAlign = formatting.hAlign === 'left' && !isEditing && isNumericValue(cell.displayValue)
    ? 'right' : (formatting.hAlign ?? 'left');

  const indent = formatting.indent ?? 0;
  const rotation = formatting.textRotation ?? 0;
  const wrapText = formatting.wrapText ?? false;

  const fontFamily = formatting.fontFamily ?? 'Default';
  const fontFamilyCSS = fontFamily === 'Default' ? undefined : `'${fontFamily}', sans-serif`;

  const inputStyle: CSSProperties = {
    fontWeight: formatting.bold ? 'bold' : 'normal',
    fontStyle: formatting.italic ? 'italic' : 'normal',
    textDecoration: formatting.underline ? 'underline' : 'none',
    color: formatting.textColor ?? 'inherit',
    textAlign: autoAlign,
    cursor: isEditing ? 'text' : 'default',
    pointerEvents: isEditing ? 'auto' : 'none',
    userSelect: isEditing ? 'text' : 'none',
    caretColor: isEditing ? 'auto' : 'transparent',
    ...(fontFamilyCSS ? { fontFamily: fontFamilyCSS } : {}),
    ...(indent > 0 ? { paddingLeft: indent * 12 } : {}),
    ...(rotation !== 0 ? { transform: `rotate(${-rotation}deg)`, transformOrigin: 'center center' } : {}),
    ...(wrapText ? { whiteSpace: 'pre-wrap', wordBreak: 'break-word' } : {}),
  };

  const vAlignMap = { top: 'flex-start', middle: 'center', bottom: 'flex-end' } as const;
  const tdAlignStyle: CSSProperties = {
    ...tdStyle,
    display: 'flex',
    alignItems: vAlignMap[formatting.vAlign ?? 'middle'],
    ...(wrapText ? { overflow: 'visible' } : {}),
    ...(gridRow !== undefined ? { gridRow } : {}),
    ...(gridColumn !== undefined ? { gridColumn } : {}),
  };

  return (
    <div
      ref={cellDivRef}
      className={`cell${extraClassName ? ' ' + extraClassName : ''}${isFillHighlighted ? ' cell--fill-highlight' : ''}`}
      style={tdAlignStyle}
      onMouseDown={onMouseDown}
      onMouseEnter={onMouseEnter}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
    >
      <input
        ref={inputRef}
        readOnly={!isEditing}
        tabIndex={-1}
        value={isEditing ? (editValue ?? cell.value) : formattedDisplayVal}
        onChange={(e) => { if (isEditing) onChange(e.target.value); }}
        onFocus={onFocus}
        onKeyDown={isEditing ? handleKeyDown : undefined}
        style={inputStyle}
      />
      {isActive && !isEditing && <div className="fill-handle" onMouseDown={onFillHandleMouseDown} />}
      {tableHeaderControls}
      {autocompleteMatch && ReactDOM.createPortal(
        <FormulaAutocomplete
          editValue={currentEditValue}
          position={acPosition}
          onSelect={handleAutocompleteSelect}
        />,
        document.body,
      )}
    </div>
  );
});

export default Cell;
