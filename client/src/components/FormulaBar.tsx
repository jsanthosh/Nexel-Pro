import React from 'react';
import { CellAddress } from '../types/spreadsheet';
import { colIndexToLetter } from '../utils/cellRange';

interface FormulaBarProps {
  activeCell: CellAddress | null;
  value: string;
  onChange: (row: number, col: number, value: string) => void;
}

export default function FormulaBar({ activeCell, value, onChange }: FormulaBarProps) {
  const cellLabel = activeCell
    ? `${colIndexToLetter(activeCell.col)}${activeCell.row + 1}`
    : '';

  return (
    <div className="formula-bar">
      <div className="cell-name-box">{cellLabel}</div>
      <div className="formula-bar-divider" />
      <input
        className="formula-input"
        value={value}
        onChange={(e) => {
          if (activeCell) onChange(activeCell.row, activeCell.col, e.target.value);
        }}
        placeholder="Enter value or formula (e.g. =SUM(A1:A10))"
      />
    </div>
  );
}
