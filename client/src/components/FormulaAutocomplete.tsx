import React, { useState, useEffect, useRef } from 'react';

const FUNCTIONS = [
  { name: 'SUM', desc: 'Adds all numbers in a range', syntax: 'SUM(range)' },
  { name: 'AVERAGE', desc: 'Returns the average of numbers', syntax: 'AVERAGE(range)' },
  { name: 'AVG', desc: 'Alias for AVERAGE', syntax: 'AVG(range)' },
  { name: 'MAX', desc: 'Returns the largest value', syntax: 'MAX(range)' },
  { name: 'MIN', desc: 'Returns the smallest value', syntax: 'MIN(range)' },
  { name: 'COUNT', desc: 'Counts cells with numbers', syntax: 'COUNT(range)' },
  { name: 'COUNTIF', desc: 'Counts cells matching criteria', syntax: 'COUNTIF(range, criteria)' },
  { name: 'IF', desc: 'Returns value based on condition', syntax: 'IF(condition, true_val, false_val)' },
  { name: 'IFERROR', desc: 'Returns alt value on error', syntax: 'IFERROR(value, error_val)' },
  { name: 'VLOOKUP', desc: 'Vertical lookup in a range', syntax: 'VLOOKUP(lookup, range, col_idx)' },
  { name: 'LEN', desc: 'Returns text length', syntax: 'LEN(text)' },
  { name: 'UPPER', desc: 'Converts to uppercase', syntax: 'UPPER(text)' },
  { name: 'LOWER', desc: 'Converts to lowercase', syntax: 'LOWER(text)' },
  { name: 'TRIM', desc: 'Removes extra spaces', syntax: 'TRIM(text)' },
  { name: 'CONCATENATE', desc: 'Joins text strings', syntax: 'CONCATENATE(text1, text2, ...)' },
  { name: 'ABS', desc: 'Returns absolute value', syntax: 'ABS(number)' },
  { name: 'ROUND', desc: 'Rounds to decimal places', syntax: 'ROUND(number, places)' },
  { name: 'FLOOR', desc: 'Rounds down to integer', syntax: 'FLOOR(number)' },
  { name: 'CEILING', desc: 'Rounds up to integer', syntax: 'CEILING(number)' },
  { name: 'CEIL', desc: 'Alias for CEILING', syntax: 'CEIL(number)' },
  { name: 'TODAY', desc: 'Returns today\'s date', syntax: 'TODAY()' },
  { name: 'NOW', desc: 'Returns current date/time', syntax: 'NOW()' },
];

interface FormulaAutocompleteProps {
  editValue: string;
  position: { x: number; y: number };
  onSelect: (funcName: string) => void;
}

/**
 * Extract the current function token being typed.
 * Only triggers after = and when typing letters (not after a cell ref or number).
 */
function getCurrentToken(value: string): string | null {
  if (!value.startsWith('=')) return null;
  const afterEq = value.slice(1);
  // Find the last operator/delimiter
  const match = afterEq.match(/[A-Z_]+$/i);
  if (!match) return null;
  // Only show if the token is preceded by = or an operator
  const before = afterEq.slice(0, afterEq.length - match[0].length);
  const lastChar = before[before.length - 1];
  if (before.length === 0 || /[=+\-*/,(>&<]/.test(lastChar)) {
    return match[0].toUpperCase();
  }
  return null;
}

export default function FormulaAutocomplete({ editValue, position, onSelect }: FormulaAutocompleteProps) {
  const [selectedIndex, setSelectedIndex] = useState(0);
  const ref = useRef<HTMLDivElement>(null);
  const token = getCurrentToken(editValue);

  const matches = token
    ? FUNCTIONS.filter(f => f.name.startsWith(token) && f.name !== token)
    : [];

  // Reset selection when matches change
  useEffect(() => { setSelectedIndex(0); }, [token]);

  if (matches.length === 0) return null;

  // Adjust position to stay in viewport
  const adjustedPos = { ...position };
  const dropdownH = Math.min(matches.length * 32 + 4, 200);
  if (adjustedPos.y + dropdownH > window.innerHeight) {
    adjustedPos.y = adjustedPos.y - dropdownH - 28;
  }

  return (
    <div
      ref={ref}
      className="formula-autocomplete"
      style={{ left: adjustedPos.x, top: adjustedPos.y }}
      onMouseDown={(e) => e.preventDefault()}
    >
      {matches.slice(0, 8).map((func, i) => (
        <div
          key={func.name}
          className={`formula-autocomplete-item${i === selectedIndex ? ' formula-autocomplete-item--active' : ''}`}
          onMouseEnter={() => setSelectedIndex(i)}
          onMouseDown={(e) => {
            e.preventDefault();
            e.stopPropagation();
            onSelect(func.name);
          }}
        >
          <span className="formula-autocomplete-name">{func.name}</span>
          <span className="formula-autocomplete-desc">{func.desc}</span>
        </div>
      ))}
    </div>
  );
}

// Exported for use by the Cell component's keydown handler
export function getAutocompleteMatch(editValue: string): { matches: typeof FUNCTIONS; token: string } | null {
  const token = getCurrentToken(editValue);
  if (!token) return null;
  const matches = FUNCTIONS.filter(f => f.name.startsWith(token) && f.name !== token);
  if (matches.length === 0) return null;
  return { matches, token };
}
