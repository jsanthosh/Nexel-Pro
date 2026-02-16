import React, { useState, useEffect, useRef, useCallback } from 'react';
import { SheetData, TableDefinition, SortDirection } from '../types/spreadsheet';
import { getUniqueColumnValues } from '../utils/tableUtils';

interface TableFilterDropdownProps {
  table: TableDefinition;
  col: number;
  cells: SheetData;
  position: { x: number; y: number };
  onFilter: (tableId: string, col: number, checkedValues: Set<string> | null) => void;
  onSort: (tableId: string, col: number, direction: SortDirection) => void;
  onClose: () => void;
}

export default function TableFilterDropdown({
  table, col, cells, position, onFilter, onSort, onClose,
}: TableFilterDropdownProps) {
  const ref = useRef<HTMLDivElement>(null);
  const uniqueValues = getUniqueColumnValues(table, col, cells);
  const currentFilter = table.filters.get(col);

  const [checked, setChecked] = useState<Set<string>>(() => {
    if (currentFilter?.checkedValues) return new Set(currentFilter.checkedValues);
    return new Set(uniqueValues);
  });

  // Use a ref-stable close callback so effect doesn't re-run on every render
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;

  useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        onCloseRef.current();
      }
    };
    // Small delay to avoid the click that opened the dropdown from immediately closing it
    const timer = setTimeout(() => {
      document.addEventListener('pointerdown', handleOutsideClick, true);
    }, 0);
    return () => {
      clearTimeout(timer);
      document.removeEventListener('pointerdown', handleOutsideClick, true);
    };
  }, []);

  const toggleValue = useCallback((val: string) => {
    setChecked(prev => {
      const next = new Set(prev);
      if (next.has(val)) next.delete(val);
      else next.add(val);
      return next;
    });
  }, []);

  const toggleAll = useCallback(() => {
    setChecked(prev => {
      if (prev.size === uniqueValues.length) return new Set<string>();
      return new Set(uniqueValues);
    });
  }, [uniqueValues]);

  const handleApply = useCallback(() => {
    if (checked.size === uniqueValues.length) {
      onFilter(table.id, col, null); // all selected = no filter
    } else {
      onFilter(table.id, col, new Set(checked));
    }
    onClose();
  }, [checked, uniqueValues.length, onFilter, table.id, col, onClose]);

  const handleClear = useCallback(() => {
    onFilter(table.id, col, null);
    onClose();
  }, [onFilter, table.id, col, onClose]);

  // Adjust position to stay within viewport
  const adjustedPos = { ...position };
  if (typeof window !== 'undefined') {
    const dropdownH = 300; // max-height from CSS
    if (adjustedPos.y + dropdownH > window.innerHeight) {
      adjustedPos.y = Math.max(10, window.innerHeight - dropdownH - 10);
    }
  }

  const curSortDir = table.sortColumn === col ? table.sortDirection : null;

  const handleSortAsc = useCallback(() => {
    onSort(table.id, col, curSortDir === 'asc' ? null : 'asc');
    onClose();
  }, [onSort, table.id, col, curSortDir, onClose]);

  const handleSortDesc = useCallback(() => {
    onSort(table.id, col, curSortDir === 'desc' ? null : 'desc');
    onClose();
  }, [onSort, table.id, col, curSortDir, onClose]);

  return (
    <div
      ref={ref}
      className="table-filter-dropdown"
      style={{ left: adjustedPos.x, top: adjustedPos.y }}
      onMouseDown={(e) => e.stopPropagation()}
      onClick={(e) => e.stopPropagation()}
    >
      {/* Sort options */}
      <button
        className={`table-filter-sort-btn${curSortDir === 'asc' ? ' table-filter-sort-btn--active' : ''}`}
        onClick={handleSortAsc}
      >
        <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <path d="M8 3v10M4 7l4-4 4 4"/>
        </svg>
        Sort A → Z
      </button>
      <button
        className={`table-filter-sort-btn${curSortDir === 'desc' ? ' table-filter-sort-btn--active' : ''}`}
        onClick={handleSortDesc}
      >
        <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <path d="M8 13V3M4 9l4 4 4-4"/>
        </svg>
        Sort Z → A
      </button>
      <div className="table-filter-separator" />

      <label className="table-filter-item" style={{ fontWeight: 600 }}>
        <input
          type="checkbox"
          checked={checked.size === uniqueValues.length}
          onChange={toggleAll}
        />
        (Select All)
      </label>
      <div className="table-filter-list">
        {uniqueValues.map(val => (
          <label key={val} className="table-filter-item">
            <input
              type="checkbox"
              checked={checked.has(val)}
              onChange={() => toggleValue(val)}
            />
            {val === '' ? '(Blanks)' : val}
          </label>
        ))}
      </div>
      <div className="table-filter-actions">
        <button className="toolbar-btn" onMouseDown={(e) => e.preventDefault()} onClick={handleClear}>Clear</button>
        <button className="toolbar-btn toolbar-btn--active" onMouseDown={(e) => e.preventDefault()} onClick={handleApply}>Apply</button>
      </div>
    </div>
  );
}
