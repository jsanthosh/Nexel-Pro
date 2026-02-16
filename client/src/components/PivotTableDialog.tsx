import React, { useState, useMemo } from 'react';
import { SheetData, CellRange } from '../types/spreadsheet';
import { rangeToString, parseRange } from '../utils/cellRange';
import { PivotFieldConfig, PivotValueFieldConfig, AggregationType } from '../types/pivot';
import { computePivotTable } from '../utils/pivotEngine';

interface PivotTableDialogProps {
  cells: SheetData;
  selectedRange: CellRange | null;
  onClose: () => void;
}

export default function PivotTableDialog({ cells, selectedRange, onClose }: PivotTableDialogProps) {
  const [rangeStr, setRangeStr] = useState(selectedRange ? rangeToString(selectedRange) : '');
  const [rowFields, setRowFields] = useState<PivotFieldConfig[]>([]);
  const [columnFields, setColumnFields] = useState<PivotFieldConfig[]>([]);
  const [valueFields, setValueFields] = useState<PivotValueFieldConfig[]>([]);

  // Parse source range and extract headers
  const sourceRange = useMemo(() => {
    try { return parseRange(rangeStr); } catch { return null; }
  }, [rangeStr]);

  const availableFields = useMemo(() => {
    if (!sourceRange) return [];
    const fields: { index: number; label: string }[] = [];
    for (let c = sourceRange.start.col; c <= sourceRange.end.col; c++) {
      const cell = cells.get(`${sourceRange.start.row}:${c}`);
      fields.push({ index: c - sourceRange.start.col, label: cell?.displayValue ?? cell?.value ?? `Col ${c + 1}` });
    }
    return fields;
  }, [sourceRange, cells]);

  const pivotResult = useMemo(() => {
    if (!sourceRange || rowFields.length === 0 || valueFields.length === 0) return null;
    return computePivotTable(cells, { sourceRange, rowFields, columnFields, valueFields });
  }, [cells, sourceRange, rowFields, columnFields, valueFields]);

  const addRowField = (index: number, label: string) => {
    if (rowFields.some(f => f.columnIndex === index)) return;
    setRowFields(prev => [...prev, { columnIndex: index, label }]);
  };

  const addColumnField = (index: number, label: string) => {
    if (columnFields.some(f => f.columnIndex === index)) return;
    setColumnFields(prev => [...prev, { columnIndex: index, label }]);
  };

  const addValueField = (index: number, label: string) => {
    if (valueFields.some(f => f.columnIndex === index)) return;
    setValueFields(prev => [...prev, { columnIndex: index, label, aggregation: 'sum' }]);
  };

  const updateAggregation = (index: number, aggregation: AggregationType) => {
    setValueFields(prev => prev.map(f =>
      f.columnIndex === index ? { ...f, aggregation } : f
    ));
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal-dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 800, maxHeight: '85vh' }}>
        <div className="modal-header">
          <span style={{ fontWeight: 600, fontSize: 14 }}>Pivot Table</span>
          <button className="modal-close-btn" onClick={onClose}>&times;</button>
        </div>

        <div className="modal-body" style={{ display: 'flex', gap: 16, minHeight: 300 }}>
          {/* Configuration */}
          <div style={{ flex: '0 0 260px', display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div className="cf-form-row">
              <label>Source Range</label>
              <input
                type="text"
                value={rangeStr}
                onChange={(e) => {
                  setRangeStr(e.target.value.toUpperCase());
                  setRowFields([]);
                  setColumnFields([]);
                  setValueFields([]);
                }}
                placeholder="A1:D20"
                className="cf-input"
              />
            </div>

            {availableFields.length > 0 && (
              <>
                <div style={{ fontSize: 12, fontWeight: 600, color: '#555' }}>Available Fields</div>
                <div className="pivot-fields-list">
                  {availableFields.map(f => (
                    <div key={f.index} className="pivot-field-item">
                      <span style={{ flex: 1, fontSize: 12 }}>{f.label}</span>
                      <button className="pivot-field-btn" onClick={() => addRowField(f.index, f.label)} title="Add as row">R</button>
                      <button className="pivot-field-btn" onClick={() => addColumnField(f.index, f.label)} title="Add as column">C</button>
                      <button className="pivot-field-btn" onClick={() => addValueField(f.index, f.label)} title="Add as value">V</button>
                    </div>
                  ))}
                </div>

                {rowFields.length > 0 && (
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: '#1a73e8', marginBottom: 4 }}>Row Fields</div>
                    {rowFields.map(f => (
                      <div key={f.columnIndex} className="pivot-config-item">
                        <span>{f.label}</span>
                        <button onClick={() => setRowFields(prev => prev.filter(x => x.columnIndex !== f.columnIndex))}>&times;</button>
                      </div>
                    ))}
                  </div>
                )}

                {columnFields.length > 0 && (
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: '#e8711a', marginBottom: 4 }}>Column Fields</div>
                    {columnFields.map(f => (
                      <div key={f.columnIndex} className="pivot-config-item">
                        <span>{f.label}</span>
                        <button onClick={() => setColumnFields(prev => prev.filter(x => x.columnIndex !== f.columnIndex))}>&times;</button>
                      </div>
                    ))}
                  </div>
                )}

                {valueFields.length > 0 && (
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, color: '#34a853', marginBottom: 4 }}>Value Fields</div>
                    {valueFields.map(f => (
                      <div key={f.columnIndex} className="pivot-config-item">
                        <span>{f.label}</span>
                        <select
                          value={f.aggregation}
                          onChange={(e) => updateAggregation(f.columnIndex, e.target.value as AggregationType)}
                          className="pivot-agg-select"
                        >
                          <option value="sum">Sum</option>
                          <option value="count">Count</option>
                          <option value="average">Average</option>
                          <option value="min">Min</option>
                          <option value="max">Max</option>
                        </select>
                        <button onClick={() => setValueFields(prev => prev.filter(x => x.columnIndex !== f.columnIndex))}>&times;</button>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>

          {/* Results */}
          <div style={{ flex: 1, overflow: 'auto' }}>
            {pivotResult && pivotResult.rows.length > 0 ? (
              <table className="pivot-table">
                <thead>
                  <tr>
                    {pivotResult.headers.map((h, i) => <th key={i}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {pivotResult.rows.map((row, ri) => (
                    <tr key={ri}>
                      {row.map((cell, ci) => <td key={ci}>{cell}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div style={{ color: '#888', fontSize: 13, padding: 32, textAlign: 'center' }}>
                {!sourceRange
                  ? 'Enter a source data range to begin.'
                  : rowFields.length === 0 || valueFields.length === 0
                    ? 'Add at least one Row field and one Value field.'
                    : 'No data to display.'}
              </div>
            )}
          </div>
        </div>

        <div className="modal-footer">
          <button className="modal-btn" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}
