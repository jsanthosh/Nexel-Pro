import React, { useState } from 'react';
import { CellRange, CellFormatting, ConditionalFormatRule, ConditionType } from '../types/spreadsheet';
import { rangeToString, parseRange } from '../utils/cellRange';

interface ConditionalFormatDialogProps {
  rules: ConditionalFormatRule[];
  selectedRange: CellRange | null;
  onAdd: (rule: ConditionalFormatRule) => void;
  onDelete: (id: string) => void;
  onClose: () => void;
}

const CONDITION_LABELS: Record<ConditionType, string> = {
  greaterThan: 'Greater than',
  lessThan: 'Less than',
  equalTo: 'Equal to',
  between: 'Between',
  textContains: 'Text contains',
  isEmpty: 'Is empty',
  isNotEmpty: 'Is not empty',
  duplicateValues: 'Duplicate values',
};

const NO_VALUE_CONDITIONS: ConditionType[] = ['isEmpty', 'isNotEmpty', 'duplicateValues'];

export default function ConditionalFormatDialog({ rules, selectedRange, onAdd, onDelete, onClose }: ConditionalFormatDialogProps) {
  const [rangeStr, setRangeStr] = useState(selectedRange ? rangeToString(selectedRange) : '');
  const [condition, setCondition] = useState<ConditionType>('greaterThan');
  const [value1, setValue1] = useState('');
  const [value2, setValue2] = useState('');
  const [bold, setBold] = useState(false);
  const [bgColor, setBgColor] = useState('#ffcccc');
  const [textColor, setTextColor] = useState('#cc0000');
  const [useBg, setUseBg] = useState(true);
  const [useText, setUseText] = useState(true);

  const handleAdd = () => {
    let range: CellRange;
    try {
      range = parseRange(rangeStr);
    } catch {
      return;
    }

    const values: string[] = [];
    if (!NO_VALUE_CONDITIONS.includes(condition)) {
      values.push(value1);
      if (condition === 'between') values.push(value2);
    }

    const formatting: Partial<CellFormatting> = {};
    if (bold) formatting.bold = true;
    if (useBg) formatting.backgroundColor = bgColor;
    if (useText) formatting.textColor = textColor;

    const rule: ConditionalFormatRule = {
      id: `cf-${Date.now()}`,
      range,
      condition,
      values,
      formatting,
      priority: rules.length,
    };

    onAdd(rule);
    setValue1('');
    setValue2('');
  };

  const needsValue = !NO_VALUE_CONDITIONS.includes(condition);
  const needsTwoValues = condition === 'between';

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal-dialog" onMouseDown={(e) => e.stopPropagation()} style={{ maxWidth: 560 }}>
        <div className="modal-header">
          <span style={{ fontWeight: 600, fontSize: 14 }}>Conditional Formatting</span>
          <button className="modal-close-btn" onClick={onClose}>&times;</button>
        </div>

        <div className="modal-body">
          {rules.length > 0 && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 8, color: '#555' }}>Active Rules</div>
              {rules.map(rule => (
                <div key={rule.id} className="cf-rule-item">
                  <div
                    className="cf-rule-preview"
                    style={{
                      backgroundColor: rule.formatting.backgroundColor ?? '#fff',
                      color: rule.formatting.textColor ?? '#000',
                      fontWeight: rule.formatting.bold ? 'bold' : 'normal',
                    }}
                  >
                    Aa
                  </div>
                  <div style={{ flex: 1, fontSize: 12 }}>
                    <div>{rangeToString(rule.range)}: {CONDITION_LABELS[rule.condition]}{rule.values.length > 0 ? ` ${rule.values.join(' - ')}` : ''}</div>
                  </div>
                  <button className="cf-delete-btn" onClick={() => onDelete(rule.id)} title="Delete rule">&times;</button>
                </div>
              ))}
            </div>
          )}

          <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 8, color: '#555' }}>New Rule</div>

          <div className="cf-form">
            <div className="cf-form-row">
              <label>Range</label>
              <input
                type="text"
                value={rangeStr}
                onChange={(e) => setRangeStr(e.target.value.toUpperCase())}
                placeholder="A1:B10"
                className="cf-input"
              />
            </div>

            <div className="cf-form-row">
              <label>Condition</label>
              <select
                value={condition}
                onChange={(e) => setCondition(e.target.value as ConditionType)}
                className="cf-select"
              >
                {(Object.keys(CONDITION_LABELS) as ConditionType[]).map(c => (
                  <option key={c} value={c}>{CONDITION_LABELS[c]}</option>
                ))}
              </select>
            </div>

            {needsValue && (
              <div className="cf-form-row">
                <label>Value{needsTwoValues ? ' (min)' : ''}</label>
                <input
                  type="text"
                  value={value1}
                  onChange={(e) => setValue1(e.target.value)}
                  placeholder="Enter value"
                  className="cf-input"
                />
              </div>
            )}

            {needsTwoValues && (
              <div className="cf-form-row">
                <label>Value (max)</label>
                <input
                  type="text"
                  value={value2}
                  onChange={(e) => setValue2(e.target.value)}
                  placeholder="Enter max value"
                  className="cf-input"
                />
              </div>
            )}

            <div className="cf-form-row">
              <label>Style</label>
              <div className="cf-style-options">
                <label className="cf-checkbox">
                  <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} />
                  <b>Bold</b>
                </label>
                <label className="cf-checkbox">
                  <input type="checkbox" checked={useBg} onChange={(e) => setUseBg(e.target.checked)} />
                  BG
                  {useBg && <input type="color" value={bgColor} onChange={(e) => setBgColor(e.target.value)} />}
                </label>
                <label className="cf-checkbox">
                  <input type="checkbox" checked={useText} onChange={(e) => setUseText(e.target.checked)} />
                  Text
                  {useText && <input type="color" value={textColor} onChange={(e) => setTextColor(e.target.value)} />}
                </label>
              </div>
            </div>
          </div>
        </div>

        <div className="modal-footer">
          <button className="modal-btn" onClick={onClose}>Close</button>
          <button className="modal-btn modal-btn--primary" onClick={handleAdd} disabled={!rangeStr}>
            Add Rule
          </button>
        </div>
      </div>
    </div>
  );
}
