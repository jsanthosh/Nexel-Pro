import React, { useRef, useState, useCallback, useEffect } from 'react';
import { CellRange, CellFormatting, NumberFormatType, DEFAULT_FORMATTING, TableDefinition, TABLE_STYLES } from '../types/spreadsheet';
import { NUMBER_FORMAT_LABELS, NUMBER_FORMAT_EXAMPLES, CURRENCIES, DATE_FORMATS } from '../utils/numberFormat';
import { rangesOverlap } from '../utils/tableUtils';

interface ToolbarProps {
  selectedRange: CellRange | null;
  activeFormatting: CellFormatting;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onFormat: (range: CellRange, formatting: Partial<CellFormatting>) => void;
  onColor: (range: CellRange, backgroundColor?: string, textColor?: string) => void;
  onExport: () => void;
  onImport: (file: File) => void;
  onAddChart: () => void;
  onOpenConditionalFormat: () => void;
  onOpenPivotTable: () => void;
  onCreateTable: (range: CellRange, styleId?: string) => void;
  tables: TableDefinition[];
  onSetTableStyle: (id: string, styleId: string) => void;
}

/* ── SVG Icon helpers (14×14, stroke-based like Excel ribbon) ── */
const icon = (children: React.ReactNode) => (
  <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
    {children}
  </svg>
);

const AlignLeftIcon = () => icon(<><line x1="2" y1="4" x2="14" y2="4"/><line x1="2" y1="8" x2="10" y2="8"/><line x1="2" y1="12" x2="14" y2="12"/></>);
const AlignCenterIcon = () => icon(<><line x1="2" y1="4" x2="14" y2="4"/><line x1="4" y1="8" x2="12" y2="8"/><line x1="2" y1="12" x2="14" y2="12"/></>);
const AlignRightIcon = () => icon(<><line x1="2" y1="4" x2="14" y2="4"/><line x1="6" y1="8" x2="14" y2="8"/><line x1="2" y1="12" x2="14" y2="12"/></>);

const AlignTopIcon = () => icon(<><line x1="2" y1="2" x2="14" y2="2"/><line x1="8" y1="2" x2="8" y2="12"/><polyline points="5,5 8,2 11,5"/></>);
const AlignMiddleIcon = () => icon(<><line x1="2" y1="8" x2="5" y2="8"/><line x1="11" y1="8" x2="14" y2="8"/><polyline points="6,5 8,3 10,5"/><line x1="8" y1="3" x2="8" y2="13"/><polyline points="6,11 8,13 10,11"/></>);
const AlignBottomIcon = () => icon(<><line x1="2" y1="14" x2="14" y2="14"/><line x1="8" y1="4" x2="8" y2="14"/><polyline points="5,11 8,14 11,11"/></>);

const WrapTextIcon = () => icon(<><polyline points="2,3 14,3"/><path d="M2 7h10a2 2 0 0 1 0 4H10" /><polyline points="11,9.5 10,11 9,9.5"/><line x1="2" y1="11" x2="6" y2="11"/></>);

const IncreaseIndentIcon = () => icon(<><line x1="6" y1="4" x2="14" y2="4"/><line x1="6" y1="8" x2="14" y2="8"/><line x1="6" y1="12" x2="14" y2="12"/><polyline points="2,6 4,8 2,10"/></>);
const DecreaseIndentIcon = () => icon(<><line x1="6" y1="4" x2="14" y2="4"/><line x1="6" y1="8" x2="14" y2="8"/><line x1="6" y1="12" x2="14" y2="12"/><polyline points="4,6 2,8 4,10"/></>);

/* Improved text rotation icon: slanted "A" with a circular arrow */
const TextRotationIcon = () => (
  <svg width="14" height="14" viewBox="0 0 16 16" fill="none" strokeLinecap="round" strokeLinejoin="round">
    {/* Slanted letter A */}
    <g stroke="currentColor" strokeWidth="1.5" transform="rotate(-25 8 10)">
      <line x1="4" y1="12" x2="7" y2="3" />
      <line x1="7" y1="3" x2="10" y2="12" />
      <line x1="5.2" y1="9" x2="8.8" y2="9" />
    </g>
    {/* Curved arrow */}
    <path d="M12 4a5 5 0 0 1 0 7" stroke="currentColor" strokeWidth="1.3" fill="none" />
    <polyline points="11,10 12.5,11.5 14,9.5" stroke="currentColor" strokeWidth="1.3" fill="none" />
  </svg>
);

const ChevronDownIcon = () => (
  <svg width="10" height="10" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="4,6 8,10 12,6" />
  </svg>
);

const ChevronRightIcon = () => (
  <svg width="8" height="8" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="6,4 10,8 6,12" />
  </svg>
);

const ROTATION_OPTIONS = [
  { label: 'No Rotation', value: 0 },
  { label: 'Angle Counterclockwise', value: 45 },
  { label: 'Angle Clockwise', value: -45 },
  { label: 'Vertical Text', value: 90 },
  { label: 'Rotate Text Up', value: 90 },
  { label: 'Rotate Text Down', value: -90 },
];

const FONT_FAMILIES = [
  'Default',
  'Arial',
  'Verdana',
  'Helvetica',
  'Tahoma',
  'Times New Roman',
  'Georgia',
  'Courier New',
  'Roboto',
  'Open Sans',
  'Lato',
  'Montserrat',
  'Poppins',
  'Inter',
  'Nunito',
  'Raleway',
  'Playfair Display',
  'Merriweather',
  'Source Code Pro',
  'Fira Code',
  'Ubuntu',
  'Oswald',
  'PT Sans',
  'Noto Sans',
];

// Simple format types without submenus
const SIMPLE_FORMATS: NumberFormatType[] = ['general', 'number', 'percentage', 'text', 'time'];

export default function Toolbar({ selectedRange, activeFormatting, canUndo, canRedo, onUndo, onRedo, onFormat, onColor, onExport, onImport, onAddChart, onOpenConditionalFormat, onOpenPivotTable, onCreateTable, tables, onSetTableStyle }: ToolbarProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const fmt = activeFormatting ?? DEFAULT_FORMATTING;
  const [showRotationMenu, setShowRotationMenu] = useState(false);
  const [showTableStylePicker, setShowTableStylePicker] = useState(false);
  const [showNFMenu, setShowNFMenu] = useState(false);
  const [showCurrencySubmenu, setShowCurrencySubmenu] = useState(false);
  const [showDateSubmenu, setShowDateSubmenu] = useState(false);
  const [customFormatInput, setCustomFormatInput] = useState('');

  const nfMenuRef = useRef<HTMLDivElement>(null);
  const currencyTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const dateTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const openCurrencySubmenu = useCallback(() => {
    if (currencyTimerRef.current) { clearTimeout(currencyTimerRef.current); currencyTimerRef.current = null; }
    if (dateTimerRef.current) { clearTimeout(dateTimerRef.current); dateTimerRef.current = null; }
    setShowCurrencySubmenu(true);
    setShowDateSubmenu(false);
  }, []);
  const closeCurrencySubmenu = useCallback(() => {
    currencyTimerRef.current = setTimeout(() => setShowCurrencySubmenu(false), 150);
  }, []);
  const openDateSubmenu = useCallback(() => {
    if (dateTimerRef.current) { clearTimeout(dateTimerRef.current); dateTimerRef.current = null; }
    if (currencyTimerRef.current) { clearTimeout(currencyTimerRef.current); currencyTimerRef.current = null; }
    setShowDateSubmenu(true);
    setShowCurrencySubmenu(false);
  }, []);
  const closeDateSubmenu = useCallback(() => {
    dateTimerRef.current = setTimeout(() => setShowDateSubmenu(false), 150);
  }, []);

  const activeTableStyleId = selectedRange
    ? (tables.find(t => rangesOverlap(t.range, selectedRange))?.styleId ?? null)
    : null;

  const withRange = (fn: (range: CellRange) => void) => {
    if (!selectedRange) return;
    fn(selectedRange);
  };

  // Close NF menu on outside click
  useEffect(() => {
    if (!showNFMenu) return;
    const handler = (e: MouseEvent) => {
      if (nfMenuRef.current && !nfMenuRef.current.contains(e.target as Node)) {
        setShowNFMenu(false);
        setShowCurrencySubmenu(false);
        setShowDateSubmenu(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [showNFMenu]);

  const selectFormat = useCallback((nf: NumberFormatType) => {
    withRange(r => onFormat(r, { numberFormat: nf }));
    setShowNFMenu(false);
    setShowCurrencySubmenu(false);
    setShowDateSubmenu(false);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedRange, onFormat]);

  const selectCurrency = useCallback((code: string) => {
    withRange(r => onFormat(r, { numberFormat: 'currency', currencyCode: code }));
    setShowNFMenu(false);
    setShowCurrencySubmenu(false);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedRange, onFormat]);

  const selectDateFormat = useCallback((dateFormatId: string) => {
    withRange(r => onFormat(r, { numberFormat: 'date', dateFormat: dateFormatId }));
    setShowNFMenu(false);
    setShowDateSubmenu(false);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedRange, onFormat]);

  const applyCustomFormat = useCallback(() => {
    if (!customFormatInput.trim()) return;
    withRange(r => onFormat(r, { numberFormat: 'custom', customFormat: customFormatInput.trim() }));
    setShowNFMenu(false);
    setCustomFormatInput('');
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedRange, onFormat, customFormatInput]);

  const currentFormatLabel = NUMBER_FORMAT_LABELS[fmt.numberFormat ?? 'general'];

  return (
    <div className="toolbar">
      {/* Undo / Redo */}
      <div className="toolbar-group">
        <button
          className="toolbar-btn toolbar-btn--icon"
          title="Undo (Ctrl+Z)"
          onMouseDown={(e) => e.preventDefault()}
          onClick={onUndo}
          disabled={!canUndo}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" /></svg>
        </button>
        <button
          className="toolbar-btn toolbar-btn--icon"
          title="Redo (Ctrl+Y)"
          onMouseDown={(e) => e.preventDefault()}
          onClick={onRedo}
          disabled={!canRedo}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 1 1-2.13-9.36L23 10" /></svg>
        </button>
      </div>

      <div className="toolbar-divider" />

      {/* Font family */}
      <div className="toolbar-group">
        <select
          className="toolbar-font-select"
          title="Font Family"
          value={fmt.fontFamily ?? 'Default'}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => withRange(r => onFormat(r, { fontFamily: e.target.value }))}
        >
          {FONT_FAMILIES.map((f) => (
            <option key={f} value={f} style={{ fontFamily: f === 'Default' ? 'inherit' : `'${f}', sans-serif` }}>
              {f}
            </option>
          ))}
        </select>
      </div>

      <div className="toolbar-divider" />

      {/* Number Format - Rich dropdown with submenus */}
      <div className="toolbar-group">
        <div className="nf-dropdown-wrapper" ref={nfMenuRef}>
          <button
            className={`nf-dropdown-trigger${showNFMenu ? ' nf-dropdown-trigger--open' : ''}`}
            title="Number Format"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => setShowNFMenu(!showNFMenu)}
          >
            <span>{currentFormatLabel}</span>
            <ChevronDownIcon />
          </button>

          {showNFMenu && (
            <div className="nf-dropdown-menu">
              <div className="nf-menu-section-label">Format</div>

              {/* Simple formats */}
              {SIMPLE_FORMATS.map(key => (
                <button
                  key={key}
                  className={`nf-menu-item${(fmt.numberFormat ?? 'general') === key ? ' nf-menu-item--active' : ''}`}
                  onClick={() => selectFormat(key)}
                >
                  <span>{NUMBER_FORMAT_LABELS[key]}</span>
                  {key !== 'general' && key !== 'text' && (
                    <span className="nf-menu-item-example">{NUMBER_FORMAT_EXAMPLES[key]}</span>
                  )}
                </button>
              ))}

              {/* Currency with submenu */}
              <div
                className="nf-submenu-wrapper"
                onMouseEnter={openCurrencySubmenu}
                onMouseLeave={closeCurrencySubmenu}
              >
                <button
                  className={`nf-submenu-item${(fmt.numberFormat ?? 'general') === 'currency' ? ' nf-menu-item--active' : ''}`}
                >
                  <span>Currency</span>
                  <ChevronRightIcon />
                </button>
                {showCurrencySubmenu && (
                  <div
                    className="nf-submenu"
                    onMouseEnter={openCurrencySubmenu}
                    onMouseLeave={closeCurrencySubmenu}
                  >
                    {CURRENCIES.map(cur => (
                      <button
                        key={cur.code}
                        className={`nf-menu-item${(fmt.currencyCode ?? 'USD') === cur.code && fmt.numberFormat === 'currency' ? ' nf-menu-item--active' : ''}`}
                        onClick={() => selectCurrency(cur.code)}
                      >
                        <span>{cur.label}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {/* Accounting with same currency options */}
              <button
                className={`nf-menu-item${(fmt.numberFormat ?? 'general') === 'accounting' ? ' nf-menu-item--active' : ''}`}
                onClick={() => selectFormat('accounting')}
              >
                <span>Accounting</span>
                <span className="nf-menu-item-example">{NUMBER_FORMAT_EXAMPLES.accounting}</span>
              </button>

              {/* Date with submenu */}
              <div
                className="nf-submenu-wrapper"
                onMouseEnter={openDateSubmenu}
                onMouseLeave={closeDateSubmenu}
              >
                <button
                  className={`nf-submenu-item${(fmt.numberFormat ?? 'general') === 'date' ? ' nf-menu-item--active' : ''}`}
                >
                  <span>Date</span>
                  <ChevronRightIcon />
                </button>
                {showDateSubmenu && (
                  <div
                    className="nf-submenu"
                    onMouseEnter={openDateSubmenu}
                    onMouseLeave={closeDateSubmenu}
                  >
                    {DATE_FORMATS.map(df => (
                      <button
                        key={df.id}
                        className={`nf-menu-item${(fmt.dateFormat ?? 'mm/dd/yyyy') === df.id && fmt.numberFormat === 'date' ? ' nf-menu-item--active' : ''}`}
                        onClick={() => selectDateFormat(df.id)}
                      >
                        <span>{df.label}</span>
                        <span className="nf-menu-item-example">{df.example}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>

              <div className="nf-menu-section-label" style={{ marginTop: 4 }}>Custom</div>
              <div className="nf-custom-input-row">
                <input
                  className="nf-custom-input"
                  placeholder="#,##0.00"
                  value={customFormatInput}
                  onChange={(e) => setCustomFormatInput(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter') applyCustomFormat(); }}
                />
                <button className="nf-custom-apply" onClick={applyCustomFormat}>Apply</button>
              </div>
            </div>
          )}
        </div>

        <button
          className="toolbar-btn toolbar-btn--icon"
          title="Decrease Decimals"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => withRange(r => onFormat(r, { decimalPlaces: Math.max(0, (fmt.decimalPlaces ?? 2) - 1) }))}
        >
          .0
        </button>
        <button
          className="toolbar-btn toolbar-btn--icon"
          title="Increase Decimals"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => withRange(r => onFormat(r, { decimalPlaces: Math.min(10, (fmt.decimalPlaces ?? 2) + 1) }))}
        >
          .00
        </button>
      </div>

      <div className="toolbar-divider" />

      {/* Font style */}
      <div className="toolbar-group">
        <button className={`toolbar-btn${fmt.bold ? ' toolbar-btn--active' : ''}`} title="Bold" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { bold: !fmt.bold }))}>
          <b>B</b>
        </button>
        <button className={`toolbar-btn${fmt.italic ? ' toolbar-btn--active' : ''}`} title="Italic" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { italic: !fmt.italic }))}>
          <i>I</i>
        </button>
        <button className={`toolbar-btn${fmt.underline ? ' toolbar-btn--active' : ''}`} title="Underline" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { underline: !fmt.underline }))}>
          <u>U</u>
        </button>
      </div>

      <div className="toolbar-divider" />

      {/* ── Alignment Group (Excel-style 2-row grid) ── */}
      <div className="toolbar-align-group">
        {/* Top row: Horizontal alignment */}
        <div className="toolbar-align-row">
          <button className={`toolbar-btn toolbar-btn--icon${fmt.hAlign === 'left' ? ' toolbar-btn--active' : ''}`} title="Align Left" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { hAlign: 'left' }))}>
            <AlignLeftIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${fmt.hAlign === 'center' ? ' toolbar-btn--active' : ''}`} title="Align Center" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { hAlign: 'center' }))}>
            <AlignCenterIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${fmt.hAlign === 'right' ? ' toolbar-btn--active' : ''}`} title="Align Right" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { hAlign: 'right' }))}>
            <AlignRightIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${(fmt.indent ?? 0) > 0 ? ' toolbar-btn--active' : ''}`} title="Decrease Indent" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { indent: Math.max(0, (fmt.indent ?? 0) - 1) }))}>
            <DecreaseIndentIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon`} title="Increase Indent" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { indent: (fmt.indent ?? 0) + 1 }))}>
            <IncreaseIndentIcon />
          </button>
        </div>
        {/* Bottom row: Vertical alignment + Wrap + Rotation */}
        <div className="toolbar-align-row">
          <button className={`toolbar-btn toolbar-btn--icon${fmt.vAlign === 'top' ? ' toolbar-btn--active' : ''}`} title="Align Top" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { vAlign: 'top' }))}>
            <AlignTopIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${fmt.vAlign === 'middle' ? ' toolbar-btn--active' : ''}`} title="Align Middle" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { vAlign: 'middle' }))}>
            <AlignMiddleIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${fmt.vAlign === 'bottom' ? ' toolbar-btn--active' : ''}`} title="Align Bottom" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { vAlign: 'bottom' }))}>
            <AlignBottomIcon />
          </button>
          <button className={`toolbar-btn toolbar-btn--icon${fmt.wrapText ? ' toolbar-btn--active' : ''}`} title="Wrap Text" onMouseDown={(e) => e.preventDefault()} onClick={() => withRange(r => onFormat(r, { wrapText: !fmt.wrapText }))}>
            <WrapTextIcon />
          </button>
          <div className="toolbar-rotation-wrapper">
            <button
              className={`toolbar-btn toolbar-btn--icon${(fmt.textRotation ?? 0) !== 0 ? ' toolbar-btn--active' : ''}`}
              title="Text Rotation"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setShowRotationMenu(!showRotationMenu)}
            >
              <TextRotationIcon />
            </button>
            {showRotationMenu && (
              <div className="toolbar-rotation-menu" onMouseLeave={() => setShowRotationMenu(false)}>
                {ROTATION_OPTIONS.map((opt) => (
                  <button
                    key={opt.value + opt.label}
                    className={`toolbar-rotation-item${(fmt.textRotation ?? 0) === opt.value ? ' toolbar-rotation-item--active' : ''}`}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => {
                      withRange(r => onFormat(r, { textRotation: opt.value }));
                      setShowRotationMenu(false);
                    }}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="toolbar-divider" />

      {/* Colors */}
      <div className="toolbar-group">
        <label className="toolbar-label" title="Background Color">
          BG
          <input type="color" defaultValue="#ffff00" onChange={(e) => withRange(r => onColor(r, e.target.value, undefined))} />
        </label>
        <label className="toolbar-label" title="Text Color">
          A
          <input type="color" defaultValue="#000000" onChange={(e) => withRange(r => onColor(r, undefined, e.target.value))} />
        </label>
      </div>

      <div className="toolbar-divider" />

      <div className="toolbar-group">
        <button className="toolbar-btn" onMouseDown={(e) => e.preventDefault()} onClick={onAddChart} title="Add Chart from selected range" disabled={!selectedRange}>
          + Chart
        </button>
        <button className="toolbar-btn" onMouseDown={(e) => e.preventDefault()} onClick={onOpenConditionalFormat} title="Conditional Formatting">
          CF Rules
        </button>
        {/* Table with style picker */}
        <div className="toolbar-table-wrapper">
          <button
            className="toolbar-btn"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              if (!selectedRange) return;
              if (tables.length > 0) {
                setShowTableStylePicker(!showTableStylePicker);
              } else {
                setShowTableStylePicker(true);
              }
            }}
            title="Insert Table with style"
            disabled={!selectedRange}
          >
            Table ▾
          </button>
          {showTableStylePicker && (
            <div
              className="table-style-picker"
              onMouseLeave={() => setShowTableStylePicker(false)}
            >
              <div className="table-style-picker-header">Choose a table style</div>
              <div className="table-style-grid">
                {TABLE_STYLES.map(style => (
                  <button
                    key={style.id}
                    className={`table-style-swatch${activeTableStyleId === style.id ? ' table-style-swatch--active' : ''}`}
                    title={style.name}
                    onClick={() => {
                      if (selectedRange) {
                        const existingTable = tables.find(t => rangesOverlap(t.range, selectedRange));
                        if (existingTable) {
                          onSetTableStyle(existingTable.id, style.id);
                        } else {
                          onCreateTable(selectedRange, style.id);
                        }
                      }
                      setShowTableStylePicker(false);
                    }}
                  >
                    <div className="swatch-header" style={{ background: style.headerBg, color: style.headerText }}>Hdr</div>
                    <div className="swatch-row" style={{ background: style.evenRowBg }} />
                    <div className="swatch-row" style={{ background: style.oddRowBg }} />
                    <div className="swatch-row" style={{ background: style.evenRowBg }} />
                    <span className="swatch-label">{style.name}</span>
                  </button>
                ))}
              </div>
              {tables.length > 0 && (
                <>
                  <div className="table-style-picker-header" style={{ marginTop: 8 }}>Restyle existing tables</div>
                  {tables.map(t => (
                    <div key={t.id} className="table-restyle-row">
                      <span className="table-restyle-name">{t.id.replace('table-', 'Table ')}</span>
                      <select
                        className="table-restyle-select"
                        value={t.styleId ?? 'blue'}
                        onChange={(e) => {
                          onSetTableStyle(t.id, e.target.value);
                        }}
                      >
                        {TABLE_STYLES.map(s => (
                          <option key={s.id} value={s.id}>{s.name}</option>
                        ))}
                      </select>
                    </div>
                  ))}
                </>
              )}
            </div>
          )}
        </div>
        <button className="toolbar-btn" onMouseDown={(e) => e.preventDefault()} onClick={onOpenPivotTable} title="Create Pivot Table" disabled={!selectedRange}>
          Pivot Table
        </button>
        <button className="toolbar-btn" onClick={onExport} title="Export CSV">
          Export CSV
        </button>
        <button className="toolbar-btn" onMouseDown={(e) => e.preventDefault()} onClick={() => fileInputRef.current?.click()} title="Import CSV or XLSX">
          Import
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".csv,.xlsx,.xls"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onImport(file);
            e.target.value = '';
          }}
        />
      </div>
    </div>
  );
}
