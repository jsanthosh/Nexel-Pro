# EXCEL FEATURE PARITY REFERENCE

> **Purpose**: This is the master feature reference for building our spreadsheet application.
> Claude Code MUST consult this file before implementing any spreadsheet feature to ensure
> we don't miss Excel's depth of customization. Every feature listed here represents what
> users expect from a professional spreadsheet.
>
> **Priority Key**: P1 = Must have for launch | P2 = Within 12 months | P3 = Roadmap
> **Status Key**: ⬜ Not started | 🟡 In progress | ✅ Done | ⏭️ Deferred

---

## 1. GRID RENDERING & SCROLLING

### 1.1 Virtual Scrolling / Viewport Rendering [P1]
- ⬜ Only render visible cells + configurable buffer zone (rows/cols ahead)
- ⬜ Recycle Canvas/DOM elements on scroll
- ⬜ Must handle 1M+ rows without performance degradation
- ⬜ Viewport-based rendering: track visible range, request data for visible + padding

### 1.2 Smooth Scrolling [P1]
- ⬜ Pixel-precise scroll position (not cell-snapping)
- ⬜ Momentum/inertia scrolling on touch devices
- ⬜ Scroll anchoring on content changes (insert/delete rows shouldn't jump viewport)
- ⬜ Horizontal + vertical scroll synchronization for frozen panes

### 1.3 Zoom [P1]
- ⬜ Range: 10% to 400%
- ⬜ Ctrl+Mouse wheel zoom
- ⬜ Pinch-to-zoom on touch
- ⬜ Fit-to-page / fit-to-width zoom presets
- ⬜ Zoom preserves scroll position (zoom into cursor location)
- ⬜ Zoom level indicator in status bar

### 1.4 Freeze Panes / Split Panes [P1]
- ⬜ Freeze top N rows
- ⬜ Freeze left N columns
- ⬜ Freeze both rows + columns simultaneously
- ⬜ Freeze at arbitrary row/col (not just first row/col)
- ⬜ Split into 4 independent scrollable quadrants
- ⬜ Synced scrolling between panes (horizontal in row panes, vertical in col panes)
- ⬜ Visual divider line for frozen area

### 1.5 Grid Lines [P1]
- ⬜ Show/hide gridlines per sheet
- ⬜ Print gridlines toggle (separate from display)
- ⬜ Gridline color customization

### 1.6 Multiple Views [P3]
- ⬜ View same workbook in multiple windows
- ⬜ Side-by-side arrangement
- ⬜ Independent scroll/zoom per view, synchronized data

### 1.7 Right-to-Left (RTL) Layout [P2]
- ⬜ Sheet direction toggle (LTR/RTL)
- ⬜ Column headers numbered right-to-left in RTL mode
- ⬜ Default text alignment follows sheet direction
- ⬜ Bidirectional text support within cells

---

## 2. CELL SELECTION & NAVIGATION

### 2.1 Single Cell Selection [P1]
- ⬜ Click to select
- ⬜ Arrow key navigation
- ⬜ Tab to move right, Enter to move down (configurable direction)
- ⬜ Name box for direct cell jump (type "Z100" + Enter)
- ⬜ Name box shows current cell address
- ⬜ Name box dropdown shows named ranges

### 2.2 Range Selection [P1]
- ⬜ Click + drag to select range
- ⬜ Shift + click to extend selection from active cell
- ⬜ Shift + Arrow to extend selection by keyboard
- ⬜ Visible selection border (blue/green outline)
- ⬜ Fill handle (small square at bottom-right of selection)

### 2.3 Multi-Range Selection [P1]
- ⬜ Ctrl + click to add disjoint cell to selection
- ⬜ Ctrl + drag for additional disjoint ranges
- ⬜ Operations (formatting, delete, etc.) apply to ALL selected ranges
- ⬜ Status bar shows SUM/COUNT/AVERAGE of all selected ranges

### 2.4 Row / Column Selection [P1]
- ⬜ Click row header to select entire row
- ⬜ Click column header to select entire column
- ⬜ Shift + click header to select range of rows/cols
- ⬜ Ctrl + click header for multi-select rows/cols
- ⬜ Ctrl + Space = select entire column; Shift + Space = select entire row

### 2.5 Select All [P1]
- ⬜ Click intersection button (top-left corner of row/col headers)
- ⬜ Ctrl+A: first press selects current region, second selects all
- ⬜ Ctrl+Shift+Space

### 2.6 Go To / Navigate [P1]
- ⬜ Ctrl+G / F5: Go To dialog
- ⬜ Go To named ranges
- ⬜ **Go To Special** (critical for power users):
  - ⬜ Blanks
  - ⬜ Formulas (with sub-filters: numbers, text, logicals, errors)
  - ⬜ Constants (with sub-filters)
  - ⬜ Comments/Notes
  - ⬜ Conditional formats
  - ⬜ Data validation cells
  - ⬜ Visible cells only
  - ⬜ Current region / Current array
  - ⬜ Objects (shapes, charts)
  - ⬜ Row/column differences
  - ⬜ Precedents / Dependents (direct or all levels)

### 2.7 Find & Replace [P1]
- ⬜ Find text in cell values
- ⬜ Find text in formulas
- ⬜ Find text in comments/notes
- ⬜ Match case option
- ⬜ Match entire cell contents
- ⬜ Search by rows or columns
- ⬜ Search in active sheet or entire workbook
- ⬜ Find All (list all matches with navigation)
- ⬜ Replace / Replace All
- ⬜ Find format (find cells with specific formatting)
- ⬜ Regex support (enhancement over Excel)

### 2.8 Keyboard Navigation [P1]
- ⬜ Ctrl+Home → A1
- ⬜ Ctrl+End → last used cell
- ⬜ Ctrl+Arrow → jump to edge of data region
- ⬜ Ctrl+Shift+Arrow → select to edge of data region
- ⬜ Page Up / Page Down → scroll by viewport height
- ⬜ Alt+Page Up/Down → scroll left/right
- ⬜ Ctrl+Page Up/Down → switch sheets

---

## 3. CELL EDITING

### 3.1 Inline Cell Editing [P1]
- ⬜ Double-click or F2 to enter edit mode
- ⬜ Cursor positioning within cell text
- ⬜ Text selection within cell (Shift+Arrow, double-click word, triple-click all)
- ⬜ Escape to cancel edit, Enter to confirm

### 3.2 Formula Bar [P1]
- ⬜ Shows full formula or text content of active cell
- ⬜ Editable directly (click into formula bar to edit)
- ⬜ Expandable / collapsible for long formulas (expand button)
- ⬜ Name box on left showing cell address or named range
- ⬜ Name box with autocomplete for named ranges
- ⬜ Function insert button (fx)

### 3.3 Auto-Complete [P1]
- ⬜ Column value autocomplete (suggest from existing values in column)
- ⬜ Function name autocomplete with dropdown
- ⬜ Function signature tooltip (parameters with types)
- ⬜ Named range autocomplete in formulas
- ⬜ Table column reference autocomplete (Table1[...)

### 3.4 Formula Syntax Highlighting [P1]
- ⬜ Color-coded cell references (each referenced range gets unique color)
- ⬜ Colored range highlights on the grid matching formula bar colors
- ⬜ Parenthesis matching (highlight matching pair)
- ⬜ Error/syntax highlighting (red for errors)
- ⬜ String literals in different color
- ⬜ Function names in different color

### 3.5 Cell Reference by Pointing [P1]
- ⬜ While editing formula, click a cell to insert its reference
- ⬜ Click + drag to insert range reference
- ⬜ Auto-creates relative references (adjustable with F4)
- ⬜ Works across sheets (click another sheet tab while editing)
- ⬜ Visual highlight of pointed cell/range

### 3.6 AutoFill [P1]
- ⬜ Drag fill handle to extend series
- ⬜ Number series detection (1,2,3... or 2,4,6...)
- ⬜ Date series (days, weekdays, months, years)
- ⬜ Custom text series (Mon, Tue, Wed... or Q1, Q2, Q3...)
- ⬜ Formula fill (adjust references)
- ⬜ AutoFill Options button (copy cells, fill series, fill formatting only, fill without formatting)
- ⬜ Custom lists for AutoFill (user-defined sequences)

### 3.7 Flash Fill [P1]
- ⬜ Pattern detection from examples
- ⬜ Auto-suggests when pattern detected (grey preview)
- ⬜ Ctrl+E to trigger
- ⬜ Works for: text extraction, combining, reformatting, case changes

### 3.8 Data Validation [P1]
- ⬜ **Validation types**:
  - ⬜ Whole number (between, not between, equal, not equal, greater than, less than, etc.)
  - ⬜ Decimal (same operators)
  - ⬜ List (dropdown from comma-separated or range reference)
  - ⬜ Date (same operators)
  - ⬜ Time (same operators)
  - ⬜ Text length (same operators)
  - ⬜ Custom formula (any formula returning TRUE/FALSE)
- ⬜ **Input message**: title + message shown when cell selected
- ⬜ **Error alert types**:
  - ⬜ Stop (prevents invalid entry)
  - ⬜ Warning (allows override with warning)
  - ⬜ Information (allows with info message)
- ⬜ In-cell dropdown arrow for list validation
- ⬜ Dependent/cascading dropdowns (via INDIRECT)
- ⬜ Circle invalid data (visual indicators)

---

## 4. ROW & COLUMN OPERATIONS

### 4.1 Insert / Delete [P1]
- ⬜ Insert single or multiple rows/columns
- ⬜ Insert cells and shift right or shift down
- ⬜ Delete rows/columns
- ⬜ Delete cells and shift left or shift up
- ⬜ **Auto-adjust all formula references on insert/delete** (critical)

### 4.2 Resize [P1]
- ⬜ Drag row/column border to resize
- ⬜ Double-click border to auto-fit content
- ⬜ Set exact height/width via dialog (in pixels or points)
- ⬜ Set default row height / column width for sheet
- ⬜ Resize multiple selected rows/cols simultaneously

### 4.3 Hide / Unhide [P1]
- ⬜ Hide selected rows/columns
- ⬜ Unhide via context menu
- ⬜ Unhide by selecting across hidden range and unhiding
- ⬜ Visual indicator for hidden rows/cols (double-line or gap in headers)

### 4.4 Group / Outline [P2]
- ⬜ Group selected rows or columns into collapsible sections
- ⬜ Nested groups (up to 8 levels)
- ⬜ Expand/collapse buttons (+/-)
- ⬜ Outline level buttons (1, 2, 3...)
- ⬜ Auto-outline (detect structure from formulas)
- ⬜ Subtotals with automatic grouping

---

## 5. CELL FORMATTING — NUMBER FORMATS

### 5.1 Built-in Format Categories [P1]
- ⬜ **General**: auto-detect display
- ⬜ **Number**: decimal places (0-30), thousands separator, negative number display (red, parentheses, minus, red with parentheses)
- ⬜ **Currency**: symbol selection (150+ symbols), symbol position, decimal places, negative display
- ⬜ **Accounting**: aligned currency symbols, aligned decimal points, dashes for zero
- ⬜ **Date**: 15+ formats (3/14/2026, 14-Mar-2026, March 14 2026, etc.), locale-aware
- ⬜ **Time**: h:mm, h:mm:ss, h:mm AM/PM, elapsed time [h]:mm:ss
- ⬜ **Percentage**: auto-multiply by 100, configurable decimals
- ⬜ **Fraction**: halves, quarters, eighths, sixteenths, up to 3 digits
- ⬜ **Scientific**: E+00 notation, configurable decimals
- ⬜ **Text**: treat as text even if numeric
- ⬜ **Special**: zip code, zip+4, phone, SSN (locale-dependent)

### 5.2 Custom Number Format Codes [P1]
- ⬜ **Format code structure**: `positive;negative;zero;text` (up to 4 sections)
- ⬜ **Digit placeholders**: `0` (display digit or zero), `#` (display digit or nothing), `?` (display digit or space for alignment)
- ⬜ **Thousands separator**: `,` after digit placeholder
- ⬜ **Scaling**: `,` at end divides by 1000 per comma (e.g., `#,##0,` for thousands)
- ⬜ **Literal text**: `"text"` in quotes or `\` before single character
- ⬜ **Color codes**: `[Red]`, `[Blue]`, `[Green]`, `[Black]`, `[White]`, `[Magenta]`, `[Cyan]`, `[Yellow]`, `[Color1]`-`[Color56]`
- ⬜ **Conditions**: `[>1000]#,##0;"small"` — condition-based formatting within format string
- ⬜ **Date/time codes**: `yyyy`, `mm`, `dd`, `hh`, `mm`, `ss`, `AM/PM`, `ddd` (Mon), `dddd` (Monday), `mmm` (Jan), `mmmm` (January)
- ⬜ **Fill character**: `*` followed by character to fill remaining cell width
- ⬜ **Text placeholder**: `@` represents the cell text in format string
- ⬜ **Escape**: `_` skips width of next character (used for alignment)

---

## 6. CELL FORMATTING — FONT & TEXT

### 6.1 Font Properties [P1]
- ⬜ Font family (dropdown with preview, recently used fonts at top)
- ⬜ Font size (preset sizes + custom entry, 1pt to 409pt)
- ⬜ Bold (Ctrl+B)
- ⬜ Italic (Ctrl+I)
- ⬜ Underline: single (Ctrl+U), double
- ⬜ Strikethrough (single, double on desktop)
- ⬜ Font color (theme colors, standard, recent, custom RGB/HSL picker)
- ⬜ Superscript / Subscript [P2]

### 6.2 Rich Text in Cells [P2]
- ⬜ Different formatting per character/word within same cell
- ⬜ Mixed fonts, sizes, colors, bold/italic within one cell
- ⬜ Preserved in .xlsx import/export

### 6.3 Text Alignment [P1]
- ⬜ **Horizontal**: Left, Center, Right, Fill (repeat char across width), Justify, Center Across Selection, Distributed
- ⬜ **Vertical**: Top, Middle, Bottom, Justify, Distributed
- ⬜ **Text rotation**: angle from -90° to +90°, plus vertical stacked text
- ⬜ **Indent**: increase/decrease levels (0-15)
- ⬜ **Wrap text**: auto-wrap with row height auto-adjust
- ⬜ **Shrink to fit**: auto-reduce font size to fit cell width [P2]
- ⬜ **Manual line break**: Alt+Enter within cell

---

## 7. CELL FORMATTING — BORDERS

### 7.1 Border Styles [P1]
- ⬜ 13 line styles: thin, medium, thick, double, hair, dotted, dashed, dash-dot, dash-dot-dot, medium dashed, medium dash-dot, medium dash-dot-dot, slant dash-dot
- ⬜ **Per-edge control**: set style + color independently for top, bottom, left, right
- ⬜ **Diagonal borders**: diagonal-up, diagonal-down
- ⬜ Border color: any color (theme, standard, custom RGB)
- ⬜ **Preset border buttons**: box, all borders, thick box, bottom, top+bottom, top+thick bottom, no border
- ⬜ Draw borders mode (freehand border drawing with mouse) [P3]

---

## 8. CELL FORMATTING — FILL & PATTERNS

### 8.1 Cell Fill [P1]
- ⬜ Solid fill color (theme colors, standard, recent, custom RGB/HSL)
- ⬜ No fill (transparent)

### 8.2 Pattern Fill [P3]
- ⬜ 18 pattern types (diagonal stripes, dots, grid, crosshatch, etc.)
- ⬜ Pattern foreground color + background color

### 8.3 Gradient Fill [P3]
- ⬜ Two-color gradient
- ⬜ Direction: horizontal, vertical, diagonal up, diagonal down, from center
- ⬜ Gradient stops

---

## 9. CELL OPERATIONS

### 9.1 Merge Cells [P1]
- ⬜ Merge & Center
- ⬜ Merge Across (merge each row separately in selection)
- ⬜ Merge Cells (without centering)
- ⬜ Unmerge Cells
- ⬜ Only top-left cell value is kept (warn user)
- ⬜ Merged cells affect sorting/filtering (warn or prevent)

### 9.2 Format Painter [P1]
- ⬜ Single-click: copy format, paste once
- ⬜ Double-click: continuous paste mode (keep pasting until Escape)
- ⬜ Works across sheets

### 9.3 Paste Special [P1]
- ⬜ **Paste options**:
  - ⬜ All
  - ⬜ Values only
  - ⬜ Formulas only
  - ⬜ Formats only
  - ⬜ Comments
  - ⬜ Validation
  - ⬜ All using source theme
  - ⬜ All except borders
  - ⬜ Column widths
  - ⬜ Formulas and number formats
  - ⬜ Values and number formats
- ⬜ **Operations**: None, Add, Subtract, Multiply, Divide (apply operation between clipboard and existing values)
- ⬜ **Skip blanks**: don't overwrite existing data with blank clipboard cells
- ⬜ **Transpose**: swap rows and columns on paste

### 9.4 Clear [P1]
- ⬜ Clear All
- ⬜ Clear Contents (values/formulas only)
- ⬜ Clear Formats only
- ⬜ Clear Comments
- ⬜ Clear Hyperlinks

### 9.5 Clipboard [P1]
- ⬜ Copy/Cut/Paste (Ctrl+C/X/V)
- ⬜ Marching ants animation on copied range
- ⬜ External paste detection: HTML tables, TSV, CSV, plain text, rich text
- ⬜ Smart parsing of pasted external data
- ⬜ Paste Options button (appears after paste with choices)

---

## 10. UNDO / REDO

### 10.1 Undo [P1]
- ⬜ At least 100 levels of undo
- ⬜ Undo dropdown showing action list
- ⬜ Multi-step undo (select action in list to undo everything up to that point)
- ⬜ Ctrl+Z shortcut

### 10.2 Redo [P1]
- ⬜ Redo after undo
- ⬜ Ctrl+Y shortcut
- ⬜ Redo stack clears on new action

### 10.3 Collaborative Undo [P1]
- ⬜ Undo only YOUR actions in co-authoring mode
- ⬜ Other users' actions unaffected by your undo
- ⬜ Per-user undo stack

---

## 11. SHEET / TAB MANAGEMENT

### 11.1 Sheet Operations [P1]
- ⬜ Add new sheet (+button and Shift+F11)
- ⬜ Delete sheet (with confirmation if data exists)
- ⬜ Rename (double-click tab; 31 char limit; prohibited chars: \ / : * ? [ ])
- ⬜ Reorder by drag
- ⬜ Duplicate sheet (Ctrl+drag or context menu)
- ⬜ Move/copy to another workbook

### 11.2 Tab Appearance [P1]
- ⬜ Tab color (color picker)
- ⬜ Tab scroll arrows (when many sheets)
- ⬜ Right-click tab scroll arrows → sheet list popup
- ⬜ Ctrl+Page Up/Down to switch sheets

### 11.3 Hide / Very Hidden [P2]
- ⬜ Hide sheet (unhide via context menu)
- ⬜ Very Hidden (unhide only via scripting/API; used for config sheets)
- ⬜ Protect workbook structure to prevent unhiding

---

## 12. COMPUTE ENGINE

### 12.1 Dependency Graph [P1]
- ⬜ Directed acyclic graph (DAG) tracking precedents and dependents per cell
- ⬜ Incremental updates on cell edit (only mark affected subgraph as dirty)
- ⬜ Cycle detection with user warning
- ⬜ Support for cross-sheet dependencies
- ⬜ Named range resolution in dependency tracking

### 12.2 Smart Recalculation [P1]
- ⬜ Only recalculate dirty cells and their downstream dependents
- ⬜ Skip unchanged branches
- ⬜ Dynamic calculation chain ordering (move formulas with dirty deps down the chain)
- ⬜ Volatile function handling: always recalculate NOW(), TODAY(), RAND(), RANDARRAY(), OFFSET(), INDIRECT()

### 12.3 Multi-Threaded Recalculation [P1]
- ⬜ Distribute independent formula chains across CPU cores / Web Workers
- ⬜ Thread-safe built-in functions parallelized automatically
- ⬜ Fallback to single-thread for non-safe operations
- ⬜ For web: Rust/WASM computation in Web Worker, rendering on main thread

### 12.4 Calculation Modes [P1]
- ⬜ Automatic (recalculate on every change)
- ⬜ Manual (F9 to recalculate; Ctrl+Alt+F9 for full recalc)
- ⬜ Automatic except data tables
- ⬜ Status bar indicator ("Calculate" when manual mode needs recalc)

### 12.5 Circular Reference Handling [P1]
- ⬜ Detection and warning dialog
- ⬜ Iterative calculation option (enable/disable)
- ⬜ Maximum iterations setting (1-32,767)
- ⬜ Maximum change threshold for convergence
- ⬜ Status bar indicator when circular reference exists

### 12.6 Precision [P1]
- ⬜ IEEE 754 double precision (15 significant digits)
- ⬜ Largest number: ±1.79769E+308
- ⬜ Smallest positive: 2.2251E-308
- ⬜ Date serial number system (1900 system default, 1904 system for Mac compat)

---

## 13. FORMULA SYSTEM — FUNCTIONS

> **Target: 400+ functions minimum for launch. Every function listed below with [P1] is required.**
> Functions are grouped by category. Implement in priority order.

### 13.1 Math & Trig [P1] (~70 functions)
- ⬜ SUM, SUMIF, SUMIFS, SUMPRODUCT
- ⬜ ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING, CEILING.MATH, FLOOR, FLOOR.MATH
- ⬜ INT, TRUNC, MOD, QUOTIENT
- ⬜ ABS, SIGN, POWER, SQRT, EXP, LN, LOG, LOG10
- ⬜ PI, RAND, RANDBETWEEN, RANDARRAY
- ⬜ SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2, SINH, COSH, TANH
- ⬜ RADIANS, DEGREES
- ⬜ FACT, FACTDOUBLE, COMBIN, COMBINA, PERMUT, PERMUTATIONA
- ⬜ GCD, LCM
- ⬜ PRODUCT, MULTINOMIAL
- ⬜ AGGREGATE (19 sub-functions with ignore options)
- ⬜ SUBTOTAL (11 functions, 101-111 for ignore hidden rows)
- ⬜ MMULT, MINVERSE, MDETERM, MUNIT (matrix operations)
- ⬜ SEQUENCE (dynamic array)
- ⬜ ROMAN, ARABIC
- ⬜ BASE, DECIMAL
- ⬜ SERIESSUM, SUMSQ

### 13.2 Lookup & Reference [P1] (~25 functions)
- ⬜ VLOOKUP, HLOOKUP
- ⬜ **XLOOKUP** (search_mode: first/last/binary; match_mode: exact/wildcard/approx; if_not_found)
- ⬜ INDEX (array form + reference form)
- ⬜ MATCH, **XMATCH** (search direction, match mode)
- ⬜ OFFSET (volatile! — returns reference offset from start)
- ⬜ INDIRECT (volatile! — text to reference)
- ⬜ CHOOSE, SWITCH
- ⬜ ROW, ROWS, COLUMN, COLUMNS
- ⬜ ADDRESS, AREAS
- ⬜ LOOKUP (vector + array form)
- ⬜ TRANSPOSE
- ⬜ **Dynamic Array functions** (all P1):
  - ⬜ FILTER (criteria, multiple criteria with * and +, if_empty)
  - ⬜ SORT (by_col, sort_order array for multi-level)
  - ⬜ SORTBY (sort by different array, multi-level)
  - ⬜ UNIQUE (by_col, exactly_once)
  - ⬜ VSTACK, HSTACK
  - ⬜ TAKE, DROP
  - ⬜ CHOOSECOLS, CHOOSEROWS
  - ⬜ TEXTSPLIT (col_delimiter, row_delimiter, ignore_empty, match_mode, pad_with)
  - ⬜ WRAPROWS, WRAPCOLS
  - ⬜ TOROW, TOCOL
  - ⬜ EXPAND

### 13.3 Text [P1] (~40 functions)
- ⬜ CONCATENATE, CONCAT, **TEXTJOIN** (delimiter, ignore_empty, ranges)
- ⬜ LEFT, RIGHT, MID, LEN, LENB
- ⬜ FIND, FINDB, SEARCH, SEARCHB (case-sensitive vs wildcards)
- ⬜ SUBSTITUTE (instance_num for Nth occurrence), REPLACE, REPLACEB
- ⬜ TEXT (format number as text with format code)
- ⬜ VALUE, NUMBERVALUE (decimal_separator, group_separator)
- ⬜ UPPER, LOWER, PROPER, TRIM, CLEAN
- ⬜ EXACT, REPT, CHAR, CODE, UNICODE, UNICHAR
- ⬜ T, N (convert to text/number)
- ⬜ FIXED (format number with decimals and commas)
- ⬜ DOLLAR, BAHTTEXT, YEN
- ⬜ **TEXTBEFORE, TEXTAFTER** (delimiter, instance_num, match_mode, match_end, if_not_found)
- ⬜ **TEXTSPLIT** (col_delimiter, row_delimiter) → dynamic array
- ⬜ VALUETOTEXT, ARRAYTOTEXT

### 13.4 Logical [P1] (~15 functions)
- ⬜ IF, IFS (up to 127 conditions)
- ⬜ AND, OR, NOT, XOR
- ⬜ IFERROR, IFNA
- ⬜ SWITCH (expression, value1/result1 pairs, default)
- ⬜ TRUE, FALSE
- ⬜ **LET** (name/value pairs, final calculation — eliminates repeated calculations)
- ⬜ **LAMBDA** (parameters, formula — user-defined functions)
- ⬜ **MAP** (array, LAMBDA — apply function to each element)
- ⬜ **REDUCE** (initial_value, array, LAMBDA — accumulate)
- ⬜ **SCAN** (initial_value, array, LAMBDA — running accumulation)
- ⬜ **MAKEARRAY** (rows, cols, LAMBDA)
- ⬜ **BYCOL** (array, LAMBDA), **BYROW** (array, LAMBDA)
- ⬜ **ISOMITTED** (for LAMBDA optional parameters)

### 13.5 Date & Time [P1] (~25 functions)
- ⬜ DATE, TIME, NOW (volatile), TODAY (volatile)
- ⬜ YEAR, MONTH, DAY, HOUR, MINUTE, SECOND
- ⬜ DATEVALUE, TIMEVALUE
- ⬜ **DATEDIF** (interval: "Y","M","D","MD","YM","YD") — undocumented in Excel but widely used
- ⬜ NETWORKDAYS, NETWORKDAYS.INTL (custom weekend config)
- ⬜ WORKDAY, WORKDAY.INTL
- ⬜ EDATE, EOMONTH
- ⬜ WEEKDAY (return_type options), WEEKNUM, ISOWEEKNUM
- ⬜ DAYS, DAYS360 (US vs European method)

### 13.6 Statistical [P1 core, P2 advanced] (~80 functions)
- ⬜ AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS
- ⬜ COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS
- ⬜ MAX, MAXA, MAXIFS, MIN, MINA, MINIFS
- ⬜ MEDIAN, MODE.SNGL, MODE.MULT
- ⬜ LARGE, SMALL
- ⬜ STDEV.S, STDEV.P, STDEVA, STDEVPA
- ⬜ VAR.S, VAR.P, VARA, VARPA
- ⬜ PERCENTILE.INC, PERCENTILE.EXC, PERCENTRANK.INC, PERCENTRANK.EXC
- ⬜ QUARTILE.INC, QUARTILE.EXC
- ⬜ RANK.EQ, RANK.AVG
- ⬜ CORREL, RSQ, SLOPE, INTERCEPT, STEYX
- ⬜ COVARIANCE.S, COVARIANCE.P
- ⬜ FORECAST.LINEAR, FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT [P2]
- ⬜ GROWTH, TREND, LINEST, LOGEST [P2]
- ⬜ FREQUENCY [P2]
- ⬜ **Distribution functions** [P2]: NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, CHISQ.DIST, CHISQ.INV, F.DIST, F.INV, BINOM.DIST, BINOM.INV, POISSON.DIST, EXPON.DIST, GAMMA.DIST, GAMMA.INV, BETA.DIST, BETA.INV, LOGNORM.DIST, LOGNORM.INV, WEIBULL.DIST, NEGBINOM.DIST, HYPGEOM.DIST, CONFIDENCE.NORM, CONFIDENCE.T
- ⬜ Z.TEST, T.TEST, F.TEST, CHISQ.TEST [P2]
- ⬜ DEVSQ, AVEDEV, GEOMEAN, HARMEAN, TRIMMEAN, SKEW, KURT [P2]
- ⬜ PROB, FISHER, FISHERINV, PHI, GAUSS [P3]

### 13.7 Financial [P1 core, P2 bonds] (~55 functions)
- ⬜ PMT, PPMT, IPMT, CUMIPMT, CUMPRINC
- ⬜ FV, PV, NPV, **XNPV** (irregular dates)
- ⬜ NPER, RATE
- ⬜ **IRR, XIRR, MIRR** (critical for investment analysis)
- ⬜ SLN, SYD, DB, DDB, VDB (depreciation) [P2]
- ⬜ EFFECT, NOMINAL
- ⬜ DOLLARDE, DOLLARFR
- ⬜ FVSCHEDULE, PDURATION, RRI
- ⬜ **Bond functions** [P2]: PRICE, YIELD, DURATION, MDURATION, ACCRINT, ACCRINTM
- ⬜ **T-bill functions** [P3]: TBILLEQ, TBILLPRICE, TBILLYIELD
- ⬜ **Coupon functions** [P3]: COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD
- ⬜ DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, YIELDDISC, YIELDMAT [P3]
- ⬜ AMORLINC, AMORDEGRC [P3]

### 13.8 Information [P1] (~20 functions)
- ⬜ ISBLANK, ISERROR, ISERR, ISNA, ISNUMBER, ISTEXT, ISLOGICAL, ISNONTEXT, ISREF, ISEVEN, ISODD
- ⬜ ISFORMULA
- ⬜ ERROR.TYPE, TYPE, N, NA
- ⬜ CELL (info_type, reference) [P2]
- ⬜ INFO (type_text) [P2]
- ⬜ SHEET, SHEETS [P2]
- ⬜ FORMULATEXT [P2]

### 13.9 Database [P2] (~12 functions)
- ⬜ DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMIN, DMAX, DGET, DPRODUCT, DSTDEV, DSTDEVP, DVAR, DVARP

### 13.10 Engineering [P2/P3] (~40 functions)
- ⬜ CONVERT (100+ unit conversions) [P2]
- ⬜ Base conversions: BIN2DEC, BIN2HEX, BIN2OCT, DEC2BIN, DEC2HEX, DEC2OCT, HEX2BIN, HEX2DEC, HEX2OCT, OCT2BIN, OCT2DEC, OCT2HEX [P2]
- ⬜ Complex number functions: COMPLEX, IMAGINARY, IMREAL, IMABS, IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG2, IMLOG10, IMSIN, IMCOS [P3]
- ⬜ ERF, ERF.PRECISE, ERFC, ERFC.PRECISE [P3]
- ⬜ BESSELI, BESSELJ, BESSELK, BESSELY [P3]
- ⬜ DELTA, GESTEP, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT [P3]

### 13.11 Web [P2]
- ⬜ ENCODEURL
- ⬜ WEBSERVICE (call web service URL from formula)
- ⬜ FILTERXML (XPath query on XML text)

### 13.12 Cube [P3]
- ⬜ CUBEMEMBER, CUBEVALUE, CUBESET, CUBERANKEDMEMBER, CUBESETCOUNT, CUBEMEMBERPROPERTY, CUBEKPIMEMBER

### 13.13 AI Function [P1]
- ⬜ =AI(prompt, [cell_references]) — our equivalent of Excel's COPILOT()
- ⬜ Text analysis, summarization, extraction, classification in cells
- ⬜ Batch processing when dragged/filled across rows

---

## 14. DYNAMIC ARRAY / SPILL ENGINE

### 14.1 Spill Behavior [P1]
- ⬜ Single formula returns multi-cell result
- ⬜ Automatic spill range (output fills adjacent cells)
- ⬜ **# operator**: reference entire spill range (e.g., A1# references all spilled results from A1)
- ⬜ **#SPILL! error**: when spill range is blocked by existing data
- ⬜ Spill range outlined with blue border
- ⬜ Spilled cells show formula in grey (indicating they're part of a spill)
- ⬜ Deleting the anchor formula clears all spilled cells

### 14.2 @ Implicit Intersection Operator [P2]
- ⬜ Pre-dynamic-arrays behavior for backward compatibility
- ⬜ @ prefix forces single-cell result from array formula
- ⬜ Auto-added when opening legacy workbooks with array-returning formulas

---

## 15. CELL REFERENCES & NAMING

### 15.1 Reference Types [P1]
- ⬜ A1 notation (column letter + row number)
- ⬜ Relative references (adjust on copy/fill)
- ⬜ Absolute references: $A$1 (fixed), $A1 (fixed col), A$1 (fixed row)
- ⬜ F4 key to cycle through reference modes while editing
- ⬜ Cross-sheet references: Sheet1!A1, 'Sheet Name With Spaces'!A1
- ⬜ R1C1 notation (R1C1, R[-1]C[2]) [P2]
- ⬜ 3D references: =SUM(Sheet1:Sheet12!A1) [P2]
- ⬜ External workbook references: ='[Book.xlsx]Sheet'!A1 [P3]

### 15.2 Named Ranges [P1]
- ⬜ Name a cell or range (Name Box or Define Name dialog)
- ⬜ Scope: workbook-level or sheet-level
- ⬜ Use names in formulas (=SUM(SalesData))
- ⬜ **Name Manager**: create, edit, delete, filter names
- ⬜ Name rules: start with letter or underscore, no spaces, no cell-address-like names
- ⬜ Named constants (name refers to a value, not a range) [P2]
- ⬜ Named formulas (name refers to a formula) [P2]
- ⬜ Dynamic named ranges (using OFFSET+COUNTA or INDEX) [P2]

### 15.3 Structured Table References [P1]
- ⬜ Table1[ColumnName] — entire column
- ⬜ Table1[@ColumnName] — current row intersection
- ⬜ Table1[#All], [#Data], [#Headers], [#Totals] — special selectors
- ⬜ Table1[[Col1]:[Col3]] — column range
- ⬜ Auto-complete when typing inside table formulas

---

## 16. CHARTS & VISUALIZATION

### 16.1 Chart Types [P1 = top 20, P2 = rest]

**P1 Chart Types (must have):**
- ⬜ Column (clustered, stacked, 100% stacked)
- ⬜ Bar (clustered, stacked, 100% stacked)
- ⬜ Line (with/without markers, smooth/straight)
- ⬜ Area (stacked, 100% stacked)
- ⬜ Pie, Doughnut
- ⬜ Scatter (XY) with smooth/straight lines, with/without markers
- ⬜ Combo chart (column + line on dual axes)
- ⬜ Histogram with Pareto
- ⬜ Waterfall

**P2 Chart Types:**
- ⬜ Bubble
- ⬜ Radar (with markers, filled)
- ⬜ Treemap
- ⬜ Sunburst
- ⬜ Box & Whisker
- ⬜ Funnel
- ⬜ Stock (OHLC, candlestick)
- ⬜ Map chart [P3]
- ⬜ Surface / Contour [P3]

### 16.2 Chart Customization (per chart) [P1]
- ⬜ Chart title (position, linked to cell, font/color)
- ⬜ Axis titles (horizontal + vertical)
- ⬜ **Axis formatting**: min/max bounds, major/minor units, log scale, display units (thousands/millions/billions), reverse order, number format, tick marks (inside/outside/cross/none)
- ⬜ Legend (position: top/bottom/left/right/custom, font, show/hide entries)
- ⬜ Data labels (value/category/series/percentage, position, number format, font, leader lines)
- ⬜ Gridlines (major/minor, style/color/width)
- ⬜ **Trendlines**: linear, exponential, logarithmic, polynomial (2-6), power, moving average; display equation; R²; forecast forward/backward
- ⬜ Error bars (standard error, percentage, std dev, custom range) [P2]
- ⬜ Series formatting (fill, border, gap width, overlap, series order)
- ⬜ Data table below chart [P2]
- ⬜ Plot area formatting (fill, border)

### 16.3 Sparklines [P1]
- ⬜ Line sparkline (data range, color, weight, high/low/first/last markers, marker color)
- ⬜ Column sparkline (positive/negative colors, axis settings)
- ⬜ Win/Loss sparkline [P2]
- ⬜ Sparkline groups (shared axis/formatting) [P2]

### 16.4 Conditional Formatting [P1]
- ⬜ **Highlight cells**: greater than, less than, between, equal to, text that contains, date occurring, duplicate values
- ⬜ **Top/Bottom rules**: top N items, top N%, bottom N items, bottom N%, above average, below average
- ⬜ **Data bars**: gradient/solid, positive/negative colors, bar direction, axis position, min/max type (auto/number/percent/percentile/formula), bar border
- ⬜ **Color scales**: 2-color, 3-color; custom colors for min/midpoint/max; type for each (auto/number/percent/percentile/formula)
- ⬜ **Icon sets**: 20+ sets (3-arrows, 4-arrows, 5-arrows, 3-flags, 3-traffic-lights, 4-ratings, 5-ratings, 3-shapes, 3-stars, 5-boxes, etc.); show icon only (hide value); reverse order; custom icon assignment per threshold
- ⬜ **Formula-based rules**: any formula returning TRUE/FALSE; can reference other cells; cross-row/column logic
- ⬜ **Rules Manager**: view all rules on sheet; priority ordering; stop if true; edit/delete rules; applies-to range
- ⬜ **Formatting options in rules**: font color, bold, italic, underline, strikethrough, fill color/pattern, border, number format

---

## 17. PIVOT TABLES

### 17.1 PivotTable Core [P1]
- ⬜ Create from table/range (auto-detect)
- ⬜ Field list panel with drag-and-drop to: Rows, Columns, Values, Filters
- ⬜ Aggregation: SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNT NUMBERS, STDEV, STDEVP, VAR, VARP
- ⬜ Distinct Count (if using data model)
- ⬜ **Show Values As** (12+ options): % of grand total, % of column total, % of row total, % of parent row/col, difference from, % difference from, running total, rank (small/large), index
- ⬜ **Grouping**: dates (year/quarter/month/week/day), numbers (start/end/step), custom text groups
- ⬜ Calculated fields (formulas using other fields)
- ⬜ Calculated items [P2]
- ⬜ Drill down (double-click value → detail rows)
- ⬜ Sort by value or label; ascending/descending
- ⬜ Filter: value filters, label filters, date filters, top N, manual item check/uncheck
- ⬜ Layout: compact/tabular/outline form; repeat labels; subtotals position; grand totals
- ⬜ Refresh data; change source
- ⬜ PivotTable styles gallery
- ⬜ GETPIVOTDATA formula [P2]

### 17.2 Slicers & Timelines [P1]
- ⬜ Visual slicer (button-based filter) connected to PivotTable(s)
- ⬜ Multi-select in slicer
- ⬜ Slicer search box
- ⬜ Slicer columns/size/style
- ⬜ Connect one slicer to multiple PivotTables
- ⬜ Timeline (date slicer) with zoom levels (days/months/quarters/years)
- ⬜ Table slicers (for regular tables, not just pivots) [P2]

---

## 18. SORTING & FILTERING

### 18.1 AutoFilter [P1]
- ⬜ Dropdown arrows on each header column
- ⬜ Check/uncheck values; Select All / Clear
- ⬜ Search within filter dropdown
- ⬜ Filter by cell color / font color
- ⬜ Number filters: equals, doesn't equal, greater than, between, top 10, above average, custom
- ⬜ Text filters: equals, contains, begins with, ends with, custom
- ⬜ Date filters: this week, last month, this quarter, year to date, between dates, etc.

### 18.2 Sorting [P1]
- ⬜ Sort ascending (A-Z, smallest to largest, oldest to newest)
- ⬜ Sort descending
- ⬜ Multi-level sort (up to 64 keys): Sort dialog with "Then by"
- ⬜ Sort by cell color, font color, icon
- ⬜ Custom sort lists (Mon/Tue/Wed..., custom user lists) [P2]
- ⬜ Case-sensitive sort option [P2]

### 18.3 Tables (ListObjects) [P1]
- ⬜ Create Table (Ctrl+T) with auto-detection
- ⬜ Structured references in formulas
- ⬜ Auto-expand on data entry adjacent to table
- ⬜ Total row (toggle; dropdown per column for aggregate function)
- ⬜ Table styles gallery (light/medium/dark)
- ⬜ Banded rows/columns toggle
- ⬜ Remove duplicates
- ⬜ Convert table to range

---

## 19. WHAT-IF ANALYSIS

### 19.1 Goal Seek [P1]
- ⬜ Set cell to value by changing one input cell
- ⬜ Iterative solving
- ⬜ Report found solution or failure

### 19.2 Scenario Manager [P2]
- ⬜ Define named scenarios (sets of input values)
- ⬜ Switch between scenarios
- ⬜ Summary report generation

### 19.3 Data Tables [P2]
- ⬜ 1-variable data table (one input, multiple formulas)
- ⬜ 2-variable data table (two inputs, one formula)

### 19.4 Solver [P2]
- ⬜ Objective: maximize/minimize/target value
- ⬜ Changing variable cells
- ⬜ Constraints (≤, ≥, =, integer, binary)
- ⬜ Solving methods: Simplex LP, GRG Nonlinear

---

## 20. AUTOMATION & SCRIPTING

### 20.1 Scripting Runtime [P1]
- ⬜ JavaScript/TypeScript scripting (equivalent of Office Scripts)
- ⬜ OR Python scripting (equivalent of Python in Excel)
- ⬜ Script editor with syntax highlighting and IntelliSense
- ⬜ Access to spreadsheet object model (Workbook → Sheet → Range → Cell)
- ⬜ Read/write cell values, formulas, formatting from script
- ⬜ User-Defined Functions callable from cells
- ⬜ Event triggers: on cell change, on sheet change, on workbook open

### 20.2 Macro Recording [P1]
- ⬜ Record user actions as script code
- ⬜ Play back recorded macro
- ⬜ Edit recorded script
- ⬜ Assign macro to button or keyboard shortcut

### 20.3 Automation Triggers [P1]
- ⬜ On edit / on change
- ⬜ On schedule (time-based)
- ⬜ On webhook (external event)
- ⬜ On file open
- ⬜ Manual trigger (button click)

### 20.4 Extension / Add-in SDK [P2]
- ⬜ Plugin architecture with sandboxed runtime
- ⬜ Custom side panel UI
- ⬜ Custom ribbon/toolbar buttons
- ⬜ Register custom functions
- ⬜ Marketplace for extensions

### 20.5 REST API [P1]
- ⬜ Programmatic read/write access to workbook data
- ⬜ CRUD operations on sheets, ranges, tables
- ⬜ Webhook subscriptions for change events
- ⬜ Authentication (API key, OAuth)

---

## 21. DATA ENGINE

### 21.1 Visual ETL (Power Query equivalent) [P1]
- ⬜ Step-by-step visual transformation editor
- ⬜ **Data sources**: CSV, JSON, XML, Excel, databases (via connectors), REST APIs, web pages
- ⬜ **Transforms**: remove/rename/reorder columns, change types, split/merge columns, replace values, fill down/up, extract text, add column from examples
- ⬜ **Row operations**: filter, remove duplicates, remove blank/error rows, keep top/bottom N
- ⬜ **Reshape**: pivot, unpivot, transpose, group by (with aggregations)
- ⬜ **Combine**: merge (join: inner, left, right, full, anti), append (union)
- ⬜ **Refresh**: manual, on open, scheduled
- ⬜ Applied steps sidebar (auditable trail)
- ⬜ Preview data at each step

### 21.2 Data Model (Power Pivot equivalent) [P2]
- ⬜ Multiple tables with relationships
- ⬜ Relationship types: 1:1, 1:many, many:many
- ⬜ Visual relationship diagram editor
- ⬜ Calculated measures (DAX-like formula language)
- ⬜ Calculated columns
- ⬜ Time intelligence functions

### 21.3 External Connections [P1]
- ⬜ Database connectors (MySQL, PostgreSQL, SQL Server, etc.)
- ⬜ REST API connector
- ⬜ CSV/JSON/XML file import
- ⬜ Real-time data streaming (WebSocket) [P2]
- ⬜ Linked/enriched data types (geography, stocks, custom) [P2]
- ⬜ Connection management UI (edit, refresh, credentials)

---

## 22. AI & COPILOT

### 22.1 AI Agent Mode [P1]
- ⬜ Multi-step autonomous task execution
- ⬜ Natural language task description → spreadsheet actions
- ⬜ Self-correcting: detect errors, fix, re-validate
- ⬜ Web search integration (pull live data into cells)

### 22.2 AI Cell Function [P1]
- ⬜ =AI(prompt, [cell_refs]) as a cell formula
- ⬜ Text analysis, summarization, classification, extraction
- ⬜ Batch processing when filled across rows
- ⬜ Configurable model/temperature per call [P2]

### 22.3 Formula AI [P1]
- ⬜ Formula completion on typing "="
- ⬜ Natural language to formula conversion
- ⬜ Explain any formula (step-by-step breakdown)
- ⬜ Error detection and fix suggestions

### 22.4 Data AI [P1]
- ⬜ One-click data cleaning (inconsistencies, formatting, spaces)
- ⬜ Chart/PivotTable recommendations from selected data
- ⬜ Anomaly/outlier detection
- ⬜ Trend analysis and insights

### 22.5 AI Chat Panel [P1]
- ⬜ Side panel for natural language interaction
- ⬜ Context-aware (knows selected data, sheet structure, formulas)
- ⬜ Can execute actions (insert formula, create chart, format cells)

---

## 23. COLLABORATION

### 23.1 Real-Time Co-Authoring [P1]
- ⬜ Multiple users editing simultaneously
- ⬜ Colored cursors with user names
- ⬜ Cell-level locking during active edit
- ⬜ Sub-second change propagation
- ⬜ CRDT-based conflict resolution (better than OT for offline)

### 23.2 Comments [P1]
- ⬜ Threaded comments on any cell
- ⬜ Reply in thread
- ⬜ @mention users (with notification)
- ⬜ Resolve / Reopen threads
- ⬜ Show/hide all comments
- ⬜ Navigate between comments

### 23.3 Version History [P1]
- ⬜ Automatic version creation on save
- ⬜ View previous versions
- ⬜ Restore any version
- ⬜ See who made changes + timestamps
- ⬜ Cell-level change history (Show Changes)
- ⬜ Filter changes by user, date range, cell range

### 23.4 Permissions [P1]
- ⬜ Share via link (view/edit/comment)
- ⬜ Share with specific users
- ⬜ Sheet-level protection
- ⬜ Range-level permissions (specific users can edit specific ranges)
- ⬜ Cell locking (lock/unlock per cell)
- ⬜ Formula hiding (hide formula bar content when protected)
- ⬜ External sharing with link expiration [P2]

---

## 24. FILE FORMATS

### 24.1 .xlsx Import/Export [P1] — NON-NEGOTIABLE
- ⬜ **Formulas**: all 400+ functions preserved; cell references; named ranges; structured references; array formulas; dynamic arrays
- ⬜ **Formatting**: number formats (including custom codes); fonts; colors; borders (per-edge); fills; alignment; rotation; merged cells
- ⬜ **Conditional formatting**: all rule types; data bars; color scales; icon sets; formula-based rules
- ⬜ **Charts**: all chart types; titles; axes; legends; data labels; trendlines; series formatting
- ⬜ **PivotTables**: field layout; aggregations; grouping; calculated fields; styles; cache
- ⬜ **Data validation**: all types; dropdown lists; input/error messages
- ⬜ **Images & shapes**: embedded images; position anchoring
- ⬜ **Named ranges**: workbook and sheet scope
- ⬜ **Print settings**: page setup; margins; headers/footers; page breaks; print area
- ⬜ **VBA modules**: preserve (even if not executable) — round-trip without loss
- ⬜ **Hyperlinks**: internal (cell/sheet) and external (URL)
- ⬜ **Comments**: threaded and legacy notes

### 24.2 Other Formats [P1]
- ⬜ CSV/TSV import with intelligent delimiter + encoding detection
- ⬜ CSV/TSV export
- ⬜ PDF export with print layout fidelity

### 24.3 Additional Formats [P2]
- ⬜ .xlsm import (macro-enabled; preserve macros)
- ⬜ .xlsb import (binary workbook) [P2]
- ⬜ .xls import (legacy 97-2003) [P2]
- ⬜ .ods import/export (OpenDocument) [P2]
- ⬜ JSON/XML import [P1]

---

## 25. PRINTING & PAGE LAYOUT

### 25.1 Print Preview [P2]
- ⬜ WYSIWYG preview
- ⬜ Page navigation
- ⬜ Zoom in preview

### 25.2 Page Setup [P2]
- ⬜ Paper size (A4, Letter, Legal, etc.)
- ⬜ Orientation (portrait/landscape)
- ⬜ Scaling (fit to page width/height, percentage)
- ⬜ Margins (normal/wide/narrow/custom)
- ⬜ Print area (set/clear)
- ⬜ Print titles (repeat rows at top, repeat columns at left)
- ⬜ Page breaks (manual insert/remove, page break preview mode)

### 25.3 Headers & Footers [P2]
- ⬜ Left/center/right sections
- ⬜ Insert: page number, total pages, date, time, file name, sheet name
- ⬜ Different first page
- ⬜ Different odd/even pages

---

## 26. SECURITY & ENTERPRISE

### 26.1 Protection [P1]
- ⬜ Sheet protection with configurable allowed actions
- ⬜ Workbook structure protection
- ⬜ Cell lock/unlock
- ⬜ Formula bar hiding for locked cells

### 26.2 Enterprise Features [P2]
- ⬜ SSO / SAML / SCIM
- ⬜ Role-based access control with admin console
- ⬜ Audit trail for all access and modifications
- ⬜ Data residency options
- ⬜ Sensitivity labels / DLP [P3]
- ⬜ SOC 2 / ISO 27001 compliance path [P2]

---

## FEATURE COUNT SUMMARY

| Layer | P1 | P2 | P3 | Total |
|-------|----|----|-----|-------|
| Grid & Cell UX | ~60 | ~18 | ~7 | ~85 |
| Compute Engine | ~15 | ~5 | ~2 | ~22 |
| Formula Functions | ~350 | ~80 | ~45 | ~475 |
| Dynamic Arrays/Spill | ~8 | ~2 | — | ~10 |
| References & Names | ~15 | ~8 | ~2 | ~25 |
| Charts & Cond Format | ~65 | ~30 | ~8 | ~103 |
| PivotTables & Analysis | ~50 | ~25 | ~5 | ~80 |
| Automation & Scripting | ~20 | ~10 | ~5 | ~35 |
| Data Engine (ETL) | ~25 | ~15 | ~5 | ~45 |
| AI & Copilot | ~18 | ~5 | ~2 | ~25 |
| Collaboration | ~25 | ~8 | ~2 | ~35 |
| File Formats | ~20 | ~10 | — | ~30 |
| Print & Page Layout | ~2 | ~15 | — | ~17 |
| Security & Enterprise | ~5 | ~8 | ~2 | ~15 |
| **TOTAL** | **~678** | **~239** | **~85** | **~1002** |

> **Note**: P1 features alone are ~678 individual items. This is the minimum viable
> product for a professional spreadsheet. Plan accordingly.
