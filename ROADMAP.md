# Nexel Pro — Roadmap to Beat Excel

## Mission

Beat Microsoft Excel as a native desktop spreadsheet across scalability, performance, feature depth, quality, and UI/UX. The native Qt/C++ app is the primary product.

## Current state (verified 2026-05-13)

**Strengths**
- ColumnStore: columnar storage with 64K-row chunks, sparse presence bitmaps, per-column shared mutexes. 20M-row load is real.
- 353 formula functions across Math, Text, Logical, Lookup, Date/Time, Statistical, Financial, Engineering.
- XlsxService: 3,406 lines — handles cell values, formulas (text only, not re-evaluated), styles, merged cells, column widths/row heights, basic conditional formatting, data validation, hyperlinks, basic chart types.
- CsvService: 400 lines — RFC 4180, UTF-8/UTF-16, delimiter auto-detect, true streaming with chunked import.
- Tests: ~7,700 lines across behavior/regression/e2e/ui.
- UI shell: MainWindow ~7K lines, ribbon toolbar, formula bar, FormatCells/Chart/CondFormat/DataValidation dialogs wired up.

**Gaps that disqualify the product from competing with Excel today**
- XLSX streaming is fake — loads full ZIP into memory. OOMs at ~100 MB files.
- XLSX import/export missing: named ranges, sheet protection, frozen panes, hidden rows/cols, images, shapes, sparklines (as OOXML), tables (only via Nexel JSON), pivot caches, comments, print settings, custom theme XML, formula-based conditional formatting, advanced chart types (waterfall, histogram, treemap, sunburst, combo, trendlines, error bars).
- Grid still uses QTableView — not a true virtual canvas. Scrolling/selection on 20M rows is bottlenecked.
- Recalc parallelism is incomplete: `TODO: Add per-thread AST pools for true parallel recalc` in Spreadsheet.cpp.
- No printing, no page setup, no PDF export.
- Sheet protection backend exists but is not enforced in setCellValue.
- Structured references (`Table1[Col]`) not parsed by formula engine.
- Hyperlinks have backend but no UI to insert/edit.
- GoToSpecialDialog is empty.
- Pivot tables: only Sum/Count/Avg/Min/Max/CountDistinct. No calculated fields, no grouping, no slicers, no drill-down, no GETPIVOTDATA.
- Client (React) is ~358 lines of prototype. Server is a skeleton.
- Test coverage is ~2–3% of source.

## Discipline rules during this roadmap

1. No new formulas (353 is enough; gap to Excel's ~505 is irrelevant until M1–M6 are done).
2. No real-time collaboration work.
3. No AI features.
4. No client (React) or server work.
5. No new chart types added until existing types are Excel-exact.
6. Every feature must round-trip through XLSX before being called "done."
7. Every UI dialog must pass a side-by-side screenshot diff against Excel 365 before being called "done."

## Milestones

### M1 — XLSX/CSV interop becomes Excel-class

**Why first:** The product already imports/exports basic Excel files, but cannot handle real workbooks (large files OOM, sheet protection lost, charts limited, no images/shapes). Without trustworthy interop, no other improvement matters.

**Sub-scope (each must be done):**

- **True SAX streaming XLSX reader** — read directly from `QuaZip`/`QZipReader` stream into ColumnStore, no full-XML-in-memory. Target: 2 GB XLSX without OOM on 8 GB machine.
- **SAX streaming XLSX writer** — write per-sheet incrementally; ColumnStore chunk-iterate; no intermediate `QByteArray` ≥ 256 MB.
- **Sheet protection** — parse `<sheetProtection>` on import, export on save, enforce in `Spreadsheet::setCellValue`.
- **Frozen panes** — parse/export `<sheetView><pane>` and selection, render in grid.
- **Hidden rows/columns** — parse `hidden="1"` on rows/cols, render correctly.
- **Named ranges** — parse `<workbook><definedNames>`, export back, expose to formula engine.
- **Formula-based conditional formatting** — currently skipped; parse and export `<cfRule type="expression">`.
- **Print settings / page setup** — parse/export `<pageSetup>`, `<pageMargins>`, `<printOptions>`, `<headerFooter>`, `<rowBreaks>`, `<colBreaks>`. (UI for editing them is M3.)
- **Custom theme XML** — parse `theme1.xml` for theme colors, fonts.
- **Tables from OOXML** — parse `<table>` elements (not only Nexel JSON metadata).
- **Hyperlinks UI** — Insert > Hyperlink dialog, Ctrl+K, blue-underline render, click handler.
- **OOXML comments** — parse/export `<comment>` and threaded comments (modern Excel 365 format).
- **Advanced chart types in OOXML** — waterfall, histogram, combo (multi-type), trendlines, error bars, data labels with position, legend position.
- **Images** — parse/export embedded images via DrawingML; render as floating widgets (placement, anchor, resize).
- **CSV polish** — date format detection, locale-aware numbers, Windows-1252 fallback, preserve formulas on export when sheet is round-tripped to CSV.

**Acceptance criteria for M1:**
- [ ] Open 50 reference real-world workbooks (sourced from Excel templates + public datasets) without data loss
- [ ] Round-trip diff: open → save → reopen → diff (semantic equivalence)
- [ ] 1 GB XLSX opens in < 60 s on M-series Mac without OOM
- [ ] Sheet protection blocks edits to locked cells when sheet is protected
- [ ] Named ranges work in formulas after import
- [ ] Conditional formatting formula rules survive round-trip
- [ ] Hyperlinks insertable from UI

**Estimated effort:** 6–8 weeks for one engineer.

### M2 — Custom virtual grid (true 20M-row interactivity)

**Why next:** Storage is already columnar; rendering is the bottleneck. QTableView materializes model rows. Replace with custom QAbstractScrollArea + QPainter canvas.

**Sub-scope:**
- Custom GridCanvas widget with QPainter, draws only visible viewport (~30 × 100 cells) directly from ColumnStore
- Frozen panes integration with M1
- Merged cell rendering, row/column header rendering, selection overlay, edit overlay, in-place editor
- Hit-testing for clicks, range-drag, fill handle, double-click auto-fit column
- Pixel-perfect Excel grid: 1px border, 80px default column width, 18px default row height
- Smooth wheel scrolling (Excel scrolls 3 rows; macOS pixel-precise)
- Cell tooltip on hover for truncated content

**Acceptance criteria:**
- [ ] Ctrl+End scrolls to last cell in 20M-row sheet in < 100 ms
- [ ] Wheel scroll at 60 fps with 1M visible-range cells styled
- [ ] Select 1M-cell range in < 50 ms
- [ ] Resize column over wide range without jank

**Estimated effort:** 6–8 weeks.

### M3 — Print, Page Setup, PDF export

**Why next:** Every Excel user prints. Without this, the product is a toy. M1 already parses print settings from XLSX; M3 builds the UI and engine.

**Sub-scope:**
- Print Preview window with multi-page navigation
- Page Setup dialog: margins, orientation, scale-to-fit (% or N pages wide × tall), header/footer tokens (`&P`, `&N`, `&D`, `&F`, `&A`), print titles (repeat rows/cols), print area, gridlines on/off, comments on/off
- Page break preview mode (interactive — drag to move breaks)
- PDF export via QPrinter
- Excel-exact margin defaults: 0.7" top/bottom, 0.7" left/right, 0.3" header/footer

**Acceptance criteria:**
- [ ] PDF round-trips visually to Excel's PDF output for 10 reference workbooks
- [ ] Header/footer tokens match Excel
- [ ] Scale-to-fit page math matches Excel

**Estimated effort:** 3–4 weeks.

### M4 — Pivot tables, done properly

**Why next:** Pivots are half the analytic value of a spreadsheet. Current implementation has the config struct but no UI or advanced math.

**Sub-scope:**
- Field list panel (Rows / Columns / Values / Filters drag-drop)
- Aggregations: Sum, Count, Average, Min, Max, StDev, StDevP, Var, VarP, CountDistinct, Product, Median
- Value field settings: Show Values As — % of Total, % of Row, % of Column, % of Parent, Running Total, Rank, Difference From
- Calculated fields and calculated items
- Grouping: dates (year/quarter/month/day), numbers (bins), text (manual groups)
- Drill-down (double-click → new sheet with source rows)
- Slicers (visual filter buttons)
- Timeline (date slicer)
- GETPIVOTDATA formula
- Round-trip pivot through XLSX (depends on M1 pivot cache work being extended)

**Acceptance criteria:**
- [ ] Pivot 1M-row source in < 2 s
- [ ] Slicer click filters pivot in < 200 ms
- [ ] All Show Values As modes match Excel output

**Estimated effort:** 5–7 weeks.

### M5 — Close the "marked done but lying" gaps

**Why next:** These are cheap individually but each one breaks user trust.

**Sub-scope:**
1. **Structured references** (`Table1[Col]`, `Table1[@Col]`, `[#Headers]`, `[#Totals]`, `[#All]`, `[#This Row]`) in FormulaParser. Auto-expand when table grows.
2. **Sheet protection enforcement** in `Spreadsheet::setCellValue`, insert/delete row/col, format changes. (M1 handles parse/export; M5 handles runtime enforcement.)
3. **Go To Special** — fully implement GoToSpecialDialog (constants, formulas, blanks, errors, comments, conditional formats, data validation, current region, last cell, visible only).
4. **Rich text in cells** — per-run character formatting (bold/italic/color/font within one cell). Requires Cell to hold `vector<TextRun>`.
5. **Comments / threaded notes** (UI + storage; M1 handles XLSX round-trip).
6. **Multi-level sort dialog** wired to existing `sortRangeMulti()`.
7. **True parallel recalc** — implement per-thread AST pools so level-parallel works on large levels.

**Acceptance criteria:**
- [ ] Each item passes a 5-test Excel-parity checklist
- [ ] All added to test_behavior.cpp

**Estimated effort:** 5–6 weeks.

### M6 — UI/UX polish pass

**Why last:** Polish on top of incomplete features is wasted.

**Sub-scope:**
- Unified SVG icon set (16/20/24/32px) — replace whatever is there
- Typography: Segoe UI Variable on Windows, SF Pro on macOS, Inter fallback. Match Excel's font scale (9pt menus, 11pt cells, 14pt section headers).
- Spacing tokens: 4/8/12/16/24 px grid. Audit every dialog.
- Ribbon spacing matches Excel: large-button row 70px, small-button group 22px row. Contextual tabs (Chart Design, PivotTable Analyze) appear only when relevant.
- Microinteractions: row/col resize cursors + live-preview line, fill handle, marching ants on copy, zoom slider, scroll bar tooltips with row number.
- Dark mode parity.
- Empty states for every dialog.
- Top 80 Excel keyboard shortcuts.

**Acceptance criteria:**
- [ ] Side-by-side screenshot diff vs Excel 365 for ribbon, formula bar, FormatCells, ChartDesign, PivotFieldList, ConditionalFormatManager. Visual difference index < 10%.
- [ ] User testing: 5 Excel power users, ≥ 4/5 say "feels professional."

**Estimated effort:** 4–6 weeks.

## Total

~29–39 weeks (7–10 months) with one engineer. ~4–5 months with two.

## Non-goals during this window

- No new formulas
- No real-time collaboration
- No AI features
- No web client work
- No new chart types until M1's advanced types ship

The discipline: **finish what you have before adding what you don't.**
