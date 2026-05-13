# M1 — Close XLSX/CSV Gaps to Excel-class

Detailed week-by-week breakdown of Milestone 1 from [ROADMAP.md](ROADMAP.md).

**Goal:** Open / save / round-trip any reasonable Excel workbook (up to 2 GB) without data loss or OOM. Streaming everywhere.

**Estimated total:** 6–8 weeks for one engineer.

---

## Week 1 — True SAX streaming reader: foundation

**Why first:** Current `importFromFileStreaming` loads full ZIP into memory and OOMs at ~100 MB. This is the biggest scalability lie in the codebase.

**Tasks**
1. Add `QuaZip` (LGPL) or `minizip-ng` as a vendored dependency. (Qt's internal QZipReader buffers fully — replace it.)
2. Refactor `XlsxService::importFromFileStreaming` to open the XLSX as a stream and lazy-read individual zip entries.
3. Replace `QXmlStreamReader` usage on `QByteArray` with `QXmlStreamReader` on `QIODevice` from the unzip stream.
4. Yield to UI thread every 10K rows via `QCoreApplication::processEvents()` so progress callback isn't just cosmetic.

**Deliverables**
- New file: `native/src/services/XlsxStreamReader.cpp` (consider extracting from XlsxService.cpp for testability).
- Replace [XlsxService.cpp:1411](native/src/services/XlsxService.cpp#L1411) (`importFromFileStreaming`).

**Acceptance**
- [ ] 500 MB XLSX opens without RAM > 800 MB at any point
- [ ] Progress callback fires within 200 ms of first row
- [ ] No regressions in `test_e2e.cpp` XLSX tests

---

## Week 2 — True SAX streaming writer + memory ceiling

**Tasks**
1. Refactor `XlsxService::exportToFile` to write each sheet incrementally to a zip output stream — never build the full `QByteArray` of all sheets simultaneously.
2. Use ColumnStore's `scanColumnValues` template to iterate per-chunk; write rows as they're scanned.
3. Limit intermediate buffer to ≤ 64 MB at any time.
4. Add memory ceiling self-test in `test_e2e.cpp`.

**Deliverables**
- Modify [XlsxService.cpp:1577-1817](native/src/services/XlsxService.cpp) (`generateSheet`, `exportToFile`).

**Acceptance**
- [ ] Export 20M-row 5-column workbook (~5 GB raw data) — RSS ≤ 1.5 GB throughout
- [ ] Output file opens cleanly in Excel
- [ ] Round-trip diff: open in Excel, save, reopen in Nexel → semantic equivalence

---

## Week 3 — Sheet protection, frozen panes, hidden rows/cols

**Why grouped:** All three are small workbook-XML additions; together they make most Excel files visually correct.

**Tasks**
1. **Sheet protection**
   - Parse `<sheetProtection>` element: `sheet`, `objects`, `scenarios`, `formatCells`, `formatColumns`, `formatRows`, `insertRows`, `insertColumns`, `deleteRows`, `deleteColumns`, `sort`, `autoFilter`, `pivotTables`, `password` (hashed).
   - Export same element from `Spreadsheet::isProtected()` state.
   - Add `Spreadsheet::isProtectedField(field)` queries.
2. **Frozen panes**
   - Parse `<sheetView><pane>` with `xSplit`, `ySplit`, `topLeftCell`, `activePane`, `state`.
   - Export same.
   - Add `Spreadsheet::setFrozenPanes(int row, int col)` (storage only — rendering in M2).
3. **Hidden rows/columns**
   - Parse `hidden="1"` on `<row>` and `<col>` elements.
   - Export same.
   - Add `Spreadsheet::setRowHidden(row, bool)`, `setColumnHidden(col, bool)`.

**Deliverables**
- Modify [XlsxService.cpp](native/src/services/XlsxService.cpp) parseSheet + generateSheet paths.
- New fields in [Spreadsheet.h](native/src/core/Spreadsheet.h).
- Tests in [tests/test_e2e.cpp](native/tests/test_e2e.cpp).

**Acceptance**
- [ ] 10 reference workbooks with sheet protection round-trip without losing flags
- [ ] Frozen-pane workbooks round-trip
- [ ] Hidden-row workbooks round-trip (rendering deferred to M2)
- [ ] Runtime enforcement of sheet protection is deferred to M5 — M1 only parses/exports

---

## Week 4 — Named ranges, formula-based conditional formatting, custom themes

**Tasks**
1. **Named ranges**
   - Parse `<workbook><definedNames><definedName name="...">A1:B10</definedName></definedNames>`.
   - Export from `Spreadsheet::getNamedRange*` state.
   - Make FormulaEngine resolve named refs at parse time.
2. **Formula-based conditional formatting**
   - Currently [XlsxService.cpp:995](native/src/services/XlsxService.cpp#L995) calls `skipCurrentElement()` — actually parse `<cfRule type="expression">` and its `<formula>` child.
   - Export same.
3. **Custom theme XML**
   - Parse `xl/theme/theme1.xml` for theme colors (dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink).
   - Wire into DocumentTheme.

**Acceptance**
- [ ] Named ranges work in formulas after import (`=SUM(SalesData)` resolves)
- [ ] Conditional formatting formula rules visible and round-trip
- [ ] Theme colors render correctly post-import

---

## Week 5 — OOXML tables, print settings, hyperlinks UI

**Tasks**
1. **Tables from OOXML**
   - Parse `xl/tables/table*.xml` and `<tablePart>` relationships. Currently only parsed via Nexel's own JSON metadata.
   - Export same — write proper `<table>` part, not only JSON sidecar.
   - Wiring to existing `SpreadsheetTable` struct in [Spreadsheet.h](native/src/core/Spreadsheet.h).
2. **Print settings** (parse + export only; UI in M3)
   - `<pageSetup>`, `<pageMargins>`, `<printOptions>`, `<headerFooter>`, `<rowBreaks>`, `<colBreaks>`, `<oddHeader>`, `<oddFooter>`.
3. **Hyperlinks UI**
   - New dialog: `HyperlinkDialog` (Ctrl+K). Two tabs: External URL, This Document.
   - Wire Insert menu → Hyperlink.
   - Render hyperlinks in cell as blue + underlined when cell has hyperlink.
   - Click handler on cells with hyperlink: open URL or navigate to cell.

**Acceptance**
- [ ] Excel tables round-trip from OOXML (not only Nexel JSON)
- [ ] Print settings preserved through round-trip
- [ ] Insert Hyperlink dialog works for both URL and internal targets
- [ ] Click on hyperlinked cell opens URL / navigates

---

## Week 6 — OOXML comments, advanced chart types

**Tasks**
1. **Comments / threaded comments**
   - Parse `xl/comments*.xml` (legacy) and `xl/threadedComments/threadedComment*.xml` (modern).
   - Export both.
   - Backend: `CellComment` storage in ColumnStore (may already exist — verify).
   - UI: red triangle indicator, hover popup, right-click → Insert Comment.
2. **Advanced chart types**
   - Waterfall, histogram, combo (multi-type), trendlines (linear, log, polynomial, power, exp, moving average), error bars (fixed/percentage/custom), data labels with position (inside/outside/center/best), legend position (right/top/bottom/left/none).
   - Both parse and export.
   - Wire into ChartDialog UI.

**Acceptance**
- [ ] Comments survive round-trip
- [ ] All chart types in [EXCEL_FEATURES.md](EXCEL_FEATURES.md) P1 list render correctly
- [ ] Chart trendlines visible after round-trip

---

## Week 7 — Images / shapes / sparklines as OOXML

**Tasks**
1. **Images**
   - Parse `xl/media/image*.png`/`.jpg`/`.gif` from zip stream.
   - Parse `xl/drawings/drawing*.xml` `<xdr:pic>` anchors.
   - Export images back: write media files, drawing XML, relationships.
   - Render images as floating widgets on grid (existing ImageWidget may help).
2. **Shapes**
   - Parse `<xdr:sp>` shape elements with geometry (rect, ellipse, arrow, etc.).
   - Export same.
   - Wire into ShapeWidget rendering.
3. **Sparklines as OOXML**
   - Currently sparklines exist as Nexel JSON. Parse and export `x14:sparklineGroups` extension element.
   - Backend already exists in [SparklineConfig.h](native/src/core/SparklineConfig.h).

**Acceptance**
- [ ] Embedded images round-trip (visually identical)
- [ ] Basic shapes (rect, ellipse, arrow) round-trip
- [ ] Sparklines round-trip from real Excel files (not only Nexel-saved)

---

## Week 8 — CSV polish, integration tests, hardening

**Tasks**
1. **CSV polish**
   - Date format detection: ISO-8601, US (MM/DD/YYYY), EU (DD/MM/YYYY). Locale via QLocale.
   - Locale-aware numbers: `1.234,56` (EU) vs `1,234.56` (US).
   - Windows-1252 fallback detection (no BOM, non-UTF-8 byte found).
   - Preserve formulas on CSV round-trip when source had formulas (write formula text with `=` prefix; Excel reads this correctly).
2. **Integration test suite**
   - Build a `tests/xlsx_corpus/` directory with 50 reference workbooks covering all features above.
   - Add `test_xlsx_corpus.cpp` that opens each, exports, reopens, asserts semantic equivalence.
3. **Hardening**
   - Fuzz the XLSX parser with malformed inputs (truncated zips, malformed XML, gigantic shared-string tables, etc.). Should never crash.
   - Add OOM guards for files that exceed allowed memory.
   - Logging of silently-ignored elements (so we know what we still skip).

**Acceptance**
- [ ] All 50 corpus workbooks pass round-trip semantic equivalence
- [ ] Fuzz suite runs 10 minutes without crash
- [ ] CSV files in 5 locales import correctly with dates and numbers preserved
- [ ] Skipped-element log empty on all 50 corpus workbooks

---

## Definition of M1 done

- [ ] All weekly acceptance criteria met
- [ ] [EXCEL_FEATURES.md](EXCEL_FEATURES.md) status updated for every XLSX-related item
- [ ] No new TODOs introduced inside XlsxService/CsvService
- [ ] [tests/test_e2e.cpp](native/tests/test_e2e.cpp) coverage of XLSX expanded by 1,000+ lines
- [ ] User test: open 5 power-user-grade real workbooks, all open without warnings and round-trip cleanly

## Risks

| Risk | Mitigation |
|------|------------|
| QuaZip / minizip licensing | Use BSD-licensed minizip-ng; if forbidden, fall back to Qt QZip with chunked extraction |
| Hidden assumptions in existing parseSheet | Add corpus test before refactoring; refactor in small commits |
| Streaming writer breaks chart export | Defer streaming for sheets with charts to Week 6 (rewrite once chart export is settled) |
| Sparkline OOXML extension is poorly documented | Reference LibreOffice source as fallback spec |

## Out of scope for M1 (moved to later milestones)

- Print Preview UI (M3)
- Custom grid rendering of frozen panes (M2)
- Runtime enforcement of sheet protection (M5)
- Pivot caches round-trip (M4 — depends on full pivot rewrite)
- VBA / macros (deferred indefinitely — security concerns)
