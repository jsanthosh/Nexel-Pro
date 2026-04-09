#include "Spreadsheet.h"
#include "PivotEngine.h"
#include "StringPool.h"
#include <algorithm>
#include <numeric>
#include <mutex>
#include <QElapsedTimer>
#include <QDate>
#include <QRegularExpression>
#include <QtConcurrent>

// Cross-platform bit intrinsics
#ifdef _MSC_VER
#include <intrin.h>
static inline int ctzll_s(uint64_t x) {
    unsigned long idx;
    _BitScanForward64(&idx, x);
    return static_cast<int>(idx);
}
#else
static inline int ctzll_s(uint64_t x) { return __builtin_ctzll(x); }
#endif

// Helper: adjust formula cell references for row/column insert/delete
// mode: 'R' = row shift, 'C' = column shift
// atIndex: 0-based row or column where insert/delete happens
// delta: positive for insert, negative for delete
static QString shiftFormulaRefs(const QString& formula, char mode, int atIndex, int delta) {
    // Quick bail-out: if formula has no letter characters, it can't contain cell references
    bool hasAlpha = false;
    for (int i = 0; i < formula.size(); ++i) {
        if (formula[i].isLetter()) { hasAlpha = true; break; }
    }
    if (!hasAlpha) return formula;

    // Pre-compiled regex (static local — compiled once, reused across calls)
    static QRegularExpression cellRefRe("(\\$?)([A-Za-z]+)(\\$?)(\\d+)");
    QString result;
    int lastEnd = 0;
    bool hasRefError = false;
    auto it = cellRefRe.globalMatch(formula);
    while (it.hasNext()) {
        auto match = it.next();
        result += formula.mid(lastEnd, match.capturedStart() - lastEnd);

        bool colAbsolute = !match.captured(1).isEmpty();
        QString colLetters = match.captured(2);
        bool rowAbsolute = !match.captured(3).isEmpty();
        int rowNum = match.captured(4).toInt(); // 1-based

        bool refDeleted = false;

        if (mode == 'R') {
            // rowNum is 1-based, atIndex is 0-based
            if (delta < 0) {
                // Deletion: check if reference points to a deleted row
                int deletedStart = atIndex + 1; // 1-based start of deleted range
                int deletedEnd = atIndex - delta; // 1-based end (exclusive), since delta is negative
                if (rowNum >= deletedStart && rowNum < deletedEnd) {
                    refDeleted = true;
                } else if (rowNum >= deletedEnd) {
                    rowNum += delta;
                }
            } else if (rowNum > atIndex) {
                rowNum += delta;
            }
        } else if (mode == 'C') {
            // Shift column references at or after atIndex
            int colIdx = 0;
            for (QChar ch : colLetters)
                colIdx = colIdx * 26 + (ch.toUpper().toLatin1() - 'A');
            if (delta < 0) {
                // Deletion: check if reference points to a deleted column
                int deletedStart = atIndex;
                int deletedEnd = atIndex - delta; // since delta is negative
                if (colIdx >= deletedStart && colIdx < deletedEnd) {
                    refDeleted = true;
                } else if (colIdx >= deletedEnd) {
                    colIdx += delta;
                    if (colIdx < 0) colIdx = 0;
                    bool wasUpper = !colLetters.isEmpty() && colLetters[0].isUpper();
                    colLetters.clear();
                    int c = colIdx + 1;
                    while (c > 0) {
                        colLetters = QChar((wasUpper ? 'A' : 'a') + (c - 1) % 26) + colLetters;
                        c = (c - 1) / 26;
                    }
                }
            } else if (colIdx >= atIndex) {
                colIdx += delta;
                if (colIdx < 0) colIdx = 0;
                bool wasUpper = !colLetters.isEmpty() && colLetters[0].isUpper();
                colLetters.clear();
                int c = colIdx + 1; // convert back to 1-based for letter encoding
                while (c > 0) {
                    colLetters = QChar((wasUpper ? 'A' : 'a') + (c - 1) % 26) + colLetters;
                    c = (c - 1) / 26;
                }
            }
        }

        if (refDeleted) {
            result += "#REF!";
            hasRefError = true;
        } else {
            result += (colAbsolute ? "$" : "") + colLetters + (rowAbsolute ? "$" : "") + QString::number(rowNum);
        }
        lastEnd = match.capturedEnd();
    }
    result += formula.mid(lastEnd);
    return result;
}

Spreadsheet::Spreadsheet()
    : m_sheetName("Sheet1"), m_rowCount(1000), m_columnCount(256),
      m_autoRecalculate(true), m_inTransaction(false) {
    m_formulaEngine = std::make_unique<FormulaEngine>(this);
    m_conditionalFormatting.setFormulaEvaluator([this](const QString& formula) -> QVariant {
        return m_formulaEngine->evaluate(formula);
    });
}

Spreadsheet::~Spreadsheet() = default;

CellProxy Spreadsheet::getCell(const CellAddress& addr) {
    return getCell(addr.row, addr.col);
}

CellProxy Spreadsheet::getCell(int row, int col) {
    // Ensure the column exists in the store (getOrCreateColumn creates it)
    m_columnStore.getOrCreateColumn(col);
    if (row > m_cachedMaxRow || col > m_cachedMaxCol) {
        m_maxRowColDirty = true;
    }
    return CellProxy(&m_columnStore, row, col);
}

CellProxy Spreadsheet::getCellIfExists(const CellAddress& addr) const {
    return getCellIfExists(addr.row, addr.col);
}

CellProxy Spreadsheet::getCellIfExists(int row, int col) const {
    if (!m_columnStore.hasCell(row, col)) {
        return CellProxy(); // invalid proxy — operator bool() returns false
    }
    return CellProxy(const_cast<ColumnStore*>(&m_columnStore), row, col);
}

QVariant Spreadsheet::getCellValue(const CellAddress& addr) {
    if (!m_columnStore.hasCell(addr.row, addr.col)) return QVariant();
    CellProxy proxy(&m_columnStore, addr.row, addr.col);
    return proxy.getValue();
}

void Spreadsheet::setCellValue(const CellAddress& addr, const QVariant& value) {
    bool wasNew = !m_columnStore.hasCell(addr.row, addr.col);

    // Clear any existing spill range from this cell (was previously a spill-producing formula)
    clearSpillRange(addr);

    // Set the value in ColumnStore
    m_columnStore.setCellValue(addr.row, addr.col, value);

    if (addr.row > m_cachedMaxRow || addr.col > m_cachedMaxCol) {
        m_maxRowColDirty = true;
    }

    // Remove from formula tracking (it's now a value cell)
    m_formulaCells.erase(CellKey{addr.row, addr.col});

    // Incremental nav index update instead of full rebuild
    if (!m_navIndexDirty && wasNew) {
        navIndexInsert(addr.row, addr.col);
    } else if (m_autoRecalculate) {
        m_navIndexDirty = true;
    }

    // Track dirty cells during batch update
    if (m_inBatchUpdate) {
        m_batchDirtyCells.insert(CellKey{addr.row, addr.col});
        return; // defer recalculation until endBatchUpdate()
    }

    // Skip dependency graph work when autoRecalculate is off (bulk import mode)
    if (m_autoRecalculate) {
        m_depGraph.removeDependencies(addr);
        if (!m_inTransaction) {
            recalculateDependents(addr);
            // Recalculate formulas with column references (e.g., =SUM(A:A))
            recalculateColumnDependents(addr.col);
        }
    }
}

void Spreadsheet::setCellFormula(const CellAddress& addr, const QString& formula) {
    QElapsedTimer _timer; _timer.start();
    bool wasNew = !m_columnStore.hasCell(addr.row, addr.col);

    // Set formula in ColumnStore
    m_columnStore.setCellFormula(addr.row, addr.col, formula);

    if (addr.row > m_cachedMaxRow || addr.col > m_cachedMaxCol) {
        m_maxRowColDirty = true;
    }

    // Track as formula cell
    m_formulaCells.insert(CellKey{addr.row, addr.col});

    // Incremental nav index update instead of full rebuild
    if (!m_navIndexDirty && wasNew) {
        navIndexInsert(addr.row, addr.col);
    } else if (!m_navIndexDirty) {
        // Existing cell changed type, nav index still valid (same position)
    } else {
        m_navIndexDirty = true;
    }

    // Clear any existing spill range from this cell before re-evaluation
    clearSpillRange(addr);

    // Single evaluation: compute result AND extract dependencies in one pass
    QVariant result = m_formulaEngine->evaluate(formula);

    // Update dependency graph from cached deps (no re-evaluation needed)
    m_depGraph.removeDependencies(addr);
    for (auto& [col, addrs] : m_colRefFormulas) {
        addrs.erase(std::remove(addrs.begin(), addrs.end(), addr), addrs.end());
    }
    for (const auto& dep : m_formulaEngine->getLastDependencies()) {
        m_depGraph.addDependency(addr, dep);
    }
    // Add range-level dependencies for large ranges (enables circular detection)
    for (const auto& range : m_formulaEngine->getLastRangeArgs()) {
        long long cellCount = static_cast<long long>(range.getRowCount()) * range.getColumnCount();
        if (cellCount >= 10000) {
            m_depGraph.addRangeDependency(addr, range);
        }
    }
    for (int col : m_formulaEngine->getLastColumnDeps()) {
        m_colRefFormulas[col].push_back(addr);
    }

    if (m_depGraph.hasCircularDependency(addr)) {
        m_columnStore.setComputedValue(addr.row, addr.col, QVariant("#CIRCULAR!"));
        return;
    }

    m_columnStore.setComputedValue(addr.row, addr.col, result);

    // Check if result is a 2D array (dynamic array / spill)
    if (result.typeId() == QMetaType::QVariantList) {
        QVariantList outerList = result.toList();
        if (!outerList.isEmpty() && outerList[0].typeId() == QMetaType::QVariantList) {
            // Convert QVariantList of QVariantList to vector<vector<QVariant>>
            std::vector<std::vector<QVariant>> array2D;
            for (const auto& rowVar : outerList) {
                QVariantList rowList = rowVar.toList();
                std::vector<QVariant> row;
                row.reserve(rowList.size());
                for (const auto& v : rowList) {
                    row.push_back(v);
                }
                array2D.push_back(std::move(row));
            }
            applySpillResult(addr, array2D);
            // Set the formula cell's computed value to the top-left element
            if (!array2D.empty() && !array2D[0].empty()) {
                m_columnStore.setComputedValue(addr.row, addr.col, array2D[0][0]);
            }
        }
    }

    qDebug() << "[setCellFormula] before recalcDeps:" << _timer.elapsed() << "ms";
    if (m_autoRecalculate && !m_inTransaction) {
        recalculateDependents(addr);
    }
    qDebug() << "[setCellFormula] TOTAL:" << _timer.elapsed() << "ms";
}

void Spreadsheet::fillRange(const CellRange& range, const QVariant& value) {
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            setCellValue(CellAddress(r, c), value);
        }
    }
}

void Spreadsheet::clearRange(const CellRange& range) {
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            m_columnStore.removeCell(r, c);
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

std::vector<CellProxy> Spreadsheet::getRange(const CellRange& range) {
    std::vector<CellProxy> result;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            result.push_back(getCell(r, c));
        }
    }
    return result;
}

void Spreadsheet::insertRow(int row, int count) {
    int maxCol = m_columnStore.maxCol();
    int totalRows = getRowCount();

    // FAST PATH for large datasets: O(chunks + one chunk's cells).
    // 1. Chunks entirely below insertion point: shift baseRow
    // 2. Chunk containing insertion point: shift cells within it, handle overflow
    // 3. Chunks entirely above: untouched
    // FAST PATH for large datasets: O(chunks) time.
    // Strategy: for the chunk containing the insert point, SPLIT it.
    // Everything before insertOffset stays in the original chunk.
    // Everything at/after insertOffset moves to a NEW chunk with baseRow = row + count.
    // All subsequent chunks just shift baseRow by +count.
    // No cascading overflow. No intra-chunk cell movement needed.
    if (totalRows > 100000) {
        for (int c = 0; c <= maxCol; ++c) {
            Column* col = m_columnStore.getColumn(c);
            if (!col) continue;
            auto& chunks = col->chunks();

            for (size_t ci = 0; ci < chunks.size(); ++ci) {
                auto& chunk = chunks[ci];
                if (chunk->baseRow >= row) {
                    // Chunk entirely at/after insertion — just shift baseRow
                    chunk->baseRow += count;
                } else if (chunk->baseRow + ColumnChunk::CHUNK_SIZE > row) {
                    // Chunk contains the insertion point — SPLIT it
                    int insertOffset = row - chunk->baseRow;

                    // Create a new chunk for data at/after the insert point
                    auto newChunk = std::make_unique<ColumnChunk>();
                    newChunk->baseRow = row + count; // shifted position

                    // Move cells from insertOffset..CHUNK_SIZE-1 to the new chunk
                    for (int off = insertOffset; off < ColumnChunk::CHUNK_SIZE; ++off) {
                        if (chunk->hasData(off)) {
                            int denseIdx = chunk->denseIndex(off);
                            auto type = static_cast<CellDataType>(chunk->types[denseIdx]);
                            double val = chunk->values[denseIdx];
                            uint16_t style = (chunk->styleIndices && denseIdx < (int)chunk->styleIndices->size())
                                ? (*chunk->styleIndices)[denseIdx] : 0;
                            int newOff = off - insertOffset; // offset within new chunk
                            switch (type) {
                                case CellDataType::Double: newChunk->setNumeric(newOff, val, style); break;
                                case CellDataType::Date: newChunk->setDate(newOff, val, style); break;
                                case CellDataType::Boolean: newChunk->setBoolean(newOff, val != 0.0, style); break;
                                case CellDataType::String: newChunk->setString(newOff, ColumnChunk::unpackId(val), style); break;
                                case CellDataType::Formula: {
                                    QString formula;
                                    if (chunk->formulas) {
                                        auto it = chunk->formulas->find(off);
                                        if (it != chunk->formulas->end()) formula = it->second;
                                    }
                                    newChunk->setFormula(newOff, formula, style);
                                    break;
                                }
                                default: break;
                            }
                            chunk->removeCell(off);
                        }
                    }

                    // Insert the new chunk right after the current one
                    if (newChunk->populatedCount > 0) {
                        chunks.insert(chunks.begin() + ci + 1, std::move(newChunk));
                        ++ci; // skip the newly inserted chunk in iteration
                    }

                    // Now shift all remaining chunks (already handled by the loop)
                }
            }
        }
        // Update metadata only — no cell scanning
        m_rowCount += count;
        m_maxRowColDirty = true;
        m_navIndexDirty = true;
        // Shift merged regions, validations, tables, row heights (all small lists)
        for (auto& mr : m_mergedRegions) {
            auto s = mr.range.getStart(); auto e = mr.range.getEnd();
            if (s.row >= row) s.row += count;
            if (e.row >= row) e.row += count;
            mr.range = CellRange(s, e);
        }
        for (auto& rule : m_validationRules) {
            auto s = rule.range.getStart(); auto e = rule.range.getEnd();
            if (s.row >= row) s.row += count;
            if (e.row >= row) e.row += count;
            rule.range = CellRange(s, e);
        }
        for (auto& tbl : m_tables) {
            auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
            if (s.row >= row) s.row += count;
            if (e.row >= row) e.row += count;
            tbl.range = CellRange(s, e);
        }
        {
            std::map<int, int> shifted;
            for (auto& [r, h] : m_rowHeights) {
                if (r >= row) shifted[r + count] = h;
                else shifted[r] = h;
            }
            m_rowHeights = std::move(shifted);
        }
        m_depGraph.shiftReferences(row, count, true);
        return;
    } else {
        // Original O(n) approach for small datasets (correct and handles all edge cases)
        for (int c = 0; c <= maxCol; ++c) {
            Column* col = m_columnStore.getColumn(c);
            if (!col) continue;

            struct CellData {
                int srcRow;
                QVariant value;
                CellDataType type;
                uint16_t styleIdx;
                QString formula;
                QString comment;
                QString hyperlink;
                CellAddress spillParent;
            };
            std::vector<CellData> toShift;

            for (auto& chunk : col->chunks()) {
                for (int offset = ColumnChunk::CHUNK_SIZE - 1; offset >= 0; --offset) {
                    if (!chunk->hasData(offset)) continue;
                    int r = chunk->baseRow + offset;
                    if (r < row) continue;
                    CellData cd;
                    cd.srcRow = r;
                    cd.type = chunk->getType(offset);
                    cd.value = chunk->getValue(offset);
                    cd.styleIdx = chunk->getStyleIndex(offset);
                    if (cd.type == CellDataType::Formula) cd.formula = chunk->getFormulaString(offset);
                    cd.comment = m_columnStore.getCellComment(r, c);
                    cd.hyperlink = m_columnStore.getCellHyperlink(r, c);
                    cd.spillParent = m_columnStore.getSpillParent(r, c);
                    toShift.push_back(cd);
                }
            }

            std::sort(toShift.begin(), toShift.end(), [](const CellData& a, const CellData& b) {
                return a.srcRow > b.srcRow;
            });
            for (const auto& cd : toShift) m_columnStore.removeCell(cd.srcRow, c);

            for (const auto& cd : toShift) {
                int newRow = cd.srcRow + count;
                switch (cd.type) {
                    case CellDataType::Double: case CellDataType::Date: case CellDataType::Boolean:
                        m_columnStore.setCellNumeric(newRow, c, cd.value.toDouble()); break;
                    case CellDataType::String:
                        m_columnStore.setCellString(newRow, c, cd.value.toString()); break;
                    case CellDataType::Formula: {
                        QString adjusted = shiftFormulaRefs(cd.formula, 'R', row, count);
                        m_columnStore.setCellFormula(newRow, c, adjusted); break;
                    }
                    case CellDataType::Error:
                        m_columnStore.setCellValue(newRow, c, cd.value); break;
                    default: break;
                }
                if (cd.styleIdx != 0) m_columnStore.setCellStyle(newRow, c, cd.styleIdx);
                if (!cd.comment.isEmpty()) m_columnStore.setCellComment(newRow, c, cd.comment);
                if (!cd.hyperlink.isEmpty()) m_columnStore.setCellHyperlink(newRow, c, cd.hyperlink);
                if (cd.spillParent.row >= 0) m_columnStore.setSpillParent(newRow, c, cd.spillParent);
            }
        }

        m_columnStore.forEachCell([&](int r, int c, CellDataType type, const QVariant&) {
            if (r < row && type == CellDataType::Formula) {
                QString formula = m_columnStore.getCellFormula(r, c);
                if (!formula.isEmpty()) {
                    QString adjusted = shiftFormulaRefs(formula, 'R', row, count);
                    if (adjusted != formula) m_columnStore.setCellFormula(r, c, adjusted);
                }
            }
        });
    }

    // Shift merged regions
    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        if (s.row >= row) s.row += count;
        if (e.row >= row) e.row += count;
        mr.range = CellRange(s, e);
    }
    // Shift validation rules
    for (auto& rule : m_validationRules) {
        auto s = rule.range.getStart(); auto e = rule.range.getEnd();
        if (s.row >= row) s.row += count;
        if (e.row >= row) e.row += count;
        rule.range = CellRange(s, e);
    }
    // Shift tables
    for (auto& tbl : m_tables) {
        auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
        if (s.row >= row) s.row += count;
        if (e.row >= row) e.row += count;
        tbl.range = CellRange(s, e);
    }
    // Shift row heights: rows >= row move up by count
    {
        std::map<int, int> shifted;
        for (auto& [r, h] : m_rowHeights) {
            if (r >= row) shifted[r + count] = h;
            else shifted[r] = h;
        }
        m_rowHeights = std::move(shifted);
    }
    m_rowCount += count;
    m_maxRowColDirty = true;
    m_navIndexDirty = true;

    // For small datasets, rebuild dependency graph immediately.
    // For large datasets, skip (too expensive — will be rebuilt lazily if needed).
    if (m_rowCount < 500000) {
        m_depGraph.clear();
        rebuildDependencyGraph();
    } else {
        m_depGraph.shiftReferences(row, count, true);
    }
}

void Spreadsheet::insertColumn(int column, int count) {
    // Shift columns right: process from rightmost to avoid overwrites
    int maxCol = m_columnStore.maxCol();
    for (int c = maxCol; c >= column; --c) {
        Column* col = m_columnStore.getColumn(c);
        if (!col) continue;

        // Move all cells from column c to column c+count
        for (auto& chunk : col->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (!chunk->hasData(offset)) continue;
                int r = chunk->baseRow + offset;
                QVariant val = chunk->getValue(offset);
                CellDataType type = chunk->getType(offset);
                uint16_t styleIdx = chunk->getStyleIndex(offset);

                if (type == CellDataType::Formula) {
                    QString formula = chunk->getFormulaString(offset);
                    QString adjusted = shiftFormulaRefs(formula, 'C', column, count);
                    m_columnStore.setCellFormula(r, c + count, adjusted);
                } else {
                    m_columnStore.setCellValue(r, c + count, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r, c + count, styleIdx);
            }
        }
        // Clear old column cells
        col->clear();
    }

    // Shift formula references in columns that weren't moved (left of insertion)
    for (int c = 0; c < column; ++c) {
        Column* col = m_columnStore.getColumn(c);
        if (!col) continue;
        for (auto& chunk : col->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (!chunk->hasData(offset) || chunk->getType(offset) != CellDataType::Formula)
                    continue;
                int r = chunk->baseRow + offset;
                QString formula = chunk->getFormulaString(offset);
                QString adjusted = shiftFormulaRefs(formula, 'C', column, count);
                if (adjusted != formula) {
                    m_columnStore.setCellFormula(r, c, adjusted);
                }
            }
        }
    }

    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        if (s.col >= column) s.col += count;
        if (e.col >= column) e.col += count;
        mr.range = CellRange(s, e);
    }
    for (auto& rule : m_validationRules) {
        auto s = rule.range.getStart(); auto e = rule.range.getEnd();
        if (s.col >= column) s.col += count;
        if (e.col >= column) e.col += count;
        rule.range = CellRange(s, e);
    }
    for (auto& tbl : m_tables) {
        auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
        if (s.col >= column) s.col += count;
        if (e.col >= column) e.col += count;
        tbl.range = CellRange(s, e);
    }
    // Shift column widths: cols >= column move right by count
    {
        std::map<int, int> shifted;
        for (auto& [c, w] : m_columnWidths) {
            if (c >= column) shifted[c + count] = w;
            else shifted[c] = w;
        }
        m_columnWidths = std::move(shifted);
    }
    m_columnCount += count;
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
    m_depGraph.clear();
    rebuildDependencyGraph();
}

void Spreadsheet::deleteRow(int row, int count) {
    int maxCol = m_columnStore.maxCol();
    int totalRows = getRowCount();

    // FAST PATH for large datasets: O(chunks + 64K) not O(cells)
    if (totalRows > 100000) {
        for (int c = 0; c <= maxCol; ++c) {
            Column* col = m_columnStore.getColumn(c);
            if (!col) continue;
            for (auto& chunk : col->chunks()) {
                if (chunk->baseRow >= row + count) {
                    // Entire chunk is below deleted range — shift baseRow up
                    chunk->baseRow -= count;
                } else if (chunk->baseRow + ColumnChunk::CHUNK_SIZE > row) {
                    // Chunk contains the deletion point — shift cells within it
                    int delOffset = row - chunk->baseRow;
                    // Remove deleted cells
                    for (int i = 0; i < count && delOffset + i < ColumnChunk::CHUNK_SIZE; ++i) {
                        chunk->removeCell(delOffset + i);
                    }
                    // Shift cells above deletion point down (fill the gap)
                    for (int off = delOffset + count; off < ColumnChunk::CHUNK_SIZE; ++off) {
                        if (chunk->hasData(off)) {
                            int denseIdx = chunk->denseIndex(off);
                            auto type = static_cast<CellDataType>(chunk->types[denseIdx]);
                            double val = chunk->values[denseIdx];
                            uint16_t style = (chunk->styleIndices && denseIdx < (int)chunk->styleIndices->size())
                                ? (*chunk->styleIndices)[denseIdx] : 0;
                            int newOff = off - count;
                            switch (type) {
                                case CellDataType::Double: chunk->setNumeric(newOff, val, style); break;
                                case CellDataType::Date: chunk->setDate(newOff, val, style); break;
                                case CellDataType::Boolean: chunk->setBoolean(newOff, val != 0.0, style); break;
                                case CellDataType::String: chunk->setString(newOff, ColumnChunk::unpackId(val), style); break;
                                case CellDataType::Formula:
                                    if (chunk->formulas) {
                                        auto it = chunk->formulas->find(off);
                                        if (it != chunk->formulas->end())
                                            chunk->setFormula(newOff, it->second, style);
                                    }
                                    break;
                                default: break;
                            }
                            chunk->removeCell(off);
                        }
                    }
                }
            }
        }
        m_rowCount -= count;
        m_maxRowColDirty = true;
        m_navIndexDirty = true;
        // Shift metadata (small lists)
        for (auto& mr : m_mergedRegions) {
            auto s = mr.range.getStart(); auto e = mr.range.getEnd();
            if (s.row >= row + count) s.row -= count;
            if (e.row >= row + count) e.row -= count;
            mr.range = CellRange(s, e);
        }
        for (auto& rule : m_validationRules) {
            auto s = rule.range.getStart(); auto e = rule.range.getEnd();
            if (s.row >= row + count) s.row -= count;
            if (e.row >= row + count) e.row -= count;
            rule.range = CellRange(s, e);
        }
        for (auto& tbl : m_tables) {
            auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
            if (s.row >= row + count) s.row -= count;
            if (e.row >= row + count) e.row -= count;
            tbl.range = CellRange(s, e);
        }
        {
            std::map<int, int> shifted;
            for (auto& [r, h] : m_rowHeights) {
                if (r >= row && r < row + count) continue;
                else if (r >= row + count) shifted[r - count] = h;
                else shifted[r] = h;
            }
            m_rowHeights = std::move(shifted);
        }
        m_depGraph.shiftReferences(row, -count, true);
        return;
    }

    // Original O(n) approach for small datasets
    for (int c = 0; c <= maxCol; ++c) {
        // Remove cells in deleted rows
        for (int r = row; r < row + count; ++r) {
            m_columnStore.removeCell(r, c);
        }

        // Collect cells below the deleted range and shift up
        Column* col = m_columnStore.getColumn(c);
        if (!col) continue;

        struct CellData {
            int srcRow;
            QVariant value;
            CellDataType type;
            uint16_t styleIdx;
            QString formula;
            QString comment;
            QString hyperlink;
            CellAddress spillParent;
        };
        std::vector<CellData> toShift;

        for (auto& chunk : col->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (!chunk->hasData(offset)) continue;
                int r = chunk->baseRow + offset;
                if (r < row + count) continue;

                CellData cd;
                cd.srcRow = r;
                cd.type = chunk->getType(offset);
                cd.value = chunk->getValue(offset);
                cd.styleIdx = chunk->getStyleIndex(offset);
                if (cd.type == CellDataType::Formula) {
                    cd.formula = chunk->getFormulaString(offset);
                }
                cd.comment = m_columnStore.getCellComment(r, c);
                cd.hyperlink = m_columnStore.getCellHyperlink(r, c);
                cd.spillParent = m_columnStore.getSpillParent(r, c);
                toShift.push_back(cd);
            }
        }

        // Sort by row ascending for clean removal
        std::sort(toShift.begin(), toShift.end(), [](const CellData& a, const CellData& b) {
            return a.srcRow < b.srcRow;
        });

        // Remove old positions
        for (const auto& cd : toShift) {
            m_columnStore.removeCell(cd.srcRow, c);
        }

        // Re-insert at shifted positions
        for (const auto& cd : toShift) {
            int newRow = cd.srcRow - count;
            switch (cd.type) {
                case CellDataType::Double:
                case CellDataType::Date:
                case CellDataType::Boolean:
                    m_columnStore.setCellNumeric(newRow, c, cd.value.toDouble());
                    break;
                case CellDataType::String:
                    m_columnStore.setCellString(newRow, c, cd.value.toString());
                    break;
                case CellDataType::Formula: {
                    QString adjusted = shiftFormulaRefs(cd.formula, 'R', row, -count);
                    m_columnStore.setCellFormula(newRow, c, adjusted);
                    break;
                }
                case CellDataType::Error:
                    m_columnStore.setCellValue(newRow, c, cd.value);
                    break;
                default:
                    break;
            }
            if (cd.styleIdx != 0) m_columnStore.setCellStyle(newRow, c, cd.styleIdx);
            if (!cd.comment.isEmpty()) m_columnStore.setCellComment(newRow, c, cd.comment);
            if (!cd.hyperlink.isEmpty()) m_columnStore.setCellHyperlink(newRow, c, cd.hyperlink);
            if (cd.spillParent.row >= 0) m_columnStore.setSpillParent(newRow, c, cd.spillParent);
        }
    }

    // Adjust formula refs in cells above deleted rows
    m_columnStore.forEachCell([&](int r, int c, CellDataType type, const QVariant&) {
        if (r < row && type == CellDataType::Formula) {
            QString formula = m_columnStore.getCellFormula(r, c);
            if (!formula.isEmpty()) {
                QString adjusted = shiftFormulaRefs(formula, 'R', row, -count);
                if (adjusted != formula) {
                    m_columnStore.setCellFormula(r, c, adjusted);
                }
            }
        }
    });

    // Remove merged regions fully in deleted rows; clamp partially overlapping ones
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [row, count](const MergedRegion& mr) {
            return mr.range.getStart().row >= row && mr.range.getEnd().row < row + count;
        }), m_mergedRegions.end());
    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        if (s.row >= row && s.row < row + count) s.row = row;
        else if (s.row >= row + count) s.row -= count;
        if (e.row >= row && e.row < row + count) e.row = row > 0 ? row - 1 : 0;
        else if (e.row >= row + count) e.row -= count;
        if (s.row > e.row) { s.row = e.row; }
        mr.range = CellRange(s, e);
    }
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [](const MergedRegion& mr) {
            return mr.range.getStart().row == mr.range.getEnd().row &&
                   mr.range.getStart().col == mr.range.getEnd().col;
        }), m_mergedRegions.end());
    for (auto& rule : m_validationRules) {
        auto s = rule.range.getStart(); auto e = rule.range.getEnd();
        if (s.row >= row + count) s.row -= count;
        if (e.row >= row + count) e.row -= count;
        rule.range = CellRange(s, e);
    }
    for (auto& tbl : m_tables) {
        auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
        if (s.row >= row + count) s.row -= count;
        if (e.row >= row + count) e.row -= count;
        tbl.range = CellRange(s, e);
    }
    // Shift row heights: remove deleted rows, shift remaining down
    {
        std::map<int, int> shifted;
        for (auto& [r, h] : m_rowHeights) {
            if (r >= row && r < row + count) continue; // deleted
            else if (r >= row + count) shifted[r - count] = h;
            else shifted[r] = h;
        }
        m_rowHeights = std::move(shifted);
    }
    m_rowCount -= count;
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
    if (m_rowCount < 500000) {
        m_depGraph.clear();
        rebuildDependencyGraph();
    } else {
        m_depGraph.shiftReferences(row, -count, true);
    }
}

void Spreadsheet::deleteColumn(int column, int count) {
    // Delete cells in the target columns
    for (int c = column; c < column + count; ++c) {
        Column* col = m_columnStore.getColumn(c);
        if (col) col->clear();
    }

    // Shift columns left
    int maxCol = m_columnStore.maxCol();
    for (int c = column + count; c <= maxCol; ++c) {
        Column* srcCol = m_columnStore.getColumn(c);
        if (!srcCol) continue;

        int destC = c - count;
        for (auto& chunk : srcCol->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (!chunk->hasData(offset)) continue;
                int r = chunk->baseRow + offset;
                QVariant val = chunk->getValue(offset);
                CellDataType type = chunk->getType(offset);
                uint16_t styleIdx = chunk->getStyleIndex(offset);

                if (type == CellDataType::Formula) {
                    QString formula = chunk->getFormulaString(offset);
                    QString adjusted = shiftFormulaRefs(formula, 'C', column, -count);
                    m_columnStore.setCellFormula(r, destC, adjusted);
                } else {
                    m_columnStore.setCellValue(r, destC, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r, destC, styleIdx);
            }
        }
        srcCol->clear();
    }

    // Adjust formula refs in columns left of deletion
    for (int c = 0; c < column; ++c) {
        Column* col = m_columnStore.getColumn(c);
        if (!col) continue;
        for (auto& chunk : col->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (!chunk->hasData(offset) || chunk->getType(offset) != CellDataType::Formula)
                    continue;
                int r = chunk->baseRow + offset;
                QString formula = chunk->getFormulaString(offset);
                QString adjusted = shiftFormulaRefs(formula, 'C', column, -count);
                if (adjusted != formula) {
                    m_columnStore.setCellFormula(r, c, adjusted);
                }
            }
        }
    }

    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [column, count](const MergedRegion& mr) {
            return mr.range.getStart().col >= column && mr.range.getEnd().col < column + count;
        }), m_mergedRegions.end());
    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        if (s.col >= column && s.col < column + count) s.col = column;
        else if (s.col >= column + count) s.col -= count;
        if (e.col >= column && e.col < column + count) e.col = column > 0 ? column - 1 : 0;
        else if (e.col >= column + count) e.col -= count;
        if (s.col > e.col) { s.col = e.col; }
        mr.range = CellRange(s, e);
    }
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [](const MergedRegion& mr) {
            return mr.range.getStart().row == mr.range.getEnd().row &&
                   mr.range.getStart().col == mr.range.getEnd().col;
        }), m_mergedRegions.end());
    for (auto& rule : m_validationRules) {
        auto s = rule.range.getStart(); auto e = rule.range.getEnd();
        if (s.col >= column + count) s.col -= count;
        if (e.col >= column + count) e.col -= count;
        rule.range = CellRange(s, e);
    }
    for (auto& tbl : m_tables) {
        auto s = tbl.range.getStart(); auto e = tbl.range.getEnd();
        if (s.col >= column + count) s.col -= count;
        if (e.col >= column + count) e.col -= count;
        tbl.range = CellRange(s, e);
    }
    // Shift column widths
    {
        std::map<int, int> shifted;
        for (auto& [c, w] : m_columnWidths) {
            if (c >= column && c < column + count) continue;
            else if (c >= column + count) shifted[c - count] = w;
            else shifted[c] = w;
        }
        m_columnWidths = std::move(shifted);
    }
    m_columnCount -= count;
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
    m_depGraph.clear();
    rebuildDependencyGraph();
}

QString Spreadsheet::getSheetName() const { return m_sheetName; }
void Spreadsheet::setSheetName(const QString& name) { m_sheetName = name; }

void Spreadsheet::updateMaxRowCol() const {
    if (!m_maxRowColDirty) return;
    m_cachedMaxRow = m_columnStore.maxRow();
    m_cachedMaxCol = m_columnStore.maxCol();
    m_maxRowColDirty = false;
}

int Spreadsheet::getMaxRow() const { updateMaxRowCol(); return m_cachedMaxRow; }
int Spreadsheet::getMaxColumn() const { updateMaxRowCol(); return m_cachedMaxCol; }

std::vector<CellAddress> Spreadsheet::getDirtyCells() const {
    // TODO: integrate with ColumnStore flags bitmap
    return {};
}

void Spreadsheet::clearDirtyFlag() {
    // TODO: integrate with ColumnStore flags bitmap
}

void Spreadsheet::startTransaction() { m_inTransaction = true; }

void Spreadsheet::commitTransaction() {
    m_inTransaction = false;
    if (m_autoRecalculate) recalculateAll();
}

void Spreadsheet::rollbackTransaction() { m_inTransaction = false; }

FormulaEngine& Spreadsheet::getFormulaEngine() { return *m_formulaEngine; }
void Spreadsheet::setAutoRecalculate(bool enabled) { m_autoRecalculate = enabled; }
bool Spreadsheet::getAutoRecalculate() const { return m_autoRecalculate; }

void Spreadsheet::beginBatchUpdate() {
    m_inBatchUpdate = true;
    m_batchDirtyCells.clear();
}

void Spreadsheet::endBatchUpdate() {
    m_inBatchUpdate = false;
    if (m_batchDirtyCells.empty()) return;

    // Collect all dirty cells and recalculate dependents in optimal order
    std::vector<CellAddress> allDirty;
    allDirty.reserve(m_batchDirtyCells.size());
    for (const auto& key : m_batchDirtyCells) {
        allDirty.emplace_back(key.row, key.col);
    }

    // Recalculate all formula cells that were directly changed
    for (const auto& addr : allDirty) {
        if (m_columnStore.getCellType(addr.row, addr.col) == CellDataType::Formula) {
            QString formula = m_columnStore.getCellFormula(addr.row, addr.col);
            m_columnStore.setComputedValue(addr.row, addr.col, m_formulaEngine->evaluate(formula));
        }
    }

    // Level-parallel recalculation: get topological levels and evaluate each level concurrently
    auto levels = m_depGraph.getRecalcLevels(allDirty);
    std::vector<CellAddress> recalculated;

    for (const auto& level : levels) {
        if (level.size() <= 4) {
            // Small level — sequential (avoids thread overhead)
            for (const auto& depAddr : level) {
                if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
                    QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
                    m_columnStore.setComputedValue(depAddr.row, depAddr.col,
                                                   m_formulaEngine->evaluate(formula));
                    recalculated.push_back(depAddr);
                }
            }
        } else {
            // Large level — sequential for now (parallel requires thread-safe AST pool)
            // TODO: Add per-thread AST pools for true parallel recalc
            for (const auto& depAddr : level) {
                if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
                    QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
                    m_columnStore.setComputedValue(depAddr.row, depAddr.col,
                                                   m_formulaEngine->evaluate(formula));
                    recalculated.push_back(depAddr);
                }
            }
        }
    }

    if (!recalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(recalculated);
    }

    m_batchDirtyCells.clear();
    m_navIndexDirty = true;
}

CellProxy Spreadsheet::getOrCreateCellFast(int row, int col) {
    // No mutex — main-thread-only bulk import path
    m_columnStore.getOrCreateColumn(col);
    // Skip nav index during bulk import — built lazily on first use
    // m_rowIndexCache per-row is too expensive at scale (4M+ rows = 500MB+)
    m_navIndexDirty = true;
    return CellProxy(&m_columnStore, row, col);
}

void Spreadsheet::finishBulkImport() {
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
    m_formulaCells.clear();
    // Clear incrementally built nav caches — they'll be rebuilt fully on first use
    m_colIndexCache.clear();
    m_rowIndexCache.clear();
    m_sortedOccupiedRows.clear();
}

void Spreadsheet::finishBulkImportWithMaxRowCol(int maxRow, int maxCol) {
    m_cachedMaxRow = maxRow;
    m_cachedMaxCol = maxCol;
    m_maxRowColDirty = false;
    m_formulaCells.clear();

    // Nav index will be built lazily on first use — skip expensive sort of
    // per-row cache (which would be 500MB+ for millions of rows)
    m_colIndexCache.clear();
    m_rowIndexCache.clear();
    m_sortedOccupiedRows.clear();
    m_navIndexDirty = true;
}

void Spreadsheet::rebuildDependencyGraph() {
    m_colRefFormulas.clear();
    m_columnStore.forEachCell([&](int row, int col, CellDataType type, const QVariant&) {
        if (type == CellDataType::Formula) {
            CellAddress addr(row, col);
            QString formula = m_columnStore.getCellFormula(row, col);
            m_formulaEngine->evaluate(formula);
            for (const auto& dep : m_formulaEngine->getLastDependencies()) {
                m_depGraph.addDependency(addr, dep);
            }
            for (int c : m_formulaEngine->getLastColumnDeps()) {
                m_colRefFormulas[c].push_back(addr);
            }
        }
    });
}

void Spreadsheet::mergeBulkCells(CellMap& source) {
    // Legacy compatibility: import from old CellMap format into ColumnStore
    for (auto& [key, cell] : source) {
        if (!cell) continue;
        if (cell->getType() == CellType::Formula) {
            m_columnStore.setCellFormula(key.row, key.col, cell->getFormula());
            m_columnStore.setComputedValue(key.row, key.col, cell->getComputedValue());
        } else {
            m_columnStore.setCellValue(key.row, key.col, cell->getValue());
        }
        if (cell->hasCustomStyle()) {
            uint16_t styleIdx = StyleTable::instance().intern(cell->getStyle());
            m_columnStore.setCellStyle(key.row, key.col, styleIdx);
        }
        if (cell->hasComment()) {
            m_columnStore.setCellComment(key.row, key.col, cell->getComment());
        }
        if (cell->hasHyperlink()) {
            m_columnStore.setCellHyperlink(key.row, key.col, cell->getHyperlink());
        }
        if (cell->isSpillCell()) {
            m_columnStore.setSpillParent(key.row, key.col, cell->getSpillParent());
        }
    }
    source.clear();
}

void Spreadsheet::setRowHeight(int row, int height) { m_rowHeights[row] = height; }
void Spreadsheet::setColumnWidth(int col, int width) { m_columnWidths[col] = width; }
int Spreadsheet::getRowHeight(int row) const { auto it = m_rowHeights.find(row); return it != m_rowHeights.end() ? it->second : 0; }
int Spreadsheet::getColumnWidth(int col) const { auto it = m_columnWidths.find(col); return it != m_columnWidths.end() ? it->second : 0; }

void Spreadsheet::setPivotConfig(std::unique_ptr<PivotConfig> config) {
    m_pivotConfig = std::move(config);
}
const PivotConfig* Spreadsheet::getPivotConfig() const { return m_pivotConfig.get(); }
bool Spreadsheet::isPivotSheet() const { return m_pivotConfig != nullptr; }

void Spreadsheet::setSparkline(const CellAddress& addr, const SparklineConfig& config) {
    m_sparklines[{addr.row, addr.col}] = config;
}

void Spreadsheet::removeSparkline(const CellAddress& addr) {
    m_sparklines.erase({addr.row, addr.col});
}

const SparklineConfig* Spreadsheet::getSparkline(const CellAddress& addr) const {
    auto it = m_sparklines.find({addr.row, addr.col});
    return it != m_sparklines.end() ? &it->second : nullptr;
}

CellSnapshot Spreadsheet::takeCellSnapshot(const CellAddress& addr) {
    CellProxy cell = getCell(addr);
    CellSnapshot snap;
    snap.addr = addr;
    snap.value = cell.getValue();
    snap.formula = cell.getFormula();
    snap.style = cell.getStyle();
    snap.type = cell.getType();
    return snap;
}

void Spreadsheet::forEachCell(std::function<void(int row, int col, const Cell&)> callback) const {
    // Bridge: ColumnStore → Cell for backward-compatible serialization
    // Creates a temporary Cell object for each populated cell
    m_columnStore.forEachCell([&](int row, int col, CellDataType type, const QVariant& value) {
        Cell tempCell;
        switch (type) {
            case CellDataType::Formula: {
                QString formula = m_columnStore.getCellFormula(row, col);
                tempCell.setFormula(formula);
                QVariant computed = m_columnStore.getComputedValue(row, col);
                if (computed.isValid()) tempCell.setComputedValue(computed);
                break;
            }
            case CellDataType::Error:
                tempCell.setError(value.toString());
                break;
            default:
                tempCell.setValue(value);
                break;
        }

        // Copy style if non-default
        uint16_t styleIdx = m_columnStore.getCellStyleIndex(row, col);
        if (styleIdx != 0) {
            tempCell.setStyle(StyleTable::instance().get(styleIdx));
        }

        // Copy comment
        QString comment = m_columnStore.getCellComment(row, col);
        if (!comment.isEmpty()) tempCell.setComment(comment);

        // Copy hyperlink
        QString hyperlink = m_columnStore.getCellHyperlink(row, col);
        if (!hyperlink.isEmpty()) tempCell.setHyperlink(hyperlink);

        // Copy spill parent
        CellAddress spillParent = m_columnStore.getSpillParent(row, col);
        if (spillParent.row >= 0) tempCell.setSpillParent(spillParent);

        callback(row, col, tempCell);
    });
}

void Spreadsheet::recalculate(const CellAddress& addr) {
    if (m_columnStore.getCellType(addr.row, addr.col) == CellDataType::Formula) {
        QString formula = m_columnStore.getCellFormula(addr.row, addr.col);
        m_columnStore.setComputedValue(addr.row, addr.col, m_formulaEngine->evaluate(formula));
    }
}

void Spreadsheet::recalculateAll() {
    // Ensure formula tracker is populated
    if (m_formulaCells.empty()) {
        m_columnStore.forEachCell([&](int row, int col, CellDataType type, const QVariant&) {
            if (type == CellDataType::Formula) {
                m_formulaCells.insert(CellKey{row, col});
            }
        });
    }

    if (m_formulaCells.empty()) return;

    // Build list of all formula cell addresses for topological sorting
    std::vector<CellAddress> allFormulaCells;
    allFormulaCells.reserve(m_formulaCells.size());
    for (const auto& key : m_formulaCells) {
        allFormulaCells.emplace_back(key.row, key.col);
    }

    // Get topological evaluation order using dependency graph (Kahn's algorithm).
    // getRecalcLevels returns levels where each level's cells are independent and
    // depend only on cells in previous levels — guaranteeing correct evaluation order.
    auto levels = m_depGraph.getRecalcLevels(allFormulaCells);

    // Track which cells were evaluated via the dependency graph
    std::unordered_set<CellKey, CellKeyHash> evaluated;

    for (const auto& level : levels) {
        for (const auto& addr : level) {
            if (m_columnStore.getCellType(addr.row, addr.col) == CellDataType::Formula) {
                QString formula = m_columnStore.getCellFormula(addr.row, addr.col);
                if (!formula.isEmpty()) {
                    m_columnStore.setComputedValue(addr.row, addr.col, m_formulaEngine->evaluate(formula));
                }
                evaluated.insert(CellKey{addr.row, addr.col});
            }
        }
    }

    // Evaluate any remaining formula cells not in the dependency graph
    // (e.g., formulas with no dependencies on other formula cells)
    for (const auto& key : m_formulaCells) {
        if (evaluated.find(key) == evaluated.end()) {
            if (m_columnStore.getCellType(key.row, key.col) == CellDataType::Formula) {
                QString formula = m_columnStore.getCellFormula(key.row, key.col);
                if (!formula.isEmpty()) {
                    m_columnStore.setComputedValue(key.row, key.col, m_formulaEngine->evaluate(formula));
                }
            }
        }
    }
}

void Spreadsheet::updateDependencies(const CellAddress& addr) {
    m_depGraph.removeDependencies(addr);

    // Remove old column-level deps for this formula
    for (auto& [col, addrs] : m_colRefFormulas) {
        addrs.erase(std::remove(addrs.begin(), addrs.end(), addr), addrs.end());
    }

    if (m_columnStore.getCellType(addr.row, addr.col) == CellDataType::Formula) {
        QString formula = m_columnStore.getCellFormula(addr.row, addr.col);
        m_formulaEngine->evaluate(formula);
        for (const auto& dep : m_formulaEngine->getLastDependencies()) {
            m_depGraph.addDependency(addr, dep);
        }
        // Register column-level dependencies (from A:A style references)
        for (int col : m_formulaEngine->getLastColumnDeps()) {
            m_colRefFormulas[col].push_back(addr);
        }
    }
}

void Spreadsheet::recalculateDependents(const CellAddress& addr) {
    auto order = m_depGraph.getRecalcOrder(addr);
    std::vector<CellAddress> recalculated;
    for (const auto& depAddr : order) {
        if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
            QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
            m_columnStore.setComputedValue(depAddr.row, depAddr.col, m_formulaEngine->evaluate(formula));
            recalculated.push_back(depAddr);
        }
    }
    if (!recalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(recalculated);
    }
}

void Spreadsheet::recalculateDependentsParallel(const CellAddress& addr) {
    // Get topological levels: cells at the same level are independent → can run in parallel
    auto levels = m_depGraph.getRecalcLevels(addr);
    std::vector<CellAddress> allRecalculated;

    for (const auto& level : levels) {
        if (level.size() <= 4) {
            // Small level — sequential is faster (avoids thread overhead)
            for (const auto& depAddr : level) {
                if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
                    QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
                    m_columnStore.setComputedValue(depAddr.row, depAddr.col,
                                                   m_formulaEngine->evaluate(formula));
                    allRecalculated.push_back(depAddr);
                }
            }
        } else {
            // Large level — sequential for now (parallel requires thread-safe AST pool)
            for (const auto& depAddr : level) {
                if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
                    QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
                    m_columnStore.setComputedValue(depAddr.row, depAddr.col,
                                                   m_formulaEngine->evaluate(formula));
                    allRecalculated.push_back(depAddr);
                }
            }
        }
    }

    if (!allRecalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(allRecalculated);
    }
}

void Spreadsheet::recalculateColumnDependents(int col) {
    auto it = m_colRefFormulas.find(col);
    if (it == m_colRefFormulas.end() || it->second.empty()) return;

    std::vector<CellAddress> recalculated;
    for (const auto& depAddr : it->second) {
        if (m_columnStore.getCellType(depAddr.row, depAddr.col) == CellDataType::Formula) {
            QString formula = m_columnStore.getCellFormula(depAddr.row, depAddr.col);
            m_columnStore.setComputedValue(depAddr.row, depAddr.col, m_formulaEngine->evaluate(formula));
            recalculated.push_back(depAddr);
            // Cascade: recalculate anything that depends on this formula cell
            recalculateDependents(depAddr);
        }
    }
    if (!recalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(recalculated);
    }
}

void Spreadsheet::sortRange(const CellRange& range, int sortColumn, bool ascending) {
    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;
    if (startRow >= endRow) return;

    int numRows = endRow - startRow + 1;

    // Extract sort keys: separate numeric and string keys for faster comparison
    struct SortKey {
        double numericKey;
        QString stringKey;
        bool isNumeric;
        bool isEmpty;
        int originalIndex;
    };
    std::vector<SortKey> keys(numRows);

    // Fast key extraction: read type + value directly from ColumnStore (no QVariant overhead)
    auto extractKey = [&](int i) {
        int r = startRow + i;
        keys[i].originalIndex = i;

        if (!m_columnStore.hasCell(r, sortColumn)) {
            keys[i].isEmpty = true;
            keys[i].isNumeric = false;
            keys[i].numericKey = 0;
            return;
        }

        CellDataType type = m_columnStore.getCellType(r, sortColumn);
        keys[i].isEmpty = (type == CellDataType::Empty);

        if (type == CellDataType::Double || type == CellDataType::Date) {
            // Direct numeric read — no QVariant, no string conversion
            keys[i].isNumeric = true;
            auto* col = m_columnStore.getColumn(sortColumn);
            auto* chunk = col ? col->getChunk(r) : nullptr;
            if (chunk) {
                int offset = r - chunk->baseRow;
                int idx = chunk->denseIndex(offset);
                keys[i].numericKey = chunk->values[idx];
            }
        } else if (type == CellDataType::Boolean) {
            keys[i].isNumeric = true;
            auto* col = m_columnStore.getColumn(sortColumn);
            auto* chunk = col ? col->getChunk(r) : nullptr;
            if (chunk) {
                int offset = r - chunk->baseRow;
                int idx = chunk->denseIndex(offset);
                keys[i].numericKey = chunk->values[idx];
            }
        } else {
            // String/Formula/Error — need string key
            QVariant val = CellProxy(&m_columnStore, r, sortColumn).getValue();
            keys[i].isNumeric = false;
            keys[i].stringKey = val.toString();
            // Try numeric conversion for string cells containing numbers
            bool ok;
            double d = keys[i].stringKey.toDouble(&ok);
            if (ok) {
                keys[i].isNumeric = true;
                keys[i].numericKey = d;
                keys[i].stringKey.clear();
            }
        }
    };

    if (numRows > 10000) {
        // Parallel key extraction
        std::vector<int> rowIndices(numRows);
        std::iota(rowIndices.begin(), rowIndices.end(), 0);
        QtConcurrent::blockingMap(rowIndices, extractKey);
    } else {
        for (int i = 0; i < numRows; ++i) extractKey(i);
    }

    // Stable sort preserves relative order of equal elements (Excel-compatible behavior)
    std::stable_sort(keys.begin(), keys.end(), [ascending](const SortKey& a, const SortKey& b) {
        if (a.isEmpty && b.isEmpty) return false;
        if (a.isEmpty) return false;
        if (b.isEmpty) return true;
        if (a.isNumeric && b.isNumeric) {
            return ascending ? (a.numericKey < b.numericKey) : (a.numericKey > b.numericKey);
        }
        // Mixed: numbers before strings
        if (a.isNumeric != b.isNumeric) return a.isNumeric;
        int cmp = a.stringKey.compare(b.stringKey, Qt::CaseInsensitive);
        return ascending ? (cmp < 0) : (cmp > 0);
    });

    // Build permutation vector
    std::vector<int> perm(numRows);
    for (int i = 0; i < numRows; ++i) {
        perm[i] = keys[i].originalIndex;
    }

    // Apply permutation to all columns in range
    m_columnStore.applyPermutation(startRow, startCol, endCol, perm);

    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::sortRangeMulti(const CellRange& range, const std::vector<std::pair<int, bool>>& sortKeys) {
    if (sortKeys.empty()) return;

    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;
    if (startRow >= endRow) return;

    int numRows = endRow - startRow + 1;
    int numKeys = static_cast<int>(sortKeys.size());

    // Extract sort keys for all sort columns at once
    struct SortKey {
        double numericKey;
        QString stringKey;
        bool isNumeric;
        bool isEmpty;
    };

    // keys[keyIndex][rowIndex]
    std::vector<std::vector<SortKey>> allKeys(numKeys, std::vector<SortKey>(numRows));

    for (int k = 0; k < numKeys; ++k) {
        int sortColumn = sortKeys[k].first;

        auto extractKey = [&](int i) {
            int r = startRow + i;
            auto& key = allKeys[k][i];

            if (!m_columnStore.hasCell(r, sortColumn)) {
                key.isEmpty = true;
                key.isNumeric = false;
                key.numericKey = 0;
                return;
            }

            CellDataType type = m_columnStore.getCellType(r, sortColumn);
            key.isEmpty = (type == CellDataType::Empty);

            if (type == CellDataType::Double || type == CellDataType::Date) {
                key.isNumeric = true;
                auto* col = m_columnStore.getColumn(sortColumn);
                auto* chunk = col ? col->getChunk(r) : nullptr;
                if (chunk) {
                    int offset = r - chunk->baseRow;
                    int idx = chunk->denseIndex(offset);
                    key.numericKey = chunk->values[idx];
                }
            } else if (type == CellDataType::Boolean) {
                key.isNumeric = true;
                auto* col = m_columnStore.getColumn(sortColumn);
                auto* chunk = col ? col->getChunk(r) : nullptr;
                if (chunk) {
                    int offset = r - chunk->baseRow;
                    int idx = chunk->denseIndex(offset);
                    key.numericKey = chunk->values[idx];
                }
            } else {
                QVariant val = CellProxy(&m_columnStore, r, sortColumn).getValue();
                key.isNumeric = false;
                key.stringKey = val.toString();
                bool ok;
                double d = key.stringKey.toDouble(&ok);
                if (ok) {
                    key.isNumeric = true;
                    key.numericKey = d;
                    key.stringKey.clear();
                }
            }
        };

        if (numRows > 10000) {
            std::vector<int> rowIndices(numRows);
            std::iota(rowIndices.begin(), rowIndices.end(), 0);
            QtConcurrent::blockingMap(rowIndices, extractKey);
        } else {
            for (int i = 0; i < numRows; ++i) extractKey(i);
        }
    }

    // Build index array and sort with multi-level comparator
    std::vector<int> indices(numRows);
    std::iota(indices.begin(), indices.end(), 0);

    std::stable_sort(indices.begin(), indices.end(), [&](int ai, int bi) {
        for (int k = 0; k < numKeys; ++k) {
            const auto& a = allKeys[k][ai];
            const auto& b = allKeys[k][bi];
            bool asc = sortKeys[k].second;

            // Empty cells always sort to the end
            if (a.isEmpty && b.isEmpty) continue;
            if (a.isEmpty) return false;
            if (b.isEmpty) return true;

            if (a.isNumeric && b.isNumeric) {
                if (a.numericKey != b.numericKey) {
                    return asc ? (a.numericKey < b.numericKey) : (a.numericKey > b.numericKey);
                }
                continue; // tie — check next key
            }
            // Mixed: numbers before strings
            if (a.isNumeric != b.isNumeric) return a.isNumeric;

            int cmp = a.stringKey.compare(b.stringKey, Qt::CaseInsensitive);
            if (cmp != 0) {
                return asc ? (cmp < 0) : (cmp > 0);
            }
            // tie — check next key
        }
        return false; // all keys equal — stable
    });

    // Apply permutation to all columns in range
    m_columnStore.applyPermutation(startRow, startCol, endCol, indices);

    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

// ============================================================================
// Sheet protection
// ============================================================================
void Spreadsheet::setProtected(bool protect, const QString& password) {
    m_isProtected = protect;
    if (!password.isEmpty()) {
        // Simple hash for password verification (not cryptographic security)
        m_protectionPasswordHash = QString::number(qHash(password));
    } else {
        m_protectionPasswordHash.clear();
    }
}

bool Spreadsheet::isProtected() const { return m_isProtected; }

bool Spreadsheet::checkProtectionPassword(const QString& password) const {
    if (m_protectionPasswordHash.isEmpty()) return true;
    return QString::number(qHash(password)) == m_protectionPasswordHash;
}

// ============================================================================
// Parallel search across all cells
// ============================================================================
std::vector<CellAddress> Spreadsheet::searchAllCells(const QString& query, bool caseSensitive,
                                                      bool wholeCell) const {
    int maxCol = m_columnStore.maxCol();
    int maxRow = m_columnStore.maxRow();
    if (maxCol < 0 || maxRow < 0) return {};

    Qt::CaseSensitivity cs = caseSensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;

    // Per-column results (to avoid contention)
    std::vector<std::vector<CellAddress>> perColResults(maxCol + 1);

    auto searchColumn = [&](int col) {
        m_columnStore.scanColumnValues(col, 0, maxRow,
            [&](int row, const QVariant& val) {
                QString str = val.toString();
                bool match = wholeCell
                    ? (str.compare(query, cs) == 0)
                    : str.contains(query, cs);
                if (match) {
                    perColResults[col].push_back(CellAddress(row, col));
                }
            });
    };

    if (maxCol > 4) {
        // Parallel search across columns
        std::vector<int> colIndices(maxCol + 1);
        std::iota(colIndices.begin(), colIndices.end(), 0);
        QtConcurrent::blockingMap(colIndices, searchColumn);
    } else {
        for (int c = 0; c <= maxCol; ++c) searchColumn(c);
    }

    // Merge results in row-major order
    std::vector<CellAddress> results;
    for (auto& colResult : perColResults) {
        results.insert(results.end(), colResult.begin(), colResult.end());
    }

    // Sort by row, then column for consistent ordering
    std::sort(results.begin(), results.end(), [](const CellAddress& a, const CellAddress& b) {
        return a.row < b.row || (a.row == b.row && a.col < b.col);
    });

    return results;
}

void Spreadsheet::insertCellsShiftRight(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int colCount = range.getEnd().col - startCol + 1;
    int maxCol = m_columnStore.maxCol();

    for (int r = startRow; r <= endRow; ++r) {
        // Shift cells right: process from rightmost to avoid overwrites
        for (int c = maxCol; c >= startCol; --c) {
            if (m_columnStore.hasCell(r, c)) {
                QVariant val = m_columnStore.getCellValue(r, c);
                uint16_t styleIdx = m_columnStore.getCellStyleIndex(r, c);
                CellDataType type = m_columnStore.getCellType(r, c);
                if (type == CellDataType::Formula) {
                    m_columnStore.setCellFormula(r, c + colCount, m_columnStore.getCellFormula(r, c));
                } else {
                    m_columnStore.setCellValue(r, c + colCount, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r, c + colCount, styleIdx);
                m_columnStore.removeCell(r, c);
            }
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::insertCellsShiftDown(const CellRange& range) {
    int startRow = range.getStart().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int rowCount = range.getEnd().row - startRow + 1;
    int maxRow = m_columnStore.maxRow();

    for (int c = startCol; c <= endCol; ++c) {
        // Shift cells down: process from bottom to avoid overwrites
        for (int r = maxRow; r >= startRow; --r) {
            if (m_columnStore.hasCell(r, c)) {
                QVariant val = m_columnStore.getCellValue(r, c);
                uint16_t styleIdx = m_columnStore.getCellStyleIndex(r, c);
                CellDataType type = m_columnStore.getCellType(r, c);
                if (type == CellDataType::Formula) {
                    m_columnStore.setCellFormula(r + rowCount, c, m_columnStore.getCellFormula(r, c));
                } else {
                    m_columnStore.setCellValue(r + rowCount, c, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r + rowCount, c, styleIdx);
                m_columnStore.removeCell(r, c);
            }
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::deleteCellsShiftLeft(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int colCount = endCol - startCol + 1;
    int maxCol = m_columnStore.maxCol();

    for (int r = startRow; r <= endRow; ++r) {
        // Remove cells in range
        for (int c = startCol; c <= endCol; ++c) {
            m_columnStore.removeCell(r, c);
        }
        // Shift cells left
        for (int c = endCol + 1; c <= maxCol; ++c) {
            if (m_columnStore.hasCell(r, c)) {
                QVariant val = m_columnStore.getCellValue(r, c);
                uint16_t styleIdx = m_columnStore.getCellStyleIndex(r, c);
                CellDataType type = m_columnStore.getCellType(r, c);
                if (type == CellDataType::Formula) {
                    m_columnStore.setCellFormula(r, c - colCount, m_columnStore.getCellFormula(r, c));
                } else {
                    m_columnStore.setCellValue(r, c - colCount, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r, c - colCount, styleIdx);
                m_columnStore.removeCell(r, c);
            }
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::deleteCellsShiftUp(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int rowCount = endRow - startRow + 1;
    int maxRow = m_columnStore.maxRow();

    for (int c = startCol; c <= endCol; ++c) {
        // Remove cells in range
        for (int r = startRow; r <= endRow; ++r) {
            m_columnStore.removeCell(r, c);
        }
        // Shift cells up
        for (int r = endRow + 1; r <= maxRow; ++r) {
            if (m_columnStore.hasCell(r, c)) {
                QVariant val = m_columnStore.getCellValue(r, c);
                uint16_t styleIdx = m_columnStore.getCellStyleIndex(r, c);
                CellDataType type = m_columnStore.getCellType(r, c);
                if (type == CellDataType::Formula) {
                    m_columnStore.setCellFormula(r - rowCount, c, m_columnStore.getCellFormula(r, c));
                } else {
                    m_columnStore.setCellValue(r - rowCount, c, val);
                }
                if (styleIdx != 0) m_columnStore.setCellStyle(r - rowCount, c, styleIdx);
                m_columnStore.removeCell(r, c);
            }
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

// ============== Document Theme ==============
void Spreadsheet::setDocumentTheme(const DocumentTheme& theme) {
    m_documentTheme = theme;

    // Re-theme existing tables
    auto newThemes = generateTableThemes(theme);
    for (auto& table : m_tables) {
        for (const auto& nt : newThemes) {
            if (nt.name == table.theme.name) {
                table.theme = nt;
                break;
            }
        }
    }
}

// ============== Table Support ==============
void Spreadsheet::addTable(const SpreadsheetTable& table) { m_tables.push_back(table); }

void Spreadsheet::removeTable(const QString& name) {
    m_tables.erase(std::remove_if(m_tables.begin(), m_tables.end(),
        [&name](const SpreadsheetTable& t) { return t.name == name; }), m_tables.end());
}

const SpreadsheetTable* Spreadsheet::getTableAt(int row, int col) const {
    for (const auto& t : m_tables) {
        if (row >= t.range.getStart().row && row <= t.range.getEnd().row &&
            col >= t.range.getStart().col && col <= t.range.getEnd().col) return &t;
    }
    return nullptr;
}

// ============== Merge Cells ==============
void Spreadsheet::mergeCells(const CellRange& range) {
    for (const auto& mr : m_mergedRegions)
        if (mr.range.intersects(range)) return;
    m_mergedRegions.push_back({range});
}

void Spreadsheet::unmergeCells(const CellRange& range) {
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [&range](const MergedRegion& mr) { return mr.range.intersects(range); }), m_mergedRegions.end());
}

const Spreadsheet::MergedRegion* Spreadsheet::getMergedRegionAt(int row, int col) const {
    for (const auto& mr : m_mergedRegions)
        if (mr.range.contains(row, col)) return &mr;
    return nullptr;
}

// ============== Data Validation ==============
void Spreadsheet::addValidationRule(const DataValidationRule& rule) { m_validationRules.push_back(rule); }

void Spreadsheet::removeValidationRule(int index) {
    if (index >= 0 && index < static_cast<int>(m_validationRules.size()))
        m_validationRules.erase(m_validationRules.begin() + index);
}

const Spreadsheet::DataValidationRule* Spreadsheet::getValidationAt(int row, int col) const {
    for (const auto& rule : m_validationRules)
        if (rule.range.contains(row, col)) return &rule;
    return nullptr;
}

// ============== Fast Cell Navigation ==============
void Spreadsheet::buildNavIndexIfNeeded() const {
    if (!m_navIndexDirty) return;
    m_colIndexCache.clear();
    m_rowIndexCache.clear();
    m_sortedOccupiedRows.clear();

    // Build per-column index by scanning chunks directly — much faster than forEachCell
    // which iterates every bit in every presence bitmap. This uses popcount to count
    // populated rows per chunk, then extracts row numbers only for occupied chunks.
    int numCols = m_columnStore.columnCount();
    for (int c = 0; c < numCols; ++c) {
        auto* col = m_columnStore.getColumn(c);
        if (!col) continue;

        auto& rowList = m_colIndexCache[c];
        for (const auto& chunk : col->chunks()) {
            if (chunk->populatedCount == 0) continue;
            // Scan presence bitmap using popcount for speed
            for (int w = 0; w < ColumnChunk::BITMAP_WORDS; ++w) {
                uint64_t word = chunk->presence[w];
                while (word) {
                    int bit = ctzll_s(word); // count trailing zeros
                    int offset = w * 64 + bit;
                    rowList.push_back(chunk->baseRow + offset);
                    word &= word - 1; // clear lowest set bit
                }
            }
        }
        // Rows are already sorted (chunks sorted by baseRow, bits in order)
    }

    m_navIndexDirty = false;
}

void Spreadsheet::navIndexInsert(int row, int col) const {
    if (m_navIndexDirty) return;
    auto& colRows = m_colIndexCache[col];
    auto it = std::lower_bound(colRows.begin(), colRows.end(), row);
    if (it == colRows.end() || *it != row) colRows.insert(it, row);

    auto it3 = std::lower_bound(m_sortedOccupiedRows.begin(), m_sortedOccupiedRows.end(), row);
    if (it3 == m_sortedOccupiedRows.end() || *it3 != row) m_sortedOccupiedRows.insert(it3, row);
}

void Spreadsheet::navIndexRemove(int row, int col) const {
    if (m_navIndexDirty) return;
    auto colIt = m_colIndexCache.find(col);
    if (colIt != m_colIndexCache.end()) {
        auto& rows = colIt->second;
        auto it = std::lower_bound(rows.begin(), rows.end(), row);
        if (it != rows.end() && *it == row) rows.erase(it);
        if (rows.empty()) m_colIndexCache.erase(colIt);
    }
    // Check if this row still has data in any column
    if (!m_columnStore.hasCell(row, col)) {
        bool rowEmpty = true;
        int numCols = m_columnStore.columnCount();
        for (int c = 0; c < numCols; ++c) {
            if (m_columnStore.hasCell(row, c)) { rowEmpty = false; break; }
        }
        if (rowEmpty) {
            auto it2 = std::lower_bound(m_sortedOccupiedRows.begin(), m_sortedOccupiedRows.end(), row);
            if (it2 != m_sortedOccupiedRows.end() && *it2 == row) m_sortedOccupiedRows.erase(it2);
        }
    }
}

const std::vector<int>& Spreadsheet::getOccupiedRowsInColumn(int col) const {
    static const std::vector<int> empty;
    buildNavIndexIfNeeded();
    auto it = m_colIndexCache.find(col);
    if (it == m_colIndexCache.end()) return empty;
    return it->second;
}

void Spreadsheet::streamColumnValues(int col, int startRow, int endRow,
                                     const std::function<void(const QVariant&)>& fn) const {
    // Use ColumnStore directly — zero materialization
    m_columnStore.scanColumnValues(col, startRow, endRow, [&fn](int, const QVariant& val) {
        fn(val);
    });
}

const std::vector<int>& Spreadsheet::getOccupiedColsInRow(int row) const {
    // Query ColumnStore directly — O(numColumns) which is fast (~20 cols)
    // Avoids maintaining a per-row cache (would be 500MB+ for millions of rows)
    thread_local std::vector<int> result;
    result.clear();
    int numCols = m_columnStore.columnCount();
    for (int c = 0; c < numCols; ++c) {
        if (m_columnStore.hasCell(row, c)) {
            result.push_back(c);
        }
    }
    return result;
}

const std::vector<int>& Spreadsheet::getOccupiedRows() const {
    buildNavIndexIfNeeded();
    return m_sortedOccupiedRows;
}

bool Spreadsheet::validateCell(int row, int col, const QString& value) const {
    const auto* rule = getValidationAt(row, col);
    if (!rule) return true;
    if (value.isEmpty()) return true;

    switch (rule->type) {
        case DataValidationRule::WholeNumber: {
            bool ok; int num = value.toInt(&ok); if (!ok) return false;
            int v1 = rule->value1.toInt(), v2 = rule->value2.toInt();
            switch (rule->op) {
                case DataValidationRule::Between: return num >= v1 && num <= v2;
                case DataValidationRule::NotBetween: return num < v1 || num > v2;
                case DataValidationRule::EqualTo: return num == v1;
                case DataValidationRule::NotEqualTo: return num != v1;
                case DataValidationRule::GreaterThan: return num > v1;
                case DataValidationRule::LessThan: return num < v1;
                case DataValidationRule::GreaterThanOrEqual: return num >= v1;
                case DataValidationRule::LessThanOrEqual: return num <= v1;
            }
            break;
        }
        case DataValidationRule::Decimal: {
            bool ok; double num = value.toDouble(&ok); if (!ok) return false;
            double v1 = rule->value1.toDouble(), v2 = rule->value2.toDouble();
            switch (rule->op) {
                case DataValidationRule::Between: return num >= v1 && num <= v2;
                case DataValidationRule::NotBetween: return num < v1 || num > v2;
                case DataValidationRule::EqualTo: return qFuzzyCompare(num, v1);
                case DataValidationRule::NotEqualTo: return !qFuzzyCompare(num, v1);
                case DataValidationRule::GreaterThan: return num > v1;
                case DataValidationRule::LessThan: return num < v1;
                case DataValidationRule::GreaterThanOrEqual: return num >= v1;
                case DataValidationRule::LessThanOrEqual: return num <= v1;
            }
            break;
        }
        case DataValidationRule::List: {
            QStringList items = rule->listItems;
            if (items.isEmpty() && !rule->listSourceRange.isEmpty()) {
                // Resolve range reference to get list items
                CellRange range(rule->listSourceRange);
                if (range.isValid()) {
                    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
                        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
                            auto cell = getCellIfExists(r, c);
                            if (cell) {
                                QString val = cell->getValue().toString().trimmed();
                                if (!val.isEmpty()) items.append(val);
                            }
                        }
                    }
                }
            }
            return items.contains(value, Qt::CaseInsensitive);
        }
        case DataValidationRule::TextLength: {
            int len = value.length(), v1 = rule->value1.toInt(), v2 = rule->value2.toInt();
            switch (rule->op) {
                case DataValidationRule::Between: return len >= v1 && len <= v2;
                case DataValidationRule::NotBetween: return len < v1 || len > v2;
                case DataValidationRule::EqualTo: return len == v1;
                case DataValidationRule::NotEqualTo: return len != v1;
                case DataValidationRule::GreaterThan: return len > v1;
                case DataValidationRule::LessThan: return len < v1;
                case DataValidationRule::GreaterThanOrEqual: return len >= v1;
                case DataValidationRule::LessThanOrEqual: return len <= v1;
            }
            break;
        }
        case DataValidationRule::Date: {
            QDate date = QDate::fromString(value, Qt::ISODate);
            if (!date.isValid()) date = QDate::fromString(value, "MM/dd/yyyy");
            if (!date.isValid()) date = QDate::fromString(value, "M/d/yyyy");
            if (!date.isValid()) return false; // Invalid date
            QDate d1 = QDate::fromString(rule->value1, Qt::ISODate);
            QDate d2 = QDate::fromString(rule->value2, Qt::ISODate);
            switch (rule->op) {
                case DataValidationRule::Between: return date >= d1 && date <= d2;
                case DataValidationRule::NotBetween: return date < d1 || date > d2;
                case DataValidationRule::EqualTo: return date == d1;
                case DataValidationRule::NotEqualTo: return date != d1;
                case DataValidationRule::GreaterThan: return date > d1;
                case DataValidationRule::LessThan: return date < d1;
                case DataValidationRule::GreaterThanOrEqual: return date >= d1;
                case DataValidationRule::LessThanOrEqual: return date <= d1;
                default: return true;
            }
            break;
        }
        case DataValidationRule::Custom: {
            if (rule->customFormula.isEmpty()) return true;
            // Evaluate the custom formula — if it returns TRUE, validation passes
            QVariant result = m_formulaEngine->evaluate(rule->customFormula);
            return result.toBool();
        }
        default: return true;
    }
    return true;
}

// Named ranges
void Spreadsheet::addNamedRange(const QString& name, const CellRange& range, int sheetIndex) {
    NamedRange nr;
    nr.name = name;
    nr.range = range;
    nr.sheetIndex = sheetIndex;
    nr.isGlobal = (sheetIndex == -1);
    m_namedRanges[name.toUpper()] = nr;
}

void Spreadsheet::removeNamedRange(const QString& name) {
    m_namedRanges.erase(name.toUpper());
}

const NamedRange* Spreadsheet::getNamedRange(const QString& name) const {
    auto it = m_namedRanges.find(name.toUpper());
    if (it != m_namedRanges.end()) return &it->second;
    return nullptr;
}

const std::map<QString, NamedRange>& Spreadsheet::getNamedRanges() const {
    return m_namedRanges;
}

// ============== Dynamic Array Spill Support ==============
void Spreadsheet::clearSpillRange(const CellAddress& formulaCell) {
    CellKey key{formulaCell.row, formulaCell.col};
    auto it = m_spillRanges.find(key);
    if (it == m_spillRanges.end()) return;

    CellRange range = it->second;
    // Clear all spill child cells (not the formula cell itself)
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            if (r == formulaCell.row && c == formulaCell.col) continue;
            if (m_columnStore.isSpillCell(r, c)) {
                CellAddress parent = m_columnStore.getSpillParent(r, c);
                if (parent.row == formulaCell.row && parent.col == formulaCell.col) {
                    m_columnStore.removeCell(r, c);
                    m_columnStore.clearSpillParent(r, c);
                }
            }
        }
    }
    m_spillRanges.erase(it);
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::applySpillResult(const CellAddress& formulaCell,
                                    const std::vector<std::vector<QVariant>>& result) {
    // 1. Clear any previous spill from this formula cell
    clearSpillRange(formulaCell);

    if (result.empty() || result[0].empty()) return;

    int rows = static_cast<int>(result.size());
    int cols = static_cast<int>(result[0].size());

    // Single cell result - no spill needed
    if (rows == 1 && cols == 1) return;

    int startRow = formulaCell.row;
    int startCol = formulaCell.col;

    // 2. Check if target area is free
    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols; ++c) {
            if (r == 0 && c == 0) continue;
            int targetRow = startRow + r;
            int targetCol = startCol + c;
            if (m_columnStore.hasCell(targetRow, targetCol)) {
                CellDataType type = m_columnStore.getCellType(targetRow, targetCol);
                if (type != CellDataType::Empty && !m_columnStore.isSpillCell(targetRow, targetCol)) {
                    // Blocked - set #SPILL! error
                    m_columnStore.setComputedValue(formulaCell.row, formulaCell.col, QVariant("#SPILL!"));
                    return;
                }
            }
        }
    }

    // 3. Fill in spill cells
    for (int r = 0; r < rows; ++r) {
        int numCols = static_cast<int>(result[r].size());
        for (int c = 0; c < numCols; ++c) {
            if (r == 0 && c == 0) continue;
            int targetRow = startRow + r;
            int targetCol = startCol + c;
            m_columnStore.setCellValue(targetRow, targetCol, result[r][c]);
            m_columnStore.setSpillParent(targetRow, targetCol, formulaCell);
        }
    }

    // Record the spill range
    CellKey key{formulaCell.row, formulaCell.col};
    m_spillRanges[key] = CellRange(startRow, startCol, startRow + rows - 1, startCol + cols - 1);
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

// ============== Row/Column Grouping (Outline) ==============

void Spreadsheet::groupRows(int startRow, int endRow) {
    for (int r = startRow; r <= endRow; ++r) {
        int level = m_rowOutlineLevels[r] + 1;
        if (level > 8) level = 8; // Excel max outline level
        m_rowOutlineLevels[r] = level;
    }
}

void Spreadsheet::ungroupRows(int startRow, int endRow) {
    for (int r = startRow; r <= endRow; ++r) {
        auto it = m_rowOutlineLevels.find(r);
        if (it != m_rowOutlineLevels.end()) {
            it->second--;
            if (it->second <= 0) {
                m_rowOutlineLevels.erase(it);
            }
        }
    }
}

void Spreadsheet::groupColumns(int startCol, int endCol) {
    for (int c = startCol; c <= endCol; ++c) {
        int level = m_colOutlineLevels[c] + 1;
        if (level > 8) level = 8;
        m_colOutlineLevels[c] = level;
    }
}

void Spreadsheet::ungroupColumns(int startCol, int endCol) {
    for (int c = startCol; c <= endCol; ++c) {
        auto it = m_colOutlineLevels.find(c);
        if (it != m_colOutlineLevels.end()) {
            it->second--;
            if (it->second <= 0) {
                m_colOutlineLevels.erase(it);
            }
        }
    }
}

int Spreadsheet::getRowOutlineLevel(int row) const {
    auto it = m_rowOutlineLevels.find(row);
    return (it != m_rowOutlineLevels.end()) ? it->second : 0;
}

int Spreadsheet::getColumnOutlineLevel(int col) const {
    auto it = m_colOutlineLevels.find(col);
    return (it != m_colOutlineLevels.end()) ? it->second : 0;
}

void Spreadsheet::setRowOutlineCollapsed(int row, bool collapsed) {
    if (collapsed) m_collapsedRowGroups.insert(row);
    else m_collapsedRowGroups.erase(row);
}

void Spreadsheet::setColumnOutlineCollapsed(int col, bool collapsed) {
    if (collapsed) m_collapsedColGroups.insert(col);
    else m_collapsedColGroups.erase(col);
}

bool Spreadsheet::isRowOutlineCollapsed(int row) const {
    return m_collapsedRowGroups.count(row) > 0;
}

bool Spreadsheet::isColumnOutlineCollapsed(int col) const {
    return m_collapsedColGroups.count(col) > 0;
}

int Spreadsheet::getMaxRowOutlineLevel() const {
    int maxLevel = 0;
    for (const auto& [row, level] : m_rowOutlineLevels) {
        if (level > maxLevel) maxLevel = level;
    }
    return maxLevel;
}

int Spreadsheet::getMaxColumnOutlineLevel() const {
    int maxLevel = 0;
    for (const auto& [col, level] : m_colOutlineLevels) {
        if (level > maxLevel) maxLevel = level;
    }
    return maxLevel;
}

void Spreadsheet::toggleRowGroup(int groupEndRow, int level) {
    // Find all consecutive rows at or above `level` ending at groupEndRow
    // Walk backward from groupEndRow to find the group start
    int groupStart = groupEndRow;
    while (groupStart > 0) {
        int prevLevel = getRowOutlineLevel(groupStart - 1);
        if (prevLevel >= level) {
            groupStart--;
        } else {
            break;
        }
    }

    // Check if this group is currently collapsed
    bool isCollapsed = m_collapsedRowGroups.count(groupEndRow) > 0;

    if (isCollapsed) {
        // Expand: restore row heights (show rows)
        m_collapsedRowGroups.erase(groupEndRow);
        for (int r = groupStart; r <= groupEndRow; ++r) {
            if (getRowOutlineLevel(r) >= level) {
                // Only restore if not part of a deeper nested collapsed group
                bool nestedCollapsed = false;
                for (int checkRow : m_collapsedRowGroups) {
                    if (checkRow >= groupStart && checkRow <= groupEndRow) {
                        // There's a nested collapsed group; check if row r is in it
                        int nestedStart = checkRow;
                        int nestedLevel = getRowOutlineLevel(checkRow);
                        while (nestedStart > groupStart && getRowOutlineLevel(nestedStart - 1) >= nestedLevel) {
                            nestedStart--;
                        }
                        if (r >= nestedStart && r <= checkRow && getRowOutlineLevel(r) >= nestedLevel) {
                            nestedCollapsed = true;
                            break;
                        }
                    }
                }
                if (!nestedCollapsed) {
                    // Restore default height (remove the 0-height marker)
                    auto it = m_rowHeights.find(r);
                    if (it != m_rowHeights.end() && it->second == 0) {
                        m_rowHeights.erase(it);
                    }
                }
            }
        }
    } else {
        // Collapse: hide rows by setting height to 0
        m_collapsedRowGroups.insert(groupEndRow);
        for (int r = groupStart; r <= groupEndRow; ++r) {
            if (getRowOutlineLevel(r) >= level) {
                m_rowHeights[r] = 0;
            }
        }
    }
}

void Spreadsheet::toggleColumnGroup(int groupEndCol, int level) {
    // Find group start by walking backward
    int groupStart = groupEndCol;
    while (groupStart > 0) {
        int prevLevel = getColumnOutlineLevel(groupStart - 1);
        if (prevLevel >= level) {
            groupStart--;
        } else {
            break;
        }
    }

    bool isCollapsed = m_collapsedColGroups.count(groupEndCol) > 0;

    if (isCollapsed) {
        // Expand: restore column widths
        m_collapsedColGroups.erase(groupEndCol);
        for (int c = groupStart; c <= groupEndCol; ++c) {
            if (getColumnOutlineLevel(c) >= level) {
                bool nestedCollapsed = false;
                for (int checkCol : m_collapsedColGroups) {
                    if (checkCol >= groupStart && checkCol <= groupEndCol) {
                        int nestedStart = checkCol;
                        int nestedLevel = getColumnOutlineLevel(checkCol);
                        while (nestedStart > groupStart && getColumnOutlineLevel(nestedStart - 1) >= nestedLevel) {
                            nestedStart--;
                        }
                        if (c >= nestedStart && c <= checkCol && getColumnOutlineLevel(c) >= nestedLevel) {
                            nestedCollapsed = true;
                            break;
                        }
                    }
                }
                if (!nestedCollapsed) {
                    auto it = m_columnWidths.find(c);
                    if (it != m_columnWidths.end() && it->second == 0) {
                        m_columnWidths.erase(it);
                    }
                }
            }
        }
    } else {
        // Collapse: hide columns by setting width to 0
        m_collapsedColGroups.insert(groupEndCol);
        for (int c = groupStart; c <= groupEndCol; ++c) {
            if (getColumnOutlineLevel(c) >= level) {
                m_columnWidths[c] = 0;
            }
        }
    }
}
