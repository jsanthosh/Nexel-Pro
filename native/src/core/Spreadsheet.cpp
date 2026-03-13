#include "Spreadsheet.h"
#include "PivotEngine.h"
#include <algorithm>
#include <mutex>
#include <QRegularExpression>

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
    m_cells.reserve(4096);
    m_conditionalFormatting.setFormulaEvaluator([this](const QString& formula) -> QVariant {
        return m_formulaEngine->evaluate(formula);
    });
}

Spreadsheet::~Spreadsheet() = default;

std::shared_ptr<Cell> Spreadsheet::getCell(const CellAddress& addr) {
    return getCell(addr.row, addr.col);
}

std::shared_ptr<Cell> Spreadsheet::getCell(int row, int col) {
    std::unique_lock lock(m_cellMutex);
    CellKey key{row, col};
    auto it = m_cells.find(key);
    if (it != m_cells.end()) {
        return it->second;
    }
    auto cell = std::make_shared<Cell>();
    m_cells.emplace(key, cell);
    if (row > m_cachedMaxRow || col > m_cachedMaxCol) {
        m_maxRowColDirty = true;
    }
    return cell;
}

std::shared_ptr<Cell> Spreadsheet::getCellIfExists(const CellAddress& addr) const {
    return getCellIfExists(addr.row, addr.col);
}

std::shared_ptr<Cell> Spreadsheet::getCellIfExists(int row, int col) const {
    std::shared_lock lock(m_cellMutex);
    CellKey key{row, col};
    auto it = m_cells.find(key);
    return (it != m_cells.end()) ? it->second : nullptr;
}

QVariant Spreadsheet::getCellValue(const CellAddress& addr) {
    auto cell = getCellIfExists(addr.row, addr.col);
    if (!cell) return QVariant();
    if (cell->getType() == CellType::Formula) {
        return cell->getComputedValue();
    }
    return cell->getValue();
}

void Spreadsheet::setCellValue(const CellAddress& addr, const QVariant& value) {
    bool wasNew = (m_cells.find(CellKey{addr.row, addr.col}) == m_cells.end());

    // Clear any existing spill range from this cell (was previously a spill-producing formula)
    clearSpillRange(addr);

    auto cell = getCell(addr);
    cell->setValue(value);
    // Only mark dirty if this cell extends beyond cached bounds
    if (addr.row > m_cachedMaxRow || addr.col > m_cachedMaxCol) {
        m_maxRowColDirty = true;
    }

    // Remove from formula tracking (it's now a value cell)
    m_formulaCells.erase(CellKey{addr.row, addr.col});

    // Incremental nav index update instead of full rebuild
    if (!m_navIndexDirty && wasNew) {
        navIndexInsert(addr.row, addr.col);
    } else if (m_autoRecalculate) {
        // If nav was already dirty (bulk mode), keep it dirty
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
    bool wasNew = (m_cells.find(CellKey{addr.row, addr.col}) == m_cells.end());
    auto cell = getCell(addr);
    cell->setFormula(formula);
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
        cell->setComputedValue(QVariant("#CIRCULAR!"));
        return;
    }

    cell->setComputedValue(result);

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
                cell->setComputedValue(array2D[0][0]);
            }
        }
    }

    if (m_autoRecalculate && !m_inTransaction) {
        recalculateDependents(addr);
    }
}

void Spreadsheet::fillRange(const CellRange& range, const QVariant& value) {
    for (const auto& addr : range.getCells()) {
        setCellValue(addr, value);
    }
}

void Spreadsheet::clearRange(const CellRange& range) {
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            CellKey key{r, c};
            auto it = m_cells.find(key);
            if (it != m_cells.end()) {
                it->second->clear();
            }
        }
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

std::vector<std::shared_ptr<Cell>> Spreadsheet::getRange(const CellRange& range) {
    std::vector<std::shared_ptr<Cell>> result;
    for (const auto& addr : range.getCells()) {
        result.push_back(getCell(addr));
    }
    return result;
}

void Spreadsheet::insertRow(int row, int count) {
    std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> toReinsert;
    std::vector<CellKey> toRemove;
    for (const auto& pair : m_cells) {
        if (pair.first.row >= row) {
            toRemove.push_back(pair.first);
            toReinsert.push_back({CellKey{pair.first.row + count, pair.first.col}, pair.second});
        }
    }
    for (const auto& key : toRemove) m_cells.erase(key);
    for (auto& [key, cell] : toReinsert) m_cells.emplace(key, std::move(cell));
    // Shift formula references
    for (auto& [key, cell] : m_cells) {
        if (cell->getType() == CellType::Formula && !cell->getFormula().isEmpty()) {
            QString adjusted = shiftFormulaRefs(cell->getFormula(), 'R', row, count);
            if (adjusted != cell->getFormula()) cell->setFormula(adjusted);
        }
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
    m_depGraph.clear();
    rebuildDependencyGraph();
}

void Spreadsheet::insertColumn(int column, int count) {
    std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> toReinsert;
    std::vector<CellKey> toRemove;
    for (const auto& pair : m_cells) {
        if (pair.first.col >= column) {
            toRemove.push_back(pair.first);
            toReinsert.push_back({CellKey{pair.first.row, pair.first.col + count}, pair.second});
        }
    }
    for (const auto& key : toRemove) m_cells.erase(key);
    for (auto& [key, cell] : toReinsert) m_cells.emplace(key, std::move(cell));
    for (auto& [key, cell] : m_cells) {
        if (cell->getType() == CellType::Formula && !cell->getFormula().isEmpty()) {
            QString adjusted = shiftFormulaRefs(cell->getFormula(), 'C', column, count);
            if (adjusted != cell->getFormula()) cell->setFormula(adjusted);
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
    std::vector<CellKey> toRemove;
    std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> toReinsert;
    for (const auto& pair : m_cells) {
        if (pair.first.row >= row && pair.first.row < row + count) {
            toRemove.push_back(pair.first);
        } else if (pair.first.row >= row + count) {
            toRemove.push_back(pair.first);
            toReinsert.push_back({CellKey{pair.first.row - count, pair.first.col}, pair.second});
        }
    }
    for (const auto& key : toRemove) m_cells.erase(key);
    for (auto& [key, cell] : toReinsert) m_cells.emplace(key, std::move(cell));
    for (auto& [key, cell] : m_cells) {
        if (cell->getType() == CellType::Formula && !cell->getFormula().isEmpty()) {
            QString adjusted = shiftFormulaRefs(cell->getFormula(), 'R', row, -count);
            if (adjusted != cell->getFormula()) cell->setFormula(adjusted);
        }
    }
    // Remove merged regions fully in deleted rows; clamp partially overlapping ones
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [row, count](const MergedRegion& mr) {
            return mr.range.getStart().row >= row && mr.range.getEnd().row < row + count;
        }), m_mergedRegions.end());
    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        // Clamp start if it falls within the deleted range
        if (s.row >= row && s.row < row + count) s.row = row;
        else if (s.row >= row + count) s.row -= count;
        // Clamp end if it falls within the deleted range
        if (e.row >= row && e.row < row + count) e.row = row > 0 ? row - 1 : 0;
        else if (e.row >= row + count) e.row -= count;
        // Ensure region is still valid (start <= end)
        if (s.row > e.row) { s.row = e.row; }
        mr.range = CellRange(s, e);
    }
    // Remove any merged regions that collapsed to invalid state
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
    m_depGraph.clear();
    rebuildDependencyGraph();
}

void Spreadsheet::deleteColumn(int column, int count) {
    std::vector<CellKey> toRemove;
    std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> toReinsert;
    for (const auto& pair : m_cells) {
        if (pair.first.col >= column && pair.first.col < column + count) {
            toRemove.push_back(pair.first);
        } else if (pair.first.col >= column + count) {
            toRemove.push_back(pair.first);
            toReinsert.push_back({CellKey{pair.first.row, pair.first.col - count}, pair.second});
        }
    }
    for (const auto& key : toRemove) m_cells.erase(key);
    for (auto& [key, cell] : toReinsert) m_cells.emplace(key, std::move(cell));
    for (auto& [key, cell] : m_cells) {
        if (cell->getType() == CellType::Formula && !cell->getFormula().isEmpty()) {
            QString adjusted = shiftFormulaRefs(cell->getFormula(), 'C', column, -count);
            if (adjusted != cell->getFormula()) cell->setFormula(adjusted);
        }
    }
    m_mergedRegions.erase(std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
        [column, count](const MergedRegion& mr) {
            return mr.range.getStart().col >= column && mr.range.getEnd().col < column + count;
        }), m_mergedRegions.end());
    for (auto& mr : m_mergedRegions) {
        auto s = mr.range.getStart(); auto e = mr.range.getEnd();
        // Clamp start if it falls within the deleted range
        if (s.col >= column && s.col < column + count) s.col = column;
        else if (s.col >= column + count) s.col -= count;
        // Clamp end if it falls within the deleted range
        if (e.col >= column && e.col < column + count) e.col = column > 0 ? column - 1 : 0;
        else if (e.col >= column + count) e.col -= count;
        if (s.col > e.col) { s.col = e.col; }
        mr.range = CellRange(s, e);
    }
    // Remove collapsed merged regions
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
    // Shift column widths: remove deleted cols, shift remaining left
    {
        std::map<int, int> shifted;
        for (auto& [c, w] : m_columnWidths) {
            if (c >= column && c < column + count) continue; // deleted
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
    std::shared_lock lock(m_cellMutex);
    m_cachedMaxRow = -1;
    m_cachedMaxCol = -1;
    for (const auto& pair : m_cells) {
        if (pair.second && pair.second->getType() != CellType::Empty) {
            if (pair.first.row > m_cachedMaxRow) m_cachedMaxRow = pair.first.row;
            if (pair.first.col > m_cachedMaxCol) m_cachedMaxCol = pair.first.col;
        }
    }
    m_maxRowColDirty = false;
}

int Spreadsheet::getMaxRow() const { updateMaxRowCol(); return m_cachedMaxRow; }
int Spreadsheet::getMaxColumn() const { updateMaxRowCol(); return m_cachedMaxCol; }

std::vector<CellAddress> Spreadsheet::getDirtyCells() const {
    std::shared_lock lock(m_cellMutex);
    std::vector<CellAddress> dirty;
    for (const auto& pair : m_cells) {
        if (pair.second->isDirty()) {
            dirty.push_back(CellAddress(pair.first.row, pair.first.col));
        }
    }
    return dirty;
}

void Spreadsheet::clearDirtyFlag() {
    for (auto& pair : m_cells) pair.second->setDirty(false);
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
    // First, recalculate the dirty cells themselves (they may be formula cells)
    std::vector<CellAddress> allDirty;
    allDirty.reserve(m_batchDirtyCells.size());
    for (const auto& key : m_batchDirtyCells) {
        allDirty.emplace_back(key.row, key.col);
    }

    // Recalculate all formula cells that were directly changed
    for (const auto& addr : allDirty) {
        auto cell = getCellIfExists(addr);
        if (cell && cell->getType() == CellType::Formula) {
            cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
        }
    }

    // Collect all dependents of dirty cells and recalculate in topological order
    std::unordered_set<uint64_t> visited;
    std::vector<CellAddress> recalcOrder;
    for (const auto& addr : allDirty) {
        auto order = m_depGraph.getRecalcOrder(addr);
        for (const auto& depAddr : order) {
            uint64_t key = (static_cast<uint64_t>(static_cast<uint32_t>(depAddr.row)) << 32)
                         | static_cast<uint64_t>(static_cast<uint32_t>(depAddr.col));
            if (visited.insert(key).second) {
                recalcOrder.push_back(depAddr);
            }
        }
    }

    std::vector<CellAddress> recalculated;
    for (const auto& depAddr : recalcOrder) {
        auto cell = getCellIfExists(depAddr);
        if (cell && cell->getType() == CellType::Formula) {
            cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
            recalculated.push_back(depAddr);
        }
    }

    if (!recalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(recalculated);
    }

    m_batchDirtyCells.clear();
    m_navIndexDirty = true;
}

Cell* Spreadsheet::getOrCreateCellFast(int row, int col) {
    // No mutex — main-thread-only bulk import path
    auto& ptr = m_cells[CellKey{row, col}];
    if (!ptr) {
        ptr = std::make_shared<Cell>();
        // Incrementally build nav index during import (appended in order, sorted at finish)
        m_colIndexCache[col].push_back(row);
        m_rowIndexCache[row].push_back(col);
    }
    return ptr.get();
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

    // Nav index was built incrementally during getOrCreateCellFast.
    // CSV rows are parsed sequentially, so vectors are already sorted —
    // std::sort on sorted data is O(n) verification.
    for (auto& [c, rows] : m_colIndexCache)
        std::sort(rows.begin(), rows.end());
    for (auto& [r, cols] : m_rowIndexCache)
        std::sort(cols.begin(), cols.end());

    m_sortedOccupiedRows.clear();
    m_sortedOccupiedRows.reserve(m_rowIndexCache.size());
    for (const auto& [r, _] : m_rowIndexCache)
        m_sortedOccupiedRows.push_back(r);
    std::sort(m_sortedOccupiedRows.begin(), m_sortedOccupiedRows.end());

    m_navIndexDirty = false;
}

void Spreadsheet::rebuildDependencyGraph() {
    m_colRefFormulas.clear();
    for (auto& [key, cell] : m_cells) {
        if (cell->getType() == CellType::Formula && !cell->getFormula().isEmpty()) {
            CellAddress addr(key.row, key.col);
            m_formulaEngine->evaluate(cell->getFormula());
            for (const auto& dep : m_formulaEngine->getLastDependencies()) {
                m_depGraph.addDependency(addr, dep);
            }
            for (int col : m_formulaEngine->getLastColumnDeps()) {
                m_colRefFormulas[col].push_back(addr);
            }
        }
    }
}

void Spreadsheet::mergeBulkCells(CellMap& source) {
    std::unique_lock lock(m_cellMutex);
    m_cells.merge(source);
    // Any remaining entries in source (duplicates) — move them over
    for (auto& [key, cell] : source) {
        m_cells[key] = std::move(cell);
    }
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
    auto cell = getCell(addr);
    CellSnapshot snap;
    snap.addr = addr;
    snap.value = cell->getValue();
    snap.formula = cell->getFormula();
    snap.style = cell->getStyle();
    snap.type = cell->getType();
    return snap;
}

void Spreadsheet::forEachCell(std::function<void(int row, int col, const Cell&)> callback) const {
    std::shared_lock lock(m_cellMutex);
    for (const auto& pair : m_cells) {
        if (pair.second && pair.second->getType() != CellType::Empty) {
            callback(pair.first.row, pair.first.col, *pair.second);
        }
    }
}

void Spreadsheet::recalculate(const CellAddress& addr) {
    auto cell = getCellIfExists(addr);
    if (cell && cell->getType() == CellType::Formula) {
        cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
    }
}

void Spreadsheet::recalculateAll() {
    // Use formula cell tracker for O(f) instead of scanning all O(n) cells
    if (!m_formulaCells.empty()) {
        for (const auto& key : m_formulaCells) {
            auto it = m_cells.find(key);
            if (it != m_cells.end() && it->second->getType() == CellType::Formula) {
                it->second->setComputedValue(m_formulaEngine->evaluate(it->second->getFormula()));
            }
        }
    } else {
        // Fallback: scan all cells (only needed if formula tracker wasn't populated)
        for (auto& pair : m_cells) {
            if (pair.second->getType() == CellType::Formula) {
                pair.second->setComputedValue(m_formulaEngine->evaluate(pair.second->getFormula()));
                m_formulaCells.insert(pair.first);
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

    auto cell = getCellIfExists(addr);
    if (cell && cell->getType() == CellType::Formula) {
        m_formulaEngine->evaluate(cell->getFormula());
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
        auto cell = getCellIfExists(depAddr);
        if (cell && cell->getType() == CellType::Formula) {
            cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
            recalculated.push_back(depAddr);
        }
    }
    if (!recalculated.empty() && onDependentsRecalculated) {
        onDependentsRecalculated(recalculated);
    }
}

void Spreadsheet::recalculateColumnDependents(int col) {
    auto it = m_colRefFormulas.find(col);
    if (it == m_colRefFormulas.end() || it->second.empty()) return;

    std::vector<CellAddress> recalculated;
    for (const auto& depAddr : it->second) {
        auto cell = getCellIfExists(depAddr);
        if (cell && cell->getType() == CellType::Formula) {
            cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
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

    struct RowData {
        QVariant sortValue;
        std::vector<std::pair<int, std::shared_ptr<Cell>>> cells;
    };

    std::vector<RowData> rows;
    for (int r = startRow; r <= endRow; ++r) {
        RowData rd;
        rd.sortValue = getCellValue(CellAddress(r, sortColumn));
        for (int c = startCol; c <= endCol; ++c) {
            auto it = m_cells.find(CellKey{r, c});
            if (it != m_cells.end()) rd.cells.push_back({c, it->second});
        }
        rows.push_back(std::move(rd));
    }

    std::sort(rows.begin(), rows.end(), [ascending](const RowData& a, const RowData& b) {
        bool aEmpty = !a.sortValue.isValid() || a.sortValue.toString().isEmpty();
        bool bEmpty = !b.sortValue.isValid() || b.sortValue.toString().isEmpty();
        if (aEmpty && bEmpty) return false;
        if (aEmpty) return false;
        if (bEmpty) return true;
        bool aOk, bOk;
        double aNum = a.sortValue.toString().toDouble(&aOk);
        double bNum = b.sortValue.toString().toDouble(&bOk);
        if (aOk && bOk) return ascending ? (aNum < bNum) : (aNum > bNum);
        int cmp = a.sortValue.toString().compare(b.sortValue.toString(), Qt::CaseInsensitive);
        return ascending ? (cmp < 0) : (cmp > 0);
    });

    for (int r = startRow; r <= endRow; ++r)
        for (int c = startCol; c <= endCol; ++c)
            m_cells.erase(CellKey{r, c});

    for (int i = 0; i < static_cast<int>(rows.size()); ++i) {
        int targetRow = startRow + i;
        for (auto& [col, cell] : rows[i].cells)
            m_cells[CellKey{targetRow, col}] = cell;
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::insertCellsShiftRight(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int colCount = range.getEnd().col - startCol + 1;
    for (int r = startRow; r <= endRow; ++r) {
        std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> ri;
        std::vector<CellKey> rm;
        for (const auto& p : m_cells) {
            if (p.first.row == r && p.first.col >= startCol) {
                rm.push_back(p.first);
                ri.push_back({CellKey{r, p.first.col + colCount}, p.second});
            }
        }
        for (const auto& k : rm) m_cells.erase(k);
        for (auto& [k, c] : ri) m_cells.emplace(k, std::move(c));
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::insertCellsShiftDown(const CellRange& range) {
    int startRow = range.getStart().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int rowCount = range.getEnd().row - startRow + 1;
    for (int c = startCol; c <= endCol; ++c) {
        std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> ri;
        std::vector<CellKey> rm;
        for (const auto& p : m_cells) {
            if (p.first.col == c && p.first.row >= startRow) {
                rm.push_back(p.first);
                ri.push_back({CellKey{p.first.row + rowCount, c}, p.second});
            }
        }
        for (const auto& k : rm) m_cells.erase(k);
        for (auto& [k, cl] : ri) m_cells.emplace(k, std::move(cl));
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::deleteCellsShiftLeft(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int colCount = endCol - startCol + 1;
    for (int r = startRow; r <= endRow; ++r) {
        for (int c = startCol; c <= endCol; ++c) m_cells.erase(CellKey{r, c});
        std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> ri;
        std::vector<CellKey> rm;
        for (const auto& p : m_cells) {
            if (p.first.row == r && p.first.col > endCol) {
                rm.push_back(p.first);
                ri.push_back({CellKey{r, p.first.col - colCount}, p.second});
            }
        }
        for (const auto& k : rm) m_cells.erase(k);
        for (auto& [k, c] : ri) m_cells.emplace(k, std::move(c));
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

void Spreadsheet::deleteCellsShiftUp(const CellRange& range) {
    int startRow = range.getStart().row, endRow = range.getEnd().row;
    int startCol = range.getStart().col, endCol = range.getEnd().col;
    int rowCount = endRow - startRow + 1;
    for (int c = startCol; c <= endCol; ++c) {
        for (int r = startRow; r <= endRow; ++r) m_cells.erase(CellKey{r, c});
        std::vector<std::pair<CellKey, std::shared_ptr<Cell>>> ri;
        std::vector<CellKey> rm;
        for (const auto& p : m_cells) {
            if (p.first.col == c && p.first.row > endRow) {
                rm.push_back(p.first);
                ri.push_back({CellKey{p.first.row - rowCount, c}, p.second});
            }
        }
        for (const auto& k : rm) m_cells.erase(k);
        for (auto& [k, cl] : ri) m_cells.emplace(k, std::move(cl));
    }
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}

// ============== Document Theme ==============
void Spreadsheet::setDocumentTheme(const DocumentTheme& theme) {
    m_documentTheme = theme;

    // Re-theme existing tables: regenerate table themes from the new document theme
    // and re-assign each table to its corresponding accent-based theme
    auto newThemes = generateTableThemes(theme);
    for (auto& table : m_tables) {
        // Try to match the table's current theme name to a generated theme
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
    // Also rebuild formula tracker during full scan (avoids separate O(n) pass)
    auto& formulaCells = const_cast<std::unordered_set<CellKey, CellKeyHash>&>(m_formulaCells);
    formulaCells.clear();
    for (const auto& [key, cell] : m_cells) {
        if (cell && cell->getType() != CellType::Empty) {
            m_colIndexCache[key.col].push_back(key.row);
            m_rowIndexCache[key.row].push_back(key.col);
            if (cell->getType() == CellType::Formula) {
                formulaCells.insert(key);
            }
        }
    }
    for (auto& [col, rows] : m_colIndexCache)
        std::sort(rows.begin(), rows.end());
    for (auto& [row, cols] : m_rowIndexCache)
        std::sort(cols.begin(), cols.end());
    m_sortedOccupiedRows.clear();
    m_sortedOccupiedRows.reserve(m_rowIndexCache.size());
    for (const auto& [row, _] : m_rowIndexCache)
        m_sortedOccupiedRows.push_back(row);
    std::sort(m_sortedOccupiedRows.begin(), m_sortedOccupiedRows.end());
    m_navIndexDirty = false;
}

void Spreadsheet::navIndexInsert(int row, int col) const {
    if (m_navIndexDirty) return; // Will rebuild fully later
    // Insert row into column index (maintain sorted order)
    auto& colRows = m_colIndexCache[col];
    auto it = std::lower_bound(colRows.begin(), colRows.end(), row);
    if (it == colRows.end() || *it != row) colRows.insert(it, row);

    // Insert col into row index (maintain sorted order)
    auto& rowCols = m_rowIndexCache[row];
    auto it2 = std::lower_bound(rowCols.begin(), rowCols.end(), col);
    if (it2 == rowCols.end() || *it2 != col) rowCols.insert(it2, col);

    // Insert into sorted occupied rows
    auto it3 = std::lower_bound(m_sortedOccupiedRows.begin(), m_sortedOccupiedRows.end(), row);
    if (it3 == m_sortedOccupiedRows.end() || *it3 != row) m_sortedOccupiedRows.insert(it3, row);
}

void Spreadsheet::navIndexRemove(int row, int col) const {
    if (m_navIndexDirty) return;
    // Remove row from column index
    auto colIt = m_colIndexCache.find(col);
    if (colIt != m_colIndexCache.end()) {
        auto& rows = colIt->second;
        auto it = std::lower_bound(rows.begin(), rows.end(), row);
        if (it != rows.end() && *it == row) rows.erase(it);
        if (rows.empty()) m_colIndexCache.erase(colIt);
    }
    // Remove col from row index
    auto rowIt = m_rowIndexCache.find(row);
    if (rowIt != m_rowIndexCache.end()) {
        auto& cols = rowIt->second;
        auto it = std::lower_bound(cols.begin(), cols.end(), col);
        if (it != cols.end() && *it == col) cols.erase(it);
        if (cols.empty()) {
            m_rowIndexCache.erase(rowIt);
            // Remove from sorted occupied rows
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
    buildNavIndexIfNeeded();
    auto colIt = m_colIndexCache.find(col);
    if (colIt == m_colIndexCache.end()) return;
    const auto& rows = colIt->second;
    auto lo = std::lower_bound(rows.begin(), rows.end(), startRow);
    auto hi = std::upper_bound(rows.begin(), rows.end(), endRow);
    for (auto it = lo; it != hi; ++it) {
        auto cellIt = m_cells.find(CellKey{*it, col});
        if (cellIt != m_cells.end() && cellIt->second) {
            const auto& cell = cellIt->second;
            fn(cell->getType() == CellType::Formula ?
                cell->getComputedValue() : cell->getValue());
        }
    }
}

const std::vector<int>& Spreadsheet::getOccupiedColsInRow(int row) const {
    static const std::vector<int> empty;
    buildNavIndexIfNeeded();
    auto it = m_rowIndexCache.find(row);
    if (it == m_rowIndexCache.end()) return empty;
    return it->second;
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
        case DataValidationRule::List: return rule->listItems.contains(value, Qt::CaseInsensitive);
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
            auto cell = getCellIfExists(r, c);
            if (cell && cell->isSpillCell() &&
                cell->getSpillParent().row == formulaCell.row &&
                cell->getSpillParent().col == formulaCell.col) {
                cell->clear();
                cell->clearSpillParent();
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
            if (r == 0 && c == 0) continue; // skip formula cell itself
            int targetRow = startRow + r;
            int targetCol = startCol + c;
            auto existing = getCellIfExists(targetRow, targetCol);
            if (existing && existing->getType() != CellType::Empty && !existing->isSpillCell()) {
                // Blocked - set #SPILL! error
                auto formulaCellPtr = getCell(formulaCell);
                formulaCellPtr->setComputedValue(QVariant("#SPILL!"));
                return;
            }
        }
    }

    // 3. Fill in spill cells
    for (int r = 0; r < rows; ++r) {
        int numCols = static_cast<int>(result[r].size());
        for (int c = 0; c < numCols; ++c) {
            if (r == 0 && c == 0) continue; // formula cell already has its value
            int targetRow = startRow + r;
            int targetCol = startCol + c;
            auto cell = getCell(targetRow, targetCol);
            cell->setValue(result[r][c]);
            cell->setSpillParent(formulaCell);
        }
    }

    // Record the spill range
    CellKey key{formulaCell.row, formulaCell.col};
    m_spillRanges[key] = CellRange(startRow, startCol, startRow + rows - 1, startCol + cols - 1);
    m_maxRowColDirty = true;
    m_navIndexDirty = true;
}
