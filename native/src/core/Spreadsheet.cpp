#include "Spreadsheet.h"
#include "PivotEngine.h"
#include <algorithm>

Spreadsheet::Spreadsheet()
    : m_sheetName("Sheet1"), m_rowCount(1000), m_columnCount(256),
      m_autoRecalculate(true), m_inTransaction(false) {
    m_formulaEngine = std::make_unique<FormulaEngine>(this);
    m_cells.reserve(4096);
}

Spreadsheet::~Spreadsheet() = default;

std::shared_ptr<Cell> Spreadsheet::getCell(const CellAddress& addr) {
    return getCell(addr.row, addr.col);
}

std::shared_ptr<Cell> Spreadsheet::getCell(int row, int col) {
    CellKey key{row, col};
    auto it = m_cells.find(key);
    if (it != m_cells.end()) {
        return it->second;
    }
    auto cell = std::make_shared<Cell>();
    m_cells.emplace(key, cell);
    m_maxRowColDirty = true;
    return cell;
}

std::shared_ptr<Cell> Spreadsheet::getCellIfExists(const CellAddress& addr) const {
    return getCellIfExists(addr.row, addr.col);
}

std::shared_ptr<Cell> Spreadsheet::getCellIfExists(int row, int col) const {
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
    auto cell = getCell(addr);
    cell->setValue(value);
    m_maxRowColDirty = true;

    // Skip dependency graph work when autoRecalculate is off (bulk import mode)
    if (m_autoRecalculate) {
        m_depGraph.removeDependencies(addr);
        if (!m_inTransaction) {
            recalculateDependents(addr);
        }
    }
}

void Spreadsheet::setCellFormula(const CellAddress& addr, const QString& formula) {
    auto cell = getCell(addr);
    cell->setFormula(formula);
    m_maxRowColDirty = true;
    updateDependencies(addr);

    if (m_depGraph.hasCircularDependency(addr)) {
        cell->setComputedValue(QVariant("#CIRCULAR!"));
        return;
    }

    if (m_autoRecalculate && !m_inTransaction) {
        recalculate(addr);
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
    m_rowCount += count;
    m_maxRowColDirty = true;
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
    m_columnCount += count;
    m_maxRowColDirty = true;
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
    m_rowCount -= count;
    m_maxRowColDirty = true;
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
    m_columnCount -= count;
    m_maxRowColDirty = true;
}

QString Spreadsheet::getSheetName() const { return m_sheetName; }
void Spreadsheet::setSheetName(const QString& name) { m_sheetName = name; }

void Spreadsheet::updateMaxRowCol() const {
    if (!m_maxRowColDirty) return;
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
    for (auto& pair : m_cells) {
        if (pair.second->getType() == CellType::Formula) {
            CellAddress addr(pair.first.row, pair.first.col);
            pair.second->setComputedValue(m_formulaEngine->evaluate(pair.second->getFormula()));
            updateDependencies(addr);
        }
    }
}

void Spreadsheet::updateDependencies(const CellAddress& addr) {
    m_depGraph.removeDependencies(addr);
    auto cell = getCellIfExists(addr);
    if (cell && cell->getType() == CellType::Formula) {
        m_formulaEngine->evaluate(cell->getFormula());
        for (const auto& dep : m_formulaEngine->getLastDependencies()) {
            m_depGraph.addDependency(addr, dep);
        }
    }
}

void Spreadsheet::recalculateDependents(const CellAddress& addr) {
    auto order = m_depGraph.getRecalcOrder(addr);
    for (const auto& depAddr : order) {
        auto cell = getCellIfExists(depAddr);
        if (cell && cell->getType() == CellType::Formula) {
            cell->setComputedValue(m_formulaEngine->evaluate(cell->getFormula()));
        }
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
