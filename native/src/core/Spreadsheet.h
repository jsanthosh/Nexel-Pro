#ifndef SPREADSHEET_H
#define SPREADSHEET_H

#include <QString>
#include <QVariant>
#include <unordered_map>
#include <map>
#include <memory>
#include <vector>
#include <functional>
#include "Cell.h"
#include "CellRange.h"
#include "TableStyle.h"
#include "FormulaEngine.h"
#include "UndoManager.h"
#include "DependencyGraph.h"
#include "ConditionalFormatting.h"
#include "SparklineConfig.h"

struct PivotConfig; // forward declaration

class Spreadsheet {
public:
    Spreadsheet();
    ~Spreadsheet();

    // Cell access and modification
    std::shared_ptr<Cell> getCell(const CellAddress& addr);
    std::shared_ptr<Cell> getCell(int row, int col);
    // Read-only cell access - returns nullptr for non-existent cells (no allocation)
    std::shared_ptr<Cell> getCellIfExists(const CellAddress& addr) const;
    std::shared_ptr<Cell> getCellIfExists(int row, int col) const;
    QVariant getCellValue(const CellAddress& addr);
    void setCellValue(const CellAddress& addr, const QVariant& value);
    void setCellFormula(const CellAddress& addr, const QString& formula);

    // Range operations
    void fillRange(const CellRange& range, const QVariant& value);
    void clearRange(const CellRange& range);
    std::vector<std::shared_ptr<Cell>> getRange(const CellRange& range);

    // Row/Column operations
    void insertRow(int row, int count = 1);
    void insertColumn(int column, int count = 1);
    void deleteRow(int row, int count = 1);
    void deleteColumn(int column, int count = 1);

    // Sheet properties
    QString getSheetName() const;
    void setSheetName(const QString& name);
    int getMaxRow() const;
    int getMaxColumn() const;
    int getRowCount() const { return m_rowCount; }
    int getColumnCount() const { return m_columnCount; }
    void setRowCount(int count) { m_rowCount = count; }
    void setColumnCount(int count) { m_columnCount = count; }

    // Dirty tracking
    std::vector<CellAddress> getDirtyCells() const;
    void clearDirtyFlag();

    // Undo/Redo
    void startTransaction();
    void commitTransaction();
    void rollbackTransaction();

    // Formula engine access
    FormulaEngine& getFormulaEngine();

    // Cell iteration (for serialization)
    void forEachCell(std::function<void(int row, int col, const Cell&)> callback) const;

    // Undo/Redo
    UndoManager& getUndoManager() { return m_undoManager; }
    CellSnapshot takeCellSnapshot(const CellAddress& addr);

    // Sorting
    void sortRange(const CellRange& range, int sortColumn, bool ascending);

    // Cell shift insert/delete
    void insertCellsShiftRight(const CellRange& range);
    void insertCellsShiftDown(const CellRange& range);
    void deleteCellsShiftLeft(const CellRange& range);
    void deleteCellsShiftUp(const CellRange& range);

    // Table support
    void addTable(const SpreadsheetTable& table);
    void removeTable(const QString& name);
    const SpreadsheetTable* getTableAt(int row, int col) const;
    const std::vector<SpreadsheetTable>& getTables() const { return m_tables; }

    // Conditional formatting
    ConditionalFormatting& getConditionalFormatting() { return m_conditionalFormatting; }
    const ConditionalFormatting& getConditionalFormatting() const { return m_conditionalFormatting; }

    // Merge cells
    struct MergedRegion {
        CellRange range;
    };
    void mergeCells(const CellRange& range);
    void unmergeCells(const CellRange& range);
    const MergedRegion* getMergedRegionAt(int row, int col) const;
    std::vector<MergedRegion>& getMergedRegions() { return m_mergedRegions; }
    const std::vector<MergedRegion>& getMergedRegions() const { return m_mergedRegions; }

    // Data validation
    struct DataValidationRule {
        CellRange range;
        enum Type { WholeNumber, Decimal, List, TextLength, Date, Custom } type = WholeNumber;
        enum Operator { Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan,
                        GreaterThanOrEqual, LessThanOrEqual } op = Between;
        QString value1;
        QString value2;
        QStringList listItems;
        QString customFormula;
        QString inputTitle;
        QString inputMessage;
        QString errorTitle;
        QString errorMessage;
        enum ErrorStyle { Stop, Warning, Information } errorStyle = Stop;
        bool showInputMessage = true;
        bool showErrorAlert = true;
    };

    void addValidationRule(const DataValidationRule& rule);
    void removeValidationRule(int index);
    const DataValidationRule* getValidationAt(int row, int col) const;
    std::vector<DataValidationRule>& getValidationRules() { return m_validationRules; }
    const std::vector<DataValidationRule>& getValidationRules() const { return m_validationRules; }
    bool validateCell(int row, int col, const QString& value) const;

    // Row/Column dimensions
    void setRowHeight(int row, int height);
    void setColumnWidth(int col, int width);
    int getRowHeight(int row) const;        // returns 0 if not set (use default)
    int getColumnWidth(int col) const;      // returns 0 if not set (use default)
    const std::map<int, int>& getRowHeights() const { return m_rowHeights; }
    const std::map<int, int>& getColumnWidths() const { return m_columnWidths; }

    // Pivot table support
    void setPivotConfig(std::unique_ptr<PivotConfig> config);
    const PivotConfig* getPivotConfig() const;
    bool isPivotSheet() const;

    // Performance settings
    void setAutoRecalculate(bool enabled);
    bool getAutoRecalculate() const;
    void reserveCells(size_t count) { m_cells.reserve(count); }

    // Sparklines
    void setSparkline(const CellAddress& addr, const SparklineConfig& config);
    void removeSparkline(const CellAddress& addr);
    const SparklineConfig* getSparkline(const CellAddress& addr) const;
    const auto& getSparklines() const { return m_sparklines; }

    // Display settings
    void setShowGridlines(bool show) { m_showGridlines = show; }
    bool showGridlines() const { return m_showGridlines; }

private:
    struct CellKey {
        int row, col;
        bool operator==(const CellKey& other) const {
            return row == other.row && col == other.col;
        }
        bool operator<(const CellKey& other) const {
            if (row != other.row) return row < other.row;
            return col < other.col;
        }
    };

    struct CellKeyHash {
        size_t operator()(const CellKey& k) const {
            // Fast hash combining row and col
            return std::hash<uint64_t>()(
                (static_cast<uint64_t>(k.row) << 32) | static_cast<uint32_t>(k.col));
        }
    };

    std::unordered_map<CellKey, std::shared_ptr<Cell>, CellKeyHash> m_cells;
    std::unique_ptr<FormulaEngine> m_formulaEngine;
    DependencyGraph m_depGraph;
    UndoManager m_undoManager;
    QString m_sheetName;
    int m_rowCount;
    int m_columnCount;
    bool m_autoRecalculate;
    bool m_inTransaction;
    // Cached max row/col (avoids O(n) scan every call)
    mutable int m_cachedMaxRow = -1;
    mutable int m_cachedMaxCol = -1;
    mutable bool m_maxRowColDirty = true;
    void updateMaxRowCol() const;
    std::vector<SpreadsheetTable> m_tables;
    ConditionalFormatting m_conditionalFormatting;
    std::vector<DataValidationRule> m_validationRules;
    std::vector<MergedRegion> m_mergedRegions;
    std::unique_ptr<PivotConfig> m_pivotConfig;
    std::map<int, int> m_rowHeights;     // row -> height in pixels
    std::map<int, int> m_columnWidths;   // col -> width in pixels
    bool m_showGridlines = true;
    std::unordered_map<CellKey, SparklineConfig, CellKeyHash> m_sparklines;

    void recalculate(const CellAddress& addr);
    void recalculateAll();
    void updateDependencies(const CellAddress& addr);
    void recalculateDependents(const CellAddress& addr);
};

#endif // SPREADSHEET_H
