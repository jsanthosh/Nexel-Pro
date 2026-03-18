#ifndef SPREADSHEETMODEL_H
#define SPREADSHEETMODEL_H

#include <QAbstractTableModel>
#include <memory>

class Spreadsheet;
class MacroEngine;

class SpreadsheetModel : public QAbstractTableModel {
    Q_OBJECT

public:
    static constexpr int SparklineRole = Qt::UserRole + 15;

    // Virtual windowing: when total rows exceed this threshold, virtual mode engages.
    static constexpr int VIRTUAL_THRESHOLD = 500'000;

    // Buffer window: Qt sees this many rows for native smooth scrolling.
    // 10K rows = 120KB QHeaderView metadata — negligible.
    static constexpr int WINDOW_SIZE = 10'000;

    // Recenter when within this many rows of the window edge.
    static constexpr int RECENTER_MARGIN = 2'000;

    SpreadsheetModel(std::shared_ptr<Spreadsheet> spreadsheet, QObject* parent = nullptr);
    ~SpreadsheetModel() = default;

    void setMacroEngine(MacroEngine* engine) { m_macroEngine = engine; }

    int rowCount(const QModelIndex& parent = QModelIndex()) const override;
    int columnCount(const QModelIndex& parent = QModelIndex()) const override;
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex& index) const override;
    bool setData(const QModelIndex& index, const QVariant& value, int role = Qt::EditRole) override;

    void resetModel() {
        beginResetModel();
        endResetModel();
    }

    // Structural column/row changes — Qt auto-shifts header section sizes
    void beginColumnRemoval(int first, int count) {
        beginRemoveColumns(QModelIndex(), first, first + count - 1);
    }
    void endColumnRemoval() { endRemoveColumns(); }

    void beginColumnInsertion(int first, int count) {
        beginInsertColumns(QModelIndex(), first, first + count - 1);
    }
    void endColumnInsertion() { endInsertColumns(); }

    void beginRowRemoval(int first, int count) {
        beginRemoveRows(QModelIndex(), first, first + count - 1);
    }
    void endRowRemoval() { endRemoveRows(); }

    void beginRowInsertion(int first, int count) {
        beginInsertRows(QModelIndex(), first, first + count - 1);
    }
    void endRowInsertion() { endInsertRows(); }

    // Suppress undo tracking (used during bulk operations like paste/delete)
    void setSuppressUndo(bool suppress) { m_suppressUndo = suppress; }

    // Highlight invalid cells mode
    void setHighlightInvalidCells(bool enabled) { m_highlightInvalid = enabled; }
    bool highlightInvalidCells() const { return m_highlightInvalid; }

    // Formula view mode: show raw formulas instead of computed values
    void setShowFormulas(bool show) { m_showFormulas = show; }
    bool showFormulas() const { return m_showFormulas; }

    // ---- Virtual windowing API ----

    // Check if virtual mode is active (total rows > threshold)
    bool isVirtualMode() const;

    // Total logical rows in the spreadsheet
    int totalLogicalRows() const;

    // Convert model row index to logical row (adds window base in virtual mode)
    int toLogicalRow(int modelRow) const;

    // Convert logical row to model row (subtracts window base, -1 if not in window)
    int toModelRow(int logicalRow) const;

    // Window base: logical row at model row 0
    int windowBase() const { return m_windowBase; }

    // Recenter window around a logical row (emits dataChanged, no reset)
    // Returns the shift applied (newBase - oldBase)
    int recenterWindow(int logicalRow);

    // Jump window to a specific base (for scrollbar drag). Uses beginResetModel.
    void jumpToBase(int newBase);

    // Visible rows in viewport (for scrollbar page step calculation)
    void setVisibleRows(int rows) { m_visibleRows = std::max(10, rows); }
    int visibleRows() const { return m_visibleRows; }

    // Legacy API (kept for compatibility, maps to windowBase)
    void setViewportStart(int logicalRow);
    void setViewportStartFast(int logicalRow);
    int viewportStart() const { return m_windowBase; }

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    MacroEngine* m_macroEngine = nullptr;
    bool m_suppressUndo = false;
    bool m_highlightInvalid = false;
    bool m_showFormulas = false;

    // Virtual windowing state
    int m_windowBase = 0;    // logical row at model row 0
    int m_visibleRows = 50;  // approximate viewport rows (for scrollbar page step)

    QString columnIndexToLetter(int column) const;
};

#endif // SPREADSHEETMODEL_H
