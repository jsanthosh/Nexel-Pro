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
    // Must call begin BEFORE data changes, end AFTER.
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

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    MacroEngine* m_macroEngine = nullptr;
    bool m_suppressUndo = false;
    bool m_highlightInvalid = false;
    bool m_showFormulas = false;

    QString columnIndexToLetter(int column) const;
};

#endif // SPREADSHEETMODEL_H
