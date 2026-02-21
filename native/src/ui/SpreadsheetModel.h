#ifndef SPREADSHEETMODEL_H
#define SPREADSHEETMODEL_H

#include <QAbstractTableModel>
#include <memory>

class Spreadsheet;

class SpreadsheetModel : public QAbstractTableModel {
    Q_OBJECT

public:
    static constexpr int SparklineRole = Qt::UserRole + 15;

    SpreadsheetModel(std::shared_ptr<Spreadsheet> spreadsheet, QObject* parent = nullptr);
    ~SpreadsheetModel() = default;

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

    // Suppress undo tracking (used during bulk operations like paste/delete)
    void setSuppressUndo(bool suppress) { m_suppressUndo = suppress; }

    // Highlight invalid cells mode
    void setHighlightInvalidCells(bool enabled) { m_highlightInvalid = enabled; }
    bool highlightInvalidCells() const { return m_highlightInvalid; }

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    bool m_suppressUndo = false;
    bool m_highlightInvalid = false;

    QString columnIndexToLetter(int column) const;
};

#endif // SPREADSHEETMODEL_H
