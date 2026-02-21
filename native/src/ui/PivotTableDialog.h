#ifndef PIVOTTABLEDIALOG_H
#define PIVOTTABLEDIALOG_H

#include <QDialog>
#include <memory>
#include "../core/PivotEngine.h"

class QListWidget;
class QComboBox;
class QCheckBox;
class QLineEdit;
class QLabel;
class QTableWidget;
class QPushButton;
class Spreadsheet;
class QTimer;

class PivotTableDialog : public QDialog {
    Q_OBJECT

public:
    explicit PivotTableDialog(std::shared_ptr<Spreadsheet> sheet,
                              const CellRange& sourceRange,
                              QWidget* parent = nullptr);

    PivotConfig getConfig() const;

private slots:
    void onAddRowField();
    void onAddColumnField();
    void onAddValueField();
    void onAddFilterField();
    void onRemoveRowField();
    void onRemoveColumnField();
    void onRemoveValueField();
    void onRemoveFilterField();
    void onAggregationChanged(int index);
    void schedulePreviewUpdate();
    void updatePreview();

private:
    void createLayout();
    void populateSourceFields();
    void moveFieldToZone(QListWidget* targetZone, const QString& prefix = "");
    void moveFieldBack(QListWidget* sourceZone);
    int getSelectedFieldColumn(QListWidget* list) const;

    std::shared_ptr<Spreadsheet> m_sourceSheet;
    CellRange m_sourceRange;
    QStringList m_sourceColumns;

    QLineEdit* m_rangeEdit = nullptr;
    QListWidget* m_sourceFieldList = nullptr;
    QListWidget* m_filterZone = nullptr;
    QListWidget* m_rowZone = nullptr;
    QListWidget* m_columnZone = nullptr;
    QListWidget* m_valueZone = nullptr;

    QComboBox* m_aggregationCombo = nullptr;
    QCheckBox* m_showGrandTotalRow = nullptr;
    QCheckBox* m_showGrandTotalCol = nullptr;
    QCheckBox* m_autoChart = nullptr;
    QComboBox* m_chartTypeCombo = nullptr;

    QTableWidget* m_previewTable = nullptr;
    QTimer* m_previewTimer = nullptr;

    PivotEngine m_engine;
};

#endif // PIVOTTABLEDIALOG_H
