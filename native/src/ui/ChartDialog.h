#ifndef CHARTDIALOG_H
#define CHARTDIALOG_H

#include <QDialog>
#include <memory>
#include "ChartWidget.h"
#include "ShapeWidget.h"

class QLineEdit;
class QComboBox;
class QCheckBox;
class QListWidget;
class QGroupBox;
class Spreadsheet;

class ChartDialog : public QDialog {
    Q_OBJECT

public:
    explicit ChartDialog(QWidget* parent = nullptr);
    ~ChartDialog() = default;

    ChartConfig getConfig() const;
    void setDataRange(const QString& range);
    void setConfig(const ChartConfig& config);
    void setSpreadsheet(std::shared_ptr<Spreadsheet> sheet);

private slots:
    void onChartTypeChanged(int index);
    void updatePreview();

private:
    void createLayout();
    void createChartTypeSelector();
    void createDataPanel();
    void createOptionsPanel();
    void createPreviewPanel();

    // Chart type selector
    QListWidget* m_chartTypeList = nullptr;

    // Data panel
    QLineEdit* m_dataRangeEdit = nullptr;
    QCheckBox* m_firstRowHeaders = nullptr;
    QCheckBox* m_firstColLabels = nullptr;

    // Options
    QLineEdit* m_titleEdit = nullptr;
    QLineEdit* m_xAxisEdit = nullptr;
    QLineEdit* m_yAxisEdit = nullptr;
    QCheckBox* m_showLegend = nullptr;
    QCheckBox* m_showGridLines = nullptr;
    QComboBox* m_themeCombo = nullptr;

    // Preview
    ChartWidget* m_preview = nullptr;
    std::shared_ptr<Spreadsheet> m_spreadsheet;
};

// Simple dialog for inserting shapes
class InsertShapeDialog : public QDialog {
    Q_OBJECT

public:
    explicit InsertShapeDialog(QWidget* parent = nullptr);
    ShapeConfig getConfig() const;

private:
    void createLayout();
    QListWidget* m_shapeList = nullptr;
    QComboBox* m_fillColorCombo = nullptr;
    QComboBox* m_strokeColorCombo = nullptr;
};

#endif // CHARTDIALOG_H
