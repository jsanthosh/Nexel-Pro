#ifndef CHARTPROPERTIESPANEL_H
#define CHARTPROPERTIESPANEL_H

#include <QWidget>
#include <QVBoxLayout>
#include <QLineEdit>
#include <QComboBox>
#include <QCheckBox>
#include <QPushButton>
#include <QScrollArea>
#include <QLabel>
#include "ChartWidget.h"

class ChartPropertiesPanel : public QWidget {
    Q_OBJECT

public:
    explicit ChartPropertiesPanel(QWidget* parent = nullptr);
    void setChart(ChartWidget* chart);
    ChartWidget* currentChart() const { return m_chart; }

signals:
    void closeRequested();

private slots:
    void onPropertyChanged();
    void onChartTypeClicked();
    void onRefreshData();
    void onSeriesColorClicked();

private:
    void createLayout();
    QWidget* createSectionHeader(const QString& title);
    void updateFromChart();
    void applyToChart();
    void rebuildSeriesSection();

    ChartWidget* m_chart = nullptr;

    // Chart type buttons
    QVector<QPushButton*> m_typeButtons;

    // Title & Labels
    QLineEdit* m_titleEdit = nullptr;
    QLineEdit* m_xAxisEdit = nullptr;
    QLineEdit* m_yAxisEdit = nullptr;

    // Style
    QComboBox* m_themeCombo = nullptr;
    QCheckBox* m_legendCheck = nullptr;
    QCheckBox* m_gridCheck = nullptr;

    // Data
    QLineEdit* m_dataRangeEdit = nullptr;

    // Series colors
    QVBoxLayout* m_seriesLayout = nullptr;
    QWidget* m_seriesContainer = nullptr;
    QVector<QPushButton*> m_seriesColorButtons;

    // === Deep customization (Sprint 4) ===
    // Legend position
    QComboBox* m_legendPosCombo = nullptr;

    // Data labels
    QComboBox* m_dataLabelCombo = nullptr;
    QCheckBox* m_labelShowValue = nullptr;
    QCheckBox* m_labelShowCategory = nullptr;
    QCheckBox* m_labelShowPercent = nullptr;

    // Axis formatting
    QLineEdit* m_yAxisMin = nullptr;
    QLineEdit* m_yAxisMax = nullptr;
    QLineEdit* m_yAxisUnits = nullptr;
    QCheckBox* m_yAxisLog = nullptr;

    // Chart variants
    QCheckBox* m_stackedCheck = nullptr;
    QCheckBox* m_percentStackedCheck = nullptr;
    QCheckBox* m_smoothLinesCheck = nullptr;
    QCheckBox* m_showMarkersCheck = nullptr;

    // Trendline
    QComboBox* m_trendlineCombo = nullptr;
    QCheckBox* m_trendShowEq = nullptr;
    QCheckBox* m_trendShowR2 = nullptr;

    QScrollArea* m_scrollArea = nullptr;
    bool m_updating = false;
};

#endif // CHARTPROPERTIESPANEL_H
