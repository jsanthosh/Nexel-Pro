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
#include <QStackedWidget>
#include "ChartWidget.h"

// Which chart element is currently selected for formatting
enum class ChartElement {
    ChartArea,      // Overall chart background
    PlotArea,       // Plot region inside axes
    Title,          // Chart title
    XAxisTitle,     // X axis title
    YAxisTitle,     // Y axis title
    Legend,         // Legend
    DataSeries,     // A data series (which one tracked by index)
    DataLabels,     // Data labels
    ValueAxis,      // Y axis (value axis)
    CategoryAxis,   // X axis (category axis)
    MajorGridlines, // Major gridlines
    Trendline       // Trendline
};

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
    void onElementChanged(int index);

private:
    void createLayout();
    QWidget* createSectionHeader(const QString& title);
    void updateFromChart();
    void applyToChart();
    void rebuildSeriesSection();
    void rebuildElementDropdown();
    void showElementOptions(ChartElement element);

    // Create context-specific option panels
    QWidget* createChartAreaOptions();
    QWidget* createTitleOptions();
    QWidget* createSeriesOptions();
    QWidget* createAxisOptions();
    QWidget* createLegendOptions();
    QWidget* createDataLabelOptions();
    QWidget* createGridlineOptions();
    QWidget* createTrendlineOptions();
    QWidget* createChartTypeSection();
    QWidget* createDataSection();

    ChartWidget* m_chart = nullptr;
    ChartElement m_currentElement = ChartElement::ChartArea;

    // Header
    QLabel* m_headerTitle = nullptr;

    // Element selector (Excel-style dropdown at top)
    QComboBox* m_elementCombo = nullptr;

    // Chart type buttons
    QVector<QPushButton*> m_typeButtons;

    // Stacked widget for context-sensitive options
    QStackedWidget* m_optionsStack = nullptr;

    // Common controls (reused across panels)
    QLineEdit* m_titleEdit = nullptr;
    QLineEdit* m_xAxisEdit = nullptr;
    QLineEdit* m_yAxisEdit = nullptr;
    QComboBox* m_themeCombo = nullptr;
    QCheckBox* m_legendCheck = nullptr;
    QCheckBox* m_gridCheck = nullptr;
    QLineEdit* m_dataRangeEdit = nullptr;

    // Series section
    QVBoxLayout* m_seriesLayout = nullptr;
    QWidget* m_seriesContainer = nullptr;
    QVector<QPushButton*> m_seriesColorButtons;

    // Context-specific controls
    // Series options
    QComboBox* m_seriesOverlapSlider = nullptr;
    QComboBox* m_gapWidthSlider = nullptr;
    QCheckBox* m_stackedCheck = nullptr;
    QCheckBox* m_percentStackedCheck = nullptr;

    // Axis options
    QLineEdit* m_axisMin = nullptr;
    QLineEdit* m_axisMax = nullptr;
    QLineEdit* m_axisMajorUnit = nullptr;
    QCheckBox* m_axisLogScale = nullptr;
    QComboBox* m_axisDisplayUnits = nullptr;
    QCheckBox* m_axisReverse = nullptr;

    // Legend options
    QComboBox* m_legendPosCombo = nullptr;

    // Data label options
    QComboBox* m_dataLabelPosCombo = nullptr;
    QCheckBox* m_labelShowValue = nullptr;
    QCheckBox* m_labelShowCategory = nullptr;
    QCheckBox* m_labelShowPercent = nullptr;
    QCheckBox* m_labelShowSeriesName = nullptr;

    // Line/marker options (for line/scatter charts)
    QCheckBox* m_smoothLinesCheck = nullptr;
    QCheckBox* m_showMarkersCheck = nullptr;

    // Trendline options
    QComboBox* m_trendlineTypeCombo = nullptr;
    QCheckBox* m_trendShowEq = nullptr;
    QCheckBox* m_trendShowR2 = nullptr;

    QScrollArea* m_scrollArea = nullptr;
    bool m_updating = false;
};

#endif // CHARTPROPERTIESPANEL_H
