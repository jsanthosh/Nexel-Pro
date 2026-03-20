#include "ChartPropertiesPanel.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QScrollArea>
#include <QPushButton>
#include <QColorDialog>
#include <QPainter>
#include <QPainterPath>
#include <QToolButton>

// --- Helpers ---

static QIcon makeChartTypeIcon(ChartType type, bool selected) {
    QPixmap pix(32, 32);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);

    QColor primary = selected ? ThemeManager::instance().currentTheme().accentDark : QColor("#667085");

    switch (type) {
        case ChartType::Column:
            p.setPen(Qt::NoPen); p.setBrush(primary);
            p.drawRect(4, 18, 6, 10); p.drawRect(13, 10, 6, 18); p.drawRect(22, 6, 6, 22);
            break;
        case ChartType::Bar:
            p.setPen(Qt::NoPen); p.setBrush(primary);
            p.drawRect(4, 4, 22, 6); p.drawRect(4, 13, 16, 6); p.drawRect(4, 22, 24, 6);
            break;
        case ChartType::Line:
            p.setPen(QPen(primary, 2)); p.setBrush(Qt::NoBrush);
            p.drawLine(4, 24, 12, 14); p.drawLine(12, 14, 20, 20); p.drawLine(20, 20, 28, 6);
            break;
        case ChartType::Area: {
            QPainterPath path;
            path.moveTo(4, 28); path.lineTo(4, 20); path.lineTo(14, 10);
            path.lineTo(22, 16); path.lineTo(28, 6); path.lineTo(28, 28); path.closeSubpath();
            QColor fill = primary; fill.setAlpha(80);
            p.setPen(QPen(primary, 1.5)); p.setBrush(fill); p.drawPath(path);
            break;
        }
        case ChartType::Scatter:
            p.setPen(QPen(primary, 1)); p.setBrush(primary);
            p.drawEllipse(QPoint(8, 22), 3, 3); p.drawEllipse(QPoint(14, 16), 3, 3);
            p.drawEllipse(QPoint(20, 12), 3, 3); p.drawEllipse(QPoint(26, 8), 3, 3);
            break;
        case ChartType::Pie:
            p.setPen(QPen(Qt::white, 1.5)); p.setBrush(primary);
            p.drawPie(4, 4, 24, 24, 0, 200 * 16);
            p.setBrush(primary.lighter(140));
            p.drawPie(4, 4, 24, 24, 200 * 16, 160 * 16);
            break;
        case ChartType::Donut:
            p.setPen(QPen(Qt::white, 1.5)); p.setBrush(primary);
            p.drawPie(4, 4, 24, 24, 0, 200 * 16);
            p.setBrush(primary.lighter(140));
            p.drawPie(4, 4, 24, 24, 200 * 16, 160 * 16);
            p.setPen(Qt::NoPen); p.setBrush(Qt::white);
            p.drawEllipse(10, 10, 12, 12);
            break;
        case ChartType::Histogram:
            p.setPen(Qt::NoPen); p.setBrush(primary);
            p.drawRect(3, 22, 5, 6); p.drawRect(9, 14, 5, 14); p.drawRect(15, 6, 5, 22);
            p.drawRect(21, 12, 5, 16); p.drawRect(27, 20, 5, 8);
            break;
    }

    p.end();
    return QIcon(pix);
}

static QIcon makeColorSwatch(const QColor& color, int size = 16) {
    QPixmap pix(size, size);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);
    p.setPen(QPen(color.darker(130), 1));
    p.setBrush(color);
    p.drawRoundedRect(1, 1, size - 2, size - 2, 3, 3);
    p.end();
    return QIcon(pix);
}

// --- Panel Implementation ---

ChartPropertiesPanel::ChartPropertiesPanel(QWidget* parent)
    : QWidget(parent) {
    createLayout();
}

QWidget* ChartPropertiesPanel::createSectionHeader(const QString& title) {
    QLabel* label = new QLabel(title);
    label->setStyleSheet(
        "QLabel { color: #667085; font-size: 10px; font-weight: 700; "
        "letter-spacing: 1.2px; padding: 0; margin: 0; }");
    return label;
}

void ChartPropertiesPanel::createLayout() {
    QVBoxLayout* outerLayout = new QVBoxLayout(this);
    outerLayout->setContentsMargins(0, 0, 0, 0);
    outerLayout->setSpacing(0);

    // Header bar with gradient
    QWidget* header = new QWidget();
    header->setFixedHeight(44);
    {
        const auto& t = ThemeManager::instance().currentTheme();
        header->setStyleSheet(QString(
            "QWidget { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
            "stop:0 %1, stop:1 %2); }").arg(t.accentDarker.name(), t.accentDark.name()));
    }
    QHBoxLayout* headerLayout = new QHBoxLayout(header);
    headerLayout->setContentsMargins(14, 0, 8, 0);

    QLabel* headerIcon = new QLabel("\xF0\x9F\x93\x8A"); // chart emoji
    headerIcon->setStyleSheet("QLabel { font-size: 16px; }");
    headerLayout->addWidget(headerIcon);

    m_headerTitle = new QLabel("Format Chart Area");
    m_headerTitle->setStyleSheet(
        "QLabel { color: white; font-size: 13px; font-weight: 600; "
        "letter-spacing: 0.3px; margin-left: 4px; }");
    headerLayout->addWidget(m_headerTitle);
    headerLayout->addStretch();

    QPushButton* closeBtn = new QPushButton("\xC3\x97"); // multiplication sign
    closeBtn->setFixedSize(26, 26);
    closeBtn->setStyleSheet(
        "QPushButton { background: transparent; color: rgba(255,255,255,0.8); font-size: 16px; "
        "font-weight: bold; border: none; border-radius: 13px; }"
        "QPushButton:hover { background: rgba(255,255,255,0.15); color: white; }");
    connect(closeBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::closeRequested);
    headerLayout->addWidget(closeBtn);

    outerLayout->addWidget(header);

    // Element selector dropdown (Excel-style: "Chart Elements" at top)
    m_elementCombo = new QComboBox();
    m_elementCombo->setFixedHeight(32);
    m_elementCombo->setStyleSheet(
        "QComboBox { background: white; border: 1px solid #E4E7EC; border-radius: 6px; "
        "padding: 4px 10px; font-size: 12px; font-weight: 600; color: #344054; }"
        "QComboBox::drop-down { border: none; width: 20px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; }");
    m_elementCombo->addItems({
        "Chart Area", "Plot Area", "Chart Title",
        "Value Axis (Y)", "Category Axis (X)",
        "Legend", "Data Labels", "Major Gridlines", "Trendline"
    });
    connect(m_elementCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onElementChanged);

    QWidget* elemWidget = new QWidget();
    elemWidget->setStyleSheet("background: #F0F2F5;");
    QHBoxLayout* elemLayout = new QHBoxLayout(elemWidget);
    elemLayout->setContentsMargins(10, 6, 10, 6);
    elemLayout->addWidget(new QLabel("Element:"));
    elemLayout->addWidget(m_elementCombo, 1);
    outerLayout->addWidget(elemWidget);

    // Scroll area for content
    m_scrollArea = new QScrollArea();
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_scrollArea->setStyleSheet(
        "QScrollArea { border: none; background: #F8FAFB; }"
        "QScrollBar:vertical { width: 5px; background: transparent; margin: 2px 0; }"
        "QScrollBar::handle:vertical { background: #C8CDD3; border-radius: 2px; min-height: 30px; }"
        "QScrollBar::handle:vertical:hover { background: #A0A8B0; }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }"
        "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }");

    QWidget* content = new QWidget();
    content->setStyleSheet("QWidget { background: #F8FAFB; }");
    QVBoxLayout* contentLayout = new QVBoxLayout(content);
    contentLayout->setContentsMargins(14, 12, 14, 14);
    contentLayout->setSpacing(0);

    // ===== Chart Type Section =====
    contentLayout->addWidget(createSectionHeader("CHART TYPE"));
    contentLayout->addSpacing(8);

    QGridLayout* typeGrid = new QGridLayout();
    typeGrid->setSpacing(6);

    struct TypeInfo { ChartType type; QString tip; };
    QVector<TypeInfo> types = {
        { ChartType::Column,    "Column" },
        { ChartType::Bar,       "Bar" },
        { ChartType::Line,      "Line" },
        { ChartType::Area,      "Area" },
        { ChartType::Scatter,   "Scatter" },
        { ChartType::Pie,       "Pie" },
        { ChartType::Donut,     "Donut" },
        { ChartType::Histogram, "Histogram" },
    };

    for (int i = 0; i < types.size(); ++i) {
        QPushButton* btn = new QPushButton();
        btn->setFixedSize(40, 40);
        btn->setIconSize(QSize(28, 28));
        btn->setIcon(makeChartTypeIcon(types[i].type, false));
        btn->setToolTip(types[i].tip);
        btn->setProperty("chartType", static_cast<int>(types[i].type));
        btn->setStyleSheet(
            "QPushButton { background: white; border: 1.5px solid #E4E7EC; border-radius: 8px; }"
            "QPushButton:hover { background: #F0FAF3; border-color: #86C49A; }"
            "QPushButton[selected=\"true\"] { background: #E8F5E9; border: 2px solid " + ThemeManager::instance().currentTheme().accentDark.name() + "; }");
        connect(btn, &QPushButton::clicked, this, &ChartPropertiesPanel::onChartTypeClicked);
        typeGrid->addWidget(btn, i / 4, i % 4);
        m_typeButtons.append(btn);
    }

    contentLayout->addLayout(typeGrid);
    contentLayout->addSpacing(14);

    // ===== Title & Labels =====
    contentLayout->addWidget(createSectionHeader("TITLE & LABELS"));
    contentLayout->addSpacing(8);

    auto makeLabel = [](const QString& text) {
        QLabel* l = new QLabel(text);
        l->setStyleSheet("QLabel { color: #475467; font-size: 11px; font-weight: 500; }");
        l->setFixedWidth(46);
        return l;
    };

    auto makeEdit = [](const QString& placeholder) {
        const auto& t = ThemeManager::instance().currentTheme();
        QLineEdit* e = new QLineEdit();
        e->setPlaceholderText(placeholder);
        e->setFixedHeight(30);
        e->setStyleSheet(QString(
            "QLineEdit { border: 1px solid #D0D5DD; border-radius: 6px; padding: 2px 10px; "
            "background: white; font-size: 11px; color: #1D2939; }"
            "QLineEdit:focus { border-color: %1; box-shadow: none; }"
            "QLineEdit::placeholder { color: #98A2B3; }").arg(t.accentPrimary.name()));
        return e;
    };

    m_titleEdit = makeEdit("Chart title");
    m_xAxisEdit = makeEdit("X axis label");
    m_yAxisEdit = makeEdit("Y axis label");

    QGridLayout* labelGrid = new QGridLayout();
    labelGrid->setSpacing(6);
    labelGrid->setColumnStretch(1, 1);
    labelGrid->addWidget(makeLabel("Title"), 0, 0, Qt::AlignTop | Qt::AlignLeft);
    labelGrid->addWidget(m_titleEdit, 0, 1);
    labelGrid->addWidget(makeLabel("X Axis"), 1, 0, Qt::AlignTop | Qt::AlignLeft);
    labelGrid->addWidget(m_xAxisEdit, 1, 1);
    labelGrid->addWidget(makeLabel("Y Axis"), 2, 0, Qt::AlignTop | Qt::AlignLeft);
    labelGrid->addWidget(m_yAxisEdit, 2, 1);
    contentLayout->addLayout(labelGrid);

    connect(m_titleEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_xAxisEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_yAxisEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Style =====
    contentLayout->addWidget(createSectionHeader("STYLE"));
    contentLayout->addSpacing(8);

    QGridLayout* styleGrid = new QGridLayout();
    styleGrid->setSpacing(6);
    styleGrid->setColumnStretch(1, 1);

    m_themeCombo = new QComboBox();
    m_themeCombo->setFixedHeight(30);
    m_themeCombo->addItems({"Document Theme", "Excel", "Material", "Solarized", "Dark", "Monochrome", "Pastel"});
    {
        const auto& t = ThemeManager::instance().currentTheme();
        QColor selBg = t.accentPrimary;
        selBg.setAlpha(30);
        m_themeCombo->setStyleSheet(QString(
            "QComboBox { border: 1px solid #D0D5DD; border-radius: 6px; padding: 2px 10px; "
            "background: white; font-size: 11px; color: #1D2939; min-height: 22px; }"
            "QComboBox:focus { border: 1px solid %1; }"
            "QComboBox::drop-down { border: none; width: 20px; }"
            "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
            "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 6px; }"
            "QComboBox QAbstractItemView { border: 1px solid #D0D5DD; border-radius: 6px; "
            "background: white; selection-background-color: rgba(%2,%3,%4,30); padding: 4px; outline: none; }")
            .arg(t.accentPrimary.name())
            .arg(t.accentPrimary.red()).arg(t.accentPrimary.green()).arg(t.accentPrimary.blue()));
    }

    styleGrid->addWidget(makeLabel("Theme"), 0, 0, Qt::AlignTop | Qt::AlignLeft);
    styleGrid->addWidget(m_themeCombo, 0, 1);
    contentLayout->addLayout(styleGrid);

    connect(m_themeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(6);

    m_legendCheck = new QCheckBox("Show Legend");
    m_legendCheck->setStyleSheet(
        "QCheckBox { color: #344054; font-size: 11px; spacing: 8px; }");
    m_gridCheck = new QCheckBox("Show Grid Lines");
    m_gridCheck->setStyleSheet(m_legendCheck->styleSheet());

    contentLayout->addWidget(m_legendCheck);
    contentLayout->addSpacing(2);
    contentLayout->addWidget(m_gridCheck);

    connect(m_legendCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_gridCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Data =====
    contentLayout->addWidget(createSectionHeader("DATA"));
    contentLayout->addSpacing(8);

    QHBoxLayout* dataLayout = new QHBoxLayout();
    dataLayout->setSpacing(6);

    m_dataRangeEdit = makeEdit("A1:D10");
    dataLayout->addWidget(m_dataRangeEdit, 1);

    QPushButton* refreshBtn = new QPushButton("Refresh");
    refreshBtn->setFixedHeight(30);
    refreshBtn->setCursor(Qt::PointingHandCursor);
    {
        const auto& t = ThemeManager::instance().currentTheme();
        refreshBtn->setStyleSheet(QString(
            "QPushButton { background: %1; color: white; border: none; border-radius: 6px; "
            "padding: 0 14px; font-size: 11px; font-weight: 600; }"
            "QPushButton:hover { background: %2; }"
            "QPushButton:pressed { background: %3; }")
            .arg(t.accentDark.name(), t.accentDarker.name(), t.menuBarHover.name()));
    }
    connect(refreshBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::onRefreshData);
    dataLayout->addWidget(refreshBtn);

    contentLayout->addLayout(dataLayout);

    connect(m_dataRangeEdit, &QLineEdit::editingFinished, this, &ChartPropertiesPanel::onRefreshData);

    contentLayout->addSpacing(14);

    // ===== Series Colors =====
    contentLayout->addWidget(createSectionHeader("SERIES COLORS"));
    contentLayout->addSpacing(8);

    m_seriesContainer = new QWidget();
    m_seriesLayout = new QVBoxLayout(m_seriesContainer);
    m_seriesLayout->setContentsMargins(0, 0, 0, 0);
    m_seriesLayout->setSpacing(4);
    contentLayout->addWidget(m_seriesContainer);

    contentLayout->addSpacing(14);

    // ===== Legend Position =====
    contentLayout->addWidget(createSectionHeader("LEGEND"));
    contentLayout->addSpacing(8);
    m_legendPosCombo = new QComboBox();
    m_legendPosCombo->addItems({"Right", "Top", "Bottom", "Left", "None"});
    m_legendPosCombo->setFixedHeight(30);
    m_legendPosCombo->setStyleSheet(m_themeCombo->styleSheet());
    contentLayout->addWidget(m_legendPosCombo);
    connect(m_legendPosCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Data Labels =====
    contentLayout->addWidget(createSectionHeader("DATA LABELS"));
    contentLayout->addSpacing(8);
    m_dataLabelPosCombo = new QComboBox();
    m_dataLabelPosCombo->addItems({"None", "Center", "Above", "Below", "Inside End", "Outside End"});
    m_dataLabelPosCombo->setFixedHeight(30);
    m_dataLabelPosCombo->setStyleSheet(m_themeCombo->styleSheet());
    contentLayout->addWidget(m_dataLabelPosCombo);
    connect(m_dataLabelPosCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onPropertyChanged);

    m_labelShowValue = new QCheckBox("Show Value");
    m_labelShowValue->setChecked(true);
    m_labelShowValue->setStyleSheet(m_legendCheck->styleSheet());
    m_labelShowCategory = new QCheckBox("Show Category");
    m_labelShowCategory->setStyleSheet(m_legendCheck->styleSheet());
    m_labelShowPercent = new QCheckBox("Show Percentage");
    m_labelShowPercent->setStyleSheet(m_legendCheck->styleSheet());
    contentLayout->addWidget(m_labelShowValue);
    contentLayout->addWidget(m_labelShowCategory);
    contentLayout->addWidget(m_labelShowPercent);
    connect(m_labelShowValue, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_labelShowCategory, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_labelShowPercent, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Axis Formatting =====
    contentLayout->addWidget(createSectionHeader("Y AXIS"));
    contentLayout->addSpacing(8);
    auto* axisGrid = new QGridLayout();
    axisGrid->setSpacing(6);
    m_axisMin = makeEdit("Auto");
    m_axisMax = makeEdit("Auto");
    axisGrid->addWidget(new QLabel("Min:"), 0, 0);
    axisGrid->addWidget(m_axisMin, 0, 1);
    axisGrid->addWidget(new QLabel("Max:"), 1, 0);
    axisGrid->addWidget(m_axisMax, 1, 1);
    contentLayout->addLayout(axisGrid);
    m_axisLogScale = new QCheckBox("Logarithmic Scale");
    m_axisLogScale->setStyleSheet(m_legendCheck->styleSheet());
    contentLayout->addWidget(m_axisLogScale);
    connect(m_axisMin, &QLineEdit::editingFinished, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_axisMax, &QLineEdit::editingFinished, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_axisLogScale, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Chart Variants =====
    contentLayout->addWidget(createSectionHeader("CHART OPTIONS"));
    contentLayout->addSpacing(8);
    m_stackedCheck = new QCheckBox("Stacked");
    m_stackedCheck->setStyleSheet(m_legendCheck->styleSheet());
    m_percentStackedCheck = new QCheckBox("100% Stacked");
    m_percentStackedCheck->setStyleSheet(m_legendCheck->styleSheet());
    m_smoothLinesCheck = new QCheckBox("Smooth Lines");
    m_smoothLinesCheck->setStyleSheet(m_legendCheck->styleSheet());
    m_showMarkersCheck = new QCheckBox("Show Markers");
    m_showMarkersCheck->setChecked(true);
    m_showMarkersCheck->setStyleSheet(m_legendCheck->styleSheet());
    contentLayout->addWidget(m_stackedCheck);
    contentLayout->addWidget(m_percentStackedCheck);
    contentLayout->addWidget(m_smoothLinesCheck);
    contentLayout->addWidget(m_showMarkersCheck);
    connect(m_stackedCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_percentStackedCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_smoothLinesCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_showMarkersCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addSpacing(14);

    // ===== Trendline =====
    contentLayout->addWidget(createSectionHeader("TRENDLINE"));
    contentLayout->addSpacing(8);
    m_trendlineTypeCombo = new QComboBox();
    m_trendlineTypeCombo->addItems({"None", "Linear", "Exponential", "Logarithmic", "Polynomial", "Power", "Moving Average"});
    m_trendlineTypeCombo->setFixedHeight(30);
    m_trendlineTypeCombo->setStyleSheet(m_themeCombo->styleSheet());
    contentLayout->addWidget(m_trendlineTypeCombo);
    m_trendShowEq = new QCheckBox("Display Equation");
    m_trendShowEq->setStyleSheet(m_legendCheck->styleSheet());
    m_trendShowR2 = new QCheckBox("Display R² Value");
    m_trendShowR2->setStyleSheet(m_legendCheck->styleSheet());
    contentLayout->addWidget(m_trendShowEq);
    contentLayout->addWidget(m_trendShowR2);
    connect(m_trendlineTypeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_trendShowEq, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_trendShowR2, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    contentLayout->addStretch();

    m_scrollArea->setWidget(content);
    outerLayout->addWidget(m_scrollArea);

    setStyleSheet("ChartPropertiesPanel { background: #F8FAFB; }");
}

void ChartPropertiesPanel::setChart(ChartWidget* chart) {
    m_chart = chart;
    m_updating = true;
    rebuildElementDropdown();
    updateFromChart();
    m_currentElement = ChartElement::ChartArea;
    showElementOptions(m_currentElement);
    m_updating = false;
}

void ChartPropertiesPanel::updateFromChart() {
    if (!m_chart) return;

    const ChartConfig& cfg = m_chart->config();

    m_titleEdit->setText(cfg.title);
    m_xAxisEdit->setText(cfg.xAxisTitle);
    m_yAxisEdit->setText(cfg.yAxisTitle);
    m_themeCombo->setCurrentIndex(cfg.themeIndex);
    m_legendCheck->setChecked(cfg.showLegend);
    m_gridCheck->setChecked(cfg.showGridLines);
    m_dataRangeEdit->setText(cfg.dataRange);

    // Deep customization
    if (m_legendPosCombo) m_legendPosCombo->setCurrentIndex(static_cast<int>(cfg.legendPosition));
    if (m_dataLabelPosCombo) m_dataLabelPosCombo->setCurrentIndex(static_cast<int>(cfg.dataLabelPosition));
    if (m_labelShowValue) m_labelShowValue->setChecked(cfg.dataLabelShowValue);
    if (m_labelShowCategory) m_labelShowCategory->setChecked(cfg.dataLabelShowCategory);
    if (m_labelShowPercent) m_labelShowPercent->setChecked(cfg.dataLabelShowPercentage);
    if (m_axisMin) m_axisMin->setText(cfg.yAxisConfig.autoMin ? "Auto" : QString::number(cfg.yAxisConfig.minValue));
    if (m_axisMax) m_axisMax->setText(cfg.yAxisConfig.autoMax ? "Auto" : QString::number(cfg.yAxisConfig.maxValue));
    if (m_axisLogScale) m_axisLogScale->setChecked(cfg.yAxisConfig.logScale);
    if (m_stackedCheck) m_stackedCheck->setChecked(cfg.stacked);
    if (m_percentStackedCheck) m_percentStackedCheck->setChecked(cfg.percentStacked);
    if (m_smoothLinesCheck) m_smoothLinesCheck->setChecked(cfg.smoothLines);
    if (m_showMarkersCheck) m_showMarkersCheck->setChecked(cfg.showMarkers);
    if (m_trendlineTypeCombo) {
        int trendIdx = 0;
        if (!cfg.trendlines.isEmpty()) trendIdx = static_cast<int>(cfg.trendlines[0].type);
        m_trendlineTypeCombo->setCurrentIndex(trendIdx);
    }
    if (m_trendShowEq && !cfg.trendlines.isEmpty()) m_trendShowEq->setChecked(cfg.trendlines[0].displayEquation);
    if (m_trendShowR2 && !cfg.trendlines.isEmpty()) m_trendShowR2->setChecked(cfg.trendlines[0].displayRSquared);

    // Update type button selection
    int typeInt = static_cast<int>(cfg.type);
    for (auto* btn : m_typeButtons) {
        bool sel = btn->property("chartType").toInt() == typeInt;
        btn->setProperty("selected", sel);
        btn->setIcon(makeChartTypeIcon(static_cast<ChartType>(btn->property("chartType").toInt()), sel));
        btn->style()->unpolish(btn);
        btn->style()->polish(btn);
    }

    rebuildSeriesSection();
}

void ChartPropertiesPanel::rebuildSeriesSection() {
    // Clear existing series widgets
    m_seriesColorButtons.clear();
    QLayoutItem* item;
    while ((item = m_seriesLayout->takeAt(0)) != nullptr) {
        delete item->widget();
        delete item;
    }

    if (!m_chart) return;

    const auto& series = m_chart->config().series;
    for (int i = 0; i < series.size(); ++i) {
        QHBoxLayout* row = new QHBoxLayout();
        row->setSpacing(8);
        row->setContentsMargins(0, 0, 0, 0);

        QPushButton* colorBtn = new QPushButton();
        colorBtn->setFixedSize(26, 26);
        colorBtn->setIcon(makeColorSwatch(series[i].color, 20));
        colorBtn->setIconSize(QSize(20, 20));
        colorBtn->setCursor(Qt::PointingHandCursor);
        {
            const auto& t = ThemeManager::instance().currentTheme();
            QColor hoverBg = t.accentPrimary;
            hoverBg.setAlpha(20);
            colorBtn->setStyleSheet(QString(
                "QPushButton { background: white; border: 1.5px solid #E4E7EC; border-radius: 6px; padding: 2px; }"
                "QPushButton:hover { border-color: %1; background: rgba(%2,%3,%4,20); }")
                .arg(t.accentPrimary.name())
                .arg(t.accentPrimary.red()).arg(t.accentPrimary.green()).arg(t.accentPrimary.blue()));
        }
        colorBtn->setProperty("seriesIndex", i);
        connect(colorBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::onSeriesColorClicked);

        QLabel* nameLabel = new QLabel(series[i].name);
        nameLabel->setStyleSheet("QLabel { color: #344054; font-size: 11px; font-weight: 500; }");

        row->addWidget(colorBtn);
        row->addWidget(nameLabel, 1);

        m_seriesColorButtons.append(colorBtn);

        QWidget* rowWidget = new QWidget();
        rowWidget->setLayout(row);
        m_seriesLayout->addWidget(rowWidget);
    }
}

void ChartPropertiesPanel::applyToChart() {
    if (!m_chart || m_updating) return;

    ChartConfig cfg = m_chart->config();
    cfg.title = m_titleEdit->text();
    cfg.xAxisTitle = m_xAxisEdit->text();
    cfg.yAxisTitle = m_yAxisEdit->text();
    int newTheme = m_themeCombo->currentIndex();
    bool themeChanged = (newTheme != cfg.themeIndex);
    cfg.themeIndex = newTheme;
    cfg.showLegend = m_legendCheck->isChecked();
    cfg.showGridLines = m_gridCheck->isChecked();

    // Deep customization
    if (m_legendPosCombo) cfg.legendPosition = static_cast<LegendPosition>(m_legendPosCombo->currentIndex());
    if (m_dataLabelPosCombo) cfg.dataLabelPosition = static_cast<DataLabelPosition>(m_dataLabelPosCombo->currentIndex());
    if (m_labelShowValue) cfg.dataLabelShowValue = m_labelShowValue->isChecked();
    if (m_labelShowCategory) cfg.dataLabelShowCategory = m_labelShowCategory->isChecked();
    if (m_labelShowPercent) cfg.dataLabelShowPercentage = m_labelShowPercent->isChecked();
    if (m_axisMin) {
        QString minText = m_axisMin->text().trimmed();
        if (minText.toLower() == "auto" || minText.isEmpty()) cfg.yAxisConfig.autoMin = true;
        else { cfg.yAxisConfig.autoMin = false; cfg.yAxisConfig.minValue = minText.toDouble(); }
    }
    if (m_axisMax) {
        QString maxText = m_axisMax->text().trimmed();
        if (maxText.toLower() == "auto" || maxText.isEmpty()) cfg.yAxisConfig.autoMax = true;
        else { cfg.yAxisConfig.autoMax = false; cfg.yAxisConfig.maxValue = maxText.toDouble(); }
    }
    if (m_axisLogScale) cfg.yAxisConfig.logScale = m_axisLogScale->isChecked();
    if (m_stackedCheck) cfg.stacked = m_stackedCheck->isChecked();
    if (m_percentStackedCheck) cfg.percentStacked = m_percentStackedCheck->isChecked();
    if (m_smoothLinesCheck) cfg.smoothLines = m_smoothLinesCheck->isChecked();
    if (m_showMarkersCheck) cfg.showMarkers = m_showMarkersCheck->isChecked();
    if (m_trendlineTypeCombo) {
        int trendIdx = m_trendlineTypeCombo->currentIndex();
        auto trendType = static_cast<TrendlineType>(trendIdx);
        if (trendType != TrendlineType::None) {
            cfg.trendlines.resize(qMax(cfg.trendlines.size(), cfg.series.size()));
            for (auto& tl : cfg.trendlines) {
                tl.type = trendType;
                if (m_trendShowEq) tl.displayEquation = m_trendShowEq->isChecked();
                if (m_trendShowR2) tl.displayRSquared = m_trendShowR2->isChecked();
            }
        } else {
            cfg.trendlines.clear();
        }
    }

    // When theme changes, re-apply theme colors to all series
    if (themeChanged) {
        auto colors = ChartWidget::themeColors(newTheme);
        for (int i = 0; i < cfg.series.size(); ++i) {
            cfg.series[i].color = colors[i % colors.size()];
        }
    }

    m_chart->setConfig(cfg);

    // Rebuild series color swatches to reflect new colors
    if (themeChanged) {
        m_updating = true;
        rebuildSeriesSection();
        m_updating = false;
    }
}

void ChartPropertiesPanel::onPropertyChanged() {
    applyToChart();
}

void ChartPropertiesPanel::onChartTypeClicked() {
    QPushButton* btn = qobject_cast<QPushButton*>(sender());
    if (!btn || !m_chart) return;

    ChartType newType = static_cast<ChartType>(btn->property("chartType").toInt());

    // Update selection visuals
    int typeInt = static_cast<int>(newType);
    for (auto* b : m_typeButtons) {
        bool sel = b->property("chartType").toInt() == typeInt;
        b->setProperty("selected", sel);
        b->setIcon(makeChartTypeIcon(static_cast<ChartType>(b->property("chartType").toInt()), sel));
        b->style()->unpolish(b);
        b->style()->polish(b);
    }

    ChartConfig cfg = m_chart->config();
    cfg.type = newType;
    m_chart->setConfig(cfg);
}

void ChartPropertiesPanel::onRefreshData() {
    if (!m_chart) return;

    QString newRange = m_dataRangeEdit->text().trimmed();
    if (newRange.isEmpty()) return;

    ChartConfig cfg = m_chart->config();
    cfg.dataRange = newRange;
    m_chart->setConfig(cfg);
    m_chart->loadDataFromRange(newRange);

    // Rebuild series section with new data
    m_updating = true;
    rebuildSeriesSection();
    m_updating = false;
}

void ChartPropertiesPanel::onSeriesColorClicked() {
    QPushButton* btn = qobject_cast<QPushButton*>(sender());
    if (!btn || !m_chart) return;

    int idx = btn->property("seriesIndex").toInt();
    ChartConfig cfg = m_chart->config();
    if (idx < 0 || idx >= cfg.series.size()) return;

    QColor current = cfg.series[idx].color;
    QColor newColor = QColorDialog::getColor(current, this, "Series Color",
                                              QColorDialog::ShowAlphaChannel);
    if (newColor.isValid()) {
        cfg.series[idx].color = newColor;
        m_chart->setConfig(cfg);

        // Update button icon
        btn->setIcon(makeColorSwatch(newColor, 20));
    }
}

void ChartPropertiesPanel::onElementChanged(int index) {
    if (m_updating || !m_chart) return;

    // Map combo index to element
    static const ChartElement elements[] = {
        ChartElement::ChartArea, ChartElement::PlotArea, ChartElement::Title,
        ChartElement::ValueAxis, ChartElement::CategoryAxis,
        ChartElement::Legend, ChartElement::DataLabels,
        ChartElement::MajorGridlines, ChartElement::Trendline
    };
    if (index < 0 || index >= 9) return;
    m_currentElement = elements[index];
    showElementOptions(m_currentElement);
}

void ChartPropertiesPanel::showElementOptions(ChartElement element) {
    if (!m_chart) return;

    ChartType chartType = m_chart->config().type;
    bool isCartesian = (chartType != ChartType::Pie && chartType != ChartType::Donut);
    bool isLine = (chartType == ChartType::Line || chartType == ChartType::Scatter);
    bool isBarCol = (chartType == ChartType::Column || chartType == ChartType::Bar ||
                     chartType == ChartType::Histogram);

    // Update header title
    static const char* elementNames[] = {
        "Chart Area", "Plot Area", "Chart Title",
        "X Axis Title", "Y Axis Title", "Legend",
        "Data Series", "Data Labels", "Value Axis",
        "Category Axis", "Major Gridlines", "Trendline"
    };
    int ei = static_cast<int>(element);
    if (ei < 12 && m_headerTitle)
        m_headerTitle->setText(QString("Format %1").arg(elementNames[ei]));

    // Show/hide sections based on element + chart type
    // Find widgets by searching the scroll content
    auto showSection = [this](const QString& sectionTitle, bool visible) {
        if (!m_scrollArea || !m_scrollArea->widget()) return;
        auto children = m_scrollArea->widget()->findChildren<QLabel*>();
        for (auto* label : children) {
            if (label->text() == sectionTitle) {
                // Hide/show this label and the next few widgets until next section
                label->setVisible(visible);
                // Also hide spacing and controls after this header
                QWidget* parent = label->parentWidget();
                if (parent && parent != m_scrollArea->widget()) {
                    parent->setVisible(visible);
                }
            }
        }
    };

    // Default: hide everything, then show relevant sections
    bool showChartType = (element == ChartElement::ChartArea);
    bool showTitleLabels = (element == ChartElement::ChartArea || element == ChartElement::Title ||
                            element == ChartElement::XAxisTitle || element == ChartElement::YAxisTitle);
    bool showStyle = (element == ChartElement::ChartArea);
    bool showData = (element == ChartElement::ChartArea || element == ChartElement::DataSeries);
    bool showSeriesColors = (element == ChartElement::ChartArea || element == ChartElement::DataSeries);
    bool showLegend = (element == ChartElement::Legend || element == ChartElement::ChartArea);
    bool showDataLabels = (element == ChartElement::DataLabels || element == ChartElement::DataSeries);
    bool showAxis = (element == ChartElement::ValueAxis || element == ChartElement::CategoryAxis) && isCartesian;
    bool showChartOpts = (element == ChartElement::ChartArea || element == ChartElement::DataSeries);
    bool showTrendline = (element == ChartElement::Trendline || element == ChartElement::DataSeries) && isCartesian;

    // Apply visibility — using object names on checkboxes and combos
    if (m_stackedCheck) m_stackedCheck->parentWidget()->setVisible(showChartOpts && isBarCol);
    if (m_percentStackedCheck) m_percentStackedCheck->setVisible(showChartOpts && isBarCol);
    if (m_smoothLinesCheck) m_smoothLinesCheck->setVisible(showChartOpts && isLine);
    if (m_showMarkersCheck) m_showMarkersCheck->setVisible(showChartOpts && isLine);

    if (m_legendPosCombo) m_legendPosCombo->parentWidget()->setVisible(showLegend);
    if (m_dataLabelPosCombo) m_dataLabelPosCombo->parentWidget()->setVisible(showDataLabels);
    if (m_labelShowValue) m_labelShowValue->setVisible(showDataLabels);
    if (m_labelShowCategory) m_labelShowCategory->setVisible(showDataLabels);
    if (m_labelShowPercent) {
        // Percentage only for pie/donut
        bool isPie = (chartType == ChartType::Pie || chartType == ChartType::Donut);
        m_labelShowPercent->setVisible(showDataLabels && isPie);
    }

    if (m_axisMin) m_axisMin->parentWidget()->setVisible(showAxis);
    if (m_axisMax) m_axisMax->setVisible(showAxis);
    if (m_axisLogScale) m_axisLogScale->setVisible(showAxis);

    if (m_trendlineTypeCombo) m_trendlineTypeCombo->parentWidget()->setVisible(showTrendline);
    if (m_trendShowEq) m_trendShowEq->setVisible(showTrendline);
    if (m_trendShowR2) m_trendShowR2->setVisible(showTrendline);

    // Add series items to dropdown when in series context
    if (element == ChartElement::DataSeries && m_elementCombo) {
        // Could add individual series entries here
    }
}

void ChartPropertiesPanel::rebuildElementDropdown() {
    if (!m_elementCombo || !m_chart) return;
    m_updating = true;

    m_elementCombo->clear();
    m_elementCombo->addItem("Chart Area");
    m_elementCombo->addItem("Plot Area");

    const auto& cfg = m_chart->config();
    if (!cfg.title.isEmpty()) m_elementCombo->addItem("Chart Title");

    bool isCartesian = (cfg.type != ChartType::Pie && cfg.type != ChartType::Donut);
    if (isCartesian) {
        m_elementCombo->addItem("Value Axis (Y)");
        m_elementCombo->addItem("Category Axis (X)");
    }

    if (cfg.showLegend) m_elementCombo->addItem("Legend");

    // Add series entries
    for (int i = 0; i < cfg.series.size(); ++i) {
        m_elementCombo->addItem(QString("Series \"%1\"").arg(cfg.series[i].name));
    }

    m_elementCombo->addItem("Data Labels");
    if (isCartesian) {
        m_elementCombo->addItem("Major Gridlines");
        m_elementCombo->addItem("Trendline");
    }

    m_updating = false;
}
