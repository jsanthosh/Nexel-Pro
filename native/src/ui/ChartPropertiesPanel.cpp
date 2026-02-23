#include "ChartPropertiesPanel.h"
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

    QColor primary = selected ? QColor("#217346") : QColor("#667085");

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
    header->setStyleSheet(
        "QWidget { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
        "stop:0 #1B5E3B, stop:1 #217346); }");
    QHBoxLayout* headerLayout = new QHBoxLayout(header);
    headerLayout->setContentsMargins(14, 0, 8, 0);

    QLabel* headerIcon = new QLabel("\xF0\x9F\x93\x8A"); // chart emoji
    headerIcon->setStyleSheet("QLabel { font-size: 16px; }");
    headerLayout->addWidget(headerIcon);

    QLabel* headerTitle = new QLabel("Chart Properties");
    headerTitle->setStyleSheet(
        "QLabel { color: white; font-size: 13px; font-weight: 600; "
        "letter-spacing: 0.3px; margin-left: 4px; }");
    headerLayout->addWidget(headerTitle);
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
            "QPushButton[selected=\"true\"] { background: #E8F5E9; border: 2px solid #217346; }");
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
        QLineEdit* e = new QLineEdit();
        e->setPlaceholderText(placeholder);
        e->setFixedHeight(30);
        e->setStyleSheet(
            "QLineEdit { border: 1px solid #D0D5DD; border-radius: 6px; padding: 2px 10px; "
            "background: white; font-size: 11px; color: #1D2939; }"
            "QLineEdit:focus { border-color: #34A853; box-shadow: none; }"
            "QLineEdit::placeholder { color: #98A2B3; }");
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
    m_themeCombo->addItems({"Excel", "Material", "Solarized", "Dark", "Monochrome", "Pastel"});
    m_themeCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D0D5DD; border-radius: 6px; padding: 2px 10px; "
        "background: white; font-size: 11px; color: #1D2939; min-height: 22px; }"
        "QComboBox:focus { border: 1px solid #34A853; }"
        "QComboBox::drop-down { border: none; width: 20px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 6px; }"
        "QComboBox QAbstractItemView { border: 1px solid #D0D5DD; border-radius: 6px; "
        "background: white; selection-background-color: #E8F5E9; padding: 4px; outline: none; }");

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
    refreshBtn->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 6px; "
        "padding: 0 14px; font-size: 11px; font-weight: 600; }"
        "QPushButton:hover { background: #1B5E3B; }"
        "QPushButton:pressed { background: #155C30; }");
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

    contentLayout->addStretch();

    m_scrollArea->setWidget(content);
    outerLayout->addWidget(m_scrollArea);

    setStyleSheet("ChartPropertiesPanel { background: #F8FAFB; }");
}

void ChartPropertiesPanel::setChart(ChartWidget* chart) {
    m_chart = chart;
    m_updating = true;
    updateFromChart();
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
        colorBtn->setStyleSheet(
            "QPushButton { background: white; border: 1.5px solid #E4E7EC; border-radius: 6px; padding: 2px; }"
            "QPushButton:hover { border-color: #34A853; background: #F0FAF3; }");
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
