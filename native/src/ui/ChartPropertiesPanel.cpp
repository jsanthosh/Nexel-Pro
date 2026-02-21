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
        "QLabel { color: #344054; font-size: 10px; font-weight: bold; "
        "letter-spacing: 1px; padding: 8px 0 4px 0; }");
    return label;
}

void ChartPropertiesPanel::createLayout() {
    QVBoxLayout* outerLayout = new QVBoxLayout(this);
    outerLayout->setContentsMargins(0, 0, 0, 0);
    outerLayout->setSpacing(0);

    // Header bar
    QWidget* header = new QWidget();
    header->setFixedHeight(40);
    header->setStyleSheet("QWidget { background: #1B5E3B; }");
    QHBoxLayout* headerLayout = new QHBoxLayout(header);
    headerLayout->setContentsMargins(12, 0, 8, 0);

    QLabel* headerTitle = new QLabel("Chart Properties");
    headerTitle->setStyleSheet("QLabel { color: white; font-size: 13px; font-weight: bold; }");
    headerLayout->addWidget(headerTitle);
    headerLayout->addStretch();

    QPushButton* closeBtn = new QPushButton("\u00D7"); // ×
    closeBtn->setFixedSize(24, 24);
    closeBtn->setStyleSheet(
        "QPushButton { background: transparent; color: white; font-size: 18px; "
        "font-weight: bold; border: none; border-radius: 12px; }"
        "QPushButton:hover { background: rgba(255,255,255,0.2); }");
    connect(closeBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::closeRequested);
    headerLayout->addWidget(closeBtn);

    outerLayout->addWidget(header);

    // Scroll area for content
    m_scrollArea = new QScrollArea();
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_scrollArea->setStyleSheet(
        "QScrollArea { border: none; background: #FAFBFC; }"
        "QScrollBar:vertical { width: 6px; background: transparent; }"
        "QScrollBar::handle:vertical { background: #C0C5CC; border-radius: 3px; min-height: 30px; }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }");

    QWidget* content = new QWidget();
    QVBoxLayout* contentLayout = new QVBoxLayout(content);
    contentLayout->setContentsMargins(12, 8, 12, 12);
    contentLayout->setSpacing(2);

    // ===== Chart Type Section =====
    contentLayout->addWidget(createSectionHeader("CHART TYPE"));

    QGridLayout* typeGrid = new QGridLayout();
    typeGrid->setSpacing(4);

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
        btn->setFixedSize(36, 36);
        btn->setIconSize(QSize(28, 28));
        btn->setIcon(makeChartTypeIcon(types[i].type, false));
        btn->setToolTip(types[i].tip);
        btn->setProperty("chartType", static_cast<int>(types[i].type));
        btn->setStyleSheet(
            "QPushButton { background: white; border: 1px solid #E0E3E8; border-radius: 6px; }"
            "QPushButton:hover { background: #F0F2F5; border-color: #B0B5BD; }"
            "QPushButton[selected=\"true\"] { background: #E8F5E9; border: 2px solid #217346; }");
        connect(btn, &QPushButton::clicked, this, &ChartPropertiesPanel::onChartTypeClicked);
        typeGrid->addWidget(btn, i / 4, i % 4);
        m_typeButtons.append(btn);
    }

    contentLayout->addLayout(typeGrid);

    // Separator
    QFrame* sep1 = new QFrame();
    sep1->setFrameShape(QFrame::HLine);
    sep1->setStyleSheet("QFrame { color: #E0E3E8; margin: 6px 0; }");
    contentLayout->addWidget(sep1);

    // ===== Title & Labels =====
    contentLayout->addWidget(createSectionHeader("TITLE & LABELS"));

    QGridLayout* labelGrid = new QGridLayout();
    labelGrid->setSpacing(6);
    labelGrid->setColumnStretch(1, 1);

    auto makeLabel = [](const QString& text) {
        QLabel* l = new QLabel(text);
        l->setStyleSheet("QLabel { color: #667085; font-size: 11px; }");
        return l;
    };

    auto makeEdit = [](const QString& placeholder) {
        QLineEdit* e = new QLineEdit();
        e->setPlaceholderText(placeholder);
        e->setFixedHeight(28);
        e->setStyleSheet(
            "QLineEdit { border: 1px solid #D0D5DD; border-radius: 4px; padding: 2px 8px; "
            "background: white; font-size: 11px; }"
            "QLineEdit:focus { border-color: #217346; }");
        return e;
    };

    m_titleEdit = makeEdit("Chart title");
    m_xAxisEdit = makeEdit("X axis label");
    m_yAxisEdit = makeEdit("Y axis label");

    labelGrid->addWidget(makeLabel("Title"), 0, 0);
    labelGrid->addWidget(m_titleEdit, 0, 1);
    labelGrid->addWidget(makeLabel("X Axis"), 1, 0);
    labelGrid->addWidget(m_xAxisEdit, 1, 1);
    labelGrid->addWidget(makeLabel("Y Axis"), 2, 0);
    labelGrid->addWidget(m_yAxisEdit, 2, 1);

    contentLayout->addLayout(labelGrid);

    connect(m_titleEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_xAxisEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_yAxisEdit, &QLineEdit::textChanged, this, &ChartPropertiesPanel::onPropertyChanged);

    // Separator
    QFrame* sep2 = new QFrame();
    sep2->setFrameShape(QFrame::HLine);
    sep2->setStyleSheet("QFrame { color: #E0E3E8; margin: 6px 0; }");
    contentLayout->addWidget(sep2);

    // ===== Style =====
    contentLayout->addWidget(createSectionHeader("STYLE"));

    QGridLayout* styleGrid = new QGridLayout();
    styleGrid->setSpacing(6);
    styleGrid->setColumnStretch(1, 1);

    m_themeCombo = new QComboBox();
    m_themeCombo->setFixedHeight(28);
    m_themeCombo->addItems({"Excel", "Material", "Solarized", "Dark", "Monochrome", "Pastel"});
    m_themeCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D0D5DD; border-radius: 4px; padding: 2px 8px; "
        "background: white; font-size: 11px; min-height: 20px; }"
        "QComboBox:focus { border: 1px solid #217346; }"
        "QComboBox::drop-down { border: none; width: 18px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 4px; }"
        "QComboBox QAbstractItemView { border: 1px solid #D0D5DD; border-radius: 4px; "
        "background: white; selection-background-color: #E8F5E9; padding: 2px; outline: none; }");

    styleGrid->addWidget(makeLabel("Theme"), 0, 0);
    styleGrid->addWidget(m_themeCombo, 0, 1);

    contentLayout->addLayout(styleGrid);

    connect(m_themeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartPropertiesPanel::onPropertyChanged);

    m_legendCheck = new QCheckBox("Show Legend");
    m_legendCheck->setStyleSheet("QCheckBox { color: #344054; font-size: 11px; spacing: 6px; }");
    m_gridCheck = new QCheckBox("Show Grid Lines");
    m_gridCheck->setStyleSheet("QCheckBox { color: #344054; font-size: 11px; spacing: 6px; }");

    contentLayout->addWidget(m_legendCheck);
    contentLayout->addWidget(m_gridCheck);

    connect(m_legendCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);
    connect(m_gridCheck, &QCheckBox::toggled, this, &ChartPropertiesPanel::onPropertyChanged);

    // Separator
    QFrame* sep3 = new QFrame();
    sep3->setFrameShape(QFrame::HLine);
    sep3->setStyleSheet("QFrame { color: #E0E3E8; margin: 6px 0; }");
    contentLayout->addWidget(sep3);

    // ===== Data =====
    contentLayout->addWidget(createSectionHeader("DATA"));

    QHBoxLayout* dataLayout = new QHBoxLayout();
    dataLayout->setSpacing(6);

    m_dataRangeEdit = makeEdit("A1:D10");
    dataLayout->addWidget(m_dataRangeEdit, 1);

    QPushButton* refreshBtn = new QPushButton("Refresh");
    refreshBtn->setFixedHeight(28);
    refreshBtn->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 0 12px; font-size: 11px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }");
    connect(refreshBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::onRefreshData);
    dataLayout->addWidget(refreshBtn);

    contentLayout->addLayout(dataLayout);

    connect(m_dataRangeEdit, &QLineEdit::editingFinished, this, &ChartPropertiesPanel::onRefreshData);

    // Separator
    QFrame* sep4 = new QFrame();
    sep4->setFrameShape(QFrame::HLine);
    sep4->setStyleSheet("QFrame { color: #E0E3E8; margin: 6px 0; }");
    contentLayout->addWidget(sep4);

    // ===== Series Colors =====
    contentLayout->addWidget(createSectionHeader("SERIES COLORS"));

    m_seriesContainer = new QWidget();
    m_seriesLayout = new QVBoxLayout(m_seriesContainer);
    m_seriesLayout->setContentsMargins(0, 0, 0, 0);
    m_seriesLayout->setSpacing(4);
    contentLayout->addWidget(m_seriesContainer);

    contentLayout->addStretch();

    m_scrollArea->setWidget(content);
    outerLayout->addWidget(m_scrollArea);

    setStyleSheet(
        "ChartPropertiesPanel { background: #FAFBFC; }"
    );
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

        QPushButton* colorBtn = new QPushButton();
        colorBtn->setFixedSize(22, 22);
        colorBtn->setIcon(makeColorSwatch(series[i].color, 18));
        colorBtn->setIconSize(QSize(18, 18));
        colorBtn->setStyleSheet(
            "QPushButton { background: transparent; border: 1px solid #D0D5DD; border-radius: 4px; padding: 1px; }"
            "QPushButton:hover { border-color: #217346; }");
        colorBtn->setProperty("seriesIndex", i);
        connect(colorBtn, &QPushButton::clicked, this, &ChartPropertiesPanel::onSeriesColorClicked);

        QLabel* nameLabel = new QLabel(series[i].name);
        nameLabel->setStyleSheet("QLabel { color: #344054; font-size: 11px; }");

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
    cfg.themeIndex = m_themeCombo->currentIndex();
    cfg.showLegend = m_legendCheck->isChecked();
    cfg.showGridLines = m_gridCheck->isChecked();

    m_chart->setConfig(cfg);
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
        btn->setIcon(makeColorSwatch(newColor, 18));
    }
}
