#include "ChartDialog.h"
#include "../core/Spreadsheet.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLabel>
#include <QLineEdit>
#include <QComboBox>
#include <QCheckBox>
#include <QListWidget>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QSplitter>
#include <QPainter>
#include <QPainterPath>

// ============== Chart Type Icons ==============

static QIcon createChartTypeIcon(ChartType type) {
    QPixmap pix(48, 48);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);

    QColor primary("#4472C4");
    QColor secondary("#ED7D31");

    switch (type) {
        case ChartType::Line: {
            p.setPen(QPen(primary, 2.5));
            p.drawLine(6, 36, 16, 20);
            p.drawLine(16, 20, 28, 28);
            p.drawLine(28, 28, 42, 10);
            p.setPen(QPen(secondary, 2.5));
            p.drawLine(6, 40, 16, 32);
            p.drawLine(16, 32, 28, 36);
            p.drawLine(28, 36, 42, 24);
            break;
        }
        case ChartType::Column: {
            p.setPen(Qt::NoPen);
            p.setBrush(primary);
            p.drawRect(6, 24, 8, 20);
            p.drawRect(20, 14, 8, 30);
            p.drawRect(34, 8, 8, 36);
            break;
        }
        case ChartType::Bar: {
            p.setPen(Qt::NoPen);
            p.setBrush(primary);
            p.drawRect(6, 6, 30, 8);
            p.drawRect(6, 20, 20, 8);
            p.drawRect(6, 34, 36, 8);
            break;
        }
        case ChartType::Scatter: {
            p.setPen(QPen(primary, 1.5));
            p.setBrush(primary.lighter(130));
            p.drawEllipse(QPoint(12, 32), 4, 4);
            p.drawEllipse(QPoint(20, 24), 4, 4);
            p.drawEllipse(QPoint(28, 18), 4, 4);
            p.drawEllipse(QPoint(36, 12), 4, 4);
            p.drawEllipse(QPoint(18, 16), 3, 3);
            break;
        }
        case ChartType::Pie: {
            p.setPen(QPen(Qt::white, 1.5));
            p.setBrush(primary);
            p.drawPie(6, 6, 36, 36, 0, 200 * 16);
            p.setBrush(secondary);
            p.drawPie(6, 6, 36, 36, 200 * 16, 100 * 16);
            p.setBrush(QColor("#A5A5A5"));
            p.drawPie(6, 6, 36, 36, 300 * 16, 60 * 16);
            break;
        }
        case ChartType::Area: {
            QPainterPath area;
            area.moveTo(6, 44);
            area.lineTo(6, 30);
            area.lineTo(18, 18);
            area.lineTo(30, 24);
            area.lineTo(42, 10);
            area.lineTo(42, 44);
            area.closeSubpath();
            QColor fill = primary;
            fill.setAlpha(120);
            p.setPen(QPen(primary, 2));
            p.setBrush(fill);
            p.drawPath(area);
            break;
        }
        case ChartType::Donut: {
            p.setPen(QPen(Qt::white, 2));
            QPainterPath outer;
            outer.addEllipse(6, 6, 36, 36);
            QPainterPath inner;
            inner.addEllipse(16, 16, 16, 16);
            QPainterPath ring = outer.subtracted(inner);

            // Just draw colored arcs
            p.setBrush(primary);
            p.drawPie(6, 6, 36, 36, 0, 200 * 16);
            p.setBrush(secondary);
            p.drawPie(6, 6, 36, 36, 200 * 16, 160 * 16);
            // White center hole
            p.setPen(Qt::NoPen);
            p.setBrush(Qt::white);
            p.drawEllipse(16, 16, 16, 16);
            break;
        }
        case ChartType::Histogram: {
            p.setPen(Qt::NoPen);
            p.setBrush(primary);
            p.drawRect(4, 32, 7, 12);
            p.drawRect(12, 20, 7, 24);
            p.drawRect(20, 8, 7, 36);
            p.drawRect(28, 16, 7, 28);
            p.drawRect(36, 28, 7, 16);
            break;
        }
    }

    p.end();
    return QIcon(pix);
}

// ============== ChartDialog ==============

ChartDialog::ChartDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Insert Chart");
    setMinimumSize(700, 500);
    resize(750, 550);
    createLayout();

    setStyleSheet(
        "QDialog { background: #FAFBFC; }"
        "QGroupBox { font-weight: bold; border: 1px solid #D0D5DD; border-radius: 6px; "
        "margin-top: 8px; padding-top: 16px; background: white; }"
        "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #344054; }"
        "QLineEdit { border: 1px solid #D0D5DD; border-radius: 4px; padding: 5px 8px; background: white; }"
        "QLineEdit:focus { border-color: #4A90D9; }"
        "QComboBox { border: 1px solid #D0D5DD; border-radius: 4px; padding: 5px 8px; "
        "background: white; min-height: 20px; }"
        "QComboBox:focus { border: 1px solid #4A90D9; }"
        "QComboBox::drop-down { border: none; width: 20px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 6px; }"
        "QComboBox QAbstractItemView { border: 1px solid #D0D5DD; border-radius: 4px; "
        "background: white; selection-background-color: #E8F0FE; padding: 2px; outline: none; }"
        "QCheckBox { spacing: 6px; }"
        "QListWidget { border: 1px solid #D0D5DD; border-radius: 6px; background: white; outline: none; }"
        "QListWidget::item { padding: 6px 8px; border-radius: 4px; }"
        "QListWidget::item:selected { background-color: #E8F0FE; color: #1A1A1A; }"
        "QListWidget::item:hover:!selected { background-color: #F5F5F5; }"
    );
}

void ChartDialog::createLayout() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(12);

    QHBoxLayout* topLayout = new QHBoxLayout();
    topLayout->setSpacing(12);

    // Left: chart type selector
    QGroupBox* typeGroup = new QGroupBox("Chart Type");
    QVBoxLayout* typeLayout = new QVBoxLayout(typeGroup);
    createChartTypeSelector();
    typeLayout->addWidget(m_chartTypeList);
    typeGroup->setFixedWidth(200);
    topLayout->addWidget(typeGroup);

    // Middle: data + options
    QVBoxLayout* midLayout = new QVBoxLayout();
    midLayout->setSpacing(8);

    QGroupBox* dataGroup = new QGroupBox("Data Source");
    createDataPanel();
    QVBoxLayout* dataLayout = new QVBoxLayout(dataGroup);
    dataLayout->addWidget(new QLabel("Data Range (e.g. A1:D10):"));
    dataLayout->addWidget(m_dataRangeEdit);
    QHBoxLayout* checkLayout = new QHBoxLayout();
    checkLayout->addWidget(m_firstRowHeaders);
    checkLayout->addWidget(m_firstColLabels);
    checkLayout->addStretch();
    dataLayout->addLayout(checkLayout);
    midLayout->addWidget(dataGroup);

    QGroupBox* optGroup = new QGroupBox("Options");
    createOptionsPanel();
    QGridLayout* optLayout = new QGridLayout(optGroup);
    optLayout->addWidget(new QLabel("Title:"), 0, 0);
    optLayout->addWidget(m_titleEdit, 0, 1);
    optLayout->addWidget(new QLabel("X Axis:"), 1, 0);
    optLayout->addWidget(m_xAxisEdit, 1, 1);
    optLayout->addWidget(new QLabel("Y Axis:"), 2, 0);
    optLayout->addWidget(m_yAxisEdit, 2, 1);
    optLayout->addWidget(new QLabel("Theme:"), 3, 0);
    optLayout->addWidget(m_themeCombo, 3, 1);
    QHBoxLayout* checkLayout2 = new QHBoxLayout();
    checkLayout2->addWidget(m_showLegend);
    checkLayout2->addWidget(m_showGridLines);
    checkLayout2->addStretch();
    optLayout->addLayout(checkLayout2, 4, 0, 1, 2);
    midLayout->addWidget(optGroup);

    topLayout->addLayout(midLayout, 1);

    // Right: preview
    QGroupBox* previewGroup = new QGroupBox("Preview");
    createPreviewPanel();
    QVBoxLayout* prevLayout = new QVBoxLayout(previewGroup);
    prevLayout->addWidget(m_preview);
    previewGroup->setMinimumWidth(250);
    topLayout->addWidget(previewGroup, 1);

    mainLayout->addLayout(topLayout, 1);

    // Buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Insert Chart");
    buttons->button(QDialogButtonBox::Ok)->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 8px 24px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }"
    );
    buttons->button(QDialogButtonBox::Cancel)->setStyleSheet(
        "QPushButton { background: #F0F2F5; border: 1px solid #D0D5DD; border-radius: 4px; "
        "padding: 8px 20px; }"
        "QPushButton:hover { background: #E8ECF0; }"
    );
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);
}

void ChartDialog::createChartTypeSelector() {
    m_chartTypeList = new QListWidget();
    m_chartTypeList->setIconSize(QSize(48, 48));

    struct ChartTypeInfo { ChartType type; QString name; };
    QVector<ChartTypeInfo> types = {
        { ChartType::Column,    "Column Chart" },
        { ChartType::Bar,       "Bar Chart" },
        { ChartType::Line,      "Line Chart" },
        { ChartType::Area,      "Area Chart" },
        { ChartType::Scatter,   "Scatter Plot" },
        { ChartType::Pie,       "Pie Chart" },
        { ChartType::Donut,     "Donut Chart" },
        { ChartType::Histogram, "Histogram" },
    };

    for (const auto& t : types) {
        auto* item = new QListWidgetItem(createChartTypeIcon(t.type), t.name, m_chartTypeList);
        item->setData(Qt::UserRole, static_cast<int>(t.type));
    }

    m_chartTypeList->setCurrentRow(0);
    connect(m_chartTypeList, &QListWidget::currentRowChanged, this, &ChartDialog::onChartTypeChanged);
}

void ChartDialog::createDataPanel() {
    m_dataRangeEdit = new QLineEdit();
    m_dataRangeEdit->setPlaceholderText("e.g. A1:D10");
    m_firstRowHeaders = new QCheckBox("First row as headers");
    m_firstRowHeaders->setChecked(true);
    m_firstColLabels = new QCheckBox("First column as labels");
    m_firstColLabels->setChecked(true);
}

void ChartDialog::createOptionsPanel() {
    m_titleEdit = new QLineEdit();
    m_titleEdit->setPlaceholderText("Chart Title");
    m_xAxisEdit = new QLineEdit();
    m_xAxisEdit->setPlaceholderText("X Axis Title");
    m_yAxisEdit = new QLineEdit();
    m_yAxisEdit->setPlaceholderText("Y Axis Title");

    m_showLegend = new QCheckBox("Show Legend");
    m_showLegend->setChecked(true);
    m_showGridLines = new QCheckBox("Show Grid Lines");
    m_showGridLines->setChecked(true);

    m_themeCombo = new QComboBox();
    m_themeCombo->setMinimumWidth(160);
    m_themeCombo->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    m_themeCombo->addItems({"Excel", "Material", "Solarized", "Dark", "Monochrome", "Pastel"});

    connect(m_titleEdit, &QLineEdit::textChanged, this, &ChartDialog::updatePreview);
    connect(m_dataRangeEdit, &QLineEdit::textChanged, this, &ChartDialog::updatePreview);
    connect(m_themeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &ChartDialog::updatePreview);
}

void ChartDialog::createPreviewPanel() {
    m_preview = new ChartWidget();
    m_preview->setMinimumSize(220, 180);
    m_preview->setMaximumHeight(280);

    // Add sample data for preview
    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Sample Chart";
    cfg.showLegend = true;
    cfg.showGridLines = true;

    ChartSeries s1;
    s1.name = "Series 1";
    s1.xValues = {1, 2, 3, 4, 5};
    s1.yValues = {25, 40, 30, 50, 35};
    s1.color = QColor("#4472C4");
    cfg.series.append(s1);

    ChartSeries s2;
    s2.name = "Series 2";
    s2.xValues = {1, 2, 3, 4, 5};
    s2.yValues = {15, 30, 45, 20, 40};
    s2.color = QColor("#ED7D31");
    cfg.series.append(s2);

    m_preview->setConfig(cfg);
}

void ChartDialog::onChartTypeChanged(int index) {
    if (index < 0) return;
    updatePreview();
}

void ChartDialog::updatePreview() {
    if (!m_preview) return;

    ChartConfig cfg = m_preview->config();
    if (m_chartTypeList && m_chartTypeList->currentItem()) {
        cfg.type = static_cast<ChartType>(m_chartTypeList->currentItem()->data(Qt::UserRole).toInt());
    }
    if (m_titleEdit) cfg.title = m_titleEdit->text().isEmpty() ? "Sample Chart" : m_titleEdit->text();
    if (m_themeCombo) cfg.themeIndex = m_themeCombo->currentIndex();
    if (m_showLegend) cfg.showLegend = m_showLegend->isChecked();
    if (m_showGridLines) cfg.showGridLines = m_showGridLines->isChecked();

    // Try to load real data from spreadsheet if a data range is set
    QString range = m_dataRangeEdit ? m_dataRangeEdit->text().trimmed() : "";
    if (!range.isEmpty() && m_spreadsheet) {
        cfg.dataRange = range;
        m_preview->setConfig(cfg);
        m_preview->setSpreadsheet(m_spreadsheet);
        m_preview->loadDataFromRange(range);
        return;
    }

    // Update series colors for new theme
    static const QVector<QVector<QColor>> palettes = {
        { QColor("#4472C4"), QColor("#ED7D31") },
        { QColor("#2196F3"), QColor("#FF5722") },
        { QColor("#268BD2"), QColor("#DC322F") },
        { QColor("#00C8FF"), QColor("#FF6384") },
        { QColor("#333333"), QColor("#999999") },
        { QColor("#A8D8EA"), QColor("#FFB7B2") },
    };
    int ti = qBound(0, cfg.themeIndex, static_cast<int>(palettes.size()) - 1);
    for (int i = 0; i < cfg.series.size() && i < palettes[ti].size(); ++i) {
        cfg.series[i].color = palettes[ti][i];
    }

    m_preview->setConfig(cfg);
}

ChartConfig ChartDialog::getConfig() const {
    ChartConfig cfg;
    if (m_chartTypeList && m_chartTypeList->currentItem()) {
        cfg.type = static_cast<ChartType>(m_chartTypeList->currentItem()->data(Qt::UserRole).toInt());
    }
    cfg.title = m_titleEdit ? m_titleEdit->text() : "";
    cfg.xAxisTitle = m_xAxisEdit ? m_xAxisEdit->text() : "";
    cfg.yAxisTitle = m_yAxisEdit ? m_yAxisEdit->text() : "";
    cfg.dataRange = m_dataRangeEdit ? m_dataRangeEdit->text() : "";
    cfg.showLegend = m_showLegend ? m_showLegend->isChecked() : true;
    cfg.showGridLines = m_showGridLines ? m_showGridLines->isChecked() : true;
    cfg.themeIndex = m_themeCombo ? m_themeCombo->currentIndex() : 0;
    return cfg;
}

void ChartDialog::setDataRange(const QString& range) {
    if (m_dataRangeEdit) m_dataRangeEdit->setText(range);
}

void ChartDialog::setSpreadsheet(std::shared_ptr<Spreadsheet> sheet) {
    m_spreadsheet = sheet;
    if (m_preview) {
        m_preview->setSpreadsheet(sheet);
    }
    // If we have a data range, load it immediately
    if (m_dataRangeEdit && !m_dataRangeEdit->text().isEmpty()) {
        updatePreview();
    }
}

void ChartDialog::setConfig(const ChartConfig& config) {
    if (m_titleEdit) m_titleEdit->setText(config.title);
    if (m_xAxisEdit) m_xAxisEdit->setText(config.xAxisTitle);
    if (m_yAxisEdit) m_yAxisEdit->setText(config.yAxisTitle);
    if (m_dataRangeEdit) m_dataRangeEdit->setText(config.dataRange);
    if (m_showLegend) m_showLegend->setChecked(config.showLegend);
    if (m_showGridLines) m_showGridLines->setChecked(config.showGridLines);
    if (m_themeCombo) m_themeCombo->setCurrentIndex(config.themeIndex);

    // Select matching chart type
    if (m_chartTypeList) {
        for (int i = 0; i < m_chartTypeList->count(); ++i) {
            auto* item = m_chartTypeList->item(i);
            if (item->data(Qt::UserRole).toInt() == static_cast<int>(config.type)) {
                m_chartTypeList->setCurrentRow(i);
                break;
            }
        }
    }

    updatePreview();
}

// ============== InsertShapeDialog ==============

static QIcon createShapeIcon(ShapeType type) {
    QPixmap pix(40, 40);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);
    p.setPen(QPen(QColor("#4A90D9"), 2));
    p.setBrush(QColor("#4A90D9").lighter(160));

    switch (type) {
        case ShapeType::Rectangle:
            p.drawRect(4, 8, 32, 24);
            break;
        case ShapeType::RoundedRect:
            p.drawRoundedRect(4, 8, 32, 24, 6, 6);
            break;
        case ShapeType::Circle: {
            int s = 28;
            p.drawEllipse(6, 6, s, s);
            break;
        }
        case ShapeType::Ellipse:
            p.drawEllipse(4, 10, 32, 20);
            break;
        case ShapeType::Triangle: {
            QPolygon tri;
            tri << QPoint(20, 4) << QPoint(4, 36) << QPoint(36, 36);
            p.drawPolygon(tri);
            break;
        }
        case ShapeType::Star: {
            QPolygonF star;
            for (int i = 0; i < 10; ++i) {
                double angle = M_PI / 2 + i * M_PI / 5;
                int r = (i % 2 == 0) ? 16 : 8;
                star << QPointF(20 + r * cos(angle), 20 - r * sin(angle));
            }
            p.drawPolygon(star);
            break;
        }
        case ShapeType::Arrow: {
            QPolygon arr;
            arr << QPoint(36, 20) << QPoint(24, 6) << QPoint(24, 14)
                << QPoint(4, 14) << QPoint(4, 26) << QPoint(24, 26) << QPoint(24, 34);
            p.drawPolygon(arr);
            break;
        }
        case ShapeType::Diamond: {
            QPolygon d;
            d << QPoint(20, 4) << QPoint(36, 20) << QPoint(20, 36) << QPoint(4, 20);
            p.drawPolygon(d);
            break;
        }
        case ShapeType::Pentagon: {
            QPolygonF poly;
            for (int i = 0; i < 5; ++i) {
                double angle = M_PI / 2 + i * 2 * M_PI / 5;
                poly << QPointF(20 + 16 * cos(angle), 20 - 16 * sin(angle));
            }
            p.drawPolygon(poly);
            break;
        }
        case ShapeType::Hexagon: {
            QPolygonF poly;
            for (int i = 0; i < 6; ++i) {
                double angle = i * M_PI / 3;
                poly << QPointF(20 + 16 * cos(angle), 20 - 16 * sin(angle));
            }
            p.drawPolygon(poly);
            break;
        }
        case ShapeType::Callout: {
            QPainterPath path;
            path.addRoundedRect(4, 4, 32, 22, 4, 4);
            path.moveTo(10, 26);
            path.lineTo(8, 36);
            path.lineTo(18, 26);
            p.drawPath(path);
            break;
        }
        case ShapeType::Line:
            p.drawLine(4, 36, 36, 4);
            break;
    }

    p.end();
    return QIcon(pix);
}

InsertShapeDialog::InsertShapeDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Insert Shape");
    setMinimumSize(350, 400);
    createLayout();

    setStyleSheet(
        "QDialog { background: #FAFBFC; }"
        "QListWidget { border: 1px solid #D0D5DD; border-radius: 6px; background: white; outline: none; }"
        "QListWidget::item { padding: 6px 8px; border-radius: 4px; }"
        "QListWidget::item:selected { background-color: #E8F0FE; }"
        "QListWidget::item:hover:!selected { background-color: #F5F5F5; }"
    );
}

void InsertShapeDialog::createLayout() {
    QVBoxLayout* layout = new QVBoxLayout(this);

    QLabel* label = new QLabel("Select a shape:");
    label->setStyleSheet("font-weight: bold; color: #344054; font-size: 13px;");
    layout->addWidget(label);

    m_shapeList = new QListWidget();
    m_shapeList->setIconSize(QSize(40, 40));
    m_shapeList->setViewMode(QListView::IconMode);
    m_shapeList->setGridSize(QSize(80, 70));
    m_shapeList->setResizeMode(QListView::Adjust);
    m_shapeList->setWordWrap(true);
    m_shapeList->setSpacing(4);

    struct ShapeInfo { ShapeType type; QString name; };
    QVector<ShapeInfo> shapes = {
        { ShapeType::Rectangle,   "Rectangle" },
        { ShapeType::RoundedRect, "Rounded" },
        { ShapeType::Circle,      "Circle" },
        { ShapeType::Ellipse,     "Ellipse" },
        { ShapeType::Triangle,    "Triangle" },
        { ShapeType::Star,        "Star" },
        { ShapeType::Arrow,       "Arrow" },
        { ShapeType::Diamond,     "Diamond" },
        { ShapeType::Pentagon,    "Pentagon" },
        { ShapeType::Hexagon,     "Hexagon" },
        { ShapeType::Callout,     "Callout" },
        { ShapeType::Line,        "Line" },
    };

    for (const auto& s : shapes) {
        auto* item = new QListWidgetItem(createShapeIcon(s.type), s.name, m_shapeList);
        item->setData(Qt::UserRole, static_cast<int>(s.type));
        item->setTextAlignment(Qt::AlignHCenter);
    }

    m_shapeList->setCurrentRow(0);
    layout->addWidget(m_shapeList, 1);

    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Insert Shape");
    buttons->button(QDialogButtonBox::Ok)->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 8px 24px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }"
    );
    buttons->button(QDialogButtonBox::Cancel)->setStyleSheet(
        "QPushButton { background: #F0F2F5; border: 1px solid #D0D5DD; border-radius: 4px; "
        "padding: 8px 20px; }"
        "QPushButton:hover { background: #E8ECF0; }"
    );
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    layout->addWidget(buttons);
}

ShapeConfig InsertShapeDialog::getConfig() const {
    ShapeConfig cfg;
    if (m_shapeList && m_shapeList->currentItem()) {
        cfg.type = static_cast<ShapeType>(m_shapeList->currentItem()->data(Qt::UserRole).toInt());
    }
    cfg.fillColor = QColor("#4A90D9");
    cfg.strokeColor = QColor("#2C5F8A");
    cfg.strokeWidth = 2;
    cfg.cornerRadius = 10;
    return cfg;
}
