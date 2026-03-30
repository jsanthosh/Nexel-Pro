#include "FormatCellsDialog.h"
#include "../core/NumberFormat.h"
#include "../core/DocumentTheme.h"
#include "Theme.h"
#include <QTabWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QListWidget>
#include <QSpinBox>
#include <QCheckBox>
#include <QComboBox>
#include <QLineEdit>
#include <QLabel>
#include <QFontComboBox>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QColorDialog>
#include <QStackedWidget>
#include <QMenu>
#include <QWidgetAction>
#include <QFrame>
#include <QPainter>

// --- Minimal color swatch widget for the theme-aware picker ---
namespace {
class FmtColorSwatch : public QWidget {
public:
    QColor color;
    bool selected = false;
    bool hovered = false;
    std::function<void()> onClick;

    FmtColorSwatch(const QColor& c, QWidget* parent) : QWidget(parent), color(c) {
        setFixedSize(18, 18);
        setCursor(Qt::PointingHandCursor);
        setAttribute(Qt::WA_Hover, true);
        setMouseTracking(true);
    }

protected:
    void paintEvent(QPaintEvent*) override {
        QPainter p(this);
        p.setRenderHint(QPainter::Antialiasing, false);
        QRect r = rect();
        p.fillRect(r, color);
        if (selected) {
            p.setPen(QPen(ThemeManager::instance().currentTheme().accentDark, 2));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        } else if (hovered) {
            p.setPen(QPen(QColor("#333333"), 2));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        } else if (color.lightness() > 220) {
            p.setPen(QPen(QColor("#D0D0D0"), 1));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        }
    }
    void enterEvent(QEnterEvent*) override { hovered = true; update(); }
    void leaveEvent(QEvent*) override { hovered = false; update(); }
    void mousePressEvent(QMouseEvent*) override { if (onClick) onClick(); }
};
} // anonymous namespace

// Theme-aware color picker (shared logic for font and fill color buttons)
void FormatCellsDialog::pickColor(const QString& title, QString& colorStr, QPushButton* btn) {
    const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();

    static const QColor standardColors[] = {
        QColor("#C00000"), QColor("#FF0000"), QColor("#FFC000"), QColor("#FFFF00"),
        QColor("#92D050"), QColor("#00B050"), QColor("#00B0F0"), QColor("#0070C0"),
        QColor("#002060"), QColor("#7030A0"),
    };
    static const QColor grayscale[] = {
        QColor("#000000"), QColor("#1A1A1A"), QColor("#333333"), QColor("#4D4D4D"),
        QColor("#666666"), QColor("#808080"), QColor("#999999"), QColor("#B3B3B3"),
        QColor("#D9D9D9"), QColor("#FFFFFF"),
    };
    static const int COLS = 10;

    QColor currentColor = dt.resolveAnyColor(colorStr);

    QMenu menu(this);
    menu.setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #E0E0E0; padding: 0px; border-radius: 6px; }");

    QWidget* container = new QWidget(&menu);
    QVBoxLayout* layout = new QVBoxLayout(container);
    layout->setContentsMargins(10, 8, 10, 8);
    layout->setSpacing(0);

    // Theme Colors header
    QLabel* themeLabel = new QLabel("Theme Colors", container);
    themeLabel->setStyleSheet("font: 11px 'Segoe UI', 'SF Pro Text', sans-serif; color: #666; padding-bottom: 4px;");
    layout->addWidget(themeLabel);

    QString pickedStr;
    bool picked = false;

    // Theme color grid: 6 rows × 10 columns
    QGridLayout* grid = new QGridLayout();
    grid->setSpacing(3);
    grid->setContentsMargins(0, 0, 0, 0);
    for (int r = 0; r < kThemeTintCount; ++r) {
        for (int c = 0; c < COLS; ++c) {
            QColor swatchColor = DocumentTheme::applyTint(dt.colors[c], kThemeTints[r]);
            FmtColorSwatch* swatch = new FmtColorSwatch(swatchColor, container);
            swatch->selected = currentColor.isValid() && (swatchColor == currentColor);
            swatch->setToolTip(themeColorName(c, kThemeTints[r]));
            swatch->onClick = [&pickedStr, &picked, &menu, c, r]() {
                pickedStr = DocumentTheme::makeThemeColorStr(c, kThemeTints[r]);
                picked = true;
                menu.close();
            };
            grid->addWidget(swatch, r, c);
        }
    }
    layout->addLayout(grid);

    // Separator
    layout->addSpacing(6);
    QFrame* sep1 = new QFrame(container);
    sep1->setFrameShape(QFrame::HLine);
    sep1->setStyleSheet("background: #E8E8E8; max-height: 1px;");
    layout->addWidget(sep1);
    layout->addSpacing(4);

    // Standard Colors
    QLabel* stdLabel = new QLabel("Standard Colors", container);
    stdLabel->setStyleSheet("font: 11px 'Segoe UI', 'SF Pro Text', sans-serif; color: #666; padding-bottom: 4px;");
    layout->addWidget(stdLabel);
    QHBoxLayout* stdRow = new QHBoxLayout();
    stdRow->setSpacing(3);
    stdRow->setContentsMargins(0, 0, 0, 0);
    for (int c = 0; c < COLS; ++c) {
        FmtColorSwatch* swatch = new FmtColorSwatch(standardColors[c], container);
        swatch->selected = currentColor.isValid() && (standardColors[c] == currentColor);
        swatch->onClick = [&pickedStr, &picked, &menu, c]() {
            pickedStr = standardColors[c].name();
            picked = true;
            menu.close();
        };
        stdRow->addWidget(swatch);
    }
    layout->addLayout(stdRow);
    layout->addSpacing(4);

    // Grayscale
    QHBoxLayout* grayRow = new QHBoxLayout();
    grayRow->setSpacing(3);
    grayRow->setContentsMargins(0, 0, 0, 0);
    for (int c = 0; c < COLS; ++c) {
        FmtColorSwatch* swatch = new FmtColorSwatch(grayscale[c], container);
        swatch->selected = currentColor.isValid() && (grayscale[c] == currentColor);
        swatch->onClick = [&pickedStr, &picked, &menu, c]() {
            pickedStr = grayscale[c].name();
            picked = true;
            menu.close();
        };
        grayRow->addWidget(swatch);
    }
    layout->addLayout(grayRow);

    // Separator + actions
    layout->addSpacing(6);
    QFrame* sep2 = new QFrame(container);
    sep2->setFrameShape(QFrame::HLine);
    sep2->setStyleSheet("background: #E8E8E8; max-height: 1px;");
    layout->addWidget(sep2);
    layout->addSpacing(4);

    // No Fill (for fill picker only)
    if (title.contains("Fill")) {
        QPushButton* noFillBtn = new QPushButton("No Fill", container);
        noFillBtn->setFixedHeight(26);
        noFillBtn->setCursor(Qt::PointingHandCursor);
        noFillBtn->setStyleSheet(
            "QPushButton { background: transparent; border: none; font: 12px 'Segoe UI', sans-serif;"
            "  color: #444; text-align: left; padding-left: 2px; }"
            "QPushButton:hover { background: #F0F4F8; border-radius: 3px; }");
        connect(noFillBtn, &QPushButton::clicked, &menu, [&pickedStr, &picked, &menu]() {
            pickedStr = "#FFFFFF";
            picked = true;
            menu.close();
        });
        layout->addWidget(noFillBtn);
    }

    // Custom Color
    QPushButton* customBtn = new QPushButton("Custom Color...", container);
    customBtn->setFixedHeight(26);
    customBtn->setCursor(Qt::PointingHandCursor);
    customBtn->setStyleSheet(
        "QPushButton { background: transparent; border: none; font: 12px 'Segoe UI', sans-serif;"
        "  color: #2980B9; text-align: left; padding-left: 2px; }"
        "QPushButton:hover { background: #F0F4F8; border-radius: 3px; }");
    connect(customBtn, &QPushButton::clicked, &menu, [&pickedStr, &picked, &menu, this, currentColor, title]() {
        menu.close();
        QColor custom = QColorDialog::getColor(currentColor, this, title);
        if (custom.isValid()) {
            pickedStr = custom.name();
            picked = true;
        }
    });
    layout->addWidget(customBtn);

    QWidgetAction* wa = new QWidgetAction(&menu);
    wa->setDefaultWidget(container);
    menu.addAction(wa);

    menu.exec(QCursor::pos());

    if (picked) {
        colorStr = pickedStr;
        QColor displayColor = dt.resolveAnyColor(pickedStr);
        btn->setStyleSheet(QString("background-color: %1;").arg(displayColor.name()));
    }
}

FormatCellsDialog::FormatCellsDialog(const CellStyle& style, QWidget* parent)
    : QDialog(parent), m_style(style) {
    setWindowTitle("Format Cells");
    setMinimumSize(520, 420);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    QTabWidget* tabs = new QTabWidget(this);

    QWidget* numberTab = new QWidget();
    createNumberTab(numberTab);
    tabs->addTab(numberTab, "Number");

    QWidget* fontTab = new QWidget();
    createFontTab(fontTab);
    tabs->addTab(fontTab, "Font");

    QWidget* alignTab = new QWidget();
    createAlignmentTab(alignTab);
    tabs->addTab(alignTab, "Alignment");

    QWidget* borderTab = new QWidget();
    createBorderTab(borderTab);
    tabs->addTab(borderTab, "Border");

    QWidget* fillTab = new QWidget();
    createFillTab(fillTab);
    tabs->addTab(fillTab, "Fill");

    QWidget* protTab = new QWidget();
    createProtectionTab(protTab);
    tabs->addTab(protTab, "Protection");

    mainLayout->addWidget(tabs);

    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);

    loadStyle(style);
}

void FormatCellsDialog::createNumberTab(QWidget* tab) {
    QHBoxLayout* layout = new QHBoxLayout(tab);

    // Category list
    m_categoryList = new QListWidget();
    m_categoryList->addItem("General");
    m_categoryList->addItem("Number");
    m_categoryList->addItem("Currency");
    m_categoryList->addItem("Accounting");
    m_categoryList->addItem("Percentage");
    m_categoryList->addItem("Date");
    m_categoryList->addItem("Time");
    m_categoryList->addItem("Text");
    m_categoryList->addItem("Custom");
    m_categoryList->setMaximumWidth(120);
    layout->addWidget(m_categoryList);

    // Options panel
    QVBoxLayout* optionsLayout = new QVBoxLayout();

    // Preview
    QGroupBox* previewBox = new QGroupBox("Preview");
    QVBoxLayout* previewLayout = new QVBoxLayout(previewBox);
    m_previewLabel = new QLabel("General");
    m_previewLabel->setStyleSheet("QLabel { padding: 8px; background: white; border: 1px solid #ccc; }");
    previewLayout->addWidget(m_previewLabel);
    optionsLayout->addWidget(previewBox);

    // Decimal places
    QHBoxLayout* decimalRow = new QHBoxLayout();
    decimalRow->addWidget(new QLabel("Decimal places:"));
    m_decimalSpin = new QSpinBox();
    m_decimalSpin->setRange(0, 10);
    m_decimalSpin->setValue(2);
    decimalRow->addWidget(m_decimalSpin);
    optionsLayout->addLayout(decimalRow);

    // Thousand separator
    m_thousandCheck = new QCheckBox("Use 1000 separator (,)");
    optionsLayout->addWidget(m_thousandCheck);

    // Currency
    QHBoxLayout* currencyRow = new QHBoxLayout();
    currencyRow->addWidget(new QLabel("Currency:"));
    m_currencyCombo = new QComboBox();
    for (const auto& c : NumberFormat::currencies()) {
        m_currencyCombo->addItem(c.label, c.code);
    }
    currencyRow->addWidget(m_currencyCombo);
    optionsLayout->addLayout(currencyRow);

    // Date format
    QHBoxLayout* dateRow = new QHBoxLayout();
    dateRow->addWidget(new QLabel("Date format:"));
    m_dateFormatCombo = new QComboBox();
    m_dateFormatCombo->addItem("MM/DD/YYYY", "mm/dd/yyyy");
    m_dateFormatCombo->addItem("DD/MM/YYYY", "dd/mm/yyyy");
    m_dateFormatCombo->addItem("YYYY-MM-DD", "yyyy-mm-dd");
    m_dateFormatCombo->addItem("MMM D, YYYY", "mmm d, yyyy");
    m_dateFormatCombo->addItem("MMMM D, YYYY", "mmmm d, yyyy");
    m_dateFormatCombo->addItem("D-MMM-YY", "d-mmm-yy");
    m_dateFormatCombo->addItem("MM/DD", "mm/dd");
    dateRow->addWidget(m_dateFormatCombo);
    optionsLayout->addLayout(dateRow);

    // Custom format
    QHBoxLayout* customRow = new QHBoxLayout();
    customRow->addWidget(new QLabel("Custom:"));
    m_customFormatEdit = new QLineEdit();
    m_customFormatEdit->setPlaceholderText("#,##0.00");
    customRow->addWidget(m_customFormatEdit);
    optionsLayout->addLayout(customRow);

    optionsLayout->addStretch();
    layout->addLayout(optionsLayout);

    // Connect category change
    connect(m_categoryList, &QListWidget::currentRowChanged, this, [this](int row) {
        QStringList types = {"General", "Number", "Currency", "Accounting",
                             "Percentage", "Date", "Time", "Text", "Custom"};
        if (row >= 0 && row < types.size()) {
            m_style.numberFormat = types[row];
            updatePreview();
        }
    });

    connect(m_decimalSpin, &QSpinBox::valueChanged, this, [this](int val) {
        m_style.decimalPlaces = val;
        updatePreview();
    });

    connect(m_thousandCheck, &QCheckBox::toggled, this, [this](bool checked) {
        m_style.useThousandsSeparator = checked;
        updatePreview();
    });

    connect(m_currencyCombo, &QComboBox::currentIndexChanged, this, [this](int idx) {
        m_style.currencyCode = m_currencyCombo->itemData(idx).toString();
        updatePreview();
    });

    connect(m_dateFormatCombo, &QComboBox::currentIndexChanged, this, [this](int idx) {
        m_style.dateFormatId = m_dateFormatCombo->itemData(idx).toString();
        updatePreview();
    });

    connect(m_customFormatEdit, &QLineEdit::textChanged, this, [this](const QString&) {
        updatePreview();
    });
}

void FormatCellsDialog::createFontTab(QWidget* tab) {
    QGridLayout* layout = new QGridLayout(tab);

    layout->addWidget(new QLabel("Font:"), 0, 0);
    m_fontFamilyCombo = new QFontComboBox();
    layout->addWidget(m_fontFamilyCombo, 0, 1, 1, 2);

    layout->addWidget(new QLabel("Size:"), 1, 0);
    m_fontSizeSpin = new QSpinBox();
    m_fontSizeSpin->setRange(6, 72);
    layout->addWidget(m_fontSizeSpin, 1, 1);

    QGroupBox* styleGroup = new QGroupBox("Style");
    QVBoxLayout* styleLayout = new QVBoxLayout(styleGroup);
    m_boldCheck = new QCheckBox("Bold");
    m_italicCheck = new QCheckBox("Italic");
    m_underlineCheck = new QCheckBox("Underline");
    m_strikethroughCheck = new QCheckBox("Strikethrough");
    styleLayout->addWidget(m_boldCheck);
    styleLayout->addWidget(m_italicCheck);
    styleLayout->addWidget(m_underlineCheck);
    styleLayout->addWidget(m_strikethroughCheck);
    layout->addWidget(styleGroup, 2, 0, 1, 3);

    QHBoxLayout* colorRow = new QHBoxLayout();
    colorRow->addWidget(new QLabel("Color:"));
    m_fontColorBtn = new QPushButton();
    m_fontColorBtn->setFixedSize(60, 24);
    colorRow->addWidget(m_fontColorBtn);
    colorRow->addStretch();
    layout->addLayout(colorRow, 3, 0, 1, 3);

    connect(m_fontColorBtn, &QPushButton::clicked, this, [this]() {
        pickColor("Font Color", m_fontColorStr, m_fontColorBtn);
    });

    layout->setRowStretch(4, 1);
}

void FormatCellsDialog::createAlignmentTab(QWidget* tab) {
    QGridLayout* layout = new QGridLayout(tab);

    layout->addWidget(new QLabel("Horizontal:"), 0, 0);
    m_hAlignCombo = new QComboBox();
    m_hAlignCombo->addItem("General", static_cast<int>(HorizontalAlignment::General));
    m_hAlignCombo->addItem("Left", static_cast<int>(HorizontalAlignment::Left));
    m_hAlignCombo->addItem("Center", static_cast<int>(HorizontalAlignment::Center));
    m_hAlignCombo->addItem("Right", static_cast<int>(HorizontalAlignment::Right));
    layout->addWidget(m_hAlignCombo, 0, 1);

    layout->addWidget(new QLabel("Vertical:"), 1, 0);
    m_vAlignCombo = new QComboBox();
    m_vAlignCombo->addItem("Top", static_cast<int>(VerticalAlignment::Top));
    m_vAlignCombo->addItem("Middle", static_cast<int>(VerticalAlignment::Middle));
    m_vAlignCombo->addItem("Bottom", static_cast<int>(VerticalAlignment::Bottom));
    layout->addWidget(m_vAlignCombo, 1, 1);

    layout->setRowStretch(2, 1);
}

void FormatCellsDialog::createFillTab(QWidget* tab) {
    QVBoxLayout* layout = new QVBoxLayout(tab);

    QHBoxLayout* colorRow = new QHBoxLayout();
    colorRow->addWidget(new QLabel("Background color:"));
    m_fillColorBtn = new QPushButton();
    m_fillColorBtn->setFixedSize(60, 24);
    colorRow->addWidget(m_fillColorBtn);
    colorRow->addStretch();
    layout->addLayout(colorRow);

    connect(m_fillColorBtn, &QPushButton::clicked, this, [this]() {
        pickColor("Fill Color", m_fillColorStr, m_fillColorBtn);
    });

    layout->addStretch();
}

void FormatCellsDialog::loadStyle(const CellStyle& style) {
    // Number tab
    QStringList types = {"General", "Number", "Currency", "Accounting",
                         "Percentage", "Date", "Time", "Text", "Custom"};
    int idx = types.indexOf(style.numberFormat);
    if (idx >= 0) m_categoryList->setCurrentRow(idx);
    m_decimalSpin->setValue(style.decimalPlaces);
    m_thousandCheck->setChecked(style.useThousandsSeparator);
    for (int i = 0; i < m_currencyCombo->count(); ++i) {
        if (m_currencyCombo->itemData(i).toString() == style.currencyCode) {
            m_currencyCombo->setCurrentIndex(i);
            break;
        }
    }
    for (int i = 0; i < m_dateFormatCombo->count(); ++i) {
        if (m_dateFormatCombo->itemData(i).toString() == style.dateFormatId) {
            m_dateFormatCombo->setCurrentIndex(i);
            break;
        }
    }

    // Font tab
    m_fontFamilyCombo->setCurrentFont(QFont(style.fontName));
    m_fontSizeSpin->setValue(style.fontSize);
    m_boldCheck->setChecked(style.bold);
    m_italicCheck->setChecked(style.italic);
    m_underlineCheck->setChecked(style.underline);
    m_strikethroughCheck->setChecked(style.strikethrough);
    m_fontColorStr = style.foregroundColor;
    const DocumentTheme& dt = m_docTheme ? *m_docTheme : defaultDocumentTheme();
    QColor fgDisplay = dt.resolveAnyColor(m_fontColorStr);
    m_fontColorBtn->setStyleSheet(
        QString("background-color: %1;").arg(fgDisplay.name()));

    // Alignment
    m_hAlignCombo->setCurrentIndex(static_cast<int>(style.hAlign));
    m_vAlignCombo->setCurrentIndex(static_cast<int>(style.vAlign));

    // Fill
    m_fillColorStr = style.backgroundColor;
    QColor bgDisplay = dt.resolveAnyColor(m_fillColorStr);
    m_fillColorBtn->setStyleSheet(
        QString("background-color: %1;").arg(bgDisplay.name()));

    // Protection
    if (m_lockedCheck) m_lockedCheck->setChecked(style.locked);
    if (m_hiddenCheck) m_hiddenCheck->setChecked(style.hidden);

    updatePreview();
}

CellStyle FormatCellsDialog::getStyle() const {
    CellStyle style = m_style;

    // Number
    style.decimalPlaces = m_decimalSpin->value();
    style.useThousandsSeparator = m_thousandCheck->isChecked();
    style.currencyCode = m_currencyCombo->currentData().toString();
    style.dateFormatId = m_dateFormatCombo->currentData().toString();

    // Font
    style.fontName = m_fontFamilyCombo->currentFont().family();
    style.fontSize = m_fontSizeSpin->value();
    style.bold = m_boldCheck->isChecked();
    style.italic = m_italicCheck->isChecked();
    style.underline = m_underlineCheck->isChecked();
    style.strikethrough = m_strikethroughCheck->isChecked();
    style.foregroundColor = m_fontColorStr;

    // Alignment
    style.hAlign = static_cast<HorizontalAlignment>(m_hAlignCombo->currentData().toInt());
    style.vAlign = static_cast<VerticalAlignment>(m_vAlignCombo->currentData().toInt());

    // Fill
    style.backgroundColor = m_fillColorStr;

    // Protection
    style.locked = m_lockedCheck->isChecked() ? 1 : 0;
    style.hidden = m_hiddenCheck->isChecked() ? 1 : 0;

    return style;
}

void FormatCellsDialog::updatePreview() {
    NumberFormatOptions opts;
    opts.type = NumberFormat::typeFromString(m_style.numberFormat);
    opts.decimalPlaces = m_decimalSpin->value();
    opts.useThousandsSeparator = m_thousandCheck->isChecked();
    opts.currencyCode = m_currencyCombo->currentData().toString();
    opts.dateFormatId = m_dateFormatCombo->currentData().toString();
    opts.customFormat = m_customFormatEdit->text();

    QString sample = "1234.56";
    if (opts.type == NumberFormatType::Percentage) sample = "0.1234";
    if (opts.type == NumberFormatType::Date) sample = "2026-02-17";

    QString formatted = NumberFormat::format(sample, opts);
    m_previewLabel->setText(formatted);
}

// ============================================================================
// Border Tab (Excel-style: presets + line style + per-edge toggles)
// ============================================================================
void FormatCellsDialog::createBorderTab(QWidget* tab) {
    QHBoxLayout* mainLayout = new QHBoxLayout(tab);

    // Left side: Line style + color
    QVBoxLayout* lineLayout = new QVBoxLayout();
    lineLayout->addWidget(new QLabel("Line"));

    // Style list
    m_borderStyleList = new QListWidget();
    m_borderStyleList->setFixedWidth(120);
    m_borderStyleList->setFixedHeight(180);
    QStringList styles = {"None", "Hair", "Thin", "Thin Dashed", "Thin Dash-Dot",
                          "Thin Dash-Dot-Dot", "Medium", "Medium Dashed",
                          "Medium Dash-Dot", "Medium Dash-Dot-Dot",
                          "Thick", "Double", "Slant Dash-Dot"};
    m_borderStyleList->addItems(styles);
    m_borderStyleList->setCurrentRow(2); // Default: Thin
    lineLayout->addWidget(new QLabel("Style:"));
    lineLayout->addWidget(m_borderStyleList);

    // Color button
    m_borderColorBtn = new QPushButton();
    m_borderColorBtn->setFixedSize(100, 26);
    m_borderColorBtn->setStyleSheet(QString("QPushButton { background: %1; border: 1px solid #ccc; border-radius: 4px; }").arg(m_borderColorStr));
    connect(m_borderColorBtn, &QPushButton::clicked, this, [this]() {
        QColor color = QColorDialog::getColor(QColor(m_borderColorStr), this, "Border Color");
        if (color.isValid()) {
            m_borderColorStr = color.name();
            m_borderColorBtn->setStyleSheet(QString("QPushButton { background: %1; border: 1px solid #ccc; border-radius: 4px; }").arg(m_borderColorStr));
        }
    });
    lineLayout->addWidget(new QLabel("Color:"));
    lineLayout->addWidget(m_borderColorBtn);
    lineLayout->addStretch();
    mainLayout->addLayout(lineLayout);

    // Right side: Presets + Edge buttons + Preview
    QVBoxLayout* rightLayout = new QVBoxLayout();

    // Presets
    rightLayout->addWidget(new QLabel("Presets"));
    QHBoxLayout* presetLayout = new QHBoxLayout();
    m_borderPresetNone = new QPushButton("None");
    m_borderPresetOutline = new QPushButton("Outline");
    m_borderPresetInside = new QPushButton("Inside");
    for (auto* btn : {m_borderPresetNone, m_borderPresetOutline, m_borderPresetInside}) {
        btn->setFixedSize(70, 30);
        btn->setStyleSheet("QPushButton { border: 1px solid #ccc; border-radius: 4px; font-size: 11px; }"
                          "QPushButton:hover { background: #e8f0fe; }");
        presetLayout->addWidget(btn);
    }
    presetLayout->addStretch();
    rightLayout->addLayout(presetLayout);

    // Preset actions
    connect(m_borderPresetNone, &QPushButton::clicked, this, [this]() {
        m_style.borderTop = m_style.borderBottom = m_style.borderLeft = m_style.borderRight = BorderStyle();
        updateBorderPreview();
    });
    connect(m_borderPresetOutline, &QPushButton::clicked, this, [this]() {
        int penStyle = m_borderStyleList->currentRow();
        BorderStyle bs; bs.enabled = true; bs.width = (penStyle >= 6 && penStyle <= 9) ? 2 : 1;
        bs.color = m_borderColorStr; bs.penStyle = penStyle;
        m_style.borderTop = m_style.borderBottom = m_style.borderLeft = m_style.borderRight = bs;
        updateBorderPreview();
    });
    connect(m_borderPresetInside, &QPushButton::clicked, this, [this]() {
        // Inside borders only (for multi-cell selection — simplified: applies to all edges)
        int penStyle = m_borderStyleList->currentRow();
        BorderStyle bs; bs.enabled = true; bs.width = 1;
        bs.color = m_borderColorStr; bs.penStyle = penStyle;
        m_style.borderTop = m_style.borderBottom = m_style.borderLeft = m_style.borderRight = bs;
        updateBorderPreview();
    });

    rightLayout->addSpacing(10);

    // Edge toggle buttons
    rightLayout->addWidget(new QLabel("Border"));
    QGridLayout* edgeGrid = new QGridLayout();

    auto makeEdgeBtn = [](const QString& label) -> QPushButton* {
        auto* btn = new QPushButton(label);
        btn->setFixedSize(30, 30);
        btn->setCheckable(true);
        btn->setStyleSheet("QPushButton { border: 1px solid #ccc; border-radius: 4px; font-size: 11px; }"
                          "QPushButton:checked { background: #d0e8ff; border-color: #4285f4; }");
        return btn;
    };

    m_borderTop = makeEdgeBtn("T");
    m_borderBottom = makeEdgeBtn("B");
    m_borderLeft = makeEdgeBtn("L");
    m_borderRight = makeEdgeBtn("R");

    edgeGrid->addWidget(m_borderTop, 0, 1);
    edgeGrid->addWidget(m_borderLeft, 1, 0);
    edgeGrid->addWidget(m_borderRight, 1, 2);
    edgeGrid->addWidget(m_borderBottom, 2, 1);

    // Center: preview area
    m_borderPreviewWidget = new QWidget();
    m_borderPreviewWidget->setFixedSize(80, 50);
    m_borderPreviewWidget->setStyleSheet("background: white; border: 1px solid #e0e0e0;");
    edgeGrid->addWidget(m_borderPreviewWidget, 1, 1);

    rightLayout->addLayout(edgeGrid);

    // Connect edge buttons
    auto toggleEdge = [this](QPushButton* btn, BorderStyle& bs) {
        connect(btn, &QPushButton::toggled, this, [this, &bs, btn](bool checked) {
            bs.enabled = checked;
            if (checked) {
                bs.color = m_borderColorStr;
                bs.width = 1;
                bs.penStyle = m_borderStyleList->currentRow();
            }
            updateBorderPreview();
        });
    };
    toggleEdge(m_borderTop, m_style.borderTop);
    toggleEdge(m_borderBottom, m_style.borderBottom);
    toggleEdge(m_borderLeft, m_style.borderLeft);
    toggleEdge(m_borderRight, m_style.borderRight);

    rightLayout->addStretch();
    mainLayout->addLayout(rightLayout);
}

void FormatCellsDialog::updateBorderPreview() {
    // Update toggle button states to match style
    if (m_borderTop) m_borderTop->setChecked(m_style.borderTop.enabled);
    if (m_borderBottom) m_borderBottom->setChecked(m_style.borderBottom.enabled);
    if (m_borderLeft) m_borderLeft->setChecked(m_style.borderLeft.enabled);
    if (m_borderRight) m_borderRight->setChecked(m_style.borderRight.enabled);
}

// ============================================================================
// Protection Tab
// ============================================================================
void FormatCellsDialog::createProtectionTab(QWidget* tab) {
    QVBoxLayout* layout = new QVBoxLayout(tab);

    m_lockedCheck = new QCheckBox("Locked");
    m_lockedCheck->setChecked(true); // Default: locked
    layout->addWidget(m_lockedCheck);

    QLabel* lockDesc = new QLabel(
        "Locking cells or hiding formulas has no effect until you protect\n"
        "the worksheet (Data menu → Protect Sheet).");
    lockDesc->setStyleSheet("color: #666; font-size: 11px;");
    lockDesc->setWordWrap(true);
    layout->addWidget(lockDesc);

    layout->addSpacing(20);

    m_hiddenCheck = new QCheckBox("Hidden");
    m_hiddenCheck->setChecked(false);
    layout->addWidget(m_hiddenCheck);

    QLabel* hideDesc = new QLabel(
        "Hides the formula in the formula bar when the cell is selected.\n"
        "The value is still displayed in the cell.");
    hideDesc->setStyleSheet("color: #666; font-size: 11px;");
    hideDesc->setWordWrap(true);
    layout->addWidget(hideDesc);

    layout->addStretch();
}
