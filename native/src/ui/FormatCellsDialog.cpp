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

    QWidget* fillTab = new QWidget();
    createFillTab(fillTab);
    tabs->addTab(fillTab, "Fill");

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
