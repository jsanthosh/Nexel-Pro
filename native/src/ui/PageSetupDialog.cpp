#include "PageSetupDialog.h"
#include "Theme.h"

#include <QTabWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFormLayout>
#include <QGroupBox>
#include <QLabel>
#include <QDoubleSpinBox>
#include <QSpinBox>
#include <QComboBox>
#include <QLineEdit>
#include <QCheckBox>
#include <QRadioButton>
#include <QButtonGroup>
#include <QPushButton>
#include <QDialogButtonBox>

namespace {

// OOXML paper size codes that real Excel uses. Numbers match
// w:pgSz / pageSetup paperSize attribute. We surface the popular
// ones; the long tail is supported on round-trip but hidden here.
struct PaperEntry { int code; const char* label; };
constexpr PaperEntry kPapers[] = {
    {1,  "Letter (8.5 x 11 in)"},
    {3,  "Tabloid (11 x 17 in)"},
    {5,  "Legal (8.5 x 14 in)"},
    {8,  "A3 (297 x 420 mm)"},
    {9,  "A4 (210 x 297 mm)"},
    {11, "A5 (148 x 210 mm)"},
    {70, "Executive (7.25 x 10.5 in)"},
};

QDoubleSpinBox* makeMarginBox(double initial) {
    auto* sp = new QDoubleSpinBox();
    sp->setRange(0.0, 10.0);
    sp->setSingleStep(0.05);
    sp->setDecimals(2);
    sp->setSuffix(" in");
    sp->setValue(initial);
    return sp;
}

} // namespace

PageSetupDialog::PageSetupDialog(const Spreadsheet::PrintSettings& init,
                                   QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Page Setup");
    setModal(true);
    resize(540, 460);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(16, 16, 16, 16);
    mainLayout->setSpacing(10);

    auto* tabs = new QTabWidget(this);

    // ---- Page tab ----
    auto* pageTab = new QWidget;
    {
        auto* layout = new QVBoxLayout(pageTab);

        auto* orientGroup = new QGroupBox("Orientation", pageTab);
        auto* orientLayout = new QHBoxLayout(orientGroup);
        m_portraitRadio  = new QRadioButton("&Portrait");
        m_landscapeRadio = new QRadioButton("&Landscape");
        auto* orientBtnGroup = new QButtonGroup(this);
        orientBtnGroup->addButton(m_portraitRadio);
        orientBtnGroup->addButton(m_landscapeRadio);
        if (init.orientation == 2) m_landscapeRadio->setChecked(true);
        else                       m_portraitRadio->setChecked(true);
        orientLayout->addWidget(m_portraitRadio);
        orientLayout->addWidget(m_landscapeRadio);
        orientLayout->addStretch();
        layout->addWidget(orientGroup);

        auto* scaleGroup = new QGroupBox("Scaling", pageTab);
        auto* scaleLayout = new QVBoxLayout(scaleGroup);
        m_scaleRadio = new QRadioButton("&Adjust to");
        auto* scaleRow = new QHBoxLayout;
        m_scaleSpin = new QSpinBox; m_scaleSpin->setRange(10, 400);
        m_scaleSpin->setSuffix(" %"); m_scaleSpin->setValue(init.scale > 0 ? init.scale : 100);
        scaleRow->addWidget(m_scaleRadio);
        scaleRow->addWidget(m_scaleSpin);
        scaleRow->addWidget(new QLabel("normal size"));
        scaleRow->addStretch();
        scaleLayout->addLayout(scaleRow);

        m_fitRadio = new QRadioButton("&Fit to");
        auto* fitRow = new QHBoxLayout;
        m_fitWidthSpin  = new QSpinBox; m_fitWidthSpin->setRange(1, 99);
        m_fitWidthSpin->setValue(init.fitToWidth  > 0 ? init.fitToWidth  : 1);
        m_fitHeightSpin = new QSpinBox; m_fitHeightSpin->setRange(1, 99);
        m_fitHeightSpin->setValue(init.fitToHeight > 0 ? init.fitToHeight : 1);
        fitRow->addWidget(m_fitRadio);
        fitRow->addWidget(m_fitWidthSpin);
        fitRow->addWidget(new QLabel("page(s) wide by"));
        fitRow->addWidget(m_fitHeightSpin);
        fitRow->addWidget(new QLabel("tall"));
        fitRow->addStretch();
        scaleLayout->addLayout(fitRow);

        auto* scaleBtnGroup = new QButtonGroup(this);
        scaleBtnGroup->addButton(m_scaleRadio);
        scaleBtnGroup->addButton(m_fitRadio);
        if (init.fitToWidth > 0 || init.fitToHeight > 0) m_fitRadio->setChecked(true);
        else                                              m_scaleRadio->setChecked(true);
        layout->addWidget(scaleGroup);

        auto* sizeGroup = new QGroupBox("Paper", pageTab);
        auto* sizeLayout = new QFormLayout(sizeGroup);
        m_paperSizeCombo = new QComboBox;
        for (const auto& p : kPapers) {
            m_paperSizeCombo->addItem(p.label, p.code);
        }
        for (int i = 0; i < m_paperSizeCombo->count(); ++i) {
            if (m_paperSizeCombo->itemData(i).toInt() == init.paperSize) {
                m_paperSizeCombo->setCurrentIndex(i); break;
            }
        }
        sizeLayout->addRow("Paper size:", m_paperSizeCombo);
        layout->addWidget(sizeGroup);
        layout->addStretch();
    }
    tabs->addTab(pageTab, "&Page");

    // ---- Margins tab ----
    auto* marginsTab = new QWidget;
    {
        auto* g = new QGroupBox("Margins (inches)", marginsTab);
        auto* form = new QFormLayout(g);
        m_topMargin    = makeMarginBox(init.topMargin);
        m_bottomMargin = makeMarginBox(init.bottomMargin);
        m_leftMargin   = makeMarginBox(init.leftMargin);
        m_rightMargin  = makeMarginBox(init.rightMargin);
        m_headerMargin = makeMarginBox(init.headerMargin);
        m_footerMargin = makeMarginBox(init.footerMargin);
        form->addRow("&Top:",    m_topMargin);
        form->addRow("&Bottom:", m_bottomMargin);
        form->addRow("&Left:",   m_leftMargin);
        form->addRow("&Right:",  m_rightMargin);
        form->addRow("&Header:", m_headerMargin);
        form->addRow("&Footer:", m_footerMargin);
        auto* outer = new QVBoxLayout(marginsTab);
        outer->addWidget(g);
        outer->addStretch();
    }
    tabs->addTab(marginsTab, "&Margins");

    // ---- Header / Footer tab ----
    auto* hfTab = new QWidget;
    {
        auto* form = new QFormLayout(hfTab);
        auto* note = new QLabel(
            "Tokens: <code>&amp;P</code> page#, "
            "<code>&amp;N</code> total pages, <code>&amp;D</code> date, "
            "<code>&amp;F</code> filename, <code>&amp;A</code> sheet name, "
            "<code>&amp;L</code>/<code>&amp;C</code>/<code>&amp;R</code> "
            "left/centre/right alignment.");
        note->setTextFormat(Qt::RichText);
        note->setWordWrap(true);
        note->setStyleSheet("color:#667085; font-size: 11px;");
        form->addRow(note);
        m_oddHeaderEdit = new QLineEdit;
        m_oddHeaderEdit->setText(init.oddHeader);
        m_oddHeaderEdit->setPlaceholderText("e.g. &CMonthly Report");
        m_oddFooterEdit = new QLineEdit;
        m_oddFooterEdit->setText(init.oddFooter);
        m_oddFooterEdit->setPlaceholderText("e.g. &LPage &P of &N");
        form->addRow("&Header:", m_oddHeaderEdit);
        form->addRow("F&ooter:", m_oddFooterEdit);
    }
    tabs->addTab(hfTab, "&Header/Footer");

    // ---- Sheet tab ----
    auto* sheetTab = new QWidget;
    {
        auto* g = new QGroupBox("Print", sheetTab);
        auto* layout = new QVBoxLayout(g);
        m_printGridlinesCheck = new QCheckBox("Print &gridlines");
        m_printHeadingsCheck  = new QCheckBox("Print row and column &headings");
        m_printGridlinesCheck->setChecked(init.printGridlines);
        m_printHeadingsCheck->setChecked(init.printHeadings);
        layout->addWidget(m_printGridlinesCheck);
        layout->addWidget(m_printHeadingsCheck);
        auto* outer = new QVBoxLayout(sheetTab);
        outer->addWidget(g);
        outer->addStretch();
    }
    tabs->addTab(sheetTab, "&Sheet");

    mainLayout->addWidget(tabs, 1);

    auto* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);

    setStyleSheet(ThemeManager::instance().dialogStylesheet());
}

Spreadsheet::PrintSettings PageSetupDialog::settings() const {
    Spreadsheet::PrintSettings s;
    s.orientation   = m_landscapeRadio->isChecked() ? 2 : 1;
    s.paperSize     = m_paperSizeCombo->currentData().toInt();
    if (m_fitRadio->isChecked()) {
        s.fitToWidth  = m_fitWidthSpin->value();
        s.fitToHeight = m_fitHeightSpin->value();
        s.scale       = 100; // fit overrides scale
    } else {
        s.scale       = m_scaleSpin->value();
        s.fitToWidth  = 0;
        s.fitToHeight = 0;
    }
    s.topMargin    = m_topMargin->value();
    s.bottomMargin = m_bottomMargin->value();
    s.leftMargin   = m_leftMargin->value();
    s.rightMargin  = m_rightMargin->value();
    s.headerMargin = m_headerMargin->value();
    s.footerMargin = m_footerMargin->value();
    s.oddHeader     = m_oddHeaderEdit->text();
    s.oddFooter     = m_oddFooterEdit->text();
    s.printGridlines = m_printGridlinesCheck->isChecked();
    s.printHeadings  = m_printHeadingsCheck->isChecked();
    return s;
}
