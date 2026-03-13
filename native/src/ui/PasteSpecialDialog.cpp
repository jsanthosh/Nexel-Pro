#include "PasteSpecialDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGroupBox>
#include <QDialogButtonBox>

PasteSpecialDialog::PasteSpecialDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Paste Special");
    setFixedSize(380, 320);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(10);
    mainLayout->setContentsMargins(12, 12, 12, 12);

    auto* topLayout = new QHBoxLayout();

    // Paste group
    auto* pasteGroup = new QGroupBox("Paste", this);
    auto* pasteLayout = new QVBoxLayout(pasteGroup);
    pasteLayout->setSpacing(4);
    m_pasteAll = new QRadioButton("All", pasteGroup);
    m_pasteValues = new QRadioButton("Values", pasteGroup);
    m_pasteFormulas = new QRadioButton("Formulas", pasteGroup);
    m_pasteFormats = new QRadioButton("Formats", pasteGroup);
    m_pasteColumnWidths = new QRadioButton("Column Widths", pasteGroup);
    m_pasteAll->setChecked(true);
    pasteLayout->addWidget(m_pasteAll);
    pasteLayout->addWidget(m_pasteValues);
    pasteLayout->addWidget(m_pasteFormulas);
    pasteLayout->addWidget(m_pasteFormats);
    pasteLayout->addWidget(m_pasteColumnWidths);
    topLayout->addWidget(pasteGroup);

    // Operation group
    auto* opGroup = new QGroupBox("Operation", this);
    auto* opLayout = new QVBoxLayout(opGroup);
    opLayout->setSpacing(4);
    m_opNone = new QRadioButton("None", opGroup);
    m_opAdd = new QRadioButton("Add", opGroup);
    m_opSubtract = new QRadioButton("Subtract", opGroup);
    m_opMultiply = new QRadioButton("Multiply", opGroup);
    m_opDivide = new QRadioButton("Divide", opGroup);
    m_opNone->setChecked(true);
    opLayout->addWidget(m_opNone);
    opLayout->addWidget(m_opAdd);
    opLayout->addWidget(m_opSubtract);
    opLayout->addWidget(m_opMultiply);
    opLayout->addWidget(m_opDivide);
    topLayout->addWidget(opGroup);

    mainLayout->addLayout(topLayout);

    // Options
    auto* optionsGroup = new QGroupBox("Options", this);
    auto* optionsLayout = new QHBoxLayout(optionsGroup);
    m_skipBlanks = new QCheckBox("Skip Blanks", optionsGroup);
    m_transpose = new QCheckBox("Transpose", optionsGroup);
    optionsLayout->addWidget(m_skipBlanks);
    optionsLayout->addWidget(m_transpose);
    mainLayout->addWidget(optionsGroup);

    // Buttons
    auto* buttonBox = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    mainLayout->addWidget(buttonBox);

    connect(buttonBox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttonBox, &QDialogButtonBox::rejected, this, &QDialog::reject);

    setStyleSheet(ThemeManager::dialogStylesheet());
}

PasteSpecialOptions PasteSpecialDialog::getOptions() const {
    PasteSpecialOptions opts;

    if (m_pasteAll->isChecked()) opts.pasteType = PasteSpecialOptions::All;
    else if (m_pasteValues->isChecked()) opts.pasteType = PasteSpecialOptions::Values;
    else if (m_pasteFormulas->isChecked()) opts.pasteType = PasteSpecialOptions::Formulas;
    else if (m_pasteFormats->isChecked()) opts.pasteType = PasteSpecialOptions::Formats;
    else if (m_pasteColumnWidths->isChecked()) opts.pasteType = PasteSpecialOptions::ColumnWidths;

    if (m_opNone->isChecked()) opts.operation = PasteSpecialOptions::OpNone;
    else if (m_opAdd->isChecked()) opts.operation = PasteSpecialOptions::Add;
    else if (m_opSubtract->isChecked()) opts.operation = PasteSpecialOptions::Subtract;
    else if (m_opMultiply->isChecked()) opts.operation = PasteSpecialOptions::Multiply;
    else if (m_opDivide->isChecked()) opts.operation = PasteSpecialOptions::Divide;

    opts.skipBlanks = m_skipBlanks->isChecked();
    opts.transpose = m_transpose->isChecked();

    return opts;
}
