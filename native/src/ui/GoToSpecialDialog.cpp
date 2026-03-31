#include "GoToSpecialDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGroupBox>
#include <QLabel>
#include <QPushButton>

GoToSpecialDialog::GoToSpecialDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Go To Special");
    setFixedSize(380, 420);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);
    mainLayout->setContentsMargins(16, 16, 16, 16);

    auto* titleLabel = new QLabel("Select:", this);
    titleLabel->setStyleSheet("font-weight: bold; font-size: 13px;");
    mainLayout->addWidget(titleLabel);

    m_typeGroup = new QButtonGroup(this);

    // Radio buttons
    m_blanksRadio = new QRadioButton("&Blanks", this);
    m_constantsRadio = new QRadioButton("&Constants", this);
    m_formulasRadio = new QRadioButton("&Formulas", this);
    m_commentsRadio = new QRadioButton("Co&mments", this);
    m_conditionalFormatsRadio = new QRadioButton("Conditional forma&ts", this);
    m_dataValidationRadio = new QRadioButton("Data &validation", this);
    m_visibleCellsRadio = new QRadioButton("&Visible cells only", this);
    m_currentRegionRadio = new QRadioButton("Current &region", this);

    m_typeGroup->addButton(m_blanksRadio, static_cast<int>(Blanks));
    m_typeGroup->addButton(m_constantsRadio, static_cast<int>(Constants));
    m_typeGroup->addButton(m_formulasRadio, static_cast<int>(Formulas));
    m_typeGroup->addButton(m_commentsRadio, static_cast<int>(Comments));
    m_typeGroup->addButton(m_conditionalFormatsRadio, static_cast<int>(ConditionalFormats));
    m_typeGroup->addButton(m_dataValidationRadio, static_cast<int>(DataValidation));
    m_typeGroup->addButton(m_visibleCellsRadio, static_cast<int>(VisibleCells));
    m_typeGroup->addButton(m_currentRegionRadio, static_cast<int>(CurrentRegion));

    m_blanksRadio->setChecked(true);

    mainLayout->addWidget(m_blanksRadio);

    // Constants with sub-checkboxes
    mainLayout->addWidget(m_constantsRadio);

    // Sub-checkboxes for Constants/Formulas filtering
    auto* subCheckLayout = new QHBoxLayout();
    subCheckLayout->setContentsMargins(24, 0, 0, 0);
    subCheckLayout->setSpacing(12);

    m_numbersCheck = new QCheckBox("Numbers", this);
    m_textCheck = new QCheckBox("Text", this);
    m_logicalsCheck = new QCheckBox("Logicals", this);
    m_errorsCheck = new QCheckBox("Errors", this);

    m_numbersCheck->setChecked(true);
    m_textCheck->setChecked(true);
    m_logicalsCheck->setChecked(true);
    m_errorsCheck->setChecked(true);

    subCheckLayout->addWidget(m_numbersCheck);
    subCheckLayout->addWidget(m_textCheck);
    subCheckLayout->addWidget(m_logicalsCheck);
    subCheckLayout->addWidget(m_errorsCheck);
    subCheckLayout->addStretch();

    mainLayout->addLayout(subCheckLayout);

    mainLayout->addWidget(m_formulasRadio);

    // Note: same sub-checkboxes apply to both Constants and Formulas
    auto* noteLabel = new QLabel("(Above checkboxes apply to Constants and Formulas)", this);
    noteLabel->setStyleSheet("color: #667085; font-size: 10px; margin-left: 24px;");
    mainLayout->addWidget(noteLabel);

    mainLayout->addWidget(m_commentsRadio);
    mainLayout->addWidget(m_conditionalFormatsRadio);
    mainLayout->addWidget(m_dataValidationRadio);
    mainLayout->addWidget(m_visibleCellsRadio);
    mainLayout->addWidget(m_currentRegionRadio);

    mainLayout->addStretch();

    // Buttons
    auto* btnLayout = new QHBoxLayout();
    btnLayout->addStretch();
    auto* okBtn = new QPushButton("OK", this);
    okBtn->setDefault(true);
    okBtn->setFixedWidth(80);
    auto* cancelBtn = new QPushButton("Cancel", this);
    cancelBtn->setFixedWidth(80);
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    mainLayout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);

    // Enable/disable sub-checkboxes based on selection
    connect(m_typeGroup, &QButtonGroup::idToggled, this, [this]() {
        updateSubCheckboxes();
    });

    updateSubCheckboxes();

    setStyleSheet(ThemeManager::dialogStylesheet());
}

GoToSpecialDialog::SelectionType GoToSpecialDialog::getSelectionType() const {
    int id = m_typeGroup->checkedId();
    if (id >= 0) return static_cast<SelectionType>(id);
    return Blanks;
}

bool GoToSpecialDialog::includeNumbers() const {
    return m_numbersCheck->isChecked();
}

bool GoToSpecialDialog::includeText() const {
    return m_textCheck->isChecked();
}

bool GoToSpecialDialog::includeLogicals() const {
    return m_logicalsCheck->isChecked();
}

bool GoToSpecialDialog::includeErrors() const {
    return m_errorsCheck->isChecked();
}

void GoToSpecialDialog::updateSubCheckboxes() {
    bool enabled = m_constantsRadio->isChecked() || m_formulasRadio->isChecked();
    m_numbersCheck->setEnabled(enabled);
    m_textCheck->setEnabled(enabled);
    m_logicalsCheck->setEnabled(enabled);
    m_errorsCheck->setEnabled(enabled);
}
