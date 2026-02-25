#include "DataValidationDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFormLayout>
#include <QDialogButtonBox>
#include <QGroupBox>

DataValidationDialog::DataValidationDialog(const CellRange& range, QWidget* parent)
    : QDialog(parent), m_range(range) {
    setWindowTitle("Data Validation");
    setMinimumSize(440, 420);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    // Range label
    QLabel* rangeLabel = new QLabel(QString("Applies to: %1").arg(range.toString()), this);
    rangeLabel->setStyleSheet("font-weight: bold; padding: 4px;");
    mainLayout->addWidget(rangeLabel);

    // Tab widget
    QTabWidget* tabs = new QTabWidget(this);

    // === Settings Tab ===
    QWidget* settingsTab = new QWidget();
    QFormLayout* settingsLayout = new QFormLayout(settingsTab);

    m_typeCombo = new QComboBox(this);
    m_typeCombo->addItem("Whole Number", static_cast<int>(Spreadsheet::DataValidationRule::WholeNumber));
    m_typeCombo->addItem("Decimal", static_cast<int>(Spreadsheet::DataValidationRule::Decimal));
    m_typeCombo->addItem("List", static_cast<int>(Spreadsheet::DataValidationRule::List));
    m_typeCombo->addItem("Text Length", static_cast<int>(Spreadsheet::DataValidationRule::TextLength));
    m_typeCombo->addItem("Custom", static_cast<int>(Spreadsheet::DataValidationRule::Custom));
    settingsLayout->addRow("Allow:", m_typeCombo);

    m_operatorCombo = new QComboBox(this);
    m_operatorCombo->addItem("between", static_cast<int>(Spreadsheet::DataValidationRule::Between));
    m_operatorCombo->addItem("not between", static_cast<int>(Spreadsheet::DataValidationRule::NotBetween));
    m_operatorCombo->addItem("equal to", static_cast<int>(Spreadsheet::DataValidationRule::EqualTo));
    m_operatorCombo->addItem("not equal to", static_cast<int>(Spreadsheet::DataValidationRule::NotEqualTo));
    m_operatorCombo->addItem("greater than", static_cast<int>(Spreadsheet::DataValidationRule::GreaterThan));
    m_operatorCombo->addItem("less than", static_cast<int>(Spreadsheet::DataValidationRule::LessThan));
    m_operatorCombo->addItem("greater than or equal to", static_cast<int>(Spreadsheet::DataValidationRule::GreaterThanOrEqual));
    m_operatorCombo->addItem("less than or equal to", static_cast<int>(Spreadsheet::DataValidationRule::LessThanOrEqual));
    settingsLayout->addRow("Data:", m_operatorCombo);

    m_value1Label = new QLabel("Minimum:", this);
    m_value1Edit = new QLineEdit(this);
    settingsLayout->addRow(m_value1Label, m_value1Edit);

    m_value2Label = new QLabel("Maximum:", this);
    m_value2Edit = new QLineEdit(this);
    settingsLayout->addRow(m_value2Label, m_value2Edit);

    m_listLabel = new QLabel("Source (comma-separated):", this);
    m_listEdit = new QTextEdit(this);
    m_listEdit->setMaximumHeight(60);
    settingsLayout->addRow(m_listLabel, m_listEdit);

    m_formulaLabel = new QLabel("Formula:", this);
    m_formulaEdit = new QLineEdit(this);
    m_formulaEdit->setPlaceholderText("e.g. =A1>0");
    settingsLayout->addRow(m_formulaLabel, m_formulaEdit);

    m_ignoreBlank = new QCheckBox("Ignore blank cells", this);
    m_ignoreBlank->setChecked(true);
    settingsLayout->addRow(m_ignoreBlank);

    tabs->addTab(settingsTab, "Settings");

    // === Input Message Tab ===
    QWidget* inputTab = new QWidget();
    QFormLayout* inputLayout = new QFormLayout(inputTab);

    m_showInputMsg = new QCheckBox("Show input message when cell is selected", this);
    m_showInputMsg->setChecked(true);
    inputLayout->addRow(m_showInputMsg);

    m_inputTitle = new QLineEdit(this);
    inputLayout->addRow("Title:", m_inputTitle);

    m_inputMessage = new QTextEdit(this);
    m_inputMessage->setMaximumHeight(80);
    inputLayout->addRow("Input message:", m_inputMessage);

    tabs->addTab(inputTab, "Input Message");

    // === Error Alert Tab ===
    QWidget* errorTab = new QWidget();
    QFormLayout* errorLayout = new QFormLayout(errorTab);

    m_showErrorAlert = new QCheckBox("Show error alert after invalid data is entered", this);
    m_showErrorAlert->setChecked(true);
    errorLayout->addRow(m_showErrorAlert);

    m_errorStyle = new QComboBox(this);
    m_errorStyle->addItem("Stop", static_cast<int>(Spreadsheet::DataValidationRule::Stop));
    m_errorStyle->addItem("Warning", static_cast<int>(Spreadsheet::DataValidationRule::Warning));
    m_errorStyle->addItem("Information", static_cast<int>(Spreadsheet::DataValidationRule::Information));
    errorLayout->addRow("Style:", m_errorStyle);

    m_errorTitle = new QLineEdit(this);
    errorLayout->addRow("Title:", m_errorTitle);

    m_errorMessage = new QTextEdit(this);
    m_errorMessage->setMaximumHeight(80);
    errorLayout->addRow("Error message:", m_errorMessage);

    tabs->addTab(errorTab, "Error Alert");

    mainLayout->addWidget(tabs);

    // Buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    mainLayout->addWidget(buttons);

    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    connect(m_typeCombo, &QComboBox::currentIndexChanged, this, &DataValidationDialog::onTypeChanged);
    connect(m_operatorCombo, &QComboBox::currentIndexChanged, this, &DataValidationDialog::onOperatorChanged);

    updateFieldVisibility();

    setStyleSheet(ThemeManager::dialogStylesheet());
}

void DataValidationDialog::onTypeChanged(int) {
    updateFieldVisibility();
}

void DataValidationDialog::onOperatorChanged(int) {
    updateFieldVisibility();
}

void DataValidationDialog::updateFieldVisibility() {
    int typeData = m_typeCombo->currentData().toInt();
    auto type = static_cast<Spreadsheet::DataValidationRule::Type>(typeData);

    bool isList = (type == Spreadsheet::DataValidationRule::List);
    bool isCustom = (type == Spreadsheet::DataValidationRule::Custom);
    bool isNumeric = !isList && !isCustom;

    m_operatorCombo->setVisible(isNumeric);
    m_value1Label->setVisible(isNumeric);
    m_value1Edit->setVisible(isNumeric);

    int opData = m_operatorCombo->currentData().toInt();
    auto op = static_cast<Spreadsheet::DataValidationRule::Operator>(opData);
    bool showValue2 = isNumeric && (op == Spreadsheet::DataValidationRule::Between ||
                                     op == Spreadsheet::DataValidationRule::NotBetween);
    m_value2Label->setVisible(showValue2);
    m_value2Edit->setVisible(showValue2);

    m_listLabel->setVisible(isList);
    m_listEdit->setVisible(isList);

    m_formulaLabel->setVisible(isCustom);
    m_formulaEdit->setVisible(isCustom);

    // Update value labels based on operator
    if (showValue2) {
        m_value1Label->setText("Minimum:");
        m_value2Label->setText("Maximum:");
    } else {
        m_value1Label->setText("Value:");
    }
}

Spreadsheet::DataValidationRule DataValidationDialog::getRule() const {
    Spreadsheet::DataValidationRule rule;
    rule.range = m_range;
    rule.type = static_cast<Spreadsheet::DataValidationRule::Type>(m_typeCombo->currentData().toInt());
    rule.op = static_cast<Spreadsheet::DataValidationRule::Operator>(m_operatorCombo->currentData().toInt());
    rule.value1 = m_value1Edit->text();
    rule.value2 = m_value2Edit->text();
    rule.customFormula = m_formulaEdit->text();

    // Parse list items
    QString listText = m_listEdit->toPlainText();
    if (!listText.isEmpty()) {
        rule.listItems = listText.split(",", Qt::SkipEmptyParts);
        for (auto& item : rule.listItems) {
            item = item.trimmed();
        }
    }

    rule.showInputMessage = m_showInputMsg->isChecked();
    rule.inputTitle = m_inputTitle->text();
    rule.inputMessage = m_inputMessage->toPlainText();

    rule.showErrorAlert = m_showErrorAlert->isChecked();
    rule.errorStyle = static_cast<Spreadsheet::DataValidationRule::ErrorStyle>(m_errorStyle->currentData().toInt());
    rule.errorTitle = m_errorTitle->text();
    rule.errorMessage = m_errorMessage->toPlainText();

    return rule;
}

void DataValidationDialog::setRule(const Spreadsheet::DataValidationRule& rule) {
    int typeIdx = m_typeCombo->findData(static_cast<int>(rule.type));
    if (typeIdx >= 0) m_typeCombo->setCurrentIndex(typeIdx);

    int opIdx = m_operatorCombo->findData(static_cast<int>(rule.op));
    if (opIdx >= 0) m_operatorCombo->setCurrentIndex(opIdx);

    m_value1Edit->setText(rule.value1);
    m_value2Edit->setText(rule.value2);
    m_formulaEdit->setText(rule.customFormula);
    m_listEdit->setPlainText(rule.listItems.join(", "));

    m_showInputMsg->setChecked(rule.showInputMessage);
    m_inputTitle->setText(rule.inputTitle);
    m_inputMessage->setPlainText(rule.inputMessage);

    m_showErrorAlert->setChecked(rule.showErrorAlert);
    int errIdx = m_errorStyle->findData(static_cast<int>(rule.errorStyle));
    if (errIdx >= 0) m_errorStyle->setCurrentIndex(errIdx);
    m_errorTitle->setText(rule.errorTitle);
    m_errorMessage->setPlainText(rule.errorMessage);

    updateFieldVisibility();
}
