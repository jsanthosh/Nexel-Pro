#include "ConditionalFormatDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGroupBox>
#include <QFormLayout>
#include <QDialogButtonBox>
#include <QMessageBox>

ConditionalFormatDialog::ConditionalFormatDialog(const CellRange& defaultRange,
                                                   ConditionalFormatting& formatting,
                                                   QWidget* parent)
    : QDialog(parent), m_defaultRange(defaultRange), m_formatting(formatting) {
    setWindowTitle("Conditional Formatting Rules Manager");
    setMinimumSize(560, 520);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    // ---- All Rules section (shows every rule on the sheet) ----
    QGroupBox* rulesGroup = new QGroupBox("All Rules (Sheet)", this);
    QVBoxLayout* rulesLayout = new QVBoxLayout(rulesGroup);

    m_ruleList = new QListWidget(this);
    m_ruleList->setMinimumHeight(140);
    rulesLayout->addWidget(m_ruleList);

    QHBoxLayout* ruleButtons = new QHBoxLayout();
    QPushButton* deleteBtn = new QPushButton("Delete Rule", this);
    QPushButton* clearAllBtn = new QPushButton("Clear All", this);
    ruleButtons->addWidget(deleteBtn);
    ruleButtons->addWidget(clearAllBtn);
    ruleButtons->addStretch();
    rulesLayout->addLayout(ruleButtons);

    mainLayout->addWidget(rulesGroup);

    // ---- New Rule section ----
    QGroupBox* newRuleGroup = new QGroupBox("Add New Rule", this);
    QFormLayout* newRuleLayout = new QFormLayout(newRuleGroup);

    m_rangeEdit = new QLineEdit(defaultRange.toString(), this);
    m_rangeEdit->setPlaceholderText("e.g. A1:D100");
    newRuleLayout->addRow("Applies to:", m_rangeEdit);

    m_conditionType = new QComboBox(this);
    m_conditionType->addItem("Cell Value Equal To", static_cast<int>(ConditionType::Equal));
    m_conditionType->addItem("Cell Value Not Equal To", static_cast<int>(ConditionType::NotEqual));
    m_conditionType->addItem("Cell Value Greater Than", static_cast<int>(ConditionType::GreaterThan));
    m_conditionType->addItem("Cell Value Less Than", static_cast<int>(ConditionType::LessThan));
    m_conditionType->addItem("Cell Value Greater Than or Equal", static_cast<int>(ConditionType::GreaterThanOrEqual));
    m_conditionType->addItem("Cell Value Less Than or Equal", static_cast<int>(ConditionType::LessThanOrEqual));
    m_conditionType->addItem("Cell Value Between", static_cast<int>(ConditionType::Between));
    m_conditionType->addItem("Cell Contains", static_cast<int>(ConditionType::CellContains));
    m_conditionType->addItem("Use a Formula", static_cast<int>(ConditionType::Formula));
    newRuleLayout->addRow("Format cells if:", m_conditionType);

    m_value1Label = new QLabel("Value:", this);
    m_value1Edit = new QLineEdit(this);
    newRuleLayout->addRow(m_value1Label, m_value1Edit);

    m_value2Label = new QLabel("And:", this);
    m_value2Edit = new QLineEdit(this);
    newRuleLayout->addRow(m_value2Label, m_value2Edit);

    m_formulaLabel = new QLabel("Formula:", this);
    m_formulaEdit = new QLineEdit(this);
    m_formulaEdit->setPlaceholderText("e.g. =A1>100");
    newRuleLayout->addRow(m_formulaLabel, m_formulaEdit);

    // Format style row
    QHBoxLayout* styleLayout = new QHBoxLayout();
    m_boldCheck = new QCheckBox("Bold", this);
    m_italicCheck = new QCheckBox("Italic", this);
    m_underlineCheck = new QCheckBox("Underline", this);
    styleLayout->addWidget(m_boldCheck);
    styleLayout->addWidget(m_italicCheck);
    styleLayout->addWidget(m_underlineCheck);

    m_fgColorBtn = new QPushButton("Font Color", this);
    m_fgColorBtn->setStyleSheet("QPushButton { border-bottom: 3px solid #000000; }");
    m_fgColorBtn->setFixedWidth(90);
    styleLayout->addWidget(m_fgColorBtn);

    m_bgColorBtn = new QPushButton("Fill Color", this);
    m_bgColorBtn->setStyleSheet(QString("QPushButton { background-color: #FFFFFF; border: 1px solid %1; }")
        .arg(ThemeManager::instance().currentTheme().dialogGroupBoxBorder.name()));
    m_bgColorBtn->setFixedWidth(90);
    styleLayout->addWidget(m_bgColorBtn);

    newRuleLayout->addRow("Style:", styleLayout);

    QHBoxLayout* actionBtns = new QHBoxLayout();
    QPushButton* addBtn = new QPushButton("Add Rule", this);
    addBtn->setProperty("primary", true);
    actionBtns->addWidget(addBtn);

    m_updateBtn = new QPushButton("Update Selected", this);
    m_updateBtn->setProperty("primary", true);
    m_updateBtn->setEnabled(false);
    actionBtns->addWidget(m_updateBtn);
    actionBtns->addStretch();

    newRuleLayout->addRow("", actionBtns);

    mainLayout->addWidget(newRuleGroup);

    // Dialog buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Close, this);
    mainLayout->addWidget(buttons);

    // Connections
    connect(addBtn, &QPushButton::clicked, this, &ConditionalFormatDialog::onAddRule);
    connect(m_updateBtn, &QPushButton::clicked, this, &ConditionalFormatDialog::onUpdateRule);
    connect(deleteBtn, &QPushButton::clicked, this, &ConditionalFormatDialog::onDeleteRule);
    connect(clearAllBtn, &QPushButton::clicked, this, [this]() {
        if (m_formatting.getAllRules().empty()) return;
        if (QMessageBox::question(this, "Clear All Rules",
                "Remove all conditional formatting rules from this sheet?",
                QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
            m_formatting.clearRules();
            populateRuleList();
        }
    });
    connect(m_ruleList, &QListWidget::currentRowChanged, this, &ConditionalFormatDialog::onRuleSelected);
    connect(m_conditionType, &QComboBox::currentIndexChanged, this, &ConditionalFormatDialog::onConditionTypeChanged);

    connect(m_fgColorBtn, &QPushButton::clicked, this, [this]() {
        QColor color = QColorDialog::getColor(m_selectedFgColor, this, "Font Color");
        if (color.isValid()) {
            m_selectedFgColor = color;
            m_fgColorBtn->setStyleSheet(
                QString("QPushButton { border-bottom: 3px solid %1; }").arg(color.name()));
        }
    });

    connect(m_bgColorBtn, &QPushButton::clicked, this, [this]() {
        QColor color = QColorDialog::getColor(m_selectedBgColor, this, "Fill Color");
        if (color.isValid()) {
            m_selectedBgColor = color;
            m_bgColorBtn->setStyleSheet(
                QString("QPushButton { background-color: %1; border: 1px solid %2; }")
                    .arg(color.name(), ThemeManager::instance().currentTheme().dialogGroupBoxBorder.name()));
        }
    });

    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::accept);

    // Initialize
    populateRuleList();
    updateValueFieldsVisibility();

    setStyleSheet(ThemeManager::dialogStylesheet());
}

void ConditionalFormatDialog::populateRuleList() {
    m_ruleList->clear();
    const auto& rules = m_formatting.getAllRules();
    for (const auto& rule : rules) {
        QString desc;
        switch (rule->getType()) {
            case ConditionType::Equal: desc = "Equal to"; break;
            case ConditionType::NotEqual: desc = "Not equal to"; break;
            case ConditionType::GreaterThan: desc = "Greater than"; break;
            case ConditionType::LessThan: desc = "Less than"; break;
            case ConditionType::GreaterThanOrEqual: desc = "Greater than or equal"; break;
            case ConditionType::LessThanOrEqual: desc = "Less than or equal"; break;
            case ConditionType::Between: desc = "Between"; break;
            case ConditionType::CellContains: desc = "Contains"; break;
            case ConditionType::Formula: desc = "Formula"; break;
        }
        const CellStyle& s = rule->getStyle();
        QString styleDesc;
        if (s.bold) styleDesc += "B ";
        if (s.italic) styleDesc += "I ";
        if (!s.backgroundColor.isEmpty() && s.backgroundColor != "#FFFFFF" && s.backgroundColor != "#ffffff")
            styleDesc += "Fill ";
        if (!s.foregroundColor.isEmpty() && s.foregroundColor != "#000000")
            styleDesc += "Color ";
        if (styleDesc.isEmpty()) styleDesc = "Style";

        auto* item = new QListWidgetItem(
            QString("%1  |  %2  [%3]")
                .arg(rule->getRange().toString(), desc, styleDesc.trimmed()));

        // Show a color swatch for the background color
        if (!s.backgroundColor.isEmpty() && s.backgroundColor != "#FFFFFF" && s.backgroundColor != "#ffffff") {
            QPixmap px(14, 14);
            px.fill(QColor(s.backgroundColor));
            item->setIcon(QIcon(px));
        }

        m_ruleList->addItem(item);
    }
}

void ConditionalFormatDialog::onAddRule() {
    auto rule = buildRuleFromUI();
    if (rule) {
        m_formatting.addRule(rule);
        populateRuleList();
    }
}

void ConditionalFormatDialog::onDeleteRule() {
    int row = m_ruleList->currentRow();
    if (row < 0) return;

    m_formatting.removeRule(static_cast<size_t>(row));
    populateRuleList();
}

void ConditionalFormatDialog::onRuleSelected(int row) {
    m_updateBtn->setEnabled(row >= 0);
    if (row < 0) return;

    const auto& rules = m_formatting.getAllRules();
    if (row >= static_cast<int>(rules.size())) return;

    const auto& rule = rules[row];

    // Populate UI from rule
    m_rangeEdit->setText(rule->getRange().toString());
    int typeIdx = m_conditionType->findData(static_cast<int>(rule->getType()));
    if (typeIdx >= 0) m_conditionType->setCurrentIndex(typeIdx);

    const CellStyle& s = rule->getStyle();
    m_boldCheck->setChecked(s.bold);
    m_italicCheck->setChecked(s.italic);
    m_underlineCheck->setChecked(s.underline);
    m_selectedBgColor = QColor(s.backgroundColor.isEmpty() ? "#FFFFFF" : s.backgroundColor);
    m_selectedFgColor = QColor(s.foregroundColor.isEmpty() ? "#000000" : s.foregroundColor);
    m_bgColorBtn->setStyleSheet(
        QString("QPushButton { background-color: %1; border: 1px solid %2; }")
            .arg(m_selectedBgColor.name(), ThemeManager::instance().currentTheme().dialogGroupBoxBorder.name()));
    m_fgColorBtn->setStyleSheet(
        QString("QPushButton { border-bottom: 3px solid %1; }").arg(m_selectedFgColor.name()));
}

void ConditionalFormatDialog::onUpdateRule() {
    int row = m_ruleList->currentRow();
    if (row < 0) return;

    auto newRule = buildRuleFromUI();
    if (!newRule) return;

    // Replace: remove old, insert new at same position
    m_formatting.removeRule(static_cast<size_t>(row));

    // Re-insert at position (add puts it at end, so we need to remove all after, add, re-add)
    // Simpler: clear and rebuild is complex, so just remove and add — it goes to end
    // For simplicity, remove old and add new (order may change, but rules still apply correctly)
    m_formatting.addRule(newRule);
    populateRuleList();

    // Select the last item (the updated rule)
    m_ruleList->setCurrentRow(m_ruleList->count() - 1);
}

void ConditionalFormatDialog::onConditionTypeChanged(int) {
    updateValueFieldsVisibility();
}

void ConditionalFormatDialog::updateValueFieldsVisibility() {
    int typeData = m_conditionType->currentData().toInt();
    ConditionType type = static_cast<ConditionType>(typeData);

    bool showValue1 = (type != ConditionType::Formula);
    bool showValue2 = (type == ConditionType::Between);
    bool showFormula = (type == ConditionType::Formula);

    m_value1Label->setVisible(showValue1);
    m_value1Edit->setVisible(showValue1);
    m_value2Label->setVisible(showValue2);
    m_value2Edit->setVisible(showValue2);
    m_formulaLabel->setVisible(showFormula);
    m_formulaEdit->setVisible(showFormula);
}

std::shared_ptr<ConditionalFormat> ConditionalFormatDialog::buildRuleFromUI() {
    // Parse range from the range edit field
    QString rangeStr = m_rangeEdit->text().trimmed();
    if (rangeStr.isEmpty()) {
        QMessageBox::warning(this, "Missing Range", "Please enter a cell range (e.g. A1:D100).");
        return nullptr;
    }

    // Parse range string — expect "A1:B10" format
    CellRange range(rangeStr);

    int typeData = m_conditionType->currentData().toInt();
    ConditionType type = static_cast<ConditionType>(typeData);

    auto rule = std::make_shared<ConditionalFormat>(range, type);

    if (type == ConditionType::Formula) {
        rule->setFormula(m_formulaEdit->text());
    } else {
        rule->setValue1(m_value1Edit->text());
        if (type == ConditionType::Between) {
            rule->setValue2(m_value2Edit->text());
        }
    }

    CellStyle style;
    style.bold = m_boldCheck->isChecked();
    style.italic = m_italicCheck->isChecked();
    style.underline = m_underlineCheck->isChecked();
    style.foregroundColor = m_selectedFgColor.name();
    style.backgroundColor = m_selectedBgColor.name();
    rule->setStyle(style);

    return rule;
}

void ConditionalFormatDialog::onApply() {
    accept();
}
