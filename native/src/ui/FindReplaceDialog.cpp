#include "FindReplaceDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>

FindReplaceDialog::FindReplaceDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Find and Replace");
    setFixedSize(480, 280);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);
    mainLayout->setContentsMargins(12, 12, 12, 12);

    // --- Row 1-2: Find / Replace text fields ---
    auto* grid = new QGridLayout();
    grid->setSpacing(6);

    grid->addWidget(new QLabel("Find:"), 0, 0);
    m_findEdit = new QLineEdit(this);
    m_findEdit->setPlaceholderText("Search text...");
    grid->addWidget(m_findEdit, 0, 1, 1, 3);

    grid->addWidget(new QLabel("Replace:"), 1, 0);
    m_replaceEdit = new QLineEdit(this);
    m_replaceEdit->setPlaceholderText("Replace with...");
    grid->addWidget(m_replaceEdit, 1, 1, 1, 3);

    // --- Row 3: Within (scope) + Look in ---
    grid->addWidget(new QLabel("Within:"), 2, 0);
    m_scopeCombo = new QComboBox(this);
    m_scopeCombo->addItem("Sheet");
    m_scopeCombo->addItem("Workbook");
    m_scopeCombo->setFixedWidth(100);
    grid->addWidget(m_scopeCombo, 2, 1);

    grid->addWidget(new QLabel("Look in:"), 2, 2);
    m_lookInCombo = new QComboBox(this);
    m_lookInCombo->addItem("Values");
    m_lookInCombo->addItem("Formulas");
    m_lookInCombo->setFixedWidth(100);
    grid->addWidget(m_lookInCombo, 2, 3);

    mainLayout->addLayout(grid);

    // --- Options row: Match case / Match entire cell ---
    auto* optionsLayout = new QHBoxLayout();
    m_matchCaseCheck = new QCheckBox("Match case", this);
    m_wholeCellCheck = new QCheckBox("Match entire cell", this);
    optionsLayout->addWidget(m_matchCaseCheck);
    optionsLayout->addWidget(m_wholeCellCheck);
    optionsLayout->addStretch();
    mainLayout->addLayout(optionsLayout);

    // --- Buttons row ---
    auto* btnLayout = new QHBoxLayout();
    btnLayout->setSpacing(6);

    auto* findPrevBtn = new QPushButton("Find Previous", this);
    auto* findNextBtn = new QPushButton("Find Next", this);
    auto* findAllBtn = new QPushButton("Find All", this);
    auto* replaceBtn = new QPushButton("Replace", this);
    auto* replaceAllBtn = new QPushButton("Replace All", this);

    findNextBtn->setDefault(true);

    btnLayout->addWidget(findPrevBtn);
    btnLayout->addWidget(findNextBtn);
    btnLayout->addWidget(findAllBtn);
    btnLayout->addWidget(replaceBtn);
    btnLayout->addWidget(replaceAllBtn);
    mainLayout->addLayout(btnLayout);

    // --- Status label ---
    m_statusLabel = new QLabel("", this);
    m_statusLabel->setStyleSheet("color: #666; font-size: 11px;");
    mainLayout->addWidget(m_statusLabel);

    connect(findNextBtn, &QPushButton::clicked, this, &FindReplaceDialog::findNext);
    connect(findPrevBtn, &QPushButton::clicked, this, &FindReplaceDialog::findPrevious);
    connect(findAllBtn, &QPushButton::clicked, this, &FindReplaceDialog::findAll);
    connect(replaceBtn, &QPushButton::clicked, this, &FindReplaceDialog::replaceOne);
    connect(replaceAllBtn, &QPushButton::clicked, this, &FindReplaceDialog::replaceAll);

    // Enter in find field triggers find next
    connect(m_findEdit, &QLineEdit::returnPressed, this, &FindReplaceDialog::findNext);

    setStyleSheet(ThemeManager::dialogStylesheet());
}

QString FindReplaceDialog::findText() const { return m_findEdit->text(); }
QString FindReplaceDialog::replaceText() const { return m_replaceEdit->text(); }
bool FindReplaceDialog::matchCase() const { return m_matchCaseCheck->isChecked(); }
bool FindReplaceDialog::matchWholeCell() const { return m_wholeCellCheck->isChecked(); }

FindReplaceDialog::SearchScope FindReplaceDialog::searchScope() const {
    return m_scopeCombo->currentIndex() == 1 ? Workbook : Sheet;
}

FindReplaceDialog::LookIn FindReplaceDialog::lookIn() const {
    return m_lookInCombo->currentIndex() == 1 ? Formulas : Values;
}

void FindReplaceDialog::setStatus(const QString& text) { m_statusLabel->setText(text); }
