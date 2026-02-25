#include "FindReplaceDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>

FindReplaceDialog::FindReplaceDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Find and Replace");
    setFixedSize(420, 220);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);
    mainLayout->setContentsMargins(12, 12, 12, 12);

    auto* grid = new QGridLayout();
    grid->setSpacing(6);

    grid->addWidget(new QLabel("Find:"), 0, 0);
    m_findEdit = new QLineEdit(this);
    m_findEdit->setPlaceholderText("Search text...");
    grid->addWidget(m_findEdit, 0, 1);

    grid->addWidget(new QLabel("Replace:"), 1, 0);
    m_replaceEdit = new QLineEdit(this);
    m_replaceEdit->setPlaceholderText("Replace with...");
    grid->addWidget(m_replaceEdit, 1, 1);

    mainLayout->addLayout(grid);

    auto* optionsLayout = new QHBoxLayout();
    m_matchCaseCheck = new QCheckBox("Match case", this);
    m_wholeCellCheck = new QCheckBox("Match entire cell", this);
    optionsLayout->addWidget(m_matchCaseCheck);
    optionsLayout->addWidget(m_wholeCellCheck);
    optionsLayout->addStretch();
    mainLayout->addLayout(optionsLayout);

    auto* btnLayout = new QHBoxLayout();
    btnLayout->setSpacing(6);

    auto* findNextBtn = new QPushButton("Find Next", this);
    auto* findPrevBtn = new QPushButton("Find Previous", this);
    auto* replaceBtn = new QPushButton("Replace", this);
    auto* replaceAllBtn = new QPushButton("Replace All", this);

    findNextBtn->setDefault(true);

    btnLayout->addWidget(findPrevBtn);
    btnLayout->addWidget(findNextBtn);
    btnLayout->addWidget(replaceBtn);
    btnLayout->addWidget(replaceAllBtn);
    mainLayout->addLayout(btnLayout);

    m_statusLabel = new QLabel("", this);
    m_statusLabel->setStyleSheet("color: #666; font-size: 11px;");
    mainLayout->addWidget(m_statusLabel);

    connect(findNextBtn, &QPushButton::clicked, this, &FindReplaceDialog::findNext);
    connect(findPrevBtn, &QPushButton::clicked, this, &FindReplaceDialog::findPrevious);
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
void FindReplaceDialog::setStatus(const QString& text) { m_statusLabel->setText(text); }
