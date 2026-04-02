#include "SortDialog.h"
#include "Theme.h"
#include <QHBoxLayout>
#include <QLabel>
#include <QFrame>

SortDialog::SortDialog(int startCol, int endCol, bool hasHeaders,
                       const QStringList& headerNames, QWidget* parent)
    : QDialog(parent), m_startCol(startCol), m_endCol(endCol), m_headerNames(headerNames) {
    setWindowTitle("Sort");
    setMinimumWidth(480);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(10);
    mainLayout->setContentsMargins(12, 12, 12, 12);

    // Header checkbox
    m_headersCheckbox = new QCheckBox("My data has headers", this);
    m_headersCheckbox->setChecked(hasHeaders);
    mainLayout->addWidget(m_headersCheckbox);

    // Separator line
    auto* sep1 = new QFrame(this);
    sep1->setFrameShape(QFrame::HLine);
    sep1->setFrameShadow(QFrame::Sunken);
    mainLayout->addWidget(sep1);

    // Column labels row
    auto* headerRow = new QHBoxLayout();
    headerRow->setSpacing(8);
    auto* colLabel = new QLabel("Column", this);
    colLabel->setFixedWidth(160);
    colLabel->setStyleSheet("font-weight: 600; color: #344054;");
    auto* orderLabel = new QLabel("Order", this);
    orderLabel->setFixedWidth(140);
    orderLabel->setStyleSheet("font-weight: 600; color: #344054;");
    headerRow->addWidget(colLabel);
    headerRow->addWidget(orderLabel);
    headerRow->addStretch();
    mainLayout->addLayout(headerRow);

    // Sort levels container
    m_levelsLayout = new QVBoxLayout();
    m_levelsLayout->setSpacing(6);
    mainLayout->addLayout(m_levelsLayout);

    // Add/Delete buttons
    auto* btnRow = new QHBoxLayout();
    btnRow->setSpacing(8);
    m_addBtn = new QPushButton("+ Add Level", this);
    m_addBtn->setFixedWidth(110);
    m_addBtn->setProperty("primary", true);
    m_deleteBtn = new QPushButton("Delete Level", this);
    m_deleteBtn->setFixedWidth(110);
    m_deleteBtn->setProperty("secondary", true);
    btnRow->addWidget(m_addBtn);
    btnRow->addWidget(m_deleteBtn);
    btnRow->addStretch();
    mainLayout->addLayout(btnRow);

    connect(m_addBtn, &QPushButton::clicked, this, &SortDialog::addLevel);
    connect(m_deleteBtn, &QPushButton::clicked, this, &SortDialog::deleteLevel);
    connect(m_headersCheckbox, &QCheckBox::toggled, this, &SortDialog::updateColumnLabels);

    mainLayout->addStretch();

    // Separator before OK/Cancel
    auto* sep2 = new QFrame(this);
    sep2->setFrameShape(QFrame::HLine);
    sep2->setFrameShadow(QFrame::Sunken);
    mainLayout->addWidget(sep2);

    // OK / Cancel
    auto* okCancelRow = new QHBoxLayout();
    okCancelRow->addStretch();
    auto* okBtn = new QPushButton("OK", this);
    okBtn->setDefault(true);
    okBtn->setFixedWidth(80);
    okBtn->setProperty("primary", true);
    auto* cancelBtn = new QPushButton("Cancel", this);
    cancelBtn->setFixedWidth(80);
    cancelBtn->setProperty("secondary", true);
    okCancelRow->addWidget(okBtn);
    okCancelRow->addWidget(cancelBtn);
    mainLayout->addLayout(okCancelRow);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);

    // Add the first sort level by default
    addLevel();

    // Apply global dialog theme
    setStyleSheet(ThemeManager::dialogStylesheet());

    m_deleteBtn->setEnabled(false);
}

void SortDialog::addLevel() {
    if ((int)m_levels.size() >= MAX_LEVELS) return;

    auto* rowWidget = new QWidget(this);
    auto* rowLayout = new QHBoxLayout(rowWidget);
    rowLayout->setContentsMargins(0, 0, 0, 0);
    rowLayout->setSpacing(8);

    // Level label
    auto* levelLabel = new QLabel(m_levels.empty() ? "Sort by" : "Then by", rowWidget);
    levelLabel->setFixedWidth(50);
    levelLabel->setStyleSheet("color: #344054; font-size: 12px;");

    // Column combo
    auto* columnCombo = new QComboBox(rowWidget);
    columnCombo->setFixedWidth(160);
    for (int c = m_startCol; c <= m_endCol; ++c) {
        columnCombo->addItem(columnLabel(c), c);
    }
    // Default: each new level selects the next column if available
    int defaultIdx = qMin((int)m_levels.size(), columnCombo->count() - 1);
    columnCombo->setCurrentIndex(defaultIdx);

    // Order combo
    auto* orderCombo = new QComboBox(rowWidget);
    orderCombo->setFixedWidth(140);
    orderCombo->addItem("A to Z");
    orderCombo->addItem("Z to A");

    rowLayout->addWidget(levelLabel);
    rowLayout->addWidget(columnCombo);
    rowLayout->addWidget(orderCombo);
    rowLayout->addStretch();

    m_levelsLayout->addWidget(rowWidget);
    m_levels.push_back({columnCombo, orderCombo, rowWidget});

    // Update button states
    m_addBtn->setEnabled((int)m_levels.size() < MAX_LEVELS);
    m_deleteBtn->setEnabled(m_levels.size() > 1);

    // Update size
    adjustSize();
}

void SortDialog::deleteLevel() {
    if (m_levels.size() <= 1) return;

    // Remove last level
    auto& last = m_levels.back();
    m_levelsLayout->removeWidget(last.rowWidget);
    delete last.rowWidget;
    m_levels.pop_back();

    // Update button states
    m_addBtn->setEnabled((int)m_levels.size() < MAX_LEVELS);
    m_deleteBtn->setEnabled(m_levels.size() > 1);

    adjustSize();
}

std::vector<SortLevel> SortDialog::getSortLevels() const {
    std::vector<SortLevel> result;
    result.reserve(m_levels.size());
    for (const auto& lvl : m_levels) {
        SortLevel sl;
        sl.column = lvl.columnCombo->currentData().toInt();
        sl.ascending = (lvl.orderCombo->currentIndex() == 0);
        result.push_back(sl);
    }
    return result;
}

bool SortDialog::hasHeaders() const {
    return m_headersCheckbox->isChecked();
}

void SortDialog::updateColumnLabels() {
    for (auto& lvl : m_levels) {
        int currentData = lvl.columnCombo->currentData().toInt();
        lvl.columnCombo->clear();
        for (int c = m_startCol; c <= m_endCol; ++c) {
            lvl.columnCombo->addItem(columnLabel(c), c);
        }
        // Restore selection
        int idx = lvl.columnCombo->findData(currentData);
        if (idx >= 0) lvl.columnCombo->setCurrentIndex(idx);
    }
}

QString SortDialog::columnLabel(int col) const {
    // If headers enabled and we have header names, show them
    if (m_headersCheckbox->isChecked()) {
        int headerIdx = col - m_startCol;
        if (headerIdx >= 0 && headerIdx < m_headerNames.size() && !m_headerNames[headerIdx].isEmpty()) {
            return m_headerNames[headerIdx];
        }
    }
    // Fallback: Excel-style column letter
    QString label;
    int c = col;
    do {
        label.prepend(QChar('A' + (c % 26)));
        c = c / 26 - 1;
    } while (c >= 0);
    return QString("Column %1").arg(label);
}
