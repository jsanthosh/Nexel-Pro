#include "RemoveDuplicatesDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QScrollArea>
#include <QGroupBox>

RemoveDuplicatesDialog::RemoveDuplicatesDialog(const QStringList& columnHeaders, QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Remove Duplicates");
    setFixedSize(350, 420);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(10);
    mainLayout->setContentsMargins(16, 16, 16, 16);

    // "My data has headers" checkbox
    m_hasHeadersCheck = new QCheckBox("My data has &headers", this);
    m_hasHeadersCheck->setChecked(true);
    mainLayout->addWidget(m_hasHeadersCheck);

    // Description label
    auto* descLabel = new QLabel("Select the columns that contain duplicates:", this);
    descLabel->setWordWrap(true);
    descLabel->setStyleSheet("color: #344054; font-size: 12px;");
    mainLayout->addWidget(descLabel);

    // Select All / Unselect All buttons
    auto* selBtnLayout = new QHBoxLayout();
    auto* selectAllBtn = new QPushButton("&Select All", this);
    selectAllBtn->setFixedWidth(100);
    auto* unselectAllBtn = new QPushButton("&Unselect All", this);
    unselectAllBtn->setFixedWidth(100);
    selBtnLayout->addWidget(selectAllBtn);
    selBtnLayout->addWidget(unselectAllBtn);
    selBtnLayout->addStretch();
    mainLayout->addLayout(selBtnLayout);

    // Scrollable column list
    auto* scrollArea = new QScrollArea(this);
    scrollArea->setWidgetResizable(true);
    scrollArea->setFrameShape(QFrame::StyledPanel);
    scrollArea->setStyleSheet("QScrollArea { border: 1px solid #D0D5DD; border-radius: 4px; }");

    auto* scrollWidget = new QWidget();
    auto* colLayout = new QVBoxLayout(scrollWidget);
    colLayout->setSpacing(4);
    colLayout->setContentsMargins(8, 8, 8, 8);

    for (int i = 0; i < columnHeaders.size(); ++i) {
        auto* check = new QCheckBox(columnHeaders[i], scrollWidget);
        check->setChecked(true);
        m_columnChecks.append(check);
        colLayout->addWidget(check);
    }
    colLayout->addStretch();

    scrollArea->setWidget(scrollWidget);
    mainLayout->addWidget(scrollArea, 1);

    // OK / Cancel buttons
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

    // Connections
    connect(selectAllBtn, &QPushButton::clicked, this, [this]() {
        for (auto* check : m_columnChecks) check->setChecked(true);
    });
    connect(unselectAllBtn, &QPushButton::clicked, this, [this]() {
        for (auto* check : m_columnChecks) check->setChecked(false);
    });
    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);

    // Update column labels when "has headers" changes
    connect(m_hasHeadersCheck, &QCheckBox::toggled, this, [this, columnHeaders](bool hasHeaders) {
        for (int i = 0; i < m_columnChecks.size() && i < columnHeaders.size(); ++i) {
            if (hasHeaders) {
                m_columnChecks[i]->setText(columnHeaders[i]);
            } else {
                // Show generic column letters
                QString colLetter;
                int c = i;
                do {
                    colLetter.prepend(QChar('A' + c % 26));
                    c = c / 26 - 1;
                } while (c >= 0);
                m_columnChecks[i]->setText(QString("Column %1").arg(colLetter));
            }
        }
    });

    setStyleSheet(ThemeManager::dialogStylesheet());
}

QVector<int> RemoveDuplicatesDialog::getSelectedColumns() const {
    QVector<int> result;
    for (int i = 0; i < m_columnChecks.size(); ++i) {
        if (m_columnChecks[i]->isChecked()) {
            result.append(i);
        }
    }
    return result;
}

bool RemoveDuplicatesDialog::hasHeaders() const {
    return m_hasHeadersCheck->isChecked();
}
