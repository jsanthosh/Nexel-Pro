#include "GoalSeekDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QLabel>

GoalSeekDialog::GoalSeekDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Goal Seek");
    setFixedSize(350, 180);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* layout = new QVBoxLayout(this);
    layout->setSpacing(10);
    layout->setContentsMargins(12, 12, 12, 12);

    auto* formLayout = new QFormLayout();
    formLayout->setSpacing(8);

    m_setCellEdit = new QLineEdit(this);
    m_setCellEdit->setPlaceholderText("e.g. B5");
    formLayout->addRow("Set cell:", m_setCellEdit);

    m_toValueEdit = new QLineEdit(this);
    m_toValueEdit->setPlaceholderText("e.g. 100");
    formLayout->addRow("To value:", m_toValueEdit);

    m_byChangingEdit = new QLineEdit(this);
    m_byChangingEdit->setPlaceholderText("e.g. A1");
    formLayout->addRow("By changing cell:", m_byChangingEdit);

    layout->addLayout(formLayout);

    auto* btnLayout = new QHBoxLayout();
    btnLayout->addStretch();
    auto* okBtn = new QPushButton("OK", this);
    okBtn->setDefault(true);
    auto* cancelBtn = new QPushButton("Cancel", this);
    btnLayout->addWidget(okBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);

    setStyleSheet(ThemeManager::dialogStylesheet());
}

QString GoalSeekDialog::getSetCell() const {
    return m_setCellEdit->text().trimmed().toUpper();
}

double GoalSeekDialog::getToValue() const {
    return m_toValueEdit->text().trimmed().toDouble();
}

QString GoalSeekDialog::getByChangingCell() const {
    return m_byChangingEdit->text().trimmed().toUpper();
}
