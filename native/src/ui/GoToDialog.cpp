#include "GoToDialog.h"
#include "Theme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QRegularExpression>

GoToDialog::GoToDialog(QWidget* parent)
    : QDialog(parent) {
    setWindowTitle("Go To");
    setFixedSize(300, 120);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* layout = new QVBoxLayout(this);
    layout->setSpacing(10);
    layout->setContentsMargins(12, 12, 12, 12);

    auto* inputLayout = new QHBoxLayout();
    inputLayout->addWidget(new QLabel("Cell reference:"));
    m_cellRefEdit = new QLineEdit(this);
    m_cellRefEdit->setPlaceholderText("e.g. A1, B25, AA100");
    inputLayout->addWidget(m_cellRefEdit);
    layout->addLayout(inputLayout);

    auto* btnLayout = new QHBoxLayout();
    btnLayout->addStretch();
    auto* goBtn = new QPushButton("Go", this);
    goBtn->setDefault(true);
    auto* cancelBtn = new QPushButton("Cancel", this);
    btnLayout->addWidget(goBtn);
    btnLayout->addWidget(cancelBtn);
    layout->addLayout(btnLayout);

    connect(goBtn, &QPushButton::clicked, this, &QDialog::accept);
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);
    connect(m_cellRefEdit, &QLineEdit::returnPressed, this, &QDialog::accept);

    setStyleSheet(ThemeManager::dialogStylesheet());
}

CellAddress GoToDialog::getAddress() const {
    QString ref = m_cellRefEdit->text().trimmed().toUpper();

    static QRegularExpression re("^([A-Z]+)(\\d+)$");
    auto match = re.match(ref);
    if (!match.hasMatch()) return CellAddress(-1, -1);

    QString letters = match.captured(1);
    int rowNum = match.captured(2).toInt();

    // Convert column letters to index (A=0, B=1, ..., Z=25, AA=26)
    int col = 0;
    for (int i = 0; i < letters.length(); ++i) {
        col = col * 26 + (letters[i].unicode() - 'A' + 1);
    }
    col -= 1;

    int row = rowNum - 1; // Convert 1-based to 0-based

    return CellAddress(row, col);
}
