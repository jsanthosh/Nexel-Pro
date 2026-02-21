#include "SparklineDialog.h"

#include <QCheckBox>
#include <QColorDialog>
#include <QComboBox>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QVBoxLayout>

SparklineDialog::SparklineDialog(QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle("Insert Sparkline");
    resize(420, 380);
    setStyleSheet(
        "QDialog { background: white; }"
        "QLabel { color: #344054; font-size: 13px; }"
        "QLineEdit { padding: 8px; border: 1px solid #D0D5DD; border-radius: 6px; font-size: 13px; }"
        "QComboBox { padding: 6px 10px; border: 1px solid #D0D5DD; border-radius: 6px; }"
        "QPushButton { padding: 8px 20px; border-radius: 6px; font-size: 13px; font-weight: 500; }"
        "QCheckBox { font-size: 13px; color: #344054; }");
    createLayout();
}

SparklineConfig SparklineDialog::getConfig() const
{
    SparklineConfig config;
    config.dataRange = m_dataRangeEdit->text();

    const QString typeText = m_typeCombo->currentText();
    if (typeText == "Line")
        config.type = SparklineType::Line;
    else if (typeText == "Column")
        config.type = SparklineType::Column;
    else
        config.type = SparklineType::WinLoss;

    config.lineColor = m_lineColor;
    config.highPointColor = m_highColor;
    config.lowPointColor = m_lowColor;
    config.showHighPoint = m_showHighCheck->isChecked();
    config.showLowPoint = m_showLowCheck->isChecked();

    return config;
}

QString SparklineDialog::getDestinationRange() const
{
    return m_destinationEdit->text();
}

void SparklineDialog::setDataRange(const QString& range)
{
    m_dataRangeEdit->setText(range);
}

void SparklineDialog::setDestination(const QString& dest)
{
    m_destinationEdit->setText(dest);
}

void SparklineDialog::createLayout()
{
    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(12);
    mainLayout->setContentsMargins(20, 20, 20, 20);

    // --- Title ---
    auto* titleLabel = new QLabel("Insert Sparkline");
    titleLabel->setStyleSheet("font-size: 16px; font-weight: 600; color: #1D2939;");
    mainLayout->addWidget(titleLabel);

    // --- Form ---
    auto* formLayout = new QFormLayout;
    formLayout->setSpacing(10);
    formLayout->setLabelAlignment(Qt::AlignRight);

    m_dataRangeEdit = new QLineEdit;
    m_dataRangeEdit->setPlaceholderText("e.g. A1:A10");
    formLayout->addRow("Data Range:", m_dataRangeEdit);

    m_destinationEdit = new QLineEdit;
    m_destinationEdit->setPlaceholderText("e.g. B1");
    formLayout->addRow("Destination Cell:", m_destinationEdit);

    m_typeCombo = new QComboBox;
    m_typeCombo->addItems({"Line", "Column", "Win-Loss"});
    formLayout->addRow("Type:", m_typeCombo);

    mainLayout->addLayout(formLayout);

    // --- Helper to create a color button ---
    auto makeColorButton = [this](const QColor& initial, QPushButton*& btn) {
        btn = new QPushButton;
        btn->setFixedSize(60, 28);
        btn->setStyleSheet(
            QString("QPushButton { background: %1; border: 1px solid #D0D5DD; border-radius: 6px; }")
                .arg(initial.name()));
        return btn;
    };

    // --- Color row ---
    auto* colorForm = new QFormLayout;
    colorForm->setSpacing(10);
    colorForm->setLabelAlignment(Qt::AlignRight);

    makeColorButton(m_lineColor, m_lineColorBtn);
    colorForm->addRow("Line Color:", m_lineColorBtn);

    makeColorButton(m_highColor, m_highColorBtn);
    colorForm->addRow("High Point Color:", m_highColorBtn);

    makeColorButton(m_lowColor, m_lowColorBtn);
    colorForm->addRow("Low Point Color:", m_lowColorBtn);

    mainLayout->addLayout(colorForm);

    // Color-button connections
    auto connectColorBtn = [this](QPushButton* btn, QColor& targetColor) {
        connect(btn, &QPushButton::clicked, this, [this, btn, &targetColor]() {
            QColor chosen = QColorDialog::getColor(targetColor, this, "Select Color");
            if (chosen.isValid()) {
                targetColor = chosen;
                btn->setStyleSheet(
                    QString("QPushButton { background: %1; border: 1px solid #D0D5DD; border-radius: 6px; }")
                        .arg(chosen.name()));
            }
        });
    };
    connectColorBtn(m_lineColorBtn, m_lineColor);
    connectColorBtn(m_highColorBtn, m_highColor);
    connectColorBtn(m_lowColorBtn, m_lowColor);

    // --- Checkboxes ---
    auto* checkLayout = new QHBoxLayout;
    m_showHighCheck = new QCheckBox("Show High Point");
    m_showHighCheck->setChecked(true);
    m_showLowCheck = new QCheckBox("Show Low Point");
    m_showLowCheck->setChecked(true);
    checkLayout->addWidget(m_showHighCheck);
    checkLayout->addWidget(m_showLowCheck);
    mainLayout->addLayout(checkLayout);

    mainLayout->addStretch();

    // --- OK / Cancel ---
    auto* buttonLayout = new QHBoxLayout;
    buttonLayout->addStretch();

    auto* cancelBtn = new QPushButton("Cancel");
    cancelBtn->setStyleSheet(
        "QPushButton { background: #F2F4F7; color: #344054; border: 1px solid #D0D5DD; }"
        "QPushButton:hover { background: #E4E7EC; }");
    connect(cancelBtn, &QPushButton::clicked, this, &QDialog::reject);

    auto* okBtn = new QPushButton("OK");
    okBtn->setDefault(true);
    okBtn->setStyleSheet(
        "QPushButton { background: #16A34A; color: white; border: none; }"
        "QPushButton:hover { background: #15803D; }");
    connect(okBtn, &QPushButton::clicked, this, &QDialog::accept);

    buttonLayout->addWidget(cancelBtn);
    buttonLayout->addWidget(okBtn);
    mainLayout->addLayout(buttonLayout);
}
