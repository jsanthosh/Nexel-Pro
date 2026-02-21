#include "ShapePropertiesDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLabel>
#include <QLineEdit>
#include <QSpinBox>
#include <QSlider>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QColorDialog>

ShapePropertiesDialog::ShapePropertiesDialog(const ShapeConfig& config, QWidget* parent)
    : QDialog(parent), m_config(config) {
    setWindowTitle("Shape Properties");
    setMinimumSize(420, 480);
    createLayout();

    setStyleSheet(
        "QDialog { background: #FAFBFC; }"
        "QGroupBox { font-weight: bold; border: 1px solid #D0D5DD; border-radius: 6px; "
        "margin-top: 8px; padding-top: 16px; background: white; }"
        "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #344054; }"
        "QLineEdit { border: 1px solid #D0D5DD; border-radius: 4px; padding: 5px 8px; background: white; }"
        "QLineEdit:focus { border-color: #4A90D9; }"
        "QSpinBox { border: 1px solid #D0D5DD; border-radius: 4px; padding: 4px 8px; background: white; }"
        "QSpinBox:focus { border-color: #4A90D9; }"
    );
}

void ShapePropertiesDialog::createLayout() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(10);

    // --- Preview ---
    QGroupBox* previewGroup = new QGroupBox("Preview");
    QVBoxLayout* prevLayout = new QVBoxLayout(previewGroup);
    m_preview = new ShapeWidget();
    m_preview->setFixedHeight(100);
    m_preview->setConfig(m_config);
    prevLayout->addWidget(m_preview);
    mainLayout->addWidget(previewGroup);

    // --- Colors ---
    QGroupBox* colorGroup = new QGroupBox("Colors");
    QGridLayout* colorLayout = new QGridLayout(colorGroup);

    colorLayout->addWidget(new QLabel("Fill Color:"), 0, 0);
    m_fillColorBtn = new QPushButton();
    m_fillColorBtn->setFixedSize(80, 28);
    m_fillColorBtn->setStyleSheet(
        QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                "QPushButton:hover { border-color: #4A90D9; }")
        .arg(m_config.fillColor.name()));
    connect(m_fillColorBtn, &QPushButton::clicked, this, &ShapePropertiesDialog::chooseFillColor);
    colorLayout->addWidget(m_fillColorBtn, 0, 1);

    colorLayout->addWidget(new QLabel("Stroke Color:"), 1, 0);
    m_strokeColorBtn = new QPushButton();
    m_strokeColorBtn->setFixedSize(80, 28);
    m_strokeColorBtn->setStyleSheet(
        QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                "QPushButton:hover { border-color: #4A90D9; }")
        .arg(m_config.strokeColor.name()));
    connect(m_strokeColorBtn, &QPushButton::clicked, this, &ShapePropertiesDialog::chooseStrokeColor);
    colorLayout->addWidget(m_strokeColorBtn, 1, 1);

    colorLayout->addWidget(new QLabel("Stroke Width:"), 2, 0);
    m_strokeWidthSpin = new QSpinBox();
    m_strokeWidthSpin->setRange(0, 20);
    m_strokeWidthSpin->setValue(m_config.strokeWidth);
    connect(m_strokeWidthSpin, &QSpinBox::valueChanged, this, &ShapePropertiesDialog::updatePreview);
    colorLayout->addWidget(m_strokeWidthSpin, 2, 1);

    colorLayout->addWidget(new QLabel("Corner Radius:"), 3, 0);
    m_cornerRadiusSpin = new QSpinBox();
    m_cornerRadiusSpin->setRange(0, 50);
    m_cornerRadiusSpin->setValue(static_cast<int>(m_config.cornerRadius));
    connect(m_cornerRadiusSpin, &QSpinBox::valueChanged, this, &ShapePropertiesDialog::updatePreview);
    colorLayout->addWidget(m_cornerRadiusSpin, 3, 1);

    colorLayout->addWidget(new QLabel("Opacity:"), 4, 0);
    QHBoxLayout* opacityLayout = new QHBoxLayout();
    m_opacitySlider = new QSlider(Qt::Horizontal);
    m_opacitySlider->setRange(10, 100);
    m_opacitySlider->setValue(static_cast<int>(m_config.opacity * 100));
    m_opacitySlider->setStyleSheet(
        "QSlider::groove:horizontal { background: #E0E3E8; height: 4px; border-radius: 2px; }"
        "QSlider::handle:horizontal { background: #4A90D9; width: 16px; height: 16px; margin: -6px 0; border-radius: 8px; }"
    );
    connect(m_opacitySlider, &QSlider::valueChanged, this, &ShapePropertiesDialog::updatePreview);
    opacityLayout->addWidget(m_opacitySlider);
    m_opacityLabel = new QLabel(QString("%1%").arg(m_opacitySlider->value()));
    m_opacityLabel->setFixedWidth(35);
    opacityLayout->addWidget(m_opacityLabel);
    colorLayout->addLayout(opacityLayout, 4, 1);

    mainLayout->addWidget(colorGroup);

    // --- Text ---
    QGroupBox* textGroup = new QGroupBox("Text");
    QGridLayout* textLayout = new QGridLayout(textGroup);

    textLayout->addWidget(new QLabel("Text:"), 0, 0);
    m_textEdit = new QLineEdit(m_config.text);
    connect(m_textEdit, &QLineEdit::textChanged, this, &ShapePropertiesDialog::updatePreview);
    textLayout->addWidget(m_textEdit, 0, 1);

    textLayout->addWidget(new QLabel("Font Size:"), 1, 0);
    m_fontSizeSpin = new QSpinBox();
    m_fontSizeSpin->setRange(6, 72);
    m_fontSizeSpin->setValue(m_config.fontSize);
    connect(m_fontSizeSpin, &QSpinBox::valueChanged, this, &ShapePropertiesDialog::updatePreview);
    textLayout->addWidget(m_fontSizeSpin, 1, 1);

    textLayout->addWidget(new QLabel("Text Color:"), 2, 0);
    m_textColorBtn = new QPushButton();
    m_textColorBtn->setFixedSize(80, 28);
    m_textColorBtn->setStyleSheet(
        QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                "QPushButton:hover { border-color: #4A90D9; }")
        .arg(m_config.textColor.name()));
    connect(m_textColorBtn, &QPushButton::clicked, this, &ShapePropertiesDialog::chooseTextColor);
    textLayout->addWidget(m_textColorBtn, 2, 1);

    mainLayout->addWidget(textGroup);

    // --- Buttons ---
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Apply");
    buttons->button(QDialogButtonBox::Ok)->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 8px 24px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }"
    );
    buttons->button(QDialogButtonBox::Cancel)->setStyleSheet(
        "QPushButton { background: #F0F2F5; border: 1px solid #D0D5DD; border-radius: 4px; padding: 8px 20px; }"
        "QPushButton:hover { background: #E8ECF0; }"
    );
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);
}

void ShapePropertiesDialog::chooseFillColor() {
    QColor color = QColorDialog::getColor(m_config.fillColor, this, "Fill Color");
    if (color.isValid()) {
        m_config.fillColor = color;
        m_fillColorBtn->setStyleSheet(
            QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                    "QPushButton:hover { border-color: #4A90D9; }").arg(color.name()));
        updatePreview();
    }
}

void ShapePropertiesDialog::chooseStrokeColor() {
    QColor color = QColorDialog::getColor(m_config.strokeColor, this, "Stroke Color");
    if (color.isValid()) {
        m_config.strokeColor = color;
        m_strokeColorBtn->setStyleSheet(
            QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                    "QPushButton:hover { border-color: #4A90D9; }").arg(color.name()));
        updatePreview();
    }
}

void ShapePropertiesDialog::chooseTextColor() {
    QColor color = QColorDialog::getColor(m_config.textColor, this, "Text Color");
    if (color.isValid()) {
        m_config.textColor = color;
        m_textColorBtn->setStyleSheet(
            QString("QPushButton { background: %1; border: 1px solid #AAA; border-radius: 4px; }"
                    "QPushButton:hover { border-color: #4A90D9; }").arg(color.name()));
        updatePreview();
    }
}

void ShapePropertiesDialog::updatePreview() {
    m_config.strokeWidth = m_strokeWidthSpin->value();
    m_config.cornerRadius = m_cornerRadiusSpin->value();
    m_config.opacity = m_opacitySlider->value() / 100.0f;
    m_config.text = m_textEdit->text();
    m_config.fontSize = m_fontSizeSpin->value();
    m_opacityLabel->setText(QString("%1%").arg(m_opacitySlider->value()));
    m_preview->setConfig(m_config);
}

ShapeConfig ShapePropertiesDialog::getConfig() const {
    ShapeConfig cfg = m_config;
    cfg.strokeWidth = m_strokeWidthSpin->value();
    cfg.cornerRadius = m_cornerRadiusSpin->value();
    cfg.opacity = m_opacitySlider->value() / 100.0f;
    cfg.text = m_textEdit->text();
    cfg.fontSize = m_fontSizeSpin->value();
    return cfg;
}
