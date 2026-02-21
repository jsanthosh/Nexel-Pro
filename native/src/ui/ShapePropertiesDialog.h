#ifndef SHAPEPROPERTIESDIALOG_H
#define SHAPEPROPERTIESDIALOG_H

#include <QDialog>
#include "ShapeWidget.h"

class QLineEdit;
class QSpinBox;
class QSlider;
class QPushButton;
class QLabel;

class ShapePropertiesDialog : public QDialog {
    Q_OBJECT

public:
    explicit ShapePropertiesDialog(const ShapeConfig& config, QWidget* parent = nullptr);
    ShapeConfig getConfig() const;

private slots:
    void chooseFillColor();
    void chooseStrokeColor();
    void chooseTextColor();
    void updatePreview();

private:
    void createLayout();

    ShapeConfig m_config;

    QPushButton* m_fillColorBtn = nullptr;
    QPushButton* m_strokeColorBtn = nullptr;
    QPushButton* m_textColorBtn = nullptr;
    QSpinBox* m_strokeWidthSpin = nullptr;
    QSlider* m_opacitySlider = nullptr;
    QLabel* m_opacityLabel = nullptr;
    QSpinBox* m_cornerRadiusSpin = nullptr;
    QLineEdit* m_textEdit = nullptr;
    QSpinBox* m_fontSizeSpin = nullptr;
    ShapeWidget* m_preview = nullptr;
};

#endif // SHAPEPROPERTIESDIALOG_H
