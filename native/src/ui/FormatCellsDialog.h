#ifndef FORMATCELLSDIALOG_H
#define FORMATCELLSDIALOG_H

#include <QDialog>
#include "../core/Cell.h"

class QTabWidget;
class QListWidget;
class QSpinBox;
class QCheckBox;
class QComboBox;
class QLineEdit;
class QLabel;
class QFontComboBox;
class QColorDialog;
class QPushButton;
struct DocumentTheme;

class FormatCellsDialog : public QDialog {
    Q_OBJECT

public:
    FormatCellsDialog(const CellStyle& style, QWidget* parent = nullptr);
    CellStyle getStyle() const;

    // Set document theme for theme-aware color picker
    void setDocumentTheme(const DocumentTheme* theme) { m_docTheme = theme; }

private:
    void createNumberTab(QWidget* tab);
    void createFontTab(QWidget* tab);
    void createAlignmentTab(QWidget* tab);
    void createBorderTab(QWidget* tab);
    void createFillTab(QWidget* tab);
    void createProtectionTab(QWidget* tab);
    void updatePreview();
    void updateBorderPreview();
    void loadStyle(const CellStyle& style);
    void pickColor(const QString& title, QString& colorStr, QPushButton* btn);

    CellStyle m_style;
    const DocumentTheme* m_docTheme = nullptr;

    // Number tab
    QListWidget* m_categoryList = nullptr;
    QSpinBox* m_decimalSpin = nullptr;
    QCheckBox* m_thousandCheck = nullptr;
    QComboBox* m_currencyCombo = nullptr;
    QComboBox* m_dateFormatCombo = nullptr;
    QLineEdit* m_customFormatEdit = nullptr;
    QLabel* m_previewLabel = nullptr;

    // Font tab
    QFontComboBox* m_fontFamilyCombo = nullptr;
    QSpinBox* m_fontSizeSpin = nullptr;
    QCheckBox* m_boldCheck = nullptr;
    QCheckBox* m_italicCheck = nullptr;
    QCheckBox* m_underlineCheck = nullptr;
    QCheckBox* m_strikethroughCheck = nullptr;
    QPushButton* m_fontColorBtn = nullptr;
    QString m_fontColorStr;

    // Alignment tab
    QComboBox* m_hAlignCombo = nullptr;
    QComboBox* m_vAlignCombo = nullptr;

    // Border tab
    QListWidget* m_borderStyleList = nullptr;
    QPushButton* m_borderColorBtn = nullptr;
    QString m_borderColorStr = "#000000";
    int m_selectedBorderStyle = 1;  // index into line styles
    QPushButton* m_borderPresetNone = nullptr;
    QPushButton* m_borderPresetOutline = nullptr;
    QPushButton* m_borderPresetInside = nullptr;
    // Per-edge toggle buttons
    QPushButton* m_borderTop = nullptr;
    QPushButton* m_borderBottom = nullptr;
    QPushButton* m_borderLeft = nullptr;
    QPushButton* m_borderRight = nullptr;
    QWidget* m_borderPreviewWidget = nullptr;

    // Fill tab
    QPushButton* m_fillColorBtn = nullptr;
    QString m_fillColorStr;

    // Protection tab
    QCheckBox* m_lockedCheck = nullptr;
    QCheckBox* m_hiddenCheck = nullptr;
};

#endif // FORMATCELLSDIALOG_H
