#ifndef TOOLBAR_H
#define TOOLBAR_H

#include <QToolBar>
#include <QColor>
#include "../core/Cell.h"
#include "../core/DocumentTheme.h"

class QFontComboBox;
class QSpinBox;
class QToolButton;

class Toolbar : public QToolBar {
    Q_OBJECT

public:
    explicit Toolbar(QWidget* parent = nullptr);

    // Creates and returns a second toolbar for row 2 (layout, borders, data, chat)
    QToolBar* createSecondaryToolbar(QWidget* parent);

    // Enable/disable save button based on dirty state
    void setSaveEnabled(bool enabled);

    // Sync toolbar state to reflect the given cell style (updates checked states, font, colors, etc.)
    void syncToStyle(const CellStyle& style);

    // Theme support
    void onThemeChanged();

    // Document theme (for color picker)
    void setDocumentTheme(const DocumentTheme* theme) { m_docTheme = theme; }
    const DocumentTheme* documentTheme() const { return m_docTheme; }

signals:
    // File
    void newDocument();
    void saveDocument();
    // Edit
    void undo();
    void redo();
    // Format painter
    void formatPainterToggled();
    // Font formatting
    void bold();
    void italic();
    void underline();
    void strikethrough();
    void fontFamilyChanged(const QString& family);
    void fontSizeChanged(int size);
    // Colors (colorStr can be "theme:4:0.4" or "#RRGGBB")
    void foregroundColorChanged(const QString& colorStr, const QColor& displayColor);
    void backgroundColorChanged(const QString& colorStr, const QColor& displayColor);
    // Alignment
    void hAlignChanged(HorizontalAlignment align);
    void vAlignChanged(VerticalAlignment align);
    // Number formatting
    void thousandSeparatorToggled();
    void numberFormatChanged(const QString& format);
    void dateFormatSelected(const QString& dateFormatId);
    void currencyFormatSelected(const QString& currencyCode);
    void accountingFormatSelected(const QString& currencyCode);
    void increaseDecimals();
    void decreaseDecimals();
    void formatCellsRequested();
    // Data
    void sortAscending();
    void sortDescending();
    void filterToggled();
    // Tables
    void tableStyleSelected(int themeIndex);
    // Borders
    void borderStyleSelected(const QString& borderType, const QColor& color, int width, int penStyle);
    // Merge cells
    void mergeCellsRequested();
    void unmergeCellsRequested();
    // Indent
    void increaseIndent();
    void decreaseIndent();
    // Text overflow
    void textOverflowChanged(TextOverflowMode mode);
    // Text rotation
    void textRotationChanged(int degrees);
    // Conditional formatting & validation
    void conditionalFormatRequested();
    void dataValidationRequested();
    // Insert chart/shape
    void insertChartRequested();
    void insertShapeRequested();
    // Picklist & Checkbox
    void insertPicklistRequested();
    void managePicklistsRequested();
    void insertCheckboxRequested();
    // Chat assistant
    void chatToggleRequested();

private:
    void createActions();
    void updateBgColorIcon();

    QFontComboBox* m_fontCombo = nullptr;
    QSpinBox* m_fontSizeSpinBox = nullptr;
    QToolButton* m_saveBtn = nullptr;
    QColor m_lastFgColor = QColor("#000000");
    QString m_lastFgColorStr = "#000000";
    QColor m_lastBgColor = QColor("#34A853");
    QString m_lastBgColorStr = "#34A853";
    QColor m_lastBorderColor = QColor("#000000");
    QString m_lastBorderColorStr = "#000000";
    const DocumentTheme* m_docTheme = nullptr;
    int m_lastBorderWidth = 1; // 1=thin, 2=medium, 3=thick
    int m_lastBorderPenStyle = 0; // 0=solid, 1=dashed, 2=dotted

    // Format button refs for sync
    QToolButton* m_boldBtn = nullptr;
    QToolButton* m_italicBtn = nullptr;
    QToolButton* m_underlineBtn = nullptr;
    QToolButton* m_strikethroughBtn = nullptr;
    QToolButton* m_fgColorBtn = nullptr;
    QToolButton* m_bgColorBtn = nullptr;
    QToolButton* m_numberFormatBtn = nullptr;

    // Alignment button refs for checked state management
    QToolButton* m_alignLeftBtn = nullptr;
    QToolButton* m_alignCenterBtn = nullptr;
    QToolButton* m_alignRightBtn = nullptr;
    QToolButton* m_vAlignTopBtn = nullptr;
    QToolButton* m_vAlignMiddleBtn = nullptr;
    QToolButton* m_vAlignBottomBtn = nullptr;
    QToolBar* m_secondaryToolbar = nullptr;
    QToolButton* m_chatBtn = nullptr;
    QToolButton* m_textOverflowBtn = nullptr;
};

#endif // TOOLBAR_H
