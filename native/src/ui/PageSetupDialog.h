#ifndef PAGESETUPDIALOG_H
#define PAGESETUPDIALOG_H

#include <QDialog>
#include "../core/Spreadsheet.h"   // for PrintSettings

class QDoubleSpinBox;
class QSpinBox;
class QComboBox;
class QLineEdit;
class QCheckBox;
class QRadioButton;

// Excel-style File > Page Setup dialog. Tabbed UI (Page / Margins / Header
// & Footer / Sheet) editing the per-Spreadsheet PrintSettings struct that
// already round-trips through XLSX via M1 Week 5b.
class PageSetupDialog : public QDialog {
    Q_OBJECT
public:
    explicit PageSetupDialog(const Spreadsheet::PrintSettings& initial,
                              QWidget* parent = nullptr);

    Spreadsheet::PrintSettings settings() const;

private:
    void buildPageTab();
    void buildMarginsTab();
    void buildHeaderFooterTab();
    void buildSheetTab();

    // Page tab
    QRadioButton*   m_portraitRadio;
    QRadioButton*   m_landscapeRadio;
    QComboBox*      m_paperSizeCombo;
    QSpinBox*       m_scaleSpin;
    QSpinBox*       m_fitWidthSpin;
    QSpinBox*       m_fitHeightSpin;
    QRadioButton*   m_scaleRadio;
    QRadioButton*   m_fitRadio;

    // Margins tab (inches)
    QDoubleSpinBox* m_topMargin;
    QDoubleSpinBox* m_bottomMargin;
    QDoubleSpinBox* m_leftMargin;
    QDoubleSpinBox* m_rightMargin;
    QDoubleSpinBox* m_headerMargin;
    QDoubleSpinBox* m_footerMargin;

    // Header / Footer tab
    QLineEdit*      m_oddHeaderEdit;
    QLineEdit*      m_oddFooterEdit;

    // Sheet tab
    QCheckBox*      m_printGridlinesCheck;
    QCheckBox*      m_printHeadingsCheck;
};

#endif // PAGESETUPDIALOG_H
