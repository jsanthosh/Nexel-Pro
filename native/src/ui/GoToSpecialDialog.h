#ifndef GOTOSPECIALDIALOG_H
#define GOTOSPECIALDIALOG_H

#include <QDialog>
#include <QRadioButton>
#include <QCheckBox>
#include <QButtonGroup>

class GoToSpecialDialog : public QDialog {
    Q_OBJECT

public:
    enum SelectionType {
        Blanks,
        Constants,
        Formulas,
        Comments,
        ConditionalFormats,
        DataValidation,
        VisibleCells,
        CurrentRegion
    };

    explicit GoToSpecialDialog(QWidget* parent = nullptr);

    SelectionType getSelectionType() const;
    bool includeNumbers() const;
    bool includeText() const;
    bool includeLogicals() const;
    bool includeErrors() const;

private:
    QButtonGroup* m_typeGroup;
    QRadioButton* m_blanksRadio;
    QRadioButton* m_constantsRadio;
    QRadioButton* m_formulasRadio;
    QRadioButton* m_commentsRadio;
    QRadioButton* m_conditionalFormatsRadio;
    QRadioButton* m_dataValidationRadio;
    QRadioButton* m_visibleCellsRadio;
    QRadioButton* m_currentRegionRadio;

    QCheckBox* m_numbersCheck;
    QCheckBox* m_textCheck;
    QCheckBox* m_logicalsCheck;
    QCheckBox* m_errorsCheck;

    void updateSubCheckboxes();
};

#endif // GOTOSPECIALDIALOG_H
