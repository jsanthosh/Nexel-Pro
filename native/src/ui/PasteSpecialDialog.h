#ifndef PASTESPECIALDIALOG_H
#define PASTESPECIALDIALOG_H

#include <QDialog>
#include <QRadioButton>
#include <QCheckBox>

struct PasteSpecialOptions {
    enum PasteType { All, Values, Formulas, Formats, ColumnWidths } pasteType = All;
    enum Operation { OpNone, Add, Subtract, Multiply, Divide } operation = OpNone;
    bool skipBlanks = false;
    bool transpose = false;
};

class PasteSpecialDialog : public QDialog {
    Q_OBJECT

public:
    explicit PasteSpecialDialog(QWidget* parent = nullptr);

    PasteSpecialOptions getOptions() const;

private:
    // Paste type
    QRadioButton* m_pasteAll;
    QRadioButton* m_pasteValues;
    QRadioButton* m_pasteFormulas;
    QRadioButton* m_pasteFormats;
    QRadioButton* m_pasteColumnWidths;

    // Operation
    QRadioButton* m_opNone;
    QRadioButton* m_opAdd;
    QRadioButton* m_opSubtract;
    QRadioButton* m_opMultiply;
    QRadioButton* m_opDivide;

    // Options
    QCheckBox* m_skipBlanks;
    QCheckBox* m_transpose;
};

#endif // PASTESPECIALDIALOG_H
