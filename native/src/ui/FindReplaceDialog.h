#ifndef FINDREPLACEDIALOG_H
#define FINDREPLACEDIALOG_H

#include <QDialog>
#include <QLineEdit>
#include <QCheckBox>
#include <QComboBox>
#include <QPushButton>
#include <QLabel>

class FindReplaceDialog : public QDialog {
    Q_OBJECT

public:
    explicit FindReplaceDialog(QWidget* parent = nullptr);

    QString findText() const;
    QString replaceText() const;
    bool matchCase() const;
    bool matchWholeCell() const;

    // Search scope: Sheet vs Workbook
    enum SearchScope { Sheet, Workbook };
    SearchScope searchScope() const;

    // Look in: Values vs Formulas
    enum LookIn { Values, Formulas };
    LookIn lookIn() const;

signals:
    void findNext();
    void findPrevious();
    void findAll();
    void replaceOne();
    void replaceAll();

private:
    QLineEdit* m_findEdit;
    QLineEdit* m_replaceEdit;
    QCheckBox* m_matchCaseCheck;
    QCheckBox* m_wholeCellCheck;
    QComboBox* m_scopeCombo;
    QComboBox* m_lookInCombo;
    QLabel* m_statusLabel;

public:
    void setStatus(const QString& text);
};

#endif // FINDREPLACEDIALOG_H
