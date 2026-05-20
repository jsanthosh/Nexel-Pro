#ifndef FUNCTIONBROWSERDIALOG_H
#define FUNCTIONBROWSERDIALOG_H

#include <QDialog>
#include <QString>

class QLineEdit;
class QListWidget;
class QLabel;
class QTextBrowser;
class QListWidgetItem;

// Excel-style fx browser. Lists every function registered in
// FormulaMetadata, filterable by name, with a details panel showing syntax,
// description, and per-parameter docs. On accept exposes the selected name
// via selectedFunction() so the caller can pipe it into FormulaBar.
class FunctionBrowserDialog : public QDialog {
    Q_OBJECT
public:
    explicit FunctionBrowserDialog(QWidget* parent = nullptr);

    QString selectedFunction() const { return m_selected; }

private slots:
    void onFilterChanged(const QString& text);
    void onSelectionChanged();
    void onItemActivated(QListWidgetItem* item);

private:
    void populateList();
    void updateDetails(const QString& funcName);

    QLineEdit*    m_filterEdit;
    QListWidget*  m_funcList;
    QLabel*       m_syntaxLabel;
    QTextBrowser* m_detailsBrowser;
    QString       m_selected;
};

#endif // FUNCTIONBROWSERDIALOG_H
