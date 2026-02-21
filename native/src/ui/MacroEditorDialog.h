#ifndef MACROEDITORDIALOG_H
#define MACROEDITORDIALOG_H

#include <QDialog>
#include <QSyntaxHighlighter>
#include <QRegularExpression>
#include "../core/MacroEngine.h"

class QPlainTextEdit;
class QListWidget;
class QListWidgetItem;
class QPushButton;
class QLineEdit;
class QTextCharFormat;

class JsSyntaxHighlighter : public QSyntaxHighlighter {
    Q_OBJECT
public:
    explicit JsSyntaxHighlighter(QTextDocument* parent = nullptr);
protected:
    void highlightBlock(const QString& text) override;
private:
    struct HighlightRule {
        QRegularExpression pattern;
        QTextCharFormat format;
    };
    QVector<HighlightRule> m_rules;
};

class MacroEditorDialog : public QDialog {
    Q_OBJECT
public:
    explicit MacroEditorDialog(MacroEngine* engine, QWidget* parent = nullptr);

private slots:
    void onRun();
    void onSave();
    void onDelete();
    void onMacroSelected(QListWidgetItem* item);
    void onRecord();
    void onLogMessage(const QString& msg);

private:
    void createLayout();
    void refreshMacroList();

    MacroEngine* m_engine;
    QListWidget* m_macroList = nullptr;
    QLineEdit* m_nameEdit = nullptr;
    QPlainTextEdit* m_codeEdit = nullptr;
    QPlainTextEdit* m_outputEdit = nullptr;
    QPushButton* m_runBtn = nullptr;
    QPushButton* m_saveBtn = nullptr;
    QPushButton* m_deleteBtn = nullptr;
    QPushButton* m_recordBtn = nullptr;
    JsSyntaxHighlighter* m_highlighter = nullptr;
};

#endif // MACROEDITORDIALOG_H
