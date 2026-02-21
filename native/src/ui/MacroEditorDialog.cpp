#include "MacroEditorDialog.h"

#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QSplitter>
#include <QPlainTextEdit>
#include <QListWidget>
#include <QPushButton>
#include <QLineEdit>
#include <QLabel>
#include <QMessageBox>
#include <QRegularExpression>
#include <QTextCharFormat>
#include <QFont>

// ---------------------------------------------------------------------------
// JsSyntaxHighlighter
// ---------------------------------------------------------------------------

JsSyntaxHighlighter::JsSyntaxHighlighter(QTextDocument* parent)
    : QSyntaxHighlighter(parent)
{
    // --- Keywords (blue, bold) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#1D4ED8"));
        fmt.setFontWeight(QFont::Bold);

        const QStringList keywords = {
            "var", "let", "const", "function", "return",
            "if", "else", "for", "while", "do",
            "switch", "case", "break", "continue", "new",
            "this", "true", "false", "null", "undefined"
        };
        for (const QString& kw : keywords) {
            HighlightRule rule;
            rule.pattern = QRegularExpression(QStringLiteral("\\b%1\\b").arg(kw));
            rule.format  = fmt;
            m_rules.append(rule);
        }
    }

    // --- Numbers (dark cyan) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#0E7490"));

        HighlightRule rule;
        rule.pattern = QRegularExpression(QStringLiteral("\\b[0-9]+\\.?[0-9]*\\b"));
        rule.format  = fmt;
        m_rules.append(rule);
    }

    // --- Sheet API calls (dark green) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#15803D"));

        const QStringList apiPatterns = {
            QStringLiteral("sheet\\.\\w+"),
            QStringLiteral("\\bgetCellValue\\b"),
            QStringLiteral("\\bsetCellValue\\b"),
            QStringLiteral("\\bsetCellFormula\\b"),
            QStringLiteral("\\balert\\b"),
            QStringLiteral("\\blog\\b")
        };
        for (const QString& p : apiPatterns) {
            HighlightRule rule;
            rule.pattern = QRegularExpression(p);
            rule.format  = fmt;
            m_rules.append(rule);
        }
    }

    // --- Strings – double-quoted (dark red) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#B91C1C"));

        HighlightRule rule;
        rule.pattern = QRegularExpression(QStringLiteral("\"[^\"]*\""));
        rule.format  = fmt;
        m_rules.append(rule);
    }

    // --- Strings – single-quoted (dark red) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#B91C1C"));

        HighlightRule rule;
        rule.pattern = QRegularExpression(QStringLiteral("'[^']*'"));
        rule.format  = fmt;
        m_rules.append(rule);
    }

    // --- Single-line comments (gray, italic) ---
    {
        QTextCharFormat fmt;
        fmt.setForeground(QColor("#6B7280"));
        fmt.setFontItalic(true);

        HighlightRule rule;
        rule.pattern = QRegularExpression(QStringLiteral("//[^\n]*"));
        rule.format  = fmt;
        m_rules.append(rule);
    }
}

void JsSyntaxHighlighter::highlightBlock(const QString& text)
{
    for (const HighlightRule& rule : std::as_const(m_rules)) {
        QRegularExpressionMatchIterator it = rule.pattern.globalMatch(text);
        while (it.hasNext()) {
            QRegularExpressionMatch match = it.next();
            setFormat(static_cast<int>(match.capturedStart()),
                      static_cast<int>(match.capturedLength()),
                      rule.format);
        }
    }
}

// ---------------------------------------------------------------------------
// MacroEditorDialog
// ---------------------------------------------------------------------------

MacroEditorDialog::MacroEditorDialog(MacroEngine* engine, QWidget* parent)
    : QDialog(parent)
    , m_engine(engine)
{
    setWindowTitle(tr("Macro Editor"));
    resize(800, 550);

    setStyleSheet(
        "QDialog { background: #F8FAFC; }"
        "QListWidget { border: 1px solid #D0D5DD; border-radius: 6px; background: white; }"
        "QLineEdit { padding: 6px; border: 1px solid #D0D5DD; border-radius: 6px; }"
        "QPlainTextEdit { border: 1px solid #D0D5DD; border-radius: 6px; "
        "  font-family: 'SF Mono', 'Menlo', 'Monaco', monospace; font-size: 13px; }"
        "QPushButton { padding: 8px 16px; border-radius: 6px; font-size: 13px; font-weight: 500; }"
    );

    createLayout();
    refreshMacroList();

    // Connect engine log signal to the output panel
    connect(m_engine, &MacroEngine::logMessage, this, &MacroEditorDialog::onLogMessage);
}

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

void MacroEditorDialog::createLayout()
{
    auto* root = new QHBoxLayout(this);
    root->setContentsMargins(12, 12, 12, 12);
    root->setSpacing(10);

    // ---- Left panel (macro list + name + save/delete) ----
    auto* leftPanel = new QVBoxLayout;
    leftPanel->setSpacing(6);

    auto* listLabel = new QLabel(tr("Saved Macros"));
    listLabel->setStyleSheet("font-weight: 600; font-size: 13px; color: #374151;");
    leftPanel->addWidget(listLabel);

    m_macroList = new QListWidget;
    m_macroList->setFixedWidth(200);
    connect(m_macroList, &QListWidget::itemClicked,
            this, &MacroEditorDialog::onMacroSelected);
    leftPanel->addWidget(m_macroList, 1);

    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(tr("Macro name..."));
    leftPanel->addWidget(m_nameEdit);

    auto* leftBtnLayout = new QHBoxLayout;
    m_saveBtn = new QPushButton(tr("Save"));
    m_saveBtn->setStyleSheet(
        "QPushButton { background: #E5E7EB; color: #374151; border: 1px solid #D0D5DD; }"
        "QPushButton:hover { background: #D1D5DB; }"
    );
    connect(m_saveBtn, &QPushButton::clicked, this, &MacroEditorDialog::onSave);
    leftBtnLayout->addWidget(m_saveBtn);

    m_deleteBtn = new QPushButton(tr("Delete"));
    m_deleteBtn->setStyleSheet(
        "QPushButton { background: #E5E7EB; color: #374151; border: 1px solid #D0D5DD; }"
        "QPushButton:hover { background: #D1D5DB; }"
    );
    connect(m_deleteBtn, &QPushButton::clicked, this, &MacroEditorDialog::onDelete);
    leftBtnLayout->addWidget(m_deleteBtn);

    leftPanel->addLayout(leftBtnLayout);

    root->addLayout(leftPanel);

    // ---- Centre + bottom (code editor, action buttons, output) ----
    auto* rightPanel = new QVBoxLayout;
    rightPanel->setSpacing(8);

    // Top toolbar row: Run and Record
    auto* toolbar = new QHBoxLayout;
    toolbar->addStretch();

    m_recordBtn = new QPushButton(tr("Record"));
    m_recordBtn->setStyleSheet(
        "QPushButton { background: #E5E7EB; color: #374151; border: 1px solid #D0D5DD; }"
        "QPushButton:hover { background: #D1D5DB; }"
    );
    connect(m_recordBtn, &QPushButton::clicked, this, &MacroEditorDialog::onRecord);
    toolbar->addWidget(m_recordBtn);

    m_runBtn = new QPushButton(tr("Run"));
    m_runBtn->setStyleSheet(
        "QPushButton { background: #16A34A; color: white; border: none; }"
        "QPushButton:hover { background: #15803D; }"
    );
    connect(m_runBtn, &QPushButton::clicked, this, &MacroEditorDialog::onRun);
    toolbar->addWidget(m_runBtn);

    rightPanel->addLayout(toolbar);

    // Code editor
    m_codeEdit = new QPlainTextEdit;
    m_codeEdit->setPlaceholderText(tr("// Write your macro here...\n"
                                      "// Use sheet.getCellValue(row, col)\n"
                                      "//     sheet.setCellValue(row, col, value)\n"
                                      "//     sheet.setCellFormula(row, col, formula)"));
    m_highlighter = new JsSyntaxHighlighter(m_codeEdit->document());
    rightPanel->addWidget(m_codeEdit, 1);

    // Output / log panel
    auto* outputLabel = new QLabel(tr("Output"));
    outputLabel->setStyleSheet("font-weight: 600; font-size: 12px; color: #6B7280;");
    rightPanel->addWidget(outputLabel);

    m_outputEdit = new QPlainTextEdit;
    m_outputEdit->setReadOnly(true);
    m_outputEdit->setFixedHeight(100);
    m_outputEdit->setStyleSheet(
        "QPlainTextEdit { background: #F1F5F9; color: #334155; }"
    );
    rightPanel->addWidget(m_outputEdit);

    root->addLayout(rightPanel, 1);
}

// ---------------------------------------------------------------------------
// Slots
// ---------------------------------------------------------------------------

void MacroEditorDialog::onRun()
{
    m_outputEdit->clear();
    const QString code = m_codeEdit->toPlainText().trimmed();
    if (code.isEmpty()) {
        m_outputEdit->appendPlainText(tr("[Error] No code to execute."));
        return;
    }

    m_outputEdit->appendPlainText(tr("--- Running macro ---"));
    auto result = m_engine->execute(code);
    if (result.success) {
        if (!result.output.isEmpty())
            m_outputEdit->appendPlainText(result.output);
        m_outputEdit->appendPlainText(tr("--- Done ---"));
    } else {
        m_outputEdit->appendPlainText(tr("[Error] %1").arg(result.error));
    }
}

void MacroEditorDialog::onSave()
{
    const QString name = m_nameEdit->text().trimmed();
    if (name.isEmpty()) {
        QMessageBox::warning(this, tr("Save Macro"),
                             tr("Please enter a name for the macro."));
        return;
    }
    const QString code = m_codeEdit->toPlainText();
    SavedMacro macro;
    macro.name = name;
    macro.code = code;
    m_engine->saveMacro(macro);
    refreshMacroList();
    m_outputEdit->appendPlainText(tr("Macro \"%1\" saved.").arg(name));
}

void MacroEditorDialog::onDelete()
{
    QListWidgetItem* item = m_macroList->currentItem();
    if (!item) {
        QMessageBox::information(this, tr("Delete Macro"),
                                 tr("Select a macro to delete."));
        return;
    }

    const QString name = item->text();
    if (QMessageBox::question(this, tr("Delete Macro"),
            tr("Delete macro \"%1\"?").arg(name),
            QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        m_engine->deleteMacro(name);
        m_nameEdit->clear();
        m_codeEdit->clear();
        refreshMacroList();
        m_outputEdit->appendPlainText(tr("Macro \"%1\" deleted.").arg(name));
    }
}

void MacroEditorDialog::onMacroSelected(QListWidgetItem* item)
{
    if (!item) return;
    const QString name = item->text();
    m_nameEdit->setText(name);
    // Find macro code by name
    for (const auto& m : m_engine->getSavedMacros()) {
        if (m.name == name) {
            m_codeEdit->setPlainText(m.code);
            break;
        }
    }
}

void MacroEditorDialog::onRecord()
{
    if (m_engine->isRecording()) {
        // Stop recording and load the captured code
        m_engine->stopRecording();
        m_codeEdit->setPlainText(m_engine->getRecordedCode());
        m_recordBtn->setText(tr("Record"));
        m_recordBtn->setStyleSheet(
            "QPushButton { background: #E5E7EB; color: #374151; border: 1px solid #D0D5DD; }"
            "QPushButton:hover { background: #D1D5DB; }"
        );
        m_outputEdit->appendPlainText(tr("Recording stopped. Code loaded into editor."));
    } else {
        // Start recording
        m_engine->startRecording();
        m_recordBtn->setText(tr("Stop Recording"));
        m_recordBtn->setStyleSheet(
            "QPushButton { background: #DC2626; color: white; border: none; }"
            "QPushButton:hover { background: #B91C1C; }"
        );
        m_outputEdit->appendPlainText(tr("Recording started. Perform actions in the spreadsheet..."));
    }
}

void MacroEditorDialog::onLogMessage(const QString& msg)
{
    m_outputEdit->appendPlainText(msg);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

void MacroEditorDialog::refreshMacroList()
{
    m_macroList->clear();
    for (const auto& macro : m_engine->getSavedMacros()) {
        m_macroList->addItem(macro.name);
    }
}
