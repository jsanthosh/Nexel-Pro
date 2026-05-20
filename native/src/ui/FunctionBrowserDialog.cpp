#include "FunctionBrowserDialog.h"
#include "../core/FormulaMetadata.h"
#include "Theme.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QSplitter>
#include <QLineEdit>
#include <QListWidget>
#include <QLabel>
#include <QTextBrowser>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QKeyEvent>

FunctionBrowserDialog::FunctionBrowserDialog(QWidget* parent) : QDialog(parent) {
    setWindowTitle("Insert Function");
    resize(720, 480);
    setWindowFlags(windowFlags() & ~Qt::WindowContextHelpButtonHint);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(16, 16, 16, 16);
    mainLayout->setSpacing(10);

    // Search field.
    m_filterEdit = new QLineEdit(this);
    m_filterEdit->setPlaceholderText("Search functions (e.g., SUM, VLOOKUP, LET)…");
    m_filterEdit->setClearButtonEnabled(true);
    mainLayout->addWidget(m_filterEdit);

    // Splitter: function list on the left, details panel on the right.
    auto* splitter = new QSplitter(Qt::Horizontal, this);
    splitter->setChildrenCollapsible(false);

    m_funcList = new QListWidget(splitter);
    m_funcList->setSortingEnabled(false);
    m_funcList->setUniformItemSizes(true);
    m_funcList->setMinimumWidth(220);

    auto* detailsBox = new QWidget(splitter);
    auto* detailsLayout = new QVBoxLayout(detailsBox);
    detailsLayout->setContentsMargins(0, 0, 0, 0);
    detailsLayout->setSpacing(6);

    m_syntaxLabel = new QLabel(detailsBox);
    m_syntaxLabel->setTextFormat(Qt::RichText);
    m_syntaxLabel->setWordWrap(true);
    m_syntaxLabel->setStyleSheet("font-family: 'SF Mono', 'Menlo', monospace; "
                                  "font-size: 13px; padding: 6px 10px; "
                                  "background: #F0F2F5; border-radius: 4px;");
    detailsLayout->addWidget(m_syntaxLabel);

    m_detailsBrowser = new QTextBrowser(detailsBox);
    m_detailsBrowser->setOpenExternalLinks(false);
    m_detailsBrowser->setStyleSheet("QTextBrowser { border: 1px solid #D0D5DD; "
                                     "border-radius: 4px; padding: 6px; }");
    detailsLayout->addWidget(m_detailsBrowser, /*stretch*/1);

    splitter->addWidget(m_funcList);
    splitter->addWidget(detailsBox);
    splitter->setStretchFactor(0, 0);
    splitter->setStretchFactor(1, 1);
    mainLayout->addWidget(splitter, 1);

    // Buttons.
    auto* btnBox = new QDialogButtonBox(QDialogButtonBox::Cancel, this);
    auto* insertBtn = new QPushButton("&Insert", this);
    insertBtn->setDefault(true);
    btnBox->addButton(insertBtn, QDialogButtonBox::AcceptRole);
    mainLayout->addWidget(btnBox);

    connect(btnBox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(btnBox, &QDialogButtonBox::rejected, this, &QDialog::reject);
    connect(m_filterEdit, &QLineEdit::textChanged,
            this, &FunctionBrowserDialog::onFilterChanged);
    connect(m_funcList, &QListWidget::currentRowChanged,
            this, [this](int) { onSelectionChanged(); });
    connect(m_funcList, &QListWidget::itemActivated,
            this, &FunctionBrowserDialog::onItemActivated);

    populateList();
    if (m_funcList->count() > 0) {
        m_funcList->setCurrentRow(0);
    }

    setStyleSheet(ThemeManager::instance().dialogStylesheet());

    // Keyboard convenience: Enter in the filter activates the first match.
    m_filterEdit->installEventFilter(this);
}

void FunctionBrowserDialog::populateList() {
    m_funcList->clear();
    const auto& reg = formulaRegistry();
    QStringList names = reg.keys();
    std::sort(names.begin(), names.end());
    for (const QString& name : names) {
        m_funcList->addItem(name);
    }
}

void FunctionBrowserDialog::onFilterChanged(const QString& text) {
    const QString needle = text.trimmed();
    int firstVisible = -1;
    for (int i = 0; i < m_funcList->count(); ++i) {
        auto* item = m_funcList->item(i);
        const bool hide = !needle.isEmpty()
                          && !item->text().contains(needle, Qt::CaseInsensitive);
        item->setHidden(hide);
        if (!hide && firstVisible < 0) firstVisible = i;
    }
    if (firstVisible >= 0) m_funcList->setCurrentRow(firstVisible);
}

void FunctionBrowserDialog::onSelectionChanged() {
    auto* item = m_funcList->currentItem();
    if (!item) {
        m_syntaxLabel->clear();
        m_detailsBrowser->clear();
        m_selected.clear();
        return;
    }
    m_selected = item->text();
    updateDetails(m_selected);
}

void FunctionBrowserDialog::onItemActivated(QListWidgetItem*) {
    accept();
}

void FunctionBrowserDialog::updateDetails(const QString& funcName) {
    const auto& reg = formulaRegistry();
    auto it = reg.find(funcName);
    if (it == reg.end()) {
        m_syntaxLabel->setText(QString("<b>%1</b>").arg(funcName.toHtmlEscaped()));
        m_detailsBrowser->setHtml("<i>No metadata for this function.</i>");
        return;
    }
    const FormulaFuncInfo& info = it.value();
    m_syntaxLabel->setText("<b>" + info.syntax.toHtmlEscaped() + "</b>");

    QString html;
    html += "<p>" + info.description.toHtmlEscaped() + "</p>";
    if (!info.params.isEmpty()) {
        html += "<h4 style='margin-top:12px;margin-bottom:4px;'>Arguments</h4>";
        html += "<table cellspacing='0' cellpadding='4' style='border-collapse:collapse;'>";
        for (const auto& p : info.params) {
            QString tag = p.optional ? " <i>(optional)</i>" : "";
            html += QString("<tr><td valign='top' style='font-family:monospace;color:#1E40AF;'>%1%2</td>"
                            "<td valign='top'>%3</td></tr>")
                    .arg(p.name.toHtmlEscaped(), tag, p.description.toHtmlEscaped());
        }
        html += "</table>";
    }
    m_detailsBrowser->setHtml(html);
}
