#include "FormulaBar.h"
#include "FormulaPopupDelegate.h"
#include "Theme.h"
#include "../core/FormulaMetadata.h"
#include <QHBoxLayout>
#include <QLabel>
#include <QKeyEvent>
#include <QApplication>

FormulaBar::FormulaBar(QWidget* parent)
    : QWidget(parent) {

    QHBoxLayout* layout = new QHBoxLayout(this);
    layout->setContentsMargins(5, 5, 5, 5);

    // Cell address label
    m_cellAddressLabel = new QLabel("A1", this);
    m_cellAddressLabel->setMinimumWidth(50);
    m_cellAddressLabel->setStyleSheet(QString("border: 1px solid %1; padding: 2px;")
        .arg(ThemeManager::instance().currentTheme().formulaInputBorder.name()));
    layout->addWidget(m_cellAddressLabel);

    // Formula/Content input
    m_formulaEdit = new QLineEdit(this);
    m_formulaEdit->setPlaceholderText("Enter formula or value...");
    layout->addWidget(m_formulaEdit);

    // Setup connections
    connect(m_formulaEdit, &QLineEdit::textChanged, this, &FormulaBar::onTextChanged);
    connect(m_formulaEdit, &QLineEdit::textEdited, this, &FormulaBar::onTextEdited);
    connect(m_formulaEdit, &QLineEdit::returnPressed, this, &FormulaBar::returnPressed);

    applyThemeStyle();

    setupAutocomplete();
}

void FormulaBar::setCellAddress(const QString& address) {
    m_cellAddressLabel->setText(address);
}

void FormulaBar::setCellContent(const QString& content) {
    hideAllPanels();
    m_formulaEdit->blockSignals(true);
    m_formulaEdit->setText(content);
    m_formulaEdit->blockSignals(false);
}

QString FormulaBar::getContent() const {
    return m_formulaEdit->text();
}

void FormulaBar::onTextChanged(const QString& text) {
    emit contentChanged(text);
}

void FormulaBar::onTextEdited(const QString& text) {
    emit contentEdited(text);
    emit formulaEditModeChanged(m_formulaEdit->hasFocus() && text.startsWith("="));
}

bool FormulaBar::isFormulaEditing() const {
    return m_formulaEdit->hasFocus() && m_formulaEdit->text().startsWith("=");
}

void FormulaBar::insertText(const QString& text) {
    m_lastInsertPos = m_formulaEdit->cursorPosition();
    m_formulaEdit->insert(text);
    m_lastInsertLen = text.length();
}

void FormulaBar::replaceLastInsertedText(const QString& newText) {
    if (m_lastInsertPos >= 0) {
        m_formulaEdit->setSelection(m_lastInsertPos, m_lastInsertLen);
        m_formulaEdit->insert(newText);
        m_lastInsertLen = newText.length();
    }
}

void FormulaBar::hideAllPanels() {
    if (m_popup) m_popup->hide();
    if (m_paramHint) m_paramHint->hide();
    if (m_detailPanel) m_detailPanel->hide();
}

void FormulaBar::setupAutocomplete() {
    m_popup = new QListWidget(window());
    m_popup->setWindowFlags(Qt::Popup | Qt::FramelessWindowHint);
    m_popup->setAttribute(Qt::WA_ShowWithoutActivating);
    m_popup->setFocusPolicy(Qt::NoFocus);
    m_popup->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_popup->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_popup->setStyleSheet(QString(
        "QListWidget { background: white; border: 1px solid %1; outline: none; "
        "border-radius: 6px; }"
        "QListWidget::item { padding: 0px; border: none; }"
        "QListWidget::item:selected { background: transparent; }"
        "QListWidget::item:hover { background: transparent; }"
    ).arg(ThemeManager::instance().currentTheme().popupBorder.name()));
    m_popup->setItemDelegate(new FormulaPopupDelegate(m_popup));
    m_popup->hide();

    m_paramHint = new QLabel(window());
    m_paramHint->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    m_paramHint->setAttribute(Qt::WA_ShowWithoutActivating);
    {
        const auto& t = ThemeManager::instance().currentTheme();
        m_paramHint->setStyleSheet(QString(
            "QLabel { background: %1; border: 1px solid %2; padding: 4px 8px; "
            "font-size: 12px; color: %3; border-radius: 3px; }")
            .arg(t.paramHintBackground.name(), t.paramHintBorder.name(), t.textPrimary.name()));
    }
    m_paramHint->setTextFormat(Qt::RichText);
    m_paramHint->hide();

    m_detailPanel = new QLabel(window());
    m_detailPanel->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    m_detailPanel->setAttribute(Qt::WA_ShowWithoutActivating);
    m_detailPanel->setStyleSheet(QString(
        "QLabel { background: white; border: 1px solid %1; padding: 12px; "
        "border-radius: 6px; }").arg(ThemeManager::instance().currentTheme().formulaInputBorder.name()));
    m_detailPanel->setTextFormat(Qt::RichText);
    m_detailPanel->setWordWrap(true);
    m_detailPanel->setFixedWidth(340);
    m_detailPanel->hide();

    // Click on popup item → show detail panel
    connect(m_popup, &QListWidget::itemClicked, this, [this](QListWidgetItem* item) {
        QString funcName = item->data(FuncNameRole).toString();
        const auto& reg = formulaRegistry();
        if (reg.contains(funcName)) {
            m_detailPanel->setText(buildDetailHtml(reg[funcName]));
            m_detailPanel->adjustSize();
            QPoint pos = m_popup->mapToGlobal(QPoint(m_popup->width() + 4, 0));
            m_detailPanel->move(pos);
            m_detailPanel->show();
        }
    });

    connect(m_formulaEdit, &QLineEdit::textEdited, this, [this]() {
        m_detailPanel->hide();
        updatePopup();
        updateParamHint();
    });

    connect(m_formulaEdit, &QLineEdit::cursorPositionChanged, this, [this]() {
        updateParamHint();
    });

    // Install event filter on BOTH popup (for key forwarding) and formula edit
    m_popup->installEventFilter(this);
    m_formulaEdit->installEventFilter(this);
}

bool FormulaBar::eventFilter(QObject* obj, QEvent* event) {
    // Handle key events from popup (Qt::Popup grabs keyboard)
    if (obj == m_popup && event->type() == QEvent::KeyPress) {
        auto* ke = static_cast<QKeyEvent*>(event);
        if (ke->key() == Qt::Key_Down) {
            int next = m_popup->currentRow() + 1;
            if (next < m_popup->count()) m_popup->setCurrentRow(next);
            return true;
        }
        if (ke->key() == Qt::Key_Up) {
            int prev = m_popup->currentRow() - 1;
            if (prev >= 0) m_popup->setCurrentRow(prev);
            return true;
        }
        if (ke->key() == Qt::Key_Return || ke->key() == Qt::Key_Enter || ke->key() == Qt::Key_Tab) {
            auto* item = m_popup->currentItem();
            if (item) {
                insertFunction(item->data(FuncNameRole).toString());
                hideAllPanels();
            }
            return true;
        }
        if (ke->key() == Qt::Key_Escape) {
            hideAllPanels();
            m_formulaEdit->setFocus();
            return true;
        }
        // Forward all other keys (typing) to formula edit
        QApplication::sendEvent(m_formulaEdit, event);
        return true;
    }
    // Handle Esc on formula edit (when popup is not visible)
    if (obj == m_formulaEdit && event->type() == QEvent::KeyPress) {
        auto* ke = static_cast<QKeyEvent*>(event);
        if (ke->key() == Qt::Key_Escape) {
            hideAllPanels();
            m_formulaEdit->clearFocus();
            return true;
        }
    }
    return QWidget::eventFilter(obj, event);
}

void FormulaBar::updatePopup() {
    QString text = m_formulaEdit->text();
    if (!text.startsWith("=") || text.length() <= 1) {
        m_popup->hide();
        return;
    }

    // Extract current token
    QString afterEq = text.mid(1);
    int lastDelim = -1;
    for (int i = afterEq.length() - 1; i >= 0; --i) {
        QChar ch = afterEq[i];
        if (ch == '(' || ch == ')' || ch == ',' || ch == '+' || ch == '-' ||
            ch == '*' || ch == '/' || ch == ':' || ch == ' ') {
            lastDelim = i;
            break;
        }
    }
    QString token = afterEq.mid(lastDelim + 1);
    if (token.isEmpty() || !token[0].isLetter()) {
        m_popup->hide();
        return;
    }

    m_popup->clear();
    const auto& reg = formulaRegistry();
    for (auto it = reg.begin(); it != reg.end(); ++it) {
        if (it->name.startsWith(token, Qt::CaseInsensitive)) {
            auto* item = new QListWidgetItem(m_popup);
            item->setData(FuncNameRole, it->name);
            item->setData(FuncDescRole, it->description);
        }
    }

    if (m_popup->count() > 0) {
        QPoint pos = m_formulaEdit->mapToGlobal(QPoint(0, m_formulaEdit->height()));
        m_popup->setFixedWidth(qMax(460, m_formulaEdit->width()));
        int visibleItems = qMin(m_popup->count(), 8);
        m_popup->setFixedHeight(visibleItems * 30 + 6);
        m_popup->move(pos);
        m_popup->show();
        m_popup->setCurrentRow(0);
    } else {
        m_popup->hide();
    }
}

void FormulaBar::updateParamHint() {
    QString text = m_formulaEdit->text();
    if (!text.startsWith("=")) {
        m_paramHint->hide();
        return;
    }

    int cursorPos = m_formulaEdit->cursorPosition();
    FormulaContext ctx = findFormulaContext(text, cursorPos);
    const auto& reg = formulaRegistry();

    if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
        const auto& info = reg[ctx.funcName];
        m_paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
        m_paramHint->adjustSize();
        QPoint hintPos = m_formulaEdit->mapToGlobal(QPoint(0, m_formulaEdit->height() + 2));
        if (m_popup && m_popup->isVisible()) {
            hintPos.setY(m_popup->mapToGlobal(QPoint(0, m_popup->height())).y() + 2);
        }
        m_paramHint->move(hintPos);
        m_paramHint->show();
    } else {
        m_paramHint->hide();
    }
}

void FormulaBar::insertFunction(const QString& funcName) {
    m_popup->hide();
    m_detailPanel->hide();
    QString text = m_formulaEdit->text();
    // Find current token to replace
    QString afterEq = text.mid(1);
    int lastDelim = -1;
    for (int i = afterEq.length() - 1; i >= 0; --i) {
        QChar ch = afterEq[i];
        if (ch == '(' || ch == ')' || ch == ',' || ch == '+' || ch == '-' ||
            ch == '*' || ch == '/' || ch == ':' || ch == ' ') {
            lastDelim = i;
            break;
        }
    }
    int tokenStart = 1 + lastDelim + 1; // position in full text
    m_formulaEdit->setText(text.left(tokenStart) + funcName + "(");
    m_formulaEdit->setCursorPosition(m_formulaEdit->text().length());
    m_formulaEdit->setFocus();
}

void FormulaBar::applyThemeStyle() {
    const auto& t = ThemeManager::instance().currentTheme();
    setStyleSheet(QString(
        "QWidget {"
        "   background-color: %1;"
        "   border-bottom: 1px solid %2;"
        "}"
        "QLineEdit {"
        "   border: 1px solid %3;"
        "   padding: 3px;"
        "   border-radius: 3px;"
        "}"
    ).arg(t.formulaBarBackground.name(), t.formulaBarBorder.name(), t.formulaInputBorder.name()));
}

void FormulaBar::onThemeChanged() {
    applyThemeStyle();
}
