#include "CellDelegate.h"
#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "FormulaPopupDelegate.h"
#include "Theme.h"
#include "../core/FormulaMetadata.h"
#include <QLineEdit>
#include <QPainter>
#include <QPainterPath>
#include <QStyleOptionViewItem>
#include <QApplication>
#include <QAbstractItemView>
#include <QListWidget>
#include <QKeyEvent>
#include <QTableView>
#include <QTimer>
#include <QLabel>
#include <QVBoxLayout>
#include <QFrame>
#include <QScrollBar>
#include <QScreen>
#include <cmath>
#include "../core/Spreadsheet.h"

// ===== Picklist tag color palettes (12 pastel bg + dark text pairs) =====
static const QColor s_tagBgColors[] = {
    QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
    QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
    QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
};
static const QColor s_tagTextColors[] = {
    QColor("#1E40AF"), QColor("#9D174D"), QColor("#5B21B6"), QColor("#065F46"),
    QColor("#92400E"), QColor("#9F1239"), QColor("#155E75"), QColor("#991B1B"),
    QColor("#374151"), QColor("#3F6212"), QColor("#3730A3"), QColor("#831843")
};
static constexpr int TAG_COLOR_COUNT = 12;

QColor CellDelegate::tagBgColor(int index) { return s_tagBgColors[index % TAG_COLOR_COUNT]; }
QColor CellDelegate::tagTextColor(int index) { return s_tagTextColors[index % TAG_COLOR_COUNT]; }

// Extract the token currently being typed (after last delimiter)
static QString extractCurrentToken(const QString& text) {
    if (!text.startsWith("=") || text.length() <= 1) return {};
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
    return afterEq.mid(lastDelim + 1);
}

CellDelegate::CellDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {
}

static void hideAllPopups(QObject* editorOrPopup) {
    // Retrieve editor from popup, or use directly
    QLineEdit* ed = qobject_cast<QLineEdit*>(editorOrPopup);
    if (!ed) ed = qobject_cast<QLineEdit*>(editorOrPopup->property("_editor").value<QObject*>());
    if (!ed) return;
    auto* p = qobject_cast<QListWidget*>(ed->property("_formulaPopup").value<QObject*>());
    auto* h = qobject_cast<QLabel*>(ed->property("_paramHint").value<QObject*>());
    auto* d = qobject_cast<QLabel*>(ed->property("_detailPanel").value<QObject*>());
    if (p) p->hide();
    if (h) h->hide();
    if (d) d->hide();
}

static void insertFunctionFromPopup(QListWidget* popup, QLineEdit* editor) {
    auto* item = popup->currentItem();
    if (!item) return;
    QString funcName = item->data(FuncNameRole).toString();
    hideAllPopups(editor);
    QString text = editor->text();
    QString token = extractCurrentToken(text);
    if (!token.isEmpty()) {
        int tokenStart = text.length() - token.length();
        editor->setText(text.left(tokenStart) + funcName + "(");
        editor->setCursorPosition(editor->text().length());
    }
    editor->setFocus();
}

bool CellDelegate::eventFilter(QObject* object, QEvent* event) {
    // When popup auto-closes (clicked outside), also hide tooltip panels
    if (auto* popupWidget = qobject_cast<QListWidget*>(object)) {
        if (event->type() == QEvent::Hide) {
            auto* ed = qobject_cast<QLineEdit*>(
                popupWidget->property("_editor").value<QObject*>());
            if (ed) {
                auto* h = qobject_cast<QLabel*>(ed->property("_paramHint").value<QObject*>());
                auto* d = qobject_cast<QLabel*>(ed->property("_detailPanel").value<QObject*>());
                if (h) h->hide();
                if (d) d->hide();
            }
        }
    }
    // Handle key events from the popup (Qt::Popup grabs keyboard)
    if (auto* popupWidget = qobject_cast<QListWidget*>(object)) {
        if (event->type() == QEvent::KeyPress) {
            auto* ke = static_cast<QKeyEvent*>(event);
            auto* ed = qobject_cast<QLineEdit*>(
                popupWidget->property("_editor").value<QObject*>());
            if (!ed) return false;

            if (ke->key() == Qt::Key_Down) {
                int next = popupWidget->currentRow() + 1;
                if (next < popupWidget->count()) popupWidget->setCurrentRow(next);
                return true;
            }
            if (ke->key() == Qt::Key_Up) {
                int prev = popupWidget->currentRow() - 1;
                if (prev >= 0) popupWidget->setCurrentRow(prev);
                return true;
            }
            if (ke->key() == Qt::Key_Return || ke->key() == Qt::Key_Enter ||
                ke->key() == Qt::Key_Tab) {
                insertFunctionFromPopup(popupWidget, ed);
                return true;
            }
            if (ke->key() == Qt::Key_Escape) {
                hideAllPopups(popupWidget);
                m_formulaEditMode = false;
                emit formulaEditModeChanged(false);
                emit closeEditor(ed, QAbstractItemDelegate::RevertModelCache);
                return true;
            }
            // Forward all other keys (typing) to the editor
            QApplication::sendEvent(ed, event);
            return true;
        }
        return false;
    }

    // During formula edit mode, block FocusOut from closing the editor
    if (m_formulaEditMode && event->type() == QEvent::FocusOut) {
        return true;
    }

    if (event->type() == QEvent::KeyPress) {
        QKeyEvent* keyEvent = static_cast<QKeyEvent*>(event);
        QLineEdit* editor = qobject_cast<QLineEdit*>(object);

        // Esc: cancel edit and revert value
        if (keyEvent->key() == Qt::Key_Escape && editor) {
            hideAllPopups(editor);
            m_formulaEditMode = false;
            emit formulaEditModeChanged(false);
            emit closeEditor(editor, QAbstractItemDelegate::RevertModelCache);
            return true;
        }

        // Arrow keys during editing: commit and move (like Excel)
        // But NOT during formula edit mode — arrows navigate in the formula text
        // Also, if the editor has selected text, let arrows deselect first (Excel behavior)
        if (!m_formulaEditMode &&
            (keyEvent->key() == Qt::Key_Up || keyEvent->key() == Qt::Key_Down ||
             keyEvent->key() == Qt::Key_Left || keyEvent->key() == Qt::Key_Right)) {
            if (editor) {
                // If editor has selected text, let the arrow key deselect it first
                if (editor->hasSelectedText())
                    return QStyledItemDelegate::eventFilter(object, event);

                // Left/Right: only commit if cursor is at boundary
                if (keyEvent->key() == Qt::Key_Left && editor->cursorPosition() > 0)
                    return QStyledItemDelegate::eventFilter(object, event);
                if (keyEvent->key() == Qt::Key_Right && editor->cursorPosition() < editor->text().length())
                    return QStyledItemDelegate::eventFilter(object, event);

                emit commitData(editor);
                emit closeEditor(editor, QAbstractItemDelegate::NoHint);

                // Navigate to adjacent cell after editor closes
                int key = keyEvent->key();
                QWidget* viewport = editor->parentWidget();
                QTableView* view = viewport ? qobject_cast<QTableView*>(viewport->parentWidget()) : nullptr;
                if (view) {
                    QTimer::singleShot(0, view, [view, key]() {
                        QModelIndex cur = view->currentIndex();
                        if (!cur.isValid() || !view->model()) return;
                        int row = cur.row(), col = cur.column();
                        if (key == Qt::Key_Up) row = qMax(0, row - 1);
                        else if (key == Qt::Key_Down) row = qMin(view->model()->rowCount() - 1, row + 1);
                        else if (key == Qt::Key_Left) col = qMax(0, col - 1);
                        else if (key == Qt::Key_Right) col = qMin(view->model()->columnCount() - 1, col + 1);
                        QModelIndex next = view->model()->index(row, col);
                        if (next.isValid()) {
                            view->setCurrentIndex(next);
                        }
                    });
                }
                return true; // consume the event
            }
        }
    }
    return QStyledItemDelegate::eventFilter(object, event);
}

QWidget* CellDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                                   const QModelIndex& index) const {
    // Checkbox cells: no editor — toggle is handled in SpreadsheetView::mousePressEvent
    if (m_spreadsheetView) {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) {
            auto cell = spreadsheet->getCellIfExists(index.row(), index.column());
            if (cell) {
                const auto& style = cell->getStyle();
                if (style.numberFormat == "Checkbox") return nullptr;
                // Picklist cells: handled by SpreadsheetView popup
                if (style.numberFormat == "Picklist") return nullptr;
            }
        }
    }

    const auto& edTheme = ThemeManager::instance().currentTheme();
    bool dark = ThemeManager::instance().isDarkTheme();
    QString edBg = dark ? "#2d2d2d" : "white";
    QString edText = dark ? "#d4d4d4" : "black";
    QString popBg = dark ? "#252526" : "white";
    QString popBorder = dark ? "#404040" : "#C0C0C0";

    QLineEdit* editor = new QLineEdit(parent);
    editor->setFrame(false);
    editor->setStyleSheet(QString(
        "QLineEdit { background: %1; color: %2; padding: 1px 2px; "
        "border: 2px solid %3; selection-background-color: #0078D4; }")
        .arg(edBg, edText, edTheme.editorBorderColor.name()));

    // Expand editor as text grows (like Excel)
    connect(editor, &QLineEdit::textChanged, this, [this, editor, index]() {
        if (!m_spreadsheetView) return;
        QStyleOptionViewItem opt;
        opt.rect = m_spreadsheetView->visualRect(index);
        opt.font = m_spreadsheetView->font();
        const_cast<CellDelegate*>(this)->updateEditorGeometry(editor, opt, index);
    });

    // --- Custom formula autocomplete popup ---
    auto* popup = new QListWidget(parent->window());
    popup->setWindowFlags(Qt::Popup | Qt::FramelessWindowHint);
    popup->setAttribute(Qt::WA_ShowWithoutActivating);
    popup->setFocusPolicy(Qt::NoFocus);
    popup->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    popup->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    popup->setStyleSheet(QString(
        "QListWidget { background: %1; border: 1px solid %2; outline: none; "
        "border-radius: 6px; }"
        "QListWidget::item { padding: 0px; border: none; }"
        "QListWidget::item:selected { background: transparent; }"
        "QListWidget::item:hover { background: transparent; }")
        .arg(popBg, popBorder));
    popup->setItemDelegate(new FormulaPopupDelegate(popup));
    popup->hide();

    // --- Parameter hint tooltip ---
    QString hintBg = dark ? "#333333" : "#FFF8DC";
    QString hintBorder = dark ? "#404040" : "#E0D8B0";
    QString hintText = dark ? "#d4d4d4" : "#333";

    auto* paramHint = new QLabel(parent->window());
    paramHint->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    paramHint->setAttribute(Qt::WA_ShowWithoutActivating);
    paramHint->setStyleSheet(QString(
        "QLabel { background: %1; border: 1px solid %2; padding: 4px 8px; "
        "font-size: 12px; color: %3; border-radius: 3px; }").arg(hintBg, hintBorder, hintText));
    paramHint->setTextFormat(Qt::RichText);
    paramHint->hide();

    // --- Detail panel (shown on click) ---
    auto* detailPanel = new QLabel(parent->window());
    detailPanel->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    detailPanel->setAttribute(Qt::WA_ShowWithoutActivating);
    detailPanel->setStyleSheet(QString(
        "QLabel { background: %1; border: 1px solid %2; padding: 12px; "
        "border-radius: 6px; color: %3; }").arg(popBg, popBorder, hintText));
    detailPanel->setTextFormat(Qt::RichText);
    detailPanel->setWordWrap(true);
    detailPanel->setFixedWidth(340);
    detailPanel->hide();

    // Populate popup items from registry (startsWith filter)
    auto populatePopup = [popup](const QString& prefix) {
        popup->clear();
        const auto& reg = formulaRegistry();
        for (auto it = reg.begin(); it != reg.end(); ++it) {
            if (it->name.startsWith(prefix, Qt::CaseInsensitive)) {
                auto* item = new QListWidgetItem(popup);
                item->setData(FuncNameRole, it->name);
                item->setData(FuncDescRole, it->description);
            }
        }
        return popup->count() > 0;
    };

    // When item clicked in popup → show detail panel
    QObject::connect(popup, &QListWidget::itemClicked, editor,
        [popup, detailPanel](QListWidgetItem* item) {
            QString funcName = item->data(FuncNameRole).toString();
            const auto& reg = formulaRegistry();
            if (reg.contains(funcName)) {
                detailPanel->setText(buildDetailHtml(reg[funcName]));
                detailPanel->adjustSize();
                // Try right side first; if clipped, move to left side
                QPoint rightPos = popup->mapToGlobal(QPoint(popup->width() + 4, 0));
                QScreen* screen = popup->screen();
                if (screen && rightPos.x() + detailPanel->width() > screen->availableGeometry().right()) {
                    QPoint leftPos = popup->mapToGlobal(QPoint(-detailPanel->width() - 4, 0));
                    detailPanel->move(leftPos);
                } else {
                    detailPanel->move(rightPos);
                }
                detailPanel->show();
            }
        });

    // Install event filter on popup (to forward keys to editor) and on editor
    popup->installEventFilter(const_cast<CellDelegate*>(this));
    popup->setProperty("_editor", QVariant::fromValue(static_cast<QObject*>(editor)));
    editor->installEventFilter(const_cast<CellDelegate*>(this));

    // Store widgets on editor for access in eventFilter
    editor->setProperty("_formulaPopup", QVariant::fromValue(static_cast<QObject*>(popup)));
    editor->setProperty("_paramHint", QVariant::fromValue(static_cast<QObject*>(paramHint)));
    editor->setProperty("_detailPanel", QVariant::fromValue(static_cast<QObject*>(detailPanel)));

    // Show popup / param hint on text change
    QObject::connect(editor, &QLineEdit::textChanged, editor,
        [this, editor, popup, paramHint, detailPanel, populatePopup](const QString& text) {
        emit formulaEditModeChanged(text.startsWith("="));
        detailPanel->hide(); // hide detail panel on any text change

        if (text.startsWith("=") && text.length() > 1) {
            QString token = extractCurrentToken(text);

            // Show autocomplete if typing a function name token
            if (!token.isEmpty() && token[0].isLetter()) {
                if (populatePopup(token)) {
                    popup->setCurrentRow(0);
                    // Defer positioning to next event loop — editor geometry
                    // may not be finalized on the very first keystroke
                    QTimer::singleShot(0, editor, [editor, popup]() {
                        if (!editor || !popup) return;
                        int popupW = qMax(460, editor->width());
                        int visibleItems = qMin(popup->count(), 8);
                        int popupH = visibleItems * 30 + 6;
                        popup->setFixedWidth(popupW);
                        popup->setFixedHeight(popupH);
                        int edH = qMax(editor->height(), 25);
                        QPoint below = editor->mapToGlobal(QPoint(0, edH + 2));
                        QPoint above = editor->mapToGlobal(QPoint(0, -popupH - 2));
                        QScreen* screen = editor->screen();
                        QPoint pos = below;
                        if (screen) {
                            QRect sr = screen->availableGeometry();
                            if (pos.y() + popupH > sr.bottom())
                                pos = above;
                            // Clamp horizontal: don't go off right edge
                            if (pos.x() + popupW > sr.right())
                                pos.setX(sr.right() - popupW);
                            // Don't go off left edge
                            if (pos.x() < sr.left())
                                pos.setX(sr.left());
                        }
                        popup->move(pos);
                        popup->show();
                    });
                } else {
                    popup->hide();
                }
            } else {
                popup->hide();
            }

            // Show param hint if inside a function call
            int cursorPos = editor->cursorPosition();
            FormulaContext ctx = findFormulaContext(text, cursorPos);
            const auto& reg = formulaRegistry();
            if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
                const auto& info = reg[ctx.funcName];
                paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
                paramHint->adjustSize();
                // Defer positioning to next event loop
                QTimer::singleShot(0, editor, [editor, popup, paramHint]() {
                    if (!editor || !paramHint) return;
                    QPoint hintPos;
                    if (popup && popup->isVisible()) {
                        hintPos = popup->mapToGlobal(QPoint(0, popup->height() + 2));
                    } else {
                        int edH = qMax(editor->height(), 25);
                        hintPos = editor->mapToGlobal(QPoint(0, edH + 2));
                    }
                    // Clamp to screen bounds
                    QScreen* screen = editor->screen();
                    if (screen) {
                        QRect sr = screen->availableGeometry();
                        if (hintPos.x() + paramHint->width() > sr.right())
                            hintPos.setX(sr.right() - paramHint->width());
                        if (hintPos.x() < sr.left())
                            hintPos.setX(sr.left());
                    }
                    paramHint->move(hintPos);
                    paramHint->show();
                });
            } else {
                paramHint->hide();
            }
        } else {
            popup->hide();
            paramHint->hide();
        }
    });

    // Also update param hint on cursor position change
    QObject::connect(editor, &QLineEdit::cursorPositionChanged, editor,
        [editor, paramHint, popup](int, int newPos) {
        QString text = editor->text();
        if (!text.startsWith("=")) { paramHint->hide(); return; }
        FormulaContext ctx = findFormulaContext(text, newPos);
        const auto& reg = formulaRegistry();
        if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
            const auto& info = reg[ctx.funcName];
            paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
            paramHint->adjustSize();
            QPoint hintPos;
            if (popup && popup->isVisible()) {
                hintPos = popup->mapToGlobal(QPoint(0, popup->height() + 2));
            } else {
                int edH = qMax(editor->height(), 25);
                hintPos = editor->mapToGlobal(QPoint(0, edH + 2));
            }
            // Clamp to screen bounds
            QScreen* screen = editor->screen();
            if (screen) {
                QRect sr = screen->availableGeometry();
                if (hintPos.x() + paramHint->width() > sr.right())
                    hintPos.setX(sr.right() - paramHint->width());
                if (hintPos.x() < sr.left())
                    hintPos.setX(sr.left());
            }
            paramHint->move(hintPos);
            paramHint->show();
        } else {
            paramHint->hide();
        }
    });

    // Clean up on editor destruction — hide immediately, then delete
    QObject::connect(editor, &QObject::destroyed, popup, [popup]() {
        popup->hide();
        popup->deleteLater();
    });
    QObject::connect(editor, &QObject::destroyed, paramHint, [paramHint]() {
        paramHint->hide();
        paramHint->deleteLater();
    });
    QObject::connect(editor, &QObject::destroyed, detailPanel, [detailPanel]() {
        detailPanel->hide();
        detailPanel->deleteLater();
    });

    return editor;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        lineEdit->setText(index.data(Qt::EditRole).toString());
        // Place cursor at end (not selecting all text) — like Excel
        // Deferred to after Qt finishes initializing the editor, which otherwise re-selects all text
        QTimer::singleShot(0, lineEdit, [lineEdit]() {
            lineEdit->deselect();
            lineEdit->setCursorPosition(lineEdit->text().length());
        });
    }
}

void CellDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                               const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        model->setData(index, lineEdit->text(), Qt::EditRole);
    }
}

void CellDelegate::updateEditorGeometry(QWidget* editor, const QStyleOptionViewItem& option,
                                       const QModelIndex& index) const {
    // Start with the cell rect
    QRect geom = option.rect;

    // Expand editor rightward as text grows (like Excel)
    if (m_spreadsheetView) {
        auto* lineEdit = qobject_cast<QLineEdit*>(editor);
        if (lineEdit) {
            QFontMetrics fm(lineEdit->font());
            int textW = fm.horizontalAdvance(lineEdit->text()) + 20; // padding for cursor + border
            if (textW > geom.width()) {
                auto sp = m_spreadsheetView->getSpreadsheet();
                int row = index.row();
                int col = index.column();
                int maxCol = m_spreadsheetView->model()->columnCount() - 1;
                int expandedRight = geom.right();
                // Expand into empty neighbor cells to the right
                for (int c = col + 1; c <= maxCol; ++c) {
                    if (sp) {
                        auto neighbor = sp->getCellIfExists(row, c);
                        if (neighbor && neighbor->getType() != CellType::Empty) break;
                    }
                    QModelIndex nIdx = m_spreadsheetView->model()->index(row, c);
                    QRect nRect = m_spreadsheetView->visualRect(nIdx);
                    expandedRight = nRect.right();
                    if (expandedRight - geom.left() >= textW) break;
                }
                geom.setRight(qMax(geom.right(), expandedRight));
            }
        }
    }

    editor->setGeometry(geom);
}

void CellDelegate::paint(QPainter* painter, const QStyleOptionViewItem& option,
                        const QModelIndex& index) const {
    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, false);

    QRect rect = option.rect;
    bool isSelected = option.state & QStyle::State_Selected;
    bool hasFocus = option.state & QStyle::State_HasFocus;

    // --- Check if this is a Picklist/Checkbox cell (skip selection tint) ---
    bool isPicklistOrCheckbox = false;
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp) {
            auto c = sp->getCellIfExists(index.row(), index.column());
            if (c) {
                const auto& fmt = c->getStyle().numberFormat;
                isPicklistOrCheckbox = (fmt == "Picklist" || fmt == "Checkbox");
            }
        }
    }

    // --- Background ---
    const auto& theme = ThemeManager::instance().currentTheme();
    QColor defaultBg = theme.gridBackground;
    QColor bgColor = defaultBg;

    // Check style overlays directly (instant visual for bulk formatting)
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        auto* mdl = m_spreadsheetView->getModel();
        if (sp && !sp->getStyleOverlays().empty()) {
            int logRow = mdl ? mdl->toLogicalRow(index.row()) : index.row();
            int col = index.column();
            for (const auto& ov : sp->getStyleOverlays()) {
                if (logRow >= ov.minRow && logRow <= ov.maxRow &&
                    col >= ov.minCol && col <= ov.maxCol) {
                    CellStyle tmpStyle;
                    ov.modifier(tmpStyle);
                    QColor ovBg(tmpStyle.backgroundColor);
                    // Resolve theme colors (e.g. "theme:4:0" → actual hex)
                    if (!ovBg.isValid() && !tmpStyle.backgroundColor.isEmpty()) {
                        ovBg = sp->getDocumentTheme().resolveColor(tmpStyle.backgroundColor);
                    }
                    if (ovBg.isValid() && ovBg != Qt::white && ovBg != QColor("#FFFFFF")) {
                        bgColor = ovBg;
                    }
                }
            }
        }
    }

    // Also check model data for non-overlay backgrounds
    if (bgColor == defaultBg) {
        QVariant bgData = index.data(Qt::BackgroundRole);
        if (bgData.isValid()) {
            QColor cellBg = bgData.value<QColor>();
            if (cellBg.isValid() && cellBg.rgb() != QColor(Qt::white).rgb()) {
                bgColor = cellBg;
            }
        }
    }

    bool isFocusCell = (option.state & QStyle::State_HasFocus);

    if (isPicklistOrCheckbox && !(isSelected && !isFocusCell)) {
        // Picklist/Checkbox cells: use theme bg when not part of a multi-select
        painter->fillRect(rect, defaultBg);
    } else if (isSelected && !isFocusCell) {
        // Multi-select (non-focus cells): paint bg color first, then semi-transparent
        // selection tint so fill colors are visible even while selected.
        // The active/focus cell keeps its original background (white/clear) — Excel behavior.
        painter->fillRect(rect, bgColor);
        QColor tint = ThemeManager::instance().currentTheme().selectionTint;
        if (bgColor != defaultBg) {
            // Cell has a custom bg color — use very light tint so color shows through
            tint.setAlpha(40);
        }
        painter->fillRect(rect, tint);
    } else {
        // Active/focus cell or unselected cell: paint original background, no tint
        painter->fillRect(rect, bgColor);
    }

    // --- Visual conditional formatting (data bar, color scale, icon set) ---
    std::optional<VisualFormatResult> visualFormat;
    int iconTextShift = 0; // pixels to shift text right for icon set
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp && !sp->getConditionalFormatting().getAllRules().empty()) {
            CellAddress addr(index.row(), index.column());
            QVariant cellValue = sp->getCellValue(addr);
            // Provide value lookup for auto-range computation
            ConditionalFormatting::ValueLookup lookup = [&sp](int r, int c) -> QVariant {
                return sp->getCellValue(CellAddress(r, c));
            };
            visualFormat = sp->getConditionalFormatting().getVisualFormat(addr, cellValue, lookup);
        }
    }

    // Color scale: override background color
    if (visualFormat && (visualFormat->type == ConditionType::ColorScale2 ||
                         visualFormat->type == ConditionType::ColorScale3)) {
        drawColorScaleBackground(painter, rect, *visualFormat);
    }

    // Data bar: draw bar after background
    if (visualFormat && visualFormat->type == ConditionType::DataBar) {
        QString cellText = index.data(Qt::DisplayRole).toString();
        int alignment = Qt::AlignVCenter | Qt::AlignLeft;
        QVariant alignData = index.data(Qt::TextAlignmentRole);
        if (alignData.isValid()) alignment = alignData.toInt();
        drawDataBar(painter, rect, *visualFormat, cellText, alignment);
    }

    // Icon set: draw icon, compute text shift
    if (visualFormat && visualFormat->type == ConditionType::IconSet) {
        drawIconSet(painter, rect, *visualFormat);
        iconTextShift = ICON_SIZE + ICON_MARGIN * 2;
    }

    // --- Formula recalc flash overlay (bright yellow, fade-in then hold then fade-out) ---
    if (m_spreadsheetView) {
        double flashProgress = m_spreadsheetView->cellAnimationProgress(index.row(), index.column());
        if (flashProgress > 0.01) {
            int alpha = static_cast<int>(130 * flashProgress);
            painter->fillRect(rect, QColor(255, 243, 100, alpha));
        }
    }

    // --- Picklist / Checkbox rendering ---
    bool skipText = false;
    if (m_spreadsheetView) {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) {
            auto cell = spreadsheet->getCellIfExists(index.row(), index.column());
            if (cell) {
                const auto& style = cell->getStyle();
                if (style.numberFormat == "Checkbox") {
                    bool checked = false;
                    auto val = cell->getValue();
                    if (val.typeId() == QMetaType::Bool) checked = val.toBool();
                    else {
                        QString s = val.toString().toLower();
                        checked = (s == "true" || s == "1");
                    }
                    drawCheckbox(painter, rect, checked);
                    skipText = true;
                } else if (style.numberFormat == "Picklist") {
                    QString val = cell->getValue().toString();
                    // White bg already painted above (isPicklistOrCheckbox branch)
                    const auto* rule = spreadsheet->getValidationAt(index.row(), index.column());
                    QStringList options = rule ? rule->listItems : QStringList();
                    QStringList colors = rule ? rule->listItemColors : QStringList();
                    Qt::Alignment align = Qt::AlignLeft | Qt::AlignVCenter;
                    QVariant alignVar = index.data(Qt::TextAlignmentRole);
                    if (alignVar.isValid())
                        align = Qt::Alignment(alignVar.toInt());
                    drawPicklistTags(painter, rect, val, options, colors, align);
                    skipText = true;
                }
            }
        }
    }

    // --- Check for hyperlink ---
    bool hasHyperlink = false;
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp) {
            auto cell = sp->getCellIfExists(index.row(), index.column());
            if (cell && cell->hasHyperlink()) {
                hasHyperlink = true;
            }
        }
    }

    // --- Skip text for visual formats that hide values ---
    if (visualFormat && visualFormat->type == ConditionType::DataBar && !visualFormat->barShowValue) {
        skipText = true;
    }
    if (visualFormat && visualFormat->type == ConditionType::IconSet && !visualFormat->iconShowValue) {
        skipText = true;
    }

    // --- Text ---
    QString text = index.data(Qt::DisplayRole).toString();
    if (!skipText && !text.isEmpty()) {
        QFont font = option.font;
        QVariant fontData = index.data(Qt::FontRole);
        if (fontData.isValid()) {
            font = fontData.value<QFont>();
        }

        // Hyperlink styling: blue underline
        if (hasHyperlink) {
            font.setUnderline(true);
        }
        painter->setFont(font);

        QColor fgColor = theme.textPrimary;
        if (hasHyperlink) {
            fgColor = QColor("#0563C1"); // Excel-style hyperlink blue
        }
        // Check style overlays for font color (instant visual for bulk operations)
        if (!hasHyperlink && m_spreadsheetView) {
            auto sp = m_spreadsheetView->getSpreadsheet();
            auto* mdl = m_spreadsheetView->getModel();
            if (sp && !sp->getStyleOverlays().empty()) {
                int logRow = mdl ? mdl->toLogicalRow(index.row()) : index.row();
                for (const auto& ov : sp->getStyleOverlays()) {
                    if (logRow >= ov.minRow && logRow <= ov.maxRow &&
                        index.column() >= ov.minCol && index.column() <= ov.maxCol) {
                        CellStyle tmpStyle;
                        ov.modifier(tmpStyle);
                        QColor ovFg(tmpStyle.foregroundColor);
                        if (!ovFg.isValid() && !tmpStyle.foregroundColor.isEmpty())
                            ovFg = sp->getDocumentTheme().resolveColor(tmpStyle.foregroundColor);
                        if (ovFg.isValid() && ovFg != Qt::black && ovFg != QColor("#000000"))
                            fgColor = ovFg;
                    }
                }
            }
        }
        if (!hasHyperlink && fgColor == theme.textPrimary) {
            QVariant fgData = index.data(Qt::ForegroundRole);
            if (fgData.isValid()) {
                QColor c = fgData.value<QColor>();
                if (c.isValid() && c.rgb() != QColor(Qt::black).rgb()) fgColor = c;
            }
        }
        painter->setPen(fgColor);

        int alignment = Qt::AlignVCenter | Qt::AlignLeft;
        QVariant alignData = index.data(Qt::TextAlignmentRole);
        if (alignData.isValid()) {
            alignment = alignData.toInt();
        }

        // Indent support: add left padding per indent level
        int indentPx = 0;
        QVariant indentData = index.data(Qt::UserRole + 10);
        if (indentData.isValid()) {
            indentPx = indentData.toInt() * 12; // 12px per indent level
        }

        // Text rotation support
        int rotation = 0;
        QVariant rotData = index.data(Qt::UserRole + 16);
        if (rotData.isValid()) {
            rotation = rotData.toInt();
        }

        QRect textRect = rect.adjusted(4 + indentPx + iconTextShift, 1, -4, -1);

        // Get text overflow mode for this cell
        TextOverflowMode overflowMode = TextOverflowMode::Overflow;
        if (m_spreadsheetView) {
            auto sp = m_spreadsheetView->getSpreadsheet();
            if (sp) {
                auto cell = sp->getCellIfExists(index.row(), index.column());
                if (cell) overflowMode = cell->getStyle().textOverflow;
            }
        }

        if (rotation == 0 && overflowMode == TextOverflowMode::Wrap) {
            // Wrap text within cell
            painter->drawText(textRect, alignment | Qt::TextWordWrap, text);
        } else if (rotation == 0 && overflowMode == TextOverflowMode::ShrinkToFit) {
            // Shrink font until text fits
            QFontMetrics fm(font);
            int textW = fm.horizontalAdvance(text);
            if (textW > textRect.width() && textRect.width() > 0) {
                double scale = static_cast<double>(textRect.width()) / textW;
                QFont shrunkFont = font;
                int newSize = qMax(5, static_cast<int>(font.pointSize() * scale));
                shrunkFont.setPointSize(newSize);
                painter->setFont(shrunkFont);
            }
            painter->drawText(textRect, alignment, text);
            painter->setFont(font); // restore
        } else if (rotation == 0) {
            // Overflow mode (default): text overflows into empty neighbors.
            // Numbers: show ### when too wide (Excel behavior) — no overflow.
            QFontMetrics fm(font);
            int textW = fm.horizontalAdvance(text);
            bool isNumeric = (alignment & Qt::AlignRight);
            if (isNumeric && textW > textRect.width() && textRect.width() > 10) {
                // Number too wide — show ### like Excel
                QString hashes = QString(textRect.width() / fm.horizontalAdvance('#'), '#');
                painter->drawText(textRect, alignment, hashes);
            } else {
                painter->drawText(textRect, alignment, text);
            }
        } else if (rotation == 270) {
            // Vertical stacked text: draw each character on its own line
            QFontMetrics fm(font);
            int charH = fm.height();
            int maxCharW = 0;
            for (const QChar& ch : text) {
                maxCharW = qMax(maxCharW, fm.horizontalAdvance(ch));
            }
            int totalH = charH * text.length();
            int startY = textRect.top() + qMax(0, (textRect.height() - totalH) / 2);
            int centerX = textRect.left() + (textRect.width() - maxCharW) / 2;
            for (int i = 0; i < text.length(); ++i) {
                QRect charRect(centerX, startY + i * charH, maxCharW, charH);
                painter->drawText(charRect, Qt::AlignCenter, QString(text[i]));
            }
        } else {
            // Angled text: rotate painter
            painter->save();
            QFontMetrics fm(font);
            int textW = fm.horizontalAdvance(text);
            int textH = fm.height();
            QPointF center = textRect.center();
            painter->translate(center);
            painter->rotate(-rotation); // Qt rotates clockwise, we want CCW for positive angles
            QRectF rotRect(-textW / 2.0, -textH / 2.0, textW, textH);
            painter->drawText(rotRect, Qt::AlignCenter, text);
            painter->restore();
        }
    }

    // --- Sparkline rendering ---
    QVariant sparkData = index.data(Qt::UserRole + 15); // SparklineRole
    if (sparkData.isValid() && sparkData.canConvert<SparklineRenderData>()) {
        auto rd = sparkData.value<SparklineRenderData>();
        if (!rd.values.isEmpty()) {
            QRect sparkRect = rect.adjusted(3, 3, -3, -3);
            drawSparkline(painter, sparkRect, rd);
        }
    }

    // --- Spill indicator: blue dotted border for spill child cells, blue arrow for spill parent ---
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp) {
            auto cell = sp->getCellIfExists(index.row(), index.column());
            if (cell && cell->isSpillCell()) {
                // Spill child: subtle blue dotted border
                painter->setPen(QPen(QColor(70, 130, 230, 120), 1, Qt::DotLine));
                painter->drawRect(rect.adjusted(0, 0, -1, -1));
            }
            // Spill parent: small blue spill indicator arrow in bottom-right
            CellAddress addr(index.row(), index.column());
            if (sp->hasSpillRange(addr)) {
                painter->setRenderHint(QPainter::Antialiasing, true);
                QPainterPath arrow;
                int ax = rect.right() - 8;
                int ay = rect.bottom() - 8;
                arrow.moveTo(ax, ay + 6);
                arrow.lineTo(ax + 6, ay + 6);
                arrow.lineTo(ax + 6, ay);
                arrow.lineTo(ax + 4, ay + 2);
                arrow.lineTo(ax + 6, ay + 6);
                painter->setPen(QPen(QColor(70, 130, 230), 1.2));
                painter->setBrush(Qt::NoBrush);
                painter->drawPath(arrow);
                painter->setRenderHint(QPainter::Antialiasing, false);
            }
        }
    }

    // --- Text overflow: draw overflowed text from left neighbor into this empty cell ---
    {
        QString cellText = index.data(Qt::DisplayRole).toString();
        if (cellText.isEmpty() && m_spreadsheetView) {
            auto sp = m_spreadsheetView->getSpreadsheet();
            auto* mdl = m_spreadsheetView->getModel();
            if (sp && mdl) {
                int row = mdl->toLogicalRow(index.row());
                int col = index.column();
                // Search leftward for a source cell with overflowing text
                for (int srcCol = col - 1; srcCol >= std::max(0, col - 20); --srcCol) {
                    auto srcCell = sp->getCellIfExists(row, srcCol);
                    if (!srcCell || srcCell->getType() == CellType::Empty) continue;
                    // Numbers never overflow (they show ### instead)
                    auto srcType = srcCell->getType();
                    if (srcType == CellType::Number || srcType == CellType::Date ||
                        srcType == CellType::Formula) break;
                    // Found a non-empty text cell — check if it overflows
                    const auto& srcStyle = srcCell->getStyle();
                    if (srcStyle.textOverflow != TextOverflowMode::Overflow) break;
                    // Check that all cells between source and current are empty
                    bool allEmpty = true;
                    for (int c = srcCol + 1; c < col; ++c) {
                        auto between = sp->getCellIfExists(row, c);
                        if (between && between->getType() != CellType::Empty) { allEmpty = false; break; }
                    }
                    if (!allEmpty) break;
                    // Get source cell text
                    QModelIndex srcIdx = m_spreadsheetView->model()->index(row, srcCol);
                    QString srcText = srcIdx.data(Qt::DisplayRole).toString();
                    if (srcText.isEmpty()) break;
                    // Get source cell font
                    QFont srcFont = option.font;
                    QVariant srcFontData = srcIdx.data(Qt::FontRole);
                    if (srcFontData.isValid()) srcFont = srcFontData.value<QFont>();
                    QFontMetrics srcFm(srcFont);
                    int srcTextW = srcFm.horizontalAdvance(srcText);
                    // Get source cell rect
                    QRect srcRect = m_spreadsheetView->visualRect(srcIdx);
                    int srcTextLeft = srcRect.left() + 4;
                    // Check indent
                    QVariant srcIndent = srcIdx.data(Qt::UserRole + 10);
                    if (srcIndent.isValid()) srcTextLeft += srcIndent.toInt() * 12;
                    int textRight = srcTextLeft + srcTextW;
                    // Does the source text reach into our cell?
                    if (textRight > rect.left()) {
                        // Get source fg color
                        QColor srcFg = theme.textPrimary;
                        QVariant srcFgData = srcIdx.data(Qt::ForegroundRole);
                        if (srcFgData.isValid()) {
                            QColor c = srcFgData.value<QColor>();
                            if (c.isValid()) srcFg = c;
                        }
                        painter->setFont(srcFont);
                        painter->setPen(srcFg);
                        // Draw text spanning from source origin across into this cell
                        QRect fullTextRect(srcTextLeft, rect.top() + 1, textRight - srcTextLeft, rect.height() - 2);
                        painter->drawText(fullTextRect, Qt::AlignVCenter | Qt::AlignLeft, srcText);
                    }
                    break; // only check the nearest non-empty cell to the left
                }
            }
        }
    }

    // --- Comment indicator: small red triangle in top-right corner ---
    if (m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp) {
            auto cell = sp->getCellIfExists(index.row(), index.column());
            if (cell && cell->hasComment()) {
                painter->setRenderHint(QPainter::Antialiasing, true);
                QPainterPath triangle;
                triangle.moveTo(rect.right() - 7, rect.top());
                triangle.lineTo(rect.right(), rect.top());
                triangle.lineTo(rect.right(), rect.top() + 7);
                triangle.closeSubpath();
                painter->setPen(Qt::NoPen);
                painter->setBrush(QColor("#FF0000"));
                painter->drawPath(triangle);
                painter->setRenderHint(QPainter::Antialiasing, false);
            }
        }
    }

    // --- Hyperlink indicator: small blue triangle in bottom-left corner ---
    if (hasHyperlink) {
        painter->setRenderHint(QPainter::Antialiasing, true);
        QPainterPath linkTriangle;
        linkTriangle.moveTo(rect.left(), rect.bottom() - 5);
        linkTriangle.lineTo(rect.left(), rect.bottom());
        linkTriangle.lineTo(rect.left() + 5, rect.bottom());
        linkTriangle.closeSubpath();
        painter->setPen(Qt::NoPen);
        painter->setBrush(QColor("#0563C1"));
        painter->drawPath(linkTriangle);
        painter->setRenderHint(QPainter::Antialiasing, false);
    }

    // --- Gridlines: single thin line on right and bottom edges ---
    // Excel behavior: gridlines are hidden when a cell has a background color
    // Also hide right gridline when text overflows into the next cell
    bool hasBgColor = bgColor != defaultBg;
    bool hideRightGridline = false;
    if (m_showGridlines && !hasBgColor && m_spreadsheetView) {
        auto sp = m_spreadsheetView->getSpreadsheet();
        if (sp) {
            int row = index.row();
            int col = index.column();
            // Check if THIS cell's text overflows right
            auto thisCell = sp->getCellIfExists(row, col);
            if (thisCell && thisCell->getType() != CellType::Empty &&
                thisCell->getStyle().textOverflow == TextOverflowMode::Overflow) {
                QString thisText = index.data(Qt::DisplayRole).toString();
                if (!thisText.isEmpty()) {
                    QFont f = option.font;
                    QVariant fd = index.data(Qt::FontRole);
                    if (fd.isValid()) f = fd.value<QFont>();
                    QFontMetrics fm(f);
                    int textW = fm.horizontalAdvance(thisText);
                    int cellTextW = rect.width() - 8;
                    if (textW > cellTextW) {
                        // Check if next cell is empty
                        auto nextCell = sp->getCellIfExists(row, col + 1);
                        if (!nextCell || nextCell->getType() == CellType::Empty)
                            hideRightGridline = true;
                    }
                }
            }
            // Check if this cell RECEIVES overflow from the left (is empty with overflow text)
            if (!hideRightGridline && (!thisCell || thisCell->getType() == CellType::Empty)) {
                for (int srcCol = col - 1; srcCol >= 0; --srcCol) {
                    auto srcCell = sp->getCellIfExists(row, srcCol);
                    if (!srcCell || srcCell->getType() == CellType::Empty) continue;
                    if (srcCell->getStyle().textOverflow != TextOverflowMode::Overflow) break;
                    // Check all cells between are empty
                    bool allEmpty = true;
                    for (int c = srcCol + 1; c < col; ++c) {
                        auto bet = sp->getCellIfExists(row, c);
                        if (bet && bet->getType() != CellType::Empty) { allEmpty = false; break; }
                    }
                    if (!allEmpty) break;
                    QModelIndex srcIdx = m_spreadsheetView->model()->index(row, srcCol);
                    QString srcText = srcIdx.data(Qt::DisplayRole).toString();
                    QFont sf = option.font;
                    QVariant sfd = srcIdx.data(Qt::FontRole);
                    if (sfd.isValid()) sf = sfd.value<QFont>();
                    QFontMetrics sfm(sf);
                    QRect srcRect = m_spreadsheetView->visualRect(srcIdx);
                    int srcTextLeft = srcRect.left() + 4;
                    QVariant si = srcIdx.data(Qt::UserRole + 10);
                    if (si.isValid()) srcTextLeft += si.toInt() * 12;
                    int textRight = srcTextLeft + sfm.horizontalAdvance(srcText);
                    if (textRight > rect.right()) hideRightGridline = true;
                    break;
                }
            }
        }
    }
    if (m_showGridlines && !hasBgColor) {
        painter->setPen(QPen(ThemeManager::instance().currentTheme().gridLineColor, 1, Qt::SolidLine));
        if (!hideRightGridline)
            painter->drawLine(rect.right(), rect.top(), rect.right(), rect.bottom());
        painter->drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom());
    }

    // --- Cell borders (user-defined) ---
    auto drawBorder = [&](const QVariant& borderData, int x1, int y1, int x2, int y2) {
        if (!borderData.isValid()) return;
        QStringList parts = borderData.toString().split(',');
        if (parts.size() >= 2) {
            int w = parts[0].toInt();
            QColor c(parts[1]);
            int ps = (parts.size() >= 3) ? parts[2].toInt() : 0;
            if (c.isValid() && w > 0) {
                Qt::PenStyle penStyle = Qt::SolidLine;
                if (ps == 1) penStyle = Qt::DashLine;
                else if (ps == 2) penStyle = Qt::DotLine;
                painter->setPen(QPen(c, w, penStyle));
                painter->drawLine(x1, y1, x2, y2);
            }
        }
    };
    drawBorder(index.data(Qt::UserRole + 11), rect.left(), rect.top(), rect.right(), rect.top());      // top
    drawBorder(index.data(Qt::UserRole + 12), rect.left(), rect.bottom(), rect.right(), rect.bottom()); // bottom
    drawBorder(index.data(Qt::UserRole + 13), rect.left(), rect.top(), rect.left(), rect.bottom());     // left
    drawBorder(index.data(Qt::UserRole + 14), rect.right(), rect.top(), rect.right(), rect.bottom());   // right

    // --- Focus border: green rectangle for the current cell ---
    // --- Focus cell border ---
    if (hasFocus) {
        painter->setClipRect(rect.adjusted(-1, -1, 1, 1));
        QPen focusPen(ThemeManager::instance().currentTheme().focusBorderColor, 2, Qt::SolidLine);
        painter->setPen(focusPen);
        painter->drawRect(rect.adjusted(1, 1, -1, -1));
    }

    painter->restore();
}

QSize CellDelegate::sizeHint(const QStyleOptionViewItem& option,
                             const QModelIndex& index) const {
    return QStyledItemDelegate::sizeHint(option, index);
}

void CellDelegate::drawSparkline(QPainter* painter, const QRect& rect,
                                  const SparklineRenderData& data) const {
    if (data.values.isEmpty() || rect.width() < 4 || rect.height() < 4) return;

    double range = data.maxVal - data.minVal;
    if (range == 0) range = 1.0;
    int n = data.values.size();

    switch (data.type) {
        case SparklineType::Line: {
            painter->setRenderHint(QPainter::Antialiasing, true);
            double stepX = static_cast<double>(rect.width()) / qMax(1, n - 1);

            QPainterPath path;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * stepX;
                double y = rect.bottom() - ((data.values[i] - data.minVal) / range) * rect.height();
                if (i == 0) path.moveTo(x, y);
                else path.lineTo(x, y);
            }
            painter->setPen(QPen(data.lineColor, data.lineWidth));
            painter->setBrush(Qt::NoBrush);
            painter->drawPath(path);

            if (data.showHighPoint && data.highIndex >= 0) {
                double x = rect.left() + data.highIndex * stepX;
                double y = rect.bottom() - ((data.values[data.highIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.highPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            if (data.showLowPoint && data.lowIndex >= 0) {
                double x = rect.left() + data.lowIndex * stepX;
                double y = rect.bottom() - ((data.values[data.lowIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.lowPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            painter->setRenderHint(QPainter::Antialiasing, false);
            break;
        }
        case SparklineType::Column: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double zeroY = rect.bottom();
            if (data.minVal < 0) {
                zeroY = rect.bottom() - ((-data.minVal) / range) * rect.height();
            }
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                double barH = (std::abs(data.values[i]) / range) * rect.height();
                if (data.values[i] >= 0) {
                    QRectF bar(x, zeroY - barH, barW, barH);
                    painter->fillRect(bar, data.lineColor);
                } else {
                    QRectF bar(x, zeroY, barW, barH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
        case SparklineType::WinLoss: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double midY = rect.top() + rect.height() / 2.0;
            double halfH = rect.height() / 2.0 - 2;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                if (data.values[i] > 0) {
                    QRectF bar(x, midY - halfH, barW, halfH);
                    painter->fillRect(bar, data.lineColor);
                } else if (data.values[i] < 0) {
                    QRectF bar(x, midY, barW, halfH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
    }
}

void CellDelegate::drawCheckbox(QPainter* painter, const QRect& rect, bool checked) const {
    int boxSize = 14;
    int x = rect.left() + (rect.width() - boxSize) / 2;
    int y = rect.top() + (rect.height() - boxSize) / 2;
    QRectF boxRect(x + 0.5, y + 0.5, boxSize - 1, boxSize - 1);

    painter->setRenderHint(QPainter::Antialiasing, true);
    if (checked) {
        // Soft green filled rounded rect
        painter->setPen(Qt::NoPen);
        painter->setBrush(ThemeManager::instance().currentTheme().checkboxChecked);
        painter->drawRoundedRect(boxRect, 4, 4);
        // Smooth checkmark
        QPainterPath check;
        check.moveTo(x + 3.5, y + 7);
        check.lineTo(x + 6, y + 10);
        check.lineTo(x + 10.5, y + 4.5);
        painter->setPen(QPen(Qt::white, 1.6, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        painter->setBrush(Qt::NoBrush);
        painter->drawPath(check);
    } else {
        // Soft outlined rounded rect
        painter->setPen(QPen(ThemeManager::instance().currentTheme().checkboxUncheckedBorder, 1.2));
        painter->setBrush(QColor("#FAFAFA"));
        painter->drawRoundedRect(boxRect, 4, 4);
    }
    painter->setRenderHint(QPainter::Antialiasing, false);
}

void CellDelegate::drawPicklistTags(QPainter* painter, const QRect& rect,
                                     const QString& value, const QStringList& allOptions,
                                     const QStringList& optionColors,
                                     Qt::Alignment alignment) const {
    painter->setRenderHint(QPainter::Antialiasing, true);

    // Draw a subtle dropdown arrow on the right
    {
        int arrowSize = 6;
        int arrowX = rect.right() - arrowSize - 6;
        int arrowY = rect.top() + (rect.height() - arrowSize / 2) / 2;
        QPainterPath arrow;
        arrow.moveTo(arrowX, arrowY);
        arrow.lineTo(arrowX + arrowSize, arrowY);
        arrow.lineTo(arrowX + arrowSize / 2.0, arrowY + arrowSize * 0.55);
        arrow.closeSubpath();
        painter->setPen(Qt::NoPen);
        painter->setBrush(QColor("#B0B4BA"));
        painter->drawPath(arrow);
    }

    if (value.isEmpty()) {
        painter->setRenderHint(QPainter::Antialiasing, false);
        return;
    }

    QStringList selected = value.split('|', Qt::SkipEmptyParts);

    QFont tagFont(painter->font().family(), 10);
    tagFont.setWeight(QFont::Medium);
    tagFont.setLetterSpacing(QFont::AbsoluteSpacing, 0.2);
    QFontMetrics fm(tagFont);
    painter->setFont(tagFont);

    int tagH = 18;
    int gap = 3;
    int maxX = rect.right() - 18; // leave room for dropdown arrow

    // Compute total width for horizontal alignment
    int totalTagW = 0;
    for (const QString& item : selected) {
        QString trimmed = item.trimmed();
        if (!trimmed.isEmpty())
            totalTagW += fm.horizontalAdvance(trimmed) + 14 + gap;
    }
    if (totalTagW > 0) totalTagW -= gap;

    int availW = maxX - rect.left() - 5;
    int x = rect.left() + 5;
    if (alignment & Qt::AlignHCenter)
        x = rect.left() + 5 + qMax(0, (availW - totalTagW) / 2);
    else if (alignment & Qt::AlignRight)
        x = rect.left() + 5 + qMax(0, availW - totalTagW);

    int y;
    if (alignment & Qt::AlignTop)
        y = rect.top() + 2;
    else if (alignment & Qt::AlignBottom)
        y = rect.bottom() - tagH - 2;
    else
        y = rect.top() + (rect.height() - tagH) / 2;

    int tagIndex = 0;
    for (const QString& item : selected) {
        QString trimmed = item.trimmed();
        if (trimmed.isEmpty()) continue;

        int colorIdx = allOptions.indexOf(trimmed);
        if (colorIdx < 0) colorIdx = static_cast<int>(qHash(trimmed) % TAG_COLOR_COUNT);

        QColor bg, fg;
        // Use custom color if provided for this option
        if (colorIdx >= 0 && colorIdx < optionColors.size() && !optionColors[colorIdx].isEmpty()) {
            bg = QColor(optionColors[colorIdx]);
            // Compute readable text color: dark text on light bg, white on dark bg
            fg = (bg.lightness() > 140) ? bg.darker(300) : QColor(Qt::white);
        } else {
            bg = tagBgColor(colorIdx);
            fg = tagTextColor(colorIdx);
        }

        int textW = fm.horizontalAdvance(trimmed);
        int tagW = textW + 14;

        // If not the first tag and doesn't fit, show "..." and stop
        if (tagIndex > 0 && x + tagW > maxX) {
            painter->setPen(QColor("#9CA3AF"));
            QFont smallFont(painter->font());
            smallFont.setPointSize(8);
            painter->setFont(smallFont);
            painter->drawText(QRectF(x, y, 20, tagH), Qt::AlignVCenter | Qt::AlignLeft, "...");
            break;
        }

        // For the first tag, clip to available width if needed
        int drawW = qMin(tagW, maxX - x);
        QRectF tagRect(x, y, drawW, tagH);
        painter->setPen(Qt::NoPen);
        painter->setBrush(bg);
        painter->drawRoundedRect(tagRect, tagH / 2.0, tagH / 2.0);

        painter->setPen(fg);
        painter->save();
        painter->setClipRect(tagRect);
        painter->drawText(QRectF(x, y, tagW, tagH), Qt::AlignCenter, trimmed);
        painter->restore();

        x += drawW + gap;
        tagIndex++;
    }

    painter->setRenderHint(QPainter::Antialiasing, false);
}

// --- Visual conditional formatting rendering ---

void CellDelegate::drawDataBar(QPainter* painter, const QRect& rect, const VisualFormatResult& vf,
                                const QString& /*text*/, int /*alignment*/) const {
    if (vf.barFraction <= 0.0) return;

    int barWidth = static_cast<int>(vf.barFraction * (rect.width() - 4));
    if (barWidth < 1) return;

    QColor barColor = vf.barColor;
    barColor.setAlpha(90); // semi-transparent so text remains readable

    QRect barRect(rect.left() + 2, rect.top() + 2, barWidth, rect.height() - 4);
    painter->fillRect(barRect, barColor);

    // Draw a thin solid border on the bar
    QColor borderColor = vf.barColor;
    borderColor.setAlpha(180);
    painter->setPen(QPen(borderColor, 1));
    painter->drawRect(barRect);
}

void CellDelegate::drawColorScaleBackground(QPainter* painter, const QRect& rect, const VisualFormatResult& vf) const {
    if (vf.scaleColor.isValid()) {
        painter->fillRect(rect, vf.scaleColor);
    }
}

void CellDelegate::drawIconSet(QPainter* painter, const QRect& rect, const VisualFormatResult& vf) const {
    if (vf.iconIndex < 0 || vf.iconIndex > 2) return;

    painter->setRenderHint(QPainter::Antialiasing, true);

    int iconX = rect.left() + ICON_MARGIN;
    int iconY = rect.top() + (rect.height() - ICON_SIZE) / 2;
    QRect iconRect(iconX, iconY, ICON_SIZE, ICON_SIZE);

    switch (vf.iconType) {
    case IconSetConfig::TrafficLights: {
        // Filled circles: red, yellow, green
        static const QColor colors[] = {QColor(220, 50, 50), QColor(240, 200, 40), QColor(60, 180, 75)};
        painter->setPen(Qt::NoPen);
        painter->setBrush(colors[vf.iconIndex]);
        painter->drawEllipse(iconRect.adjusted(1, 1, -1, -1));
        break;
    }
    case IconSetConfig::Arrows3: {
        // Triangles: down (red), right (yellow), up (green)
        static const QColor colors[] = {QColor(220, 50, 50), QColor(240, 180, 40), QColor(60, 180, 75)};
        painter->setPen(Qt::NoPen);
        painter->setBrush(colors[vf.iconIndex]);
        QPainterPath arrow;
        double cx = iconRect.center().x();
        double cy = iconRect.center().y();
        double s = ICON_SIZE * 0.35;
        if (vf.iconIndex == 0) {
            // Down arrow
            arrow.moveTo(cx, cy + s);
            arrow.lineTo(cx - s, cy - s * 0.5);
            arrow.lineTo(cx + s, cy - s * 0.5);
        } else if (vf.iconIndex == 1) {
            // Right arrow
            arrow.moveTo(cx + s, cy);
            arrow.lineTo(cx - s * 0.5, cy - s);
            arrow.lineTo(cx - s * 0.5, cy + s);
        } else {
            // Up arrow
            arrow.moveTo(cx, cy - s);
            arrow.lineTo(cx - s, cy + s * 0.5);
            arrow.lineTo(cx + s, cy + s * 0.5);
        }
        arrow.closeSubpath();
        painter->drawPath(arrow);
        break;
    }
    case IconSetConfig::Flags3: {
        // Simple flag shapes with colors
        static const QColor colors[] = {QColor(220, 50, 50), QColor(240, 200, 40), QColor(60, 180, 75)};
        // Flag pole
        painter->setPen(QPen(QColor(80, 80, 80), 1.5));
        int poleX = iconRect.left() + 3;
        painter->drawLine(poleX, iconRect.top() + 2, poleX, iconRect.bottom() - 2);
        // Flag triangle
        painter->setPen(Qt::NoPen);
        painter->setBrush(colors[vf.iconIndex]);
        QPainterPath flag;
        flag.moveTo(poleX + 1, iconRect.top() + 2);
        flag.lineTo(iconRect.right() - 2, iconRect.top() + 5);
        flag.lineTo(poleX + 1, iconRect.top() + 8);
        flag.closeSubpath();
        painter->drawPath(flag);
        break;
    }
    case IconSetConfig::Stars3: {
        // Stars: empty, half, full
        static const QColor starColor(255, 185, 15);
        painter->setPen(QPen(starColor, 1));
        if (vf.iconIndex == 0) {
            painter->setBrush(Qt::NoBrush);
        } else if (vf.iconIndex == 1) {
            painter->setBrush(QColor(255, 230, 150));
        } else {
            painter->setBrush(starColor);
        }
        // Draw simple star shape
        QPainterPath star;
        double cx = iconRect.center().x();
        double cy = iconRect.center().y();
        double outerR = ICON_SIZE * 0.38;
        double innerR = outerR * 0.38;
        for (int i = 0; i < 5; ++i) {
            double angle = -M_PI / 2.0 + i * 2.0 * M_PI / 5.0;
            double ax = cx + outerR * std::cos(angle);
            double ay = cy + outerR * std::sin(angle);
            if (i == 0) star.moveTo(ax, ay);
            else star.lineTo(ax, ay);
            double angle2 = angle + M_PI / 5.0;
            star.lineTo(cx + innerR * std::cos(angle2), cy + innerR * std::sin(angle2));
        }
        star.closeSubpath();
        painter->drawPath(star);
        break;
    }
    case IconSetConfig::Checkmarks3: {
        // Checkmarks: X (red), ~ (yellow), checkmark (green)
        QFont iconFont = painter->font();
        iconFont.setPointSize(11);
        iconFont.setBold(true);
        painter->setFont(iconFont);
        if (vf.iconIndex == 0) {
            painter->setPen(QColor(220, 50, 50));
            painter->drawText(iconRect, Qt::AlignCenter, QString::fromUtf8("\xe2\x9c\x97")); // X mark
        } else if (vf.iconIndex == 1) {
            painter->setPen(QColor(200, 160, 40));
            painter->drawText(iconRect, Qt::AlignCenter, QString::fromUtf8("~"));
        } else {
            painter->setPen(QColor(60, 160, 60));
            painter->drawText(iconRect, Qt::AlignCenter, QString::fromUtf8("\xe2\x9c\x93")); // check mark
        }
        break;
    }
    }

    painter->setRenderHint(QPainter::Antialiasing, false);
}
