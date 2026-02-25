#ifndef FORMULAPOPUPDELEGATE_H
#define FORMULAPOPUPDELEGATE_H

#include <QStyledItemDelegate>
#include <QPainter>
#include <QApplication>
#include "Theme.h"

// Custom roles for formula popup items
enum FormulaPopupRole {
    FuncNameRole = Qt::UserRole,
    FuncDescRole = Qt::UserRole + 1
};

// Custom delegate that renders each popup row as:
//   [fx]  FUNCNAME        Description text
// with bold name, gray description, and a styled "fx" badge
class FormulaPopupDelegate : public QStyledItemDelegate {
public:
    using QStyledItemDelegate::QStyledItemDelegate;

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
               const QModelIndex& index) const override {
        painter->save();
        painter->setRenderHint(QPainter::Antialiasing, true);

        bool selected = option.state & QStyle::State_Selected;
        bool hovered = option.state & QStyle::State_MouseOver;

        // Background
        const auto& theme = ThemeManager::instance().currentTheme();
        if (selected) {
            painter->fillRect(option.rect, theme.popupItemSelected);
        } else if (hovered) {
            painter->fillRect(option.rect, theme.popupItemSelected.lighter(120));
        } else {
            painter->fillRect(option.rect, theme.popupBackground);
        }

        QRect rect = option.rect.adjusted(8, 0, -8, 0);
        int y = rect.top();
        int h = rect.height();

        // Draw "fx" badge
        QFont fxFont("SF Pro Text", 9);
        fxFont.setItalic(true);
        fxFont.setBold(true);
        painter->setFont(fxFont);
        QRect fxRect(rect.left(), y + (h - 18) / 2, 22, 18);
        painter->setPen(Qt::NoPen);
        painter->setBrush(theme.popupItemSelected);
        painter->drawRoundedRect(fxRect, 3, 3);
        painter->setPen(theme.accentPrimary);
        painter->drawText(fxRect, Qt::AlignCenter, "fx");

        // Draw function name (bold)
        int nameX = fxRect.right() + 10;
        QString funcName = index.data(FuncNameRole).toString();
        QFont nameFont("SF Pro Text", 12);
        nameFont.setBold(true);
        painter->setFont(nameFont);
        painter->setPen(QColor("#202124"));
        QFontMetrics nameFm(nameFont);
        int nameWidth = nameFm.horizontalAdvance(funcName);
        painter->drawText(QRect(nameX, y, nameWidth + 4, h), Qt::AlignVCenter, funcName);

        // Draw description (gray, normal weight)
        int descX = nameX + 130;  // Fixed offset for clean alignment
        QString description = index.data(FuncDescRole).toString();
        QFont descFont("SF Pro Text", 11);
        descFont.setBold(false);
        painter->setFont(descFont);
        painter->setPen(QColor("#5F6368"));
        QFontMetrics descFm(descFont);
        QString elidedDesc = descFm.elidedText(description, Qt::ElideRight, rect.right() - descX);
        painter->drawText(QRect(descX, y, rect.right() - descX, h), Qt::AlignVCenter, elidedDesc);

        // Bottom separator line
        painter->setPen(theme.popupBorder);
        painter->drawLine(option.rect.left() + 8, option.rect.bottom(),
                          option.rect.right() - 8, option.rect.bottom());

        painter->restore();
    }

    QSize sizeHint(const QStyleOptionViewItem& /*option*/,
                   const QModelIndex& /*index*/) const override {
        return QSize(420, 30);
    }
};

#endif // FORMULAPOPUPDELEGATE_H
