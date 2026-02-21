#ifndef CELLDELEGATE_H
#define CELLDELEGATE_H

#include <QStyledItemDelegate>
#include "../core/SparklineConfig.h"

class CellDelegate : public QStyledItemDelegate {
    Q_OBJECT

public:
    explicit CellDelegate(QObject* parent = nullptr);

    QWidget* createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                         const QModelIndex& index) const override;
    void setEditorData(QWidget* editor, const QModelIndex& index) const override;
    void setModelData(QWidget* editor, QAbstractItemModel* model,
                     const QModelIndex& index) const override;
    void updateEditorGeometry(QWidget* editor, const QStyleOptionViewItem& option,
                             const QModelIndex& index) const override;

    bool eventFilter(QObject* object, QEvent* event) override;

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
              const QModelIndex& index) const override;
    QSize sizeHint(const QStyleOptionViewItem& option,
                   const QModelIndex& index) const override;

    void setShowGridlines(bool show) { m_showGridlines = show; }
    bool showGridlines() const { return m_showGridlines; }

signals:
    void formulaEditModeChanged(bool active) const;

private:
    bool m_showGridlines = true;
    void drawSparkline(QPainter* painter, const QRect& rect, const SparklineRenderData& data) const;
};

#endif // CELLDELEGATE_H
