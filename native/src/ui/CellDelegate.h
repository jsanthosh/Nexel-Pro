#ifndef CELLDELEGATE_H
#define CELLDELEGATE_H

#include <QStyledItemDelegate>
#include "../core/SparklineConfig.h"
#include "../core/ConditionalFormatting.h"

class SpreadsheetView;

class CellDelegate : public QStyledItemDelegate {
    Q_OBJECT

public:
    explicit CellDelegate(QObject* parent = nullptr);
    void setSpreadsheetView(SpreadsheetView* view) { m_spreadsheetView = view; }

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

    void setFormulaEditMode(bool active) { m_formulaEditMode = active; }
    void onThemeChanged() {} // QPainter reads theme at paint time

signals:
    void formulaEditModeChanged(bool active) const;

private:
    bool m_showGridlines = true;
    bool m_formulaEditMode = false;
    SpreadsheetView* m_spreadsheetView = nullptr;
    void drawSparkline(QPainter* painter, const QRect& rect, const SparklineRenderData& data) const;
    void drawCheckbox(QPainter* painter, const QRect& rect, bool checked) const;
    void drawPicklistTags(QPainter* painter, const QRect& rect, const QString& value,
                          const QStringList& allOptions,
                          const QStringList& optionColors = QStringList(),
                          Qt::Alignment alignment = Qt::AlignLeft | Qt::AlignVCenter) const;
    static QColor tagBgColor(int index);
    static QColor tagTextColor(int index);

    // Visual conditional formatting renderers
    void drawDataBar(QPainter* painter, const QRect& rect, const VisualFormatResult& vf,
                     const QString& text, int alignment) const;
    void drawColorScaleBackground(QPainter* painter, const QRect& rect, const VisualFormatResult& vf) const;
    void drawIconSet(QPainter* painter, const QRect& rect, const VisualFormatResult& vf) const;
    static constexpr int ICON_SIZE = 16;
    static constexpr int ICON_MARGIN = 4;
};

#endif // CELLDELEGATE_H
