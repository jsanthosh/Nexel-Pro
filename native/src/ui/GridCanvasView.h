#ifndef GRIDCANVASVIEW_H
#define GRIDCANVASVIEW_H

// ============================================================================
// GridCanvasView — Custom QPainter grid (M2 foundation)
// ============================================================================
//
// Replacement for QTableView. The existing SpreadsheetView is built on
// QTableView, which materialises model rows for every populated cell on the
// active area — fine for 100K rows, sluggish at 20M because Qt manages a
// large per-cell widget/index pipeline. This view does not use a model at
// all: it inherits QAbstractScrollArea and paints the visible viewport
// (~30 rows × ~20 cols) directly from ColumnStore via Spreadsheet's data
// API. The cost of rendering is bounded by the viewport, not the document.
//
// Scope of THIS file (foundation commit):
//   • Viewport rendering: cells, gridlines, row+col headers
//   • Vertical + horizontal scrollbar setup and tracking
//   • Single-cell selection by mouse click; arrow-key navigation
//   • Live update when the underlying Spreadsheet changes (signal hookup)
//
// Out of scope here, added in follow-up commits:
//   • Multi-cell range selection, rubber-band drag
//   • In-place cell editing
//   • Merged cells, frozen panes, fill handle, custom column widths,
//     formula bar integration, copy/paste, undo/redo, fancy styling
//
// The existing SpreadsheetView remains the production view. This widget is
// wired behind a View menu toggle so we can A/B test perf on real workloads
// without risking regressions.

#include <QAbstractScrollArea>
#include <memory>

class Spreadsheet;
class QLineEdit;

class GridCanvasView : public QAbstractScrollArea {
    Q_OBJECT

public:
    explicit GridCanvasView(QWidget* parent = nullptr);

    void setSpreadsheet(std::shared_ptr<Spreadsheet> sheet);
    std::shared_ptr<Spreadsheet> spreadsheet() const { return m_sheet; }

    int currentRow() const { return m_currentRow; }
    int currentColumn() const { return m_currentCol; }
    void setCurrentCell(int row, int col);

    // Range selection. The "anchor" is the corner from which Shift-arrows
    // and drag operations extend; the "current" cell is the active one in
    // the range (where editing would happen). Both Excel and our existing
    // SpreadsheetView use the same model.
    int selectionTop()    const { return qMin(m_anchorRow, m_currentRow); }
    int selectionBottom() const { return qMax(m_anchorRow, m_currentRow); }
    int selectionLeft()   const { return qMin(m_anchorCol, m_currentCol); }
    int selectionRight()  const { return qMax(m_anchorCol, m_currentCol); }
    bool hasMultiCellSelection() const {
        return m_anchorRow != m_currentRow || m_anchorCol != m_currentCol;
    }

    // Set anchor + current to (row, col). Default for non-shift mouse clicks.
    void selectCell(int row, int col);
    // Keep anchor where it is; move current. Default for shift+click / shift+arrow.
    void extendSelectionTo(int row, int col);

    // Cell editing. beginEdit shows an inline QLineEdit over the active
    // cell, pre-populated with the cell's current text. If initialChar is
    // non-empty (the user started typing without entering edit mode first),
    // the editor's contents are replaced with that character. commitEdit
    // writes the editor's text back to the spreadsheet and hides the editor;
    // cancelEdit discards changes.
    void beginEdit(const QString& initialChar = QString());
    void commitEdit();
    void cancelEdit();
    bool isEditing() const;

signals:
    void currentCellChanged(int row, int col);
    void selectionChanged(int top, int left, int bottom, int right);
    void cellEdited(int row, int col, const QString& newText);

protected:
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;
    void wheelEvent(QWheelEvent* event) override;

private:
    // Layout constants. Excel defaults are 64 px columns × 20 px rows on
    // 100% zoom; we use a slightly larger row height for readability with
    // modern font rendering.
    static constexpr int kDefaultRowHeight  = 22;
    static constexpr int kDefaultColWidth   = 80;
    static constexpr int kRowHeaderWidth    = 48;
    static constexpr int kColHeaderHeight   = 22;
    // Virtual document extent. Real Excel max is 1,048,576 rows × 16,384
    // columns. We mirror that so scrollbars cover the same range as Excel.
    static constexpr int kVirtualRowCount   = 1048576;
    static constexpr int kVirtualColumnCount = 16384;

    void updateScrollRanges();
    QString colLetter(int col) const;

    // Convert pixel offset (within viewport, exclusive of headers) to a
    // logical row/col index given the current scroll position. Returns the
    // row/col index, or -1 if the point lies in a header.
    int rowAt(int y) const;
    int colAt(int x) const;

    // Cell rect in viewport coordinates (after header offset + scroll).
    // Returns an invalid QRect if the cell isn't currently visible.
    QRect cellRect(int row, int col) const;

    void paintHeaders(QPainter& p, int firstRow, int lastRow,
                      int firstCol, int lastCol);
    void paintCells  (QPainter& p, int firstRow, int lastRow,
                      int firstCol, int lastCol);
    void paintSelection(QPainter& p);

    // Auto-scroll the viewport so (row, col) is visible. Called after every
    // selection change driven by keyboard or drag.
    void scrollToCell(int row, int col);

    // Reposition the in-place editor over the current cell. Called when
    // scroll position changes or the active cell moves while editing.
    void repositionEditor();

    std::shared_ptr<Spreadsheet> m_sheet;
    int m_currentRow = 0;
    int m_currentCol = 0;
    int m_anchorRow  = 0;
    int m_anchorCol  = 0;
    bool m_dragSelecting = false;
    QLineEdit* m_editor = nullptr;
};

#endif // GRIDCANVASVIEW_H
