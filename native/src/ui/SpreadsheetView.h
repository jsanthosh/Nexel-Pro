#ifndef SPREADSHEETVIEW_H
#define SPREADSHEETVIEW_H

#include <QTableView>
#include <QAbstractTableModel>
#include <QColor>
#include <QMenu>
#include <QVariantAnimation>
#include <QTimer>
#include <memory>
#include <functional>
#include "../core/Cell.h"
#include "../core/CellRange.h"

class QLabel;
class Spreadsheet;
class SpreadsheetModel;
class CellDelegate;
class MacroEngine;

class SpreadsheetView : public QTableView {
    Q_OBJECT

public:
    explicit SpreadsheetView(QWidget* parent = nullptr);
    ~SpreadsheetView() = default;

    void setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet);
    std::shared_ptr<Spreadsheet> getSpreadsheet() const;
    void setAllSheets(const std::vector<std::shared_ptr<Spreadsheet>>& sheets) { m_allSheets = sheets; }
    SpreadsheetModel* getModel() const { return m_model; }

    // Editing operations
    void cut();
    void copy();
    void paste();
    void pasteSpecial();
    void deleteSelection();
    void selectAll() override;
    void clearAll();
    void clearContent();
    void clearFormats();

    // Style operations
    void applyBold();
    void applyItalic();
    void applyUnderline();
    void applyStrikethrough();
    void applyFontFamily(const QString& family);
    void applyFontSize(int size);
    void applyForegroundColor(const QString& colorStr);
    void applyBackgroundColor(const QString& colorStr);
    void applyThousandSeparator();
    void applyNumberFormat(const QString& format);
    void applyDateFormat(const QString& dateFormatId);
    void applyCurrencyFormat(const QString& currencyCode);
    void applyAccountingFormat(const QString& currencyCode);
    void increaseDecimals();
    void decreaseDecimals();

    // Alignment
    void applyHAlign(HorizontalAlignment align);
    void applyVAlign(VerticalAlignment align);

    // Indent
    void increaseIndent();
    void decreaseIndent();

    // Text rotation
    void applyTextRotation(int degrees);
    void applyTextOverflow(TextOverflowMode mode);

    // Borders
    void applyBorderStyle(const QString& borderType, const QColor& color = QColor("#000000"), int width = 1, int penStyle = 0);

    // Merge cells
    void mergeCells();
    void unmergeCells();

    // Format painter
    void activateFormatPainter();

    // Sorting
    void sortAscending();
    void sortDescending();

    // Insert/Delete with shift
    void insertCellsShiftRight();
    void insertCellsShiftDown();
    void insertEntireRow();
    void insertEntireColumn();
    void deleteCellsShiftLeft();
    void deleteCellsShiftUp();
    void deleteEntireRow();
    void deleteEntireColumn();

    // Picklist & Checkbox
    void insertPicklist(const QStringList& options, const QStringList& colors = {});
    void showCreatePicklistDialog();
    void insertCheckbox();
    void toggleCheckbox(int row, int col);
    void showPicklistPopup(const QModelIndex& index);
    QStringList resolvePicklistFromRange(const QString& listSourceRange) const;
    void openPicklistManageDialog(int row, int col);
    void openPicklistManagerDialog();

    // Tables
    void applyTableStyle(int themeIndex);

    // Autofit
    void autofitColumn(int column);
    void autofitRow(int row);
    void autofitSelectedColumns();
    void autofitSelectedRows();

    // Macro recording
    void setMacroEngine(MacroEngine* engine) { m_macroEngine = engine; }

    // Formula editing: insert cell reference on click
    void setFormulaEditMode(bool active);
    bool isFormulaEditMode() const { return m_formulaEditMode; }
    void insertCellReference(const QString& ref);

    // Freeze Panes
    void setFrozenRow(int row);
    void setFrozenColumn(int col);

    // Auto Filter
    void toggleAutoFilter();
    void enableAutoFilter(const CellRange& range);
    bool isFilterActive() const { return m_filterActive; }
    void clearAllFilters();

    // Dimension management
    void setRowHeight(int row, int height);
    void setColumnWidth(int col, int width);
    void applyStoredDimensions();

    // Formula recalc flash animation
    double cellAnimationProgress(int row, int col) const;

    // Gridline visibility
    void setGridlinesVisible(bool visible);

    // Chart data range highlighting
    void setChartRangeHighlight(const CellRange& range,
                                const QVector<QPair<int, QColor>>& seriesColumns,
                                const QColor& categoryColor);
    void clearChartRangeHighlight();

    // UI Operations
    void refreshView();
    void zoomIn();
    void zoomOut();
    void resetZoom();

    // Theme support
    void onThemeChanged();

    // Auto-detect contiguous data region from a cell
    CellRange detectDataRegion(int startRow, int startCol) const;

    // Pivot report filter dropdown
    void showPivotFilterDropdown(int filterIndex);

    // Formula view toggle (Ctrl+`)
    void toggleFormulaView();
    bool showFormulas() const { return m_showFormulas; }

    // Bulk loading guard: disables cell-scanning navigation (Ctrl+Arrow etc.)
    void setBulkLoading(bool loading) { m_bulkLoading = loading; }

    // Cell comments
    void insertOrEditComment();
    void deleteComment();

    // Hyperlinks
    void insertOrEditHyperlink();
    void removeHyperlink();
    void openHyperlink(int row, int col);

    // Trace precedents/dependents
    void tracePrecedents();
    void traceDependents();
    void clearTraceArrows();

signals:
    void cellSelected(int row, int col, const QString& content, const QString& address);
    void formatCellsRequested();
    void cellReferenceInserted(const QString& ref);
    void cellReferenceReplaced(const QString& newRef);
    void pivotFilterChanged(int filterIndex, QStringList selectedValues);
    void requestSwitchToSheet(int index);

protected:
    void keyPressEvent(QKeyEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void currentChanged(const QModelIndex& current, const QModelIndex& previous) override;
    void selectionChanged(const QItemSelection& selected, const QItemSelection& deselected) override;
    bool viewportEvent(QEvent* event) override;
    void closeEditor(QWidget* editor, QAbstractItemDelegate::EndEditHint hint) override;
    void commitData(QWidget* editor) override;

private slots:
    void onCellClicked(const QModelIndex& index);
    void onCellDoubleClicked(const QModelIndex& index);
    void onDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight);
    void onHorizontalSectionResized(int logicalIndex, int oldSize, int newSize);
    void onVerticalSectionResized(int logicalIndex, int oldSize, int newSize);

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    std::vector<std::shared_ptr<Spreadsheet>> m_allSheets;  // All workbook sheets (for cross-sheet picklist)
    SpreadsheetModel* m_model = nullptr;
    CellDelegate* m_delegate = nullptr;
    MacroEngine* m_macroEngine = nullptr;
    int m_zoomLevel;
    int m_baseFontSize = 11;

    // Format painter state
    bool m_formatPainterActive = false;
    CellStyle m_copiedStyle;

    // Internal clipboard (retains formatting)
    struct ClipboardCell {
        QVariant value;
        CellStyle style;
        CellType type;
        QString formula;
        CellAddress sourceAddr;  // original cell position for formula ref adjustment
    };
    std::vector<std::vector<ClipboardCell>> m_internalClipboard;
    QString m_internalClipboardText;

    // Fill series state
    bool m_fillDragging = false;
    QModelIndex m_fillDragStart;
    QPoint m_fillDragCurrent;
    QRect m_fillHandleRect;

    // Multi-resize guard
    bool m_resizingMultiple = false;

    // Resize tooltip
    QLabel* m_resizeTooltip = nullptr;
    QTimer* m_resizeTooltipTimer = nullptr;
    void showResizeTooltip(const QPoint& globalPos, const QString& text);
    void hideResizeTooltip();

    // Auto filter state
    bool m_filterActive = false;
    int m_filterHeaderRow = 0;
    CellRange m_filterRange;
    std::map<int, QSet<QString>> m_columnFilters; // col -> set of visible values (empty = all visible)
    void applyFilters();
    void showFilterDropdown(int column);

    // Formula edit mode: when active, clicking cells inserts references
    bool m_formulaEditMode = false;
    QModelIndex m_formulaEditCell;  // The cell being edited with the formula

    // Formula range drag state
    bool m_formulaRangeDragging = false;
    QModelIndex m_formulaRangeStart;
    QModelIndex m_formulaRangeEnd;
    int m_formulaRefInsertPos = -1;
    int m_formulaRefInsertLen = 0;

    void initializeView();
    void applyGridStylesheet();
    void setupConnections();
    void emitCellSelected(const QModelIndex& index);
    void setupHeaderContextMenus();
    void showCellContextMenu(const QPoint& pos);

    // Fill series helpers
    bool isOverFillHandle(const QPoint& pos) const;
    void performFillSeries();
    QRect getSelectionBoundingRect() const;

    // Efficient style application: only process occupied cells for large selections
    void applyStyleChange(std::function<void(CellStyle&)> modifier, const QList<int>& roles);
    // Check if ALL cells in selection match a style predicate (for Excel-style toggle)
    bool selectionAllMatch(std::function<bool(const CellStyle&)> predicate) const;

    // Freeze pane overlay views
    int m_frozenRow = -1;
    int m_frozenCol = -1;
    QTableView* m_frozenCornerView = nullptr;
    QTableView* m_frozenRowView = nullptr;
    QTableView* m_frozenColView = nullptr;
    QWidget* m_freezeHLine = nullptr;
    QWidget* m_freezeVLine = nullptr;
    QList<QMetaObject::Connection> m_freezeConnections;

    QTableView* createFreezeOverlay();
    void setupFreezeViews();
    void destroyFreezeViews();
    void updateFreezeGeometry();

    // Formula cell flash: simple timer-driven hold + fade
    struct CellAnim {
        double progress = 1.0;    // 1.0 = full highlight, 0.0 = gone
        int elapsedMs = 0;        // time since flash started
    };
    static constexpr int FLASH_HOLD_MS = 250;    // hold full yellow for 0.25s
    static constexpr int FLASH_FADE_MS = 500;     // fade out over 0.5s
    static constexpr int FLASH_TICK_MS = 30;      // timer tick interval
    QMap<QPair<int,int>, CellAnim> m_cellAnimations;
    QTimer* m_flashTimer = nullptr;
    void startCellFlashAnimation(int row, int col);
    void onFlashTimerTick();

    // Chart data range highlight state
    struct ChartRangeHighlight {
        CellRange fullRange;
        QVector<QPair<int, QColor>> seriesColumns; // col index → series color
        QColor categoryColor;
    };
    ChartRangeHighlight m_chartHighlight;
    bool m_chartHighlightActive = false;

    // Formula view toggle
    bool m_showFormulas = false;

    // Bulk loading: when true, skip operations that scan m_cells (race with bg thread)
    bool m_bulkLoading = false;

    // Picklist popup re-entry guard
    bool m_picklistPopupOpen = false;

    // Trace arrows state
    bool m_showPrecedents = false;
    bool m_showDependents = false;
    CellAddress m_traceCell;
    std::vector<CellAddress> m_tracedCells;
    void drawTraceArrows(QPainter& painter);
};

#endif // SPREADSHEETVIEW_H
