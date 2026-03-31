#ifndef SPREADSHEETVIEW_H
#define SPREADSHEETVIEW_H

#include <QTableView>
#include <QAbstractTableModel>
#include <QColor>
#include <QMenu>
#include <QVariantAnimation>
#include <QTimer>
#include <QScrollBar>
#include <QPlainTextEdit>
#include <memory>
#include <functional>
#include "../core/Cell.h"
#include "../core/CellRange.h"
#include "../core/Spreadsheet.h"

class QLabel;
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

    // Sheet protection
    void protectSheet(const QString& password = QString());
    void unprotectSheet(const QString& password = QString());
    bool isSheetProtected() const;

    // Sorting
    void sortAscending();
    void sortDescending();
    void showSortDialog();

    // Insert/Delete with shift
    void insertCellsShiftRight();
    void insertCellsShiftDown();
    void insertEntireRow();
    void insertEntireColumn();
    void deleteCellsShiftLeft();
    void deleteCellsShiftUp();
    void deleteEntireRow();
    void deleteEntireColumn();

    // Double-click cursor positioning
    bool wasEditTriggeredByDoubleClick() const { return m_editTriggeredByDoubleClick; }
    QPoint lastDoubleClickPos() const { return m_lastDoubleClickPos; }
    void clearDoubleClickFlag() { m_editTriggeredByDoubleClick = false; }

    // Picklist & Checkbox
    void insertPicklist(const QStringList& options, const QStringList& colors = {});
    void showCreatePicklistDialog();
    void insertCheckbox();
    void toggleCheckbox(int row, int col);
    void showPicklistPopup(const QModelIndex& index);
    void showValidationDropdown(const QModelIndex& idx, const Spreadsheet::DataValidationRule* rule);
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

    // Virtual scrollbar management (for large datasets)
    void refreshVirtualScrollBar();
    void scrollTo(const QModelIndex& index, ScrollHint hint = EnsureVisible) override;

    // UI Operations
    void refreshView();
    void zoomIn();
    void zoomOut();
    void resetZoom();
    void setZoomLevel(int percent);
    int getZoomLevel() const { return m_zoomLevel; }

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

    // Go To Special
    void goToSpecial();

    // Remove Duplicates
    void removeDuplicates();

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

    // Outline/grouping support
    void groupSelectedRows();
    void ungroupSelectedRows();
    void groupSelectedColumns();
    void ungroupSelectedColumns();

signals:
    void cellSelected(int row, int col, const QString& content, const QString& address);
    void formatCellsRequested();
    void goToRequested();
    void cellReferenceInserted(const QString& ref);
    void cellReferenceReplaced(const QString& newRef);
    void pivotFilterChanged(int filterIndex, QStringList selectedValues);
    void requestSwitchToSheet(int index);
    void virtualViewportChanged();  // emitted when virtual scroll shifts viewport
    void zoomChanged(int percent);  // emitted when zoom level changes
    void editModeChanged(const QString& mode);  // "Ready", "Edit", or "Enter"

protected:
    void keyPressEvent(QKeyEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseDoubleClickEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void scrollContentsBy(int dx, int dy) override;
    void wheelEvent(QWheelEvent* event) override;
    void currentChanged(const QModelIndex& current, const QModelIndex& previous) override;
    void selectionChanged(const QItemSelection& selected, const QItemSelection& deselected) override;
    bool viewportEvent(QEvent* event) override;
    void closeEditor(QWidget* editor, QAbstractItemDelegate::EndEditHint hint) override;
    void commitData(QWidget* editor) override;
    void paintOutlineGutter(QPainter& painter);

private slots:
    void onCellClicked(const QModelIndex& index);
    void onCellDoubleClicked(const QModelIndex& index);
    void onDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight);
    void onHorizontalSectionResized(int logicalIndex, int oldSize, int newSize);
    void onVerticalSectionResized(int logicalIndex, int oldSize, int newSize);

    // Navigate to a logical row in virtual mode (updates viewport if needed)
    void navigateToLogicalRow(int logicalRow, int col);

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

    // Virtual scroll: independent scrollbar for full-range navigation
    QScrollBar* m_virtualScrollBar = nullptr;
    bool m_recentering = false;  // guard to prevent recursive recenter in scrollContentsBy
    int m_virtualFocusLogicalRow = -1;  // logical row of focused cell (persists across viewport shifts)
    int m_virtualFocusCol = -1;
    void setupVirtualScrollBar();
    void updateVirtualScrollBarRange();

    // Virtual scroll debounce
    QTimer* m_scrollDebounceTimer = nullptr;
    int m_pendingViewportStart = -1;

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

    // Translate model row to logical row (virtual mode safe)
    int logicalRow(int modelRow) const;
    int logicalRow(const QModelIndex& idx) const;

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

    // Double-click cursor positioning
    QPoint m_lastDoubleClickPos;
    bool m_editTriggeredByDoubleClick = false;

    // Marching ants for clipboard (Ctrl+C animated dashed border)
    QTimer* m_marchingAntsTimer = nullptr;
    int m_marchingAntsOffset = 0;
    QRect m_clipboardRange;  // row/col bounds of copied range (top, left, bottom, right stored as x=left, y=top, width=right, height=bottom)
    bool m_hasClipboardRange = false;
    void clearClipboardRange();
    void onMarchingAntsTick();

    // Trace arrows state
    bool m_showPrecedents = false;
    bool m_showDependents = false;
    CellAddress m_traceCell;
    std::vector<CellAddress> m_tracedCells;
    void drawTraceArrows(QPainter& painter);

    // Outline gutter state
    static constexpr int OUTLINE_GUTTER_WIDTH = 16; // pixels per outline level
    int outlineGutterTotalWidth() const;
    void handleOutlineGutterClick(const QPoint& pos);

    // Multiline cell editor (QPlainTextEdit overlay for Alt+Enter)
    QPlainTextEdit* m_multilineEditor = nullptr;
    QModelIndex m_multilineIndex;
    void openMultilineEditor(const QModelIndex& idx, const QString& text, int cursorPos);
    void commitMultilineEditor();
    void cancelMultilineEditor();
};

#endif // SPREADSHEETVIEW_H
