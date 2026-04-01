#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "CellDelegate.h"
#include "PasteSpecialDialog.h"
#include "SortDialog.h"
#include "GoToSpecialDialog.h"
#include "RemoveDuplicatesDialog.h"
#include "Theme.h"
#include "../core/Spreadsheet.h"
#include <algorithm>
#include "../core/UndoManager.h"
#include "../core/MacroEngine.h"
#include "../core/FillSeries.h"
#include "../core/TableStyle.h"
#include "../core/PivotEngine.h"
#include "../services/DocumentService.h"
#include <QHeaderView>
#include <QKeyEvent>
#include <QClipboard>
#include <QApplication>
#include <QPainter>
#include <QMenu>
#include <QFontMetrics>
#include <QSet>
#include <QLineEdit>
#include <QListWidget>
#include <QDate>
#include <QTime>
#include <QCheckBox>
#include <QRegularExpression>
#include <QMimeData>
#include <QVBoxLayout>
#include <QDialogButtonBox>
#include <QScrollArea>
#include <QDialog>
#include <QLabel>
#include <QPushButton>
#include <QAbstractButton>
#include <QScrollBar>
#include <QTextEdit>
#include <QColorDialog>
#include <QComboBox>
#include <QFormLayout>
#include <QRadioButton>
#include <QShortcut>
#include <QEventLoop>
#include <QDir>
#include <QInputDialog>
#include <QToolTip>
#include <QHelpEvent>
#include <QDesktopServices>
#include <QUrl>
#include <QMessageBox>
#include <algorithm>
#include <cmath>

// Forward declaration — defined further below
static QString adjustFormulaReferences(const QString& formula, int rowOffset, int colOffset);

SpreadsheetView::SpreadsheetView(QWidget* parent)
    : QTableView(parent), m_zoomLevel(100) {

    m_spreadsheet = DocumentService::instance().getCurrentSpreadsheet();
    if (!m_spreadsheet) {
        m_spreadsheet = std::make_shared<Spreadsheet>();
    }

    initializeView();
    setupConnections();
}

void SpreadsheetView::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet) {
    destroyFreezeViews();
    m_frozenRow = -1;
    m_frozenCol = -1;

    // Clear auto-filter state from previous sheet
    m_filterActive = false;
    m_columnFilters.clear();
    disconnect(horizontalHeader(), &QHeaderView::sectionClicked, this, nullptr);

    // Detach view from old model BEFORE deleting it — avoids Qt processing
    // teardown signals on a large model (e.g. 90K rows), which causes lag.
    setModel(nullptr);
    clearSpans();
    delete m_model;
    m_model = nullptr;

    m_spreadsheet = spreadsheet;

    // Register callback for formula recalculation flash animation
    if (m_spreadsheet) {
        m_spreadsheet->onDependentsRecalculated = [this](const std::vector<CellAddress>& cells) {
            for (const auto& addr : cells) {
                startCellFlashAnimation(addr.row, addr.col);
            }
        };
    }

    m_model = new SpreadsheetModel(m_spreadsheet, this);

    // In virtual mode, Qt sees a WINDOW_SIZE buffer of rows and scrolls natively.
    // A separate virtual scrollbar shows position in the full dataset.
    if (m_model->isVirtualMode()) {
        int rh = verticalHeader()->defaultSectionSize();
        if (rh <= 0) rh = 25;
        int visible = viewport()->height() / rh + 2;
        m_model->setVisibleRows(std::max(20, visible));
        // Use Qt's native scrollbar for smooth pixel scrolling within the buffer
        setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    } else {
        setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    }

    setModel(m_model);

    // Set up virtual scrollbar for full-range navigation
    if (m_model->isVirtualMode()) {
        setupVirtualScrollBar();
    } else if (m_virtualScrollBar) {
        m_virtualScrollBar->hide();
    }

    // Auto-size row header width based on max row number
    if (m_spreadsheet) {
        int totalRows = m_spreadsheet->getRowCount();
        QString maxLabel = QString::number(totalRows);
        QFontMetrics fm(verticalHeader()->font());
        int labelWidth = fm.horizontalAdvance(maxLabel) + 20; // padding
        verticalHeader()->setFixedWidth(std::max(40, labelWidth));
    }

    // Re-apply merged regions from the new spreadsheet
    if (m_spreadsheet) {
        for (const auto& mr : m_spreadsheet->getMergedRegions()) {
            int r0 = mr.range.getStart().row;
            int c0 = mr.range.getStart().col;
            int rowSpan = mr.range.getEnd().row - r0 + 1;
            int colSpan = mr.range.getEnd().col - c0 + 1;
            setSpan(r0, c0, rowSpan, colSpan);
        }
    }
}

std::shared_ptr<Spreadsheet> SpreadsheetView::getSpreadsheet() const {
    return m_spreadsheet;
}

int SpreadsheetView::logicalRow(int modelRow) const {
    return (m_model && m_model->isVirtualMode()) ? m_model->toLogicalRow(modelRow) : modelRow;
}
int SpreadsheetView::logicalRow(const QModelIndex& idx) const {
    return logicalRow(idx.row());
}

void SpreadsheetView::refreshVirtualScrollBar() {
    if (m_model && m_model->isVirtualMode()) {
        setupVirtualScrollBar();
    }
}

void SpreadsheetView::setupVirtualScrollBar() {
    if (!m_virtualScrollBar) {
        m_virtualScrollBar = new QScrollBar(Qt::Vertical, this);
        connect(m_virtualScrollBar, &QScrollBar::valueChanged, this, [this](int value) {
            if (!m_model || !m_model->isVirtualMode()) return;
            // Jump the window to show the target logical row
            m_model->jumpToBase(value);
            // Scroll Qt's view to the top of the window
            verticalScrollBar()->setValue(0);
            // Update focus tracking
            if (m_virtualFocusLogicalRow >= 0) {
                int newModelRow = m_model->toModelRow(m_virtualFocusLogicalRow);
                if (newModelRow >= 0 && newModelRow < m_model->rowCount()) {
                    QModelIndex newIdx = m_model->index(newModelRow, m_virtualFocusCol);
                    selectionModel()->setCurrentIndex(newIdx, QItemSelectionModel::ClearAndSelect);
                } else {
                    selectionModel()->clearSelection();
                    selectionModel()->clearCurrentIndex();
                }
            }
            emit virtualViewportChanged();
        });
    }
    updateVirtualScrollBarRange();
    m_virtualScrollBar->setValue(m_model ? m_model->viewportStart() : 0);
    m_virtualScrollBar->show();
    // Position it on the right edge of the widget
    int sbWidth = m_virtualScrollBar->sizeHint().width();
    m_virtualScrollBar->setGeometry(width() - sbWidth, 0, sbWidth, height());
}

void SpreadsheetView::updateVirtualScrollBarRange() {
    if (!m_virtualScrollBar || !m_model) return;
    int totalRows = m_model->totalLogicalRows();
    int visRows = m_model->visibleRows();
    m_virtualScrollBar->setRange(0, std::max(0, totalRows - visRows));
    m_virtualScrollBar->setPageStep(visRows);
    m_virtualScrollBar->setSingleStep(3);
}

void SpreadsheetView::initializeView() {
    m_model = new SpreadsheetModel(m_spreadsheet, this);
    setModel(m_model);

    m_delegate = new CellDelegate(this);
    m_delegate->setSpreadsheetView(this);
    setItemDelegate(m_delegate);

    // Highlight headers when corresponding rows/columns are selected (Excel behavior)
    horizontalHeader()->setHighlightSections(true);
    verticalHeader()->setHighlightSections(true);

    // Column/row sizing
    horizontalHeader()->setDefaultSectionSize(80);
    verticalHeader()->setDefaultSectionSize(25);
    horizontalHeader()->setStretchLastSection(false);
    verticalHeader()->setStretchLastSection(false);
    horizontalHeader()->setMinimumSectionSize(30);
    verticalHeader()->setMinimumSectionSize(14);
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);

    // Delegate handles all painting — disable QTableView gridlines
    setShowGrid(false);
    setSelectionBehavior(QAbstractItemView::SelectItems);
    setSelectionMode(QAbstractItemView::ExtendedSelection);

    setFont(QFont("Arial", 11));

    // Apply theme-aware stylesheet
    applyGridStylesheet();

    // Ensure corner button (top-left) triggers select all
    QAbstractButton* cornerButton = findChild<QAbstractButton*>();
    if (cornerButton) {
        connect(cornerButton, &QAbstractButton::clicked, this, &QTableView::selectAll);
    }

    // Marching ants timer for clipboard range animation
    m_marchingAntsTimer = new QTimer(this);
    connect(m_marchingAntsTimer, &QTimer::timeout, this, &SpreadsheetView::onMarchingAntsTick);

    // Enable mouse tracking for fill handle cursor changes
    viewport()->setMouseTracking(true);

    // Cell context menu (right-click)
    setContextMenuPolicy(Qt::CustomContextMenu);
    connect(this, &QTableView::customContextMenuRequested,
            this, &SpreadsheetView::showCellContextMenu);

    // Setup header context menus
    setupHeaderContextMenus();

    // Double-click column header to autofit column width
    horizontalHeader()->setSectionsClickable(true);
    connect(horizontalHeader(), &QHeaderView::sectionDoubleClicked, this, [this](int logicalIndex) {
        autofitColumn(logicalIndex);
    });

    // Double-click row header to autofit row height
    verticalHeader()->setSectionsClickable(true);
    connect(verticalHeader(), &QHeaderView::sectionDoubleClicked, this, [this](int logicalIndex) {
        autofitRow(logicalIndex);
    });
}

void SpreadsheetView::applyGridStylesheet() {
    const auto& t = ThemeManager::instance().currentTheme();
    setStyleSheet(QString(
        "QTableView {"
        "   background-color: %1;"
        "   border: none;"
        "   outline: none;"
        "}"
        "QTableView::item {"
        "   padding: 0px;"
        "   border: none;"
        "   background-color: transparent;"
        "}"
        "QTableView::item:selected {"
        "   background-color: transparent;"
        "}"
        "QTableView::item:focus {"
        "   border: none;"
        "   outline: none;"
        "}"
        "QHeaderView::section {"
        "   background-color: %2;"
        "   padding: 2px 4px;"
        "   border: none;"
        "   border-right: 1px solid %3;"
        "   border-bottom: 1px solid %3;"
        "   font-size: 11px;"
        "   color: %4;"
        "}"
        "QHeaderView::section:checked {"
        "   background-color: %5;"
        "   color: %6;"
        "   font-weight: 600;"
        "}"
        "QHeaderView {"
        "   background-color: %2;"
        "}"
        "QTableCornerButton::section {"
        "   background-color: %2;"
        "   border: none;"
        "   border-right: 1px solid %3;"
        "   border-bottom: 1px solid %3;"
        "}"
    ).arg(
        t.gridBackground.name(),
        t.headerBackground.name(),
        t.headerBorder.name(),
        t.headerText.name(),
        t.headerSelectedBackground.name(),
        t.headerSelectedText.name()
    ));
}

void SpreadsheetView::onThemeChanged() {
    applyGridStylesheet();
    if (m_delegate) m_delegate->onThemeChanged();

    // Update freeze lines if they exist
    if (m_freezeHLine) {
        const auto& t = ThemeManager::instance().currentTheme();
        m_freezeHLine->setStyleSheet(QString("background: %1;").arg(t.freezeLineColor.name()));
    }
    if (m_freezeVLine) {
        const auto& t = ThemeManager::instance().currentTheme();
        m_freezeVLine->setStyleSheet(QString("background: %1;").arg(t.freezeLineColor.name()));
    }

    viewport()->update();
}

void SpreadsheetView::setupConnections() {
    connect(this, &QTableView::clicked, this, &SpreadsheetView::onCellClicked);
    connect(this, &QTableView::doubleClicked, this, &SpreadsheetView::onCellDoubleClicked);
    if (m_model) {
        connect(m_model, &QAbstractTableModel::dataChanged, this, &SpreadsheetView::onDataChanged);
    }

    // Formula edit mode from cell editor
    if (m_delegate) {
        connect(m_delegate, &CellDelegate::formulaEditModeChanged,
                this, &SpreadsheetView::setFormulaEditMode);
    }

    // Multi-select resize
    connect(horizontalHeader(), &QHeaderView::sectionResized,
            this, &SpreadsheetView::onHorizontalSectionResized);
    connect(verticalHeader(), &QHeaderView::sectionResized,
            this, &SpreadsheetView::onVerticalSectionResized);

}

void SpreadsheetView::setupHeaderContextMenus() {
    horizontalHeader()->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(horizontalHeader(), &QHeaderView::customContextMenuRequested,
            this, [this](const QPoint& pos) {
        QMenu menu;
        int clickedCol = horizontalHeader()->logicalIndexAt(pos);

        QSet<int> selCols;
        QModelIndexList selIdx = selectionModel()->selectedColumns();
        if (selIdx.size() > 1) {
            for (const auto& idx : selIdx) selCols.insert(idx.column());
        } else {
            QModelIndexList allSel = selectionModel()->selectedIndexes();
            for (const auto& idx : allSel) selCols.insert(idx.column());
        }
        if (selCols.isEmpty()) selCols.insert(clickedCol);
        int colCount = selCols.size();

        // Cut / Copy / Paste
        menu.addAction("Cut", this, &SpreadsheetView::cut, QKeySequence::Cut);
        menu.addAction("Copy", this, &SpreadsheetView::copy, QKeySequence::Copy);
        menu.addAction("Paste", this, &SpreadsheetView::paste, QKeySequence::Paste);
        menu.addSeparator();

        // Insert / Delete
        QString insertLabel = colCount > 1 ? QString("Insert %1 Columns").arg(colCount) : "Insert Column";
        QString deleteLabel = colCount > 1 ? QString("Delete %1 Columns").arg(colCount) : "Delete Column";
        menu.addAction(insertLabel, [this]() { insertEntireColumn(); });
        menu.addAction(deleteLabel, [this]() { deleteEntireColumn(); });
        menu.addAction("Clear Contents", this, &SpreadsheetView::clearContent);
        menu.addSeparator();

        // Column Width
        menu.addAction("Column &Width...", [this, clickedCol]() {
            bool ok;
            int currentWidth = horizontalHeader()->sectionSize(clickedCol);
            int newWidth = QInputDialog::getInt(this, "Column Width",
                "Column width (pixels):", currentWidth,
                10, 1000, 1, &ok);
            if (ok) {
                for (int c = 0; c < model()->columnCount(); ++c) {
                    if (selectionModel()->isColumnSelected(c, QModelIndex())) {
                        setColumnWidth(c, newWidth);
                        if (m_spreadsheet) m_spreadsheet->setColumnWidth(c, newWidth);
                    }
                }
                if (!selectionModel()->isColumnSelected(clickedCol, QModelIndex())) {
                    setColumnWidth(clickedCol, newWidth);
                    if (m_spreadsheet) m_spreadsheet->setColumnWidth(clickedCol, newWidth);
                }
            }
        });
        menu.addAction("AutoFit Column Width", this, &SpreadsheetView::autofitSelectedColumns);
        menu.addAction("&Default Width...", [this]() {
            bool ok;
            int defaultWidth = horizontalHeader()->defaultSectionSize();
            int newWidth = QInputDialog::getInt(this, "Default Column Width",
                "Default width (pixels):", defaultWidth, 10, 500, 1, &ok);
            if (ok) {
                horizontalHeader()->setDefaultSectionSize(newWidth);
            }
        });
        menu.addSeparator();

        // Hide / Unhide
        menu.addAction("Hide", [this, selCols]() {
            for (int c : selCols) setColumnHidden(c, true);
        });
        menu.addAction("Unhide", [this, clickedCol]() {
            // Unhide columns around the clicked position
            for (int c = qMax(0, clickedCol - 1); c <= qMin(model()->columnCount() - 1, clickedCol + 1); ++c) {
                if (isColumnHidden(c)) setColumnHidden(c, false);
            }
        });

        menu.addSeparator();
        menu.addAction("Format Cells...", [this]() {
            emit formatCellsRequested();
        }, QKeySequence(Qt::CTRL | Qt::Key_1));

        menu.exec(horizontalHeader()->mapToGlobal(pos));
    });

    verticalHeader()->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(verticalHeader(), &QHeaderView::customContextMenuRequested,
            this, [this](const QPoint& pos) {
        QMenu menu;
        int clickedRow = verticalHeader()->logicalIndexAt(pos);

        QSet<int> selRows;
        QModelIndexList selIdx = selectionModel()->selectedRows();
        if (selIdx.size() > 1) {
            for (const auto& idx : selIdx) selRows.insert(logicalRow(idx));
        } else {
            QModelIndexList allSel = selectionModel()->selectedIndexes();
            for (const auto& idx : allSel) selRows.insert(logicalRow(idx));
        }
        if (selRows.isEmpty()) selRows.insert(clickedRow);
        int rowCount = selRows.size();

        // Cut / Copy / Paste
        menu.addAction("Cut", this, &SpreadsheetView::cut, QKeySequence::Cut);
        menu.addAction("Copy", this, &SpreadsheetView::copy, QKeySequence::Copy);
        menu.addAction("Paste", this, &SpreadsheetView::paste, QKeySequence::Paste);
        menu.addSeparator();

        // Insert / Delete
        QString insertLabel = rowCount > 1 ? QString("Insert %1 Rows").arg(rowCount) : "Insert Row";
        QString deleteLabel = rowCount > 1 ? QString("Delete %1 Rows").arg(rowCount) : "Delete Row";
        menu.addAction(insertLabel, [this]() { insertEntireRow(); });
        menu.addAction(deleteLabel, [this]() { deleteEntireRow(); });
        menu.addAction("Clear Contents", this, &SpreadsheetView::clearContent);
        menu.addSeparator();

        // Row Height
        menu.addAction("Row &Height...", [this, clickedRow]() {
            bool ok;
            int currentHeight = verticalHeader()->sectionSize(clickedRow);
            int newHeight = QInputDialog::getInt(this, "Row Height",
                "Row height (pixels):", currentHeight,
                5, 500, 1, &ok);
            if (ok) {
                for (int r = 0; r < model()->rowCount(); ++r) {
                    if (selectionModel()->isRowSelected(r, QModelIndex())) {
                        setRowHeight(r, newHeight);
                        if (m_spreadsheet) m_spreadsheet->setRowHeight(r, newHeight);
                    }
                }
                if (!selectionModel()->isRowSelected(clickedRow, QModelIndex())) {
                    setRowHeight(clickedRow, newHeight);
                    if (m_spreadsheet) m_spreadsheet->setRowHeight(clickedRow, newHeight);
                }
            }
        });
        menu.addAction("AutoFit Row Height", this, &SpreadsheetView::autofitSelectedRows);
        menu.addAction("De&fault Height...", [this]() {
            bool ok;
            int defaultHeight = verticalHeader()->defaultSectionSize();
            int newHeight = QInputDialog::getInt(this, "Default Row Height",
                "Default height (pixels):", defaultHeight, 5, 500, 1, &ok);
            if (ok) {
                verticalHeader()->setDefaultSectionSize(newHeight);
            }
        });
        menu.addSeparator();

        // Hide / Unhide
        menu.addAction("Hide", [this, selRows]() {
            for (int r : selRows) {
                int modelRow = m_model ? m_model->toModelRow(r) : r;
                if (modelRow >= 0) setRowHidden(modelRow, true);
            }
        });
        menu.addAction("Unhide", [this, clickedRow]() {
            for (int r = qMax(0, clickedRow - 1); r <= qMin(model()->rowCount() - 1, clickedRow + 1); ++r) {
                if (isRowHidden(r)) setRowHidden(r, false);
            }
        });

        menu.addSeparator();
        menu.addAction("Format Cells...", [this]() {
            emit formatCellsRequested();
        }, QKeySequence(Qt::CTRL | Qt::Key_1));

        menu.exec(verticalHeader()->mapToGlobal(pos));
    });
}

void SpreadsheetView::emitCellSelected(const QModelIndex& index) {
    if (!index.isValid() || !m_spreadsheet) return;

    int row = m_model ? m_model->toLogicalRow(index.row()) : index.row();
    CellAddress addr(row, index.column());
    auto cell = m_spreadsheet->getCell(addr);

    QString content;
    if (cell->getType() == CellType::Formula) {
        content = cell->getFormula();
    } else {
        // For date-formatted cells, show the date in MM/dd/yyyy in the formula bar
        const auto& style = cell->getStyle();
        if (style.numberFormat == "Date") {
            bool ok;
            double serial = cell->getValue().toDouble(&ok);
            if (ok && serial > 0 && serial < 200000) {
                static const QDate epoch(1899, 12, 30);
                QDate date = epoch.addDays(static_cast<int>(serial));
                if (date.isValid()) {
                    content = date.toString("MM/dd/yyyy");
                } else {
                    content = cell->getValue().toString();
                }
            } else {
                content = cell->getValue().toString();
            }
        } else {
            content = cell->getValue().toString();
        }
    }

    emit cellSelected(index.row(), index.column(), content, addr.toString());
}

void SpreadsheetView::currentChanged(const QModelIndex& current, const QModelIndex& previous) {
    QTableView::currentChanged(current, previous);

    // Update virtual focus tracking when user clicks/navigates to a new cell
    if (m_model && m_model->isVirtualMode() && current.isValid()) {
        m_virtualFocusLogicalRow = m_model->toLogicalRow(current.row());
        m_virtualFocusCol = current.column();
    }

    // Force repaint of previous cell to clear its focus border and fill handle
    if (previous.isValid()) {
        QRect prevRect = visualRect(previous);
        // Expand to cover 2px focus border + fill handle (8px square at corner)
        viewport()->update(prevRect.adjusted(-2, -2, 7, 7));
    }
    // Also invalidate the old fill handle rect
    if (!m_fillHandleRect.isNull()) {
        viewport()->update(m_fillHandleRect.adjusted(-2, -2, 2, 2));
    }

    emitCellSelected(current);
}

void SpreadsheetView::selectionChanged(const QItemSelection& selected, const QItemSelection& deselected) {
    // Invalidate old fill handle rect before selection updates
    if (!m_fillHandleRect.isNull()) {
        viewport()->update(m_fillHandleRect.adjusted(-2, -2, 2, 2));
    }
    QTableView::selectionChanged(selected, deselected);
    // Full viewport repaint so the selection border rectangle draws correctly
    viewport()->update();
}

// ============== Clipboard operations ==============

void SpreadsheetView::cut() {
    if (m_spreadsheet && m_spreadsheet->isProtected()) {
        // Check if any selected cell is locked
        QModelIndexList selected = selectionModel()->selectedIndexes();
        for (const auto& idx : selected) {
            int row = m_model ? m_model->toLogicalRow(idx.row()) : idx.row();
            auto cell = m_spreadsheet->getCell(row, idx.column());
            if (cell->getStyle().locked) {
                QMessageBox::warning(this, "Protected Sheet",
                    "The cell or chart you're trying to change is on a protected sheet.\n"
                    "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
                return;
            }
        }
    }
    // Excel behavior: cut only copies and marks cells with marching ants.
    // Source cells are cleared only when paste actually happens.
    copy();
    m_isCutOperation = true;
}

void SpreadsheetView::copy() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

    // A new copy cancels any pending cut operation
    m_isCutOperation = false;

    std::sort(selected.begin(), selected.end(), [](const QModelIndex& a, const QModelIndex& b) {
        if (a.row() != b.row()) return a.row() < b.row();
        return a.column() < b.column();
    });

    // Find bounding box
    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    // Build internal clipboard with formatting
    int rows = maxRow - minRow + 1;
    int cols = maxCol - minCol + 1;
    m_internalClipboard.clear();
    m_internalClipboard.resize(rows, std::vector<ClipboardCell>(cols));

    for (const auto& idx : selected) {
        int r = idx.row() - minRow;
        int c = idx.column() - minCol;
        CellAddress addr(logicalRow(idx), idx.column());
        auto cell = m_spreadsheet->getCell(addr);
        m_internalClipboard[r][c].value = cell->getValue();
        m_internalClipboard[r][c].style = cell->getStyle();
        m_internalClipboard[r][c].type = cell->getType();
        m_internalClipboard[r][c].formula = cell->getFormula();
        m_internalClipboard[r][c].sourceAddr = addr;
        m_internalClipboard[r][c].comment = cell->getComment();
        m_internalClipboard[r][c].hyperlink = cell->getHyperlink();
    }

    // Also set system clipboard text for cross-app paste
    QString data;
    int lastRow = selected.first().row();
    bool firstInRow = true;
    for (const auto& index : selected) {
        if (index.row() != lastRow) {
            data += "\n";
            lastRow = index.row();
            firstInRow = true;
        }
        if (!firstInRow) {
            data += "\t";
        }
        data += index.data().toString();
        firstInRow = false;
    }

    m_internalClipboardText = data;
    QApplication::clipboard()->setText(data);

    // Start marching ants animation around copied range
    m_clipboardRange = QRect(minCol, minRow, maxCol, maxRow); // stores col/row bounds
    m_hasClipboardRange = true;
    m_marchingAntsOffset = 0;
    if (m_marchingAntsTimer) m_marchingAntsTimer->start(100);
}

void SpreadsheetView::paste() {
    QString data = QApplication::clipboard()->text();
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int startRow = logicalRow(current);
    int startCol = current.column();

    // Protection check: block paste into locked cells on a protected sheet
    if (m_spreadsheet->isProtected()) {
        auto cell = m_spreadsheet->getCell(startRow, startCol);
        if (cell->getStyle().locked) {
            QMessageBox::warning(this, "Protected Sheet",
                "The cell or chart you're trying to change is on a protected sheet.\n"
                "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
            return;
        }
    }

    std::vector<CellSnapshot> before, after;
    m_model->setSuppressUndo(true);
    m_spreadsheet->beginBatchUpdate();

    // Check if system clipboard matches our internal clipboard (same-app paste with formatting)
    bool useInternalClipboard = !m_internalClipboard.empty() && data == m_internalClipboardText;

    if (useInternalClipboard) {
        for (int r = 0; r < static_cast<int>(m_internalClipboard.size()); ++r) {
            for (int c = 0; c < static_cast<int>(m_internalClipboard[r].size()); ++c) {
                CellAddress addr(startRow + r, startCol + c);
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                const auto& clipCell = m_internalClipboard[r][c];
                if (clipCell.type == CellType::Formula && !clipCell.formula.isEmpty()) {
                    int rowOff = addr.row - clipCell.sourceAddr.row;
                    int colOff = addr.col - clipCell.sourceAddr.col;
                    QString adjusted = adjustFormulaReferences(clipCell.formula, rowOff, colOff);
                    m_spreadsheet->setCellFormula(addr, adjusted);
                } else if (clipCell.value.isValid() && !clipCell.value.toString().isEmpty()) {
                    m_spreadsheet->setCellValue(addr, clipCell.value);
                }
                // Apply formatting
                auto cell = m_spreadsheet->getCell(addr);
                cell->setStyle(clipCell.style);
                // Restore comment and hyperlink
                if (!clipCell.comment.isEmpty()) cell->setComment(clipCell.comment);
                if (!clipCell.hyperlink.isEmpty()) cell->setHyperlink(clipCell.hyperlink);

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
    } else {
        // External paste: try HTML table first (from web browsers, Excel)
        const QMimeData* mimeData = QApplication::clipboard()->mimeData();
        bool htmlHandled = false;

        if (mimeData->hasHtml()) {
            QString html = mimeData->html();
            if (html.contains("<table", Qt::CaseInsensitive) || html.contains("<tr", Qt::CaseInsensitive)) {
                // Parse HTML table
                QRegularExpression rowRx("<tr[^>]*>(.*?)</tr>",
                    QRegularExpression::DotMatchesEverythingOption | QRegularExpression::CaseInsensitiveOption);
                QRegularExpression cellRx("<t[dh][^>]*>(.*?)</t[dh]>",
                    QRegularExpression::DotMatchesEverythingOption | QRegularExpression::CaseInsensitiveOption);

                auto rowMatches = rowRx.globalMatch(html);
                int r = 0;
                while (rowMatches.hasNext()) {
                    auto rowMatch = rowMatches.next();
                    QString rowHtml = rowMatch.captured(1);
                    auto cellMatches = cellRx.globalMatch(rowHtml);
                    int c = 0;
                    while (cellMatches.hasNext()) {
                        auto cellMatch = cellMatches.next();
                        QString cellText = cellMatch.captured(1);
                        // Strip HTML tags from cell content
                        cellText.remove(QRegularExpression("<[^>]*>"));
                        cellText = cellText.trimmed();
                        // Decode HTML entities
                        cellText.replace("&amp;", "&");
                        cellText.replace("&lt;", "<");
                        cellText.replace("&gt;", ">");
                        cellText.replace("&nbsp;", " ");
                        cellText.replace("&quot;", "\"");

                        CellAddress addr(startRow + r, startCol + c);
                        before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                        QModelIndex index = m_model->index(startRow + r, startCol + c);
                        m_model->setData(index, cellText);

                        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                        c++;
                    }
                    r++;
                }
                htmlHandled = true;
            }
        }

        if (!htmlHandled) {
            // Fallback: plain text TSV paste
            QStringList rows = data.split("\n");
            for (int r = 0; r < rows.size(); ++r) {
                QStringList cols = rows[r].split("\t");
                for (int c = 0; c < cols.size(); ++c) {
                    CellAddress addr(startRow + r, startCol + c);
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                    QModelIndex index = m_model->index(startRow + r, startCol + c);
                    m_model->setData(index, cols[c]);

                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            }
        }
    }
    m_spreadsheet->endBatchUpdate();
    m_model->setSuppressUndo(false);

    // If this is a cut-paste, also clear the source cells
    std::vector<CellSnapshot> cutBefore, cutAfter;
    if (m_isCutOperation && !m_internalClipboard.empty()) {
        m_model->setSuppressUndo(true);
        for (int r = 0; r < static_cast<int>(m_internalClipboard.size()); ++r) {
            for (int c = 0; c < static_cast<int>(m_internalClipboard[r].size()); ++c) {
                const auto& clipCell = m_internalClipboard[r][c];
                CellAddress srcAddr = clipCell.sourceAddr;
                // Don't clear the source cell if it's the same as a destination cell
                // (paste in place should not destroy data)
                bool isDestination = false;
                for (int pr = 0; pr < static_cast<int>(m_internalClipboard.size()); ++pr) {
                    for (int pc = 0; pc < static_cast<int>(m_internalClipboard[pr].size()); ++pc) {
                        if (srcAddr.row == startRow + pr && srcAddr.col == startCol + pc) {
                            isDestination = true;
                            break;
                        }
                    }
                    if (isDestination) break;
                }
                if (isDestination) continue;

                cutBefore.push_back(m_spreadsheet->takeCellSnapshot(srcAddr));
                auto cell = m_spreadsheet->getCell(srcAddr);
                cell->setValue(QVariant());
                cell->setFormula(QString());
                cutAfter.push_back(m_spreadsheet->takeCellSnapshot(srcAddr));
            }
        }
        m_model->setSuppressUndo(false);
        m_isCutOperation = false;
        m_internalClipboard.clear();
        m_internalClipboardText.clear();
    }

    // Combine paste + cut-source-clear into a single undo command
    if (!cutBefore.empty()) {
        auto compound = std::make_unique<CompoundUndoCommand>("Cut & Paste");
        compound->addChild(std::make_unique<MultiCellEditCommand>(before, after, "Paste"));
        compound->addChild(std::make_unique<MultiCellEditCommand>(cutBefore, cutAfter, "Clear Cut Source"));
        m_spreadsheet->getUndoManager().pushCommand(std::move(compound));
    } else {
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<MultiCellEditCommand>(before, after, "Paste"));
    }

    // Clear marching ants after paste (Excel behavior)
    clearClipboardRange();

    if (m_model) {
        m_model->resetModel();
    }
}

void SpreadsheetView::pasteSpecial() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    PasteSpecialDialog dlg(this);
    if (dlg.exec() != QDialog::Accepted) return;

    PasteSpecialOptions opts = dlg.getOptions();

    QString data = QApplication::clipboard()->text();
    int startRow = logicalRow(current);
    int startCol = current.column();

    std::vector<CellSnapshot> before, after;
    m_model->setSuppressUndo(true);

    bool useInternalClipboard = !m_internalClipboard.empty() && data == m_internalClipboardText;

    // Determine clipboard dimensions
    int clipRows = 0, clipCols = 0;
    std::vector<QStringList> externalRows;
    if (useInternalClipboard) {
        clipRows = static_cast<int>(m_internalClipboard.size());
        for (const auto& row : m_internalClipboard)
            clipCols = qMax(clipCols, static_cast<int>(row.size()));
    } else {
        QStringList lines = data.split("\n");
        for (const auto& line : lines) {
            externalRows.push_back(line.split("\t"));
            clipCols = qMax(clipCols, externalRows.back().size());
        }
        clipRows = static_cast<int>(externalRows.size());
    }

    // With transpose, swap rows/cols iteration
    int pasteRows = opts.transpose ? clipCols : clipRows;
    int pasteCols = opts.transpose ? clipRows : clipCols;

    for (int pr = 0; pr < pasteRows; ++pr) {
        for (int pc = 0; pc < pasteCols; ++pc) {
            // Map paste coords back to clipboard coords
            int cr = opts.transpose ? pc : pr;
            int cc = opts.transpose ? pr : pc;

            CellAddress addr(startRow + pr, startCol + pc);

            // Get clipboard cell data
            QVariant clipValue;
            CellStyle clipStyle;
            CellType clipType = CellType::Empty;
            QString clipFormula;
            CellAddress clipSourceAddr;

            if (useInternalClipboard) {
                if (cr >= static_cast<int>(m_internalClipboard.size())) continue;
                if (cc >= static_cast<int>(m_internalClipboard[cr].size())) continue;
                const auto& clipCell = m_internalClipboard[cr][cc];
                clipValue = clipCell.value;
                clipStyle = clipCell.style;
                clipType = clipCell.type;
                clipFormula = clipCell.formula;
                clipSourceAddr = clipCell.sourceAddr;
            } else {
                if (cr >= static_cast<int>(externalRows.size())) continue;
                if (cc >= externalRows[cr].size()) continue;
                clipValue = externalRows[cr][cc];
                clipType = CellType::Text;
            }

            // Skip blanks: if clipboard cell is empty, don't overwrite
            if (opts.skipBlanks && clipType == CellType::Empty &&
                (!clipValue.isValid() || clipValue.toString().isEmpty())) {
                continue;
            }

            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            if (opts.pasteType == PasteSpecialOptions::ColumnWidths) {
                // Apply column width from clipboard style
                if (useInternalClipboard) {
                    setColumnWidth(startCol + pc, clipStyle.columnWidth);
                }
            } else if (opts.pasteType == PasteSpecialOptions::Formats) {
                // Only paste styles, not values
                if (useInternalClipboard) {
                    auto cell = m_spreadsheet->getCell(addr);
                    cell->setStyle(clipStyle);
                }
            } else if (opts.pasteType == PasteSpecialOptions::Values) {
                // Paste values only (resolve formulas to their computed values)
                QVariant valToPaste = clipValue;
                if (useInternalClipboard && clipType == CellType::Formula) {
                    // For formula cells, use the computed value from the clipboard source
                    auto srcCell = m_spreadsheet->getCellIfExists(clipSourceAddr);
                    if (srcCell) valToPaste = srcCell->getComputedValue();
                }
                if (opts.operation != PasteSpecialOptions::OpNone) {
                    auto existingCell = m_spreadsheet->getCellIfExists(addr);
                    double existing = existingCell ? existingCell->getValue().toDouble() : 0.0;
                    double clipVal = valToPaste.toDouble();
                    switch (opts.operation) {
                        case PasteSpecialOptions::Add: valToPaste = existing + clipVal; break;
                        case PasteSpecialOptions::Subtract: valToPaste = existing - clipVal; break;
                        case PasteSpecialOptions::Multiply: valToPaste = existing * clipVal; break;
                        case PasteSpecialOptions::Divide:
                            valToPaste = (clipVal != 0.0) ? existing / clipVal : QVariant("#DIV/0!");
                            break;
                        default: break;
                    }
                }
                m_spreadsheet->setCellValue(addr, valToPaste);
            } else if (opts.pasteType == PasteSpecialOptions::Formulas) {
                // Paste formulas (with reference adjustment) but not styles
                if (useInternalClipboard && clipType == CellType::Formula && !clipFormula.isEmpty()) {
                    int rowOff = addr.row - clipSourceAddr.row;
                    int colOff = addr.col - clipSourceAddr.col;
                    QString adjusted = adjustFormulaReferences(clipFormula, rowOff, colOff);
                    m_spreadsheet->setCellFormula(addr, adjusted);
                } else {
                    QVariant valToPaste = clipValue;
                    if (opts.operation != PasteSpecialOptions::OpNone) {
                        auto existingCell = m_spreadsheet->getCellIfExists(addr);
                        double existing = existingCell ? existingCell->getValue().toDouble() : 0.0;
                        double clipVal = valToPaste.toDouble();
                        switch (opts.operation) {
                            case PasteSpecialOptions::Add: valToPaste = existing + clipVal; break;
                            case PasteSpecialOptions::Subtract: valToPaste = existing - clipVal; break;
                            case PasteSpecialOptions::Multiply: valToPaste = existing * clipVal; break;
                            case PasteSpecialOptions::Divide:
                                valToPaste = (clipVal != 0.0) ? existing / clipVal : QVariant("#DIV/0!");
                                break;
                            default: break;
                        }
                    }
                    if (valToPaste.isValid() && !valToPaste.toString().isEmpty())
                        m_spreadsheet->setCellValue(addr, valToPaste);
                }
            } else {
                // All: paste everything (same as normal paste)
                if (useInternalClipboard && clipType == CellType::Formula && !clipFormula.isEmpty()) {
                    int rowOff = addr.row - clipSourceAddr.row;
                    int colOff = addr.col - clipSourceAddr.col;
                    QString adjusted = adjustFormulaReferences(clipFormula, rowOff, colOff);
                    m_spreadsheet->setCellFormula(addr, adjusted);
                } else {
                    QVariant valToPaste = clipValue;
                    if (opts.operation != PasteSpecialOptions::OpNone) {
                        auto existingCell = m_spreadsheet->getCellIfExists(addr);
                        double existing = existingCell ? existingCell->getValue().toDouble() : 0.0;
                        double clipVal = valToPaste.toDouble();
                        switch (opts.operation) {
                            case PasteSpecialOptions::Add: valToPaste = existing + clipVal; break;
                            case PasteSpecialOptions::Subtract: valToPaste = existing - clipVal; break;
                            case PasteSpecialOptions::Multiply: valToPaste = existing * clipVal; break;
                            case PasteSpecialOptions::Divide:
                                valToPaste = (clipVal != 0.0) ? existing / clipVal : QVariant("#DIV/0!");
                                break;
                            default: break;
                        }
                    }
                    if (valToPaste.isValid() && !valToPaste.toString().isEmpty())
                        m_spreadsheet->setCellValue(addr, valToPaste);
                }
                if (useInternalClipboard) {
                    auto cell = m_spreadsheet->getCell(addr);
                    cell->setStyle(clipStyle);
                }
            }

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        }
    }

    m_model->setSuppressUndo(false);
    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Paste Special"));

    if (m_model) {
        m_model->resetModel();
    }
}

void SpreadsheetView::deleteSelection() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

    // Protection check: block delete on locked cells of a protected sheet
    if (m_spreadsheet->isProtected()) {
        for (const auto& idx : selected) {
            int row = m_model ? m_model->toLogicalRow(idx.row()) : idx.row();
            auto cell = m_spreadsheet->getCell(row, idx.column());
            if (cell->getStyle().locked) {
                QMessageBox::warning(this, "Protected Sheet",
                    "The cell or chart you're trying to change is on a protected sheet.\n"
                    "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
                return;
            }
        }
    }

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);
    for (const auto& index : selected) {
        CellAddress addr(logicalRow(index), index.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));

        auto cell = m_spreadsheet->getCell(addr);
        QString fmt = cell->getStyle().numberFormat;

        // If cell is a Checkbox or Picklist, also clear the special format
        if (fmt == "Checkbox" || fmt == "Picklist") {
            CellStyle style = cell->getStyle();
            style.numberFormat = "General";
            cell->setStyle(style);
            cell->setValue(QVariant());
            // Remove associated validation rule for picklists
            if (fmt == "Picklist") {
                auto& rules = m_spreadsheet->getValidationRules();
                for (int ri = (int)rules.size() - 1; ri >= 0; --ri) {
                    if (rules[ri].range.contains(addr)) {
                        m_spreadsheet->removeValidationRule(ri);
                        break;
                    }
                }
            }
        } else {
            m_model->setData(index, "");
        }

        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }
    m_model->setSuppressUndo(false);

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Delete"));

    // Refresh view for checkbox/picklist cells that were cleared
    if (m_model) {
        QModelIndex tl = m_model->index(selected.first().row(), selected.first().column());
        QModelIndex br = m_model->index(selected.last().row(), selected.last().column());
        emit m_model->dataChanged(tl, br);
    }
}

void SpreadsheetView::selectAll() {
    // Excel-style two-step Ctrl+A:
    // 1st press: select the contiguous data region around the current cell
    // 2nd press (if data region already selected): select entire sheet
    if (!m_spreadsheet || !currentIndex().isValid()) {
        QTableView::selectAll();
        return;
    }

    int curRow = logicalRow(currentIndex());
    int curCol = currentIndex().column();
    CellRange dataRegion = detectDataRegion(curRow, curCol);

    // Check if current selection already matches the data region
    QItemSelection currentSel = selectionModel()->selection();
    bool alreadyMatchesRegion = false;
    if (!currentSel.isEmpty()) {
        QItemSelectionRange range = currentSel.first();
        alreadyMatchesRegion = (range.top() == dataRegion.getStart().row
            && range.left() == dataRegion.getStart().col
            && range.bottom() == dataRegion.getEnd().row
            && range.right() == dataRegion.getEnd().col);
    }

    // If data region is just a single empty cell, or selection already matches region, select all
    if (alreadyMatchesRegion || dataRegion.isSingleCell()) {
        QTableView::selectAll();
    } else {
        // Select the data region
        QModelIndex topLeft = model()->index(dataRegion.getStart().row, dataRegion.getStart().col);
        QModelIndex bottomRight = model()->index(dataRegion.getEnd().row, dataRegion.getEnd().col);
        selectionModel()->select(QItemSelection(topLeft, bottomRight),
            QItemSelectionModel::ClearAndSelect);
        selectionModel()->setCurrentIndex(
            model()->index(currentIndex().row(), currentIndex().column()),
            QItemSelectionModel::NoUpdate);
    }
}

// ============== Style operations ==============

// Efficient style application: for large selections (select all), only iterate occupied cells
void SpreadsheetView::applyStyleChange(std::function<void(CellStyle&)> modifier, const QList<int>& roles) {
    if (!m_spreadsheet) return;

    QItemSelection sel = selectionModel()->selection();
    if (sel.isEmpty()) return;

    // Compute bounding box from selection ranges — O(ranges) not O(cells)
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    int totalCells = 0;
    for (const auto& range : sel) {
        minRow = qMin(minRow, range.top());
        maxRow = qMax(maxRow, range.bottom());
        minCol = qMin(minCol, range.left());
        maxCol = qMax(maxCol, range.right());
        totalCells += (range.bottom() - range.top() + 1) * (range.right() - range.left() + 1);
    }

    // In virtual mode, translate model rows to logical rows
    if (m_model && m_model->isVirtualMode()) {
        minRow = m_model->toLogicalRow(minRow);
        maxRow = m_model->toLogicalRow(maxRow);
    }

    // Detect full-column or large selection:
    // If selection spans all visible rows (full column click) OR >1000 cells,
    // use the instant overlay path. In virtual mode, a column click = all model rows.
    int modelRowCount = m_model ? m_model->rowCount() : 100;
    bool isFullColumnOrRow = (maxRow - minRow + 1 >= modelRowCount - 1);
    bool isLargeEnough = totalCells > 1000;
    bool useFastPath = isFullColumnOrRow || totalCells > 5000;

    // For full column selection in virtual mode, extend to ALL logical rows
    if (isFullColumnOrRow && m_model && m_model->isVirtualMode()) {
        minRow = 0;
        maxRow = m_model->totalLogicalRows() - 1;
    }

    static constexpr int LARGE_SELECTION_THRESHOLD = 5000;
    bool isLargeSelection = totalCells > LARGE_SELECTION_THRESHOLD && !useFastPath;
    bool isMassiveSelection = useFastPath;

    // Debug logging removed for performance

    std::vector<CellSnapshot> before, after;

    if (isMassiveSelection) {
        // FRONTEND FIRST: add style overlay → instant visual for ALL cells in region
        Spreadsheet::StyleOverlay overlay;
        overlay.minRow = minRow;
        overlay.maxRow = maxRow;
        overlay.minCol = minCol;
        overlay.maxCol = maxCol;
        overlay.modifier = modifier;
        m_spreadsheet->addStyleOverlay(overlay);

        // Only update sheet default style if ALL columns are selected (true "select all")
        int totalCols = m_model ? m_model->columnCount() : m_spreadsheet->getColumnCount();
        if (minCol == 0 && maxCol >= totalCols - 1) {
            CellStyle defaultStyle = m_spreadsheet->getDefaultCellStyle();
            modifier(defaultStyle);
            m_spreadsheet->setDefaultCellStyle(defaultStyle);
        }

        // Tell Qt that visible cell data changed → triggers immediate repaint
        if (m_model) {
            int visRows = m_model->rowCount();
            int visCols = m_model->columnCount();
            QModelIndex topLeft = m_model->index(0, 0);
            QModelIndex bottomRight = m_model->index(visRows - 1, visCols - 1);
            emit m_model->dataChanged(topLeft, bottomRight);
        }
        viewport()->repaint();

        // BACKEND LATER: apply to actual cell styles in background chunks
        auto modifierCopy = modifier;
        auto spreadsheet = m_spreadsheet;
        int bgMinCol = minCol, bgMaxCol = maxCol;
        int chunkStart = minRow;
        int bgMaxRow = maxRow;
        static constexpr int CHUNK_SIZE = 50000;

        auto* timer = new QTimer(this);
        timer->setInterval(0);
        connect(timer, &QTimer::timeout, this, [=]() mutable {
            if (chunkStart > bgMaxRow) {
                // All done — clear overlays now that backend has caught up
                // The overlay was only needed for instant visual feedback
                spreadsheet->clearStyleOverlays();
                timer->stop();
                timer->deleteLater();
                return;
            }
            int chunkEnd = std::min(chunkStart + CHUNK_SIZE - 1, bgMaxRow);
            auto& store = spreadsheet->getColumnStore();
            for (int col = bgMinCol; col <= bgMaxCol; ++col) {
                store.scanColumnValues(col, chunkStart, chunkEnd, [&](int row, const QVariant&) {
                    uint16_t styleIdx = store.getCellStyleIndex(row, col);
                    CellStyle style = StyleTable::instance().get(styleIdx);
                    modifierCopy(style);
                    uint16_t newIdx = StyleTable::instance().intern(style);
                    if (newIdx != styleIdx) {
                        store.setCellStyle(row, col, newIdx);
                    }
                });
            }
            chunkStart = chunkEnd + 1;
        });
        timer->start();
    } else if (isLargeSelection) {
        // Only apply to occupied cells within the selection bounds
        m_spreadsheet->forEachCell([&](int row, int col, const Cell&) {
            if (row < minRow || row > maxRow || col < minCol || col > maxCol) return;

            CellAddress addr(row, col);
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        });

        // Also apply to merged cell top-left cells that overlap selection
        // (merged cells may be empty and skipped by forEachCell)
        for (const auto& region : m_spreadsheet->getMergedRegions()) {
            int mrTop = region.range.getStart().row;
            int mrLeft = region.range.getStart().col;
            int mrBottom = region.range.getEnd().row;
            int mrRight = region.range.getEnd().col;

            if (mrBottom >= minRow && mrTop <= maxRow && mrRight >= minCol && mrLeft <= maxCol) {
                CellAddress addr(mrTop, mrLeft);
                auto cell = m_spreadsheet->getCell(addr); // creates if empty
                CellStyle style = cell->getStyle();
                CellStyle oldStyle = style;
                modifier(style);
                if (!(oldStyle.bold == style.bold && oldStyle.italic == style.italic &&
                      oldStyle.backgroundColor == style.backgroundColor &&
                      oldStyle.foregroundColor == style.foregroundColor)) {
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));
                    cell->setStyle(style);
                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            }
        }
    } else {
        QModelIndexList selected = selectionModel()->selectedIndexes();
        // Track which cells we've already styled (avoid duplicates from merged regions)
        QSet<QPair<int,int>> styled;

        for (const auto& index : selected) {
            int lr = logicalRow(index);
            int lc = index.column();

            // If this cell is part of a merged region, style the top-left cell instead
            auto* mr = m_spreadsheet->getMergedRegionAt(lr, lc);
            if (mr) {
                lr = mr->range.getStart().row;
                lc = mr->range.getStart().col;
            }

            auto key = qMakePair(lr, lc);
            if (styled.contains(key)) continue;
            styled.insert(key);

            CellAddress addr(lr, lc);
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        }

        // Also handle merged cells whose top-left falls within selection bounds
        // but wasn't included in selectedIndexes (hidden by Qt span)
        for (const auto& region : m_spreadsheet->getMergedRegions()) {
            int mrTop = region.range.getStart().row;
            int mrLeft = region.range.getStart().col;
            int mrBottom = region.range.getEnd().row;
            int mrRight = region.range.getEnd().col;

            // Check if any part of the merge overlaps the selection bounds
            if (mrBottom >= minRow && mrTop <= maxRow && mrRight >= minCol && mrLeft <= maxCol) {
                auto key = qMakePair(mrTop, mrLeft);
                if (styled.contains(key)) continue;
                styled.insert(key);

                CellAddress addr(mrTop, mrLeft);
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                auto cell = m_spreadsheet->getCell(addr);
                CellStyle style = cell->getStyle();
                modifier(style);
                cell->setStyle(style);

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
    }

    if (!before.empty()) {
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<StyleChangeCommand>(before, after), m_spreadsheet.get());
    }

    // If the selection covers the full grid, also update the sheet-level default style
    // so that empty cells (not yet created) also show the formatting
    if (minRow == 0 && minCol == 0 &&
        maxRow >= m_model->rowCount() - 1 && maxCol >= m_model->columnCount() - 1) {
        CellStyle defaultStyle = m_spreadsheet->hasDefaultCellStyle()
            ? m_spreadsheet->getDefaultCellStyle() : CellStyle();
        modifier(defaultStyle);
        m_spreadsheet->setDefaultCellStyle(defaultStyle);
    }

    // Use dataChanged instead of resetModel to preserve the selection
    if (m_model) {
        QModelIndex topLeft = m_model->index(minRow, minCol);
        QModelIndex bottomRight = m_model->index(maxRow, maxCol);
        emit m_model->dataChanged(topLeft, bottomRight, {roles.begin(), roles.end()});
    }
    // Force viewport repaint to ensure visual update
    viewport()->update();
}

// Helper: get the selection bounding range string for macro recording (e.g. "A1:D10")
static QString selectionRangeStr(QItemSelectionModel* sel) {
    QModelIndexList selected = sel->selectedIndexes();
    if (selected.isEmpty()) return {};
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }
    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    return range.toString();
}

// Returns true only if the predicate is true for ALL cells in the selection.
// Used for Excel-style toggle: if any cell doesn't match → apply to all; if all match → unapply all.
bool SpreadsheetView::selectionAllMatch(std::function<bool(const CellStyle&)> predicate) const {
    if (!m_spreadsheet) return false;

    // Only check visible cells in viewport — fast and sufficient for toggle detection.
    // Checking all 9M cells for toggle would hang the app.
    QItemSelection sel = selectionModel()->selection();
    if (sel.isEmpty()) return false;

    // Check at most 50 visible cells
    int checked = 0;
    for (const auto& range : sel) {
        for (int r = range.top(); r <= range.bottom() && checked < 50; ++r) {
            for (int c = range.left(); c <= range.right() && checked < 50; ++c) {
                int logicalRow = (m_model && m_model->isVirtualMode())
                    ? m_model->toLogicalRow(r) : r;
                auto cell = m_spreadsheet->getCellIfExists(logicalRow, c);
                CellStyle style;
                if (cell) {
                    style = cell->getStyle();
                } else if (m_spreadsheet->hasDefaultCellStyle()) {
                    style = m_spreadsheet->getDefaultCellStyle();
                }
                // Also check style overlays
                for (const auto& ov : m_spreadsheet->getStyleOverlays()) {
                    if (logicalRow >= ov.minRow && logicalRow <= ov.maxRow &&
                        c >= ov.minCol && c <= ov.maxCol) {
                        ov.modifier(style);
                    }
                }
                if (!predicate(style)) return false;
                checked++;
            }
        }
    }
    return checked > 0;
}

void SpreadsheetView::applyBold() {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setBold(\"%1\", true);").arg(selectionRangeStr(selectionModel())));
    }
    // Excel behavior: if ANY cell is not bold → all become bold; only if ALL bold → all unbold
    bool allBold = selectionAllMatch([](const CellStyle& s) { return s.bold; });
    bool newVal = !allBold;
    applyStyleChange([newVal](CellStyle& s) { s.bold = newVal; }, {Qt::FontRole});
}

void SpreadsheetView::applyItalic() {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setItalic(\"%1\", true);").arg(selectionRangeStr(selectionModel())));
    }
    bool allItalic = selectionAllMatch([](const CellStyle& s) { return s.italic; });
    bool newVal = !allItalic;
    applyStyleChange([newVal](CellStyle& s) { s.italic = newVal; }, {Qt::FontRole});
}

void SpreadsheetView::applyUnderline() {
    bool allUnderline = selectionAllMatch([](const CellStyle& s) { return s.underline; });
    bool newVal = !allUnderline;
    applyStyleChange([newVal](CellStyle& s) { s.underline = newVal; }, {Qt::FontRole});
}

void SpreadsheetView::applyStrikethrough() {
    bool allStrike = selectionAllMatch([](const CellStyle& s) { return s.strikethrough; });
    bool newVal = !allStrike;
    applyStyleChange([newVal](CellStyle& s) { s.strikethrough = newVal; }, {Qt::FontRole});
}

void SpreadsheetView::applyFontFamily(const QString& family) {
    QString fam = family;  // copy for lambda capture
    applyStyleChange([fam](CellStyle& s) { s.fontName = fam; }, {Qt::FontRole});
}

void SpreadsheetView::applyFontSize(int size) {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setFontSize(\"%1\", %2);").arg(selectionRangeStr(selectionModel())).arg(size));
    }
    applyStyleChange([size](CellStyle& s) { s.fontSize = size; }, {Qt::FontRole});
}

void SpreadsheetView::applyForegroundColor(const QString& colorStr) {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setForegroundColor(\"%1\", \"%2\");").arg(selectionRangeStr(selectionModel()), colorStr));
    }
    QString color = colorStr;  // copy for lambda capture
    applyStyleChange([color](CellStyle& s) { s.foregroundColor = color; }, {Qt::ForegroundRole});
}

void SpreadsheetView::applyBackgroundColor(const QString& colorStr) {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setBackgroundColor(\"%1\", \"%2\");").arg(selectionRangeStr(selectionModel()), colorStr));
    }
    QString color = colorStr;  // copy for lambda capture
    applyStyleChange([color](CellStyle& s) { s.backgroundColor = color; }, {Qt::BackgroundRole});
}

void SpreadsheetView::applyThousandSeparator() {
    applyStyleChange([](CellStyle& s) {
        s.useThousandsSeparator = !s.useThousandsSeparator;
        if (s.numberFormat == "General") s.numberFormat = "Number";
    }, {Qt::DisplayRole});
}

void SpreadsheetView::applyNumberFormat(const QString& format) {
    applyStyleChange([&format](CellStyle& s) {
        s.numberFormat = format;
    }, {Qt::DisplayRole});
}

void SpreadsheetView::applyDateFormat(const QString& dateFormatId) {
    applyStyleChange([&dateFormatId](CellStyle& s) {
        s.numberFormat = "Date";
        s.dateFormatId = dateFormatId;
    }, {Qt::DisplayRole});
}

void SpreadsheetView::applyCurrencyFormat(const QString& currencyCode) {
    applyStyleChange([&currencyCode](CellStyle& s) {
        s.numberFormat = "Currency";
        s.currencyCode = currencyCode;
    }, {Qt::DisplayRole});
}

void SpreadsheetView::applyAccountingFormat(const QString& currencyCode) {
    applyStyleChange([&currencyCode](CellStyle& s) {
        s.numberFormat = "Accounting";
        s.currencyCode = currencyCode;
    }, {Qt::DisplayRole});
}

void SpreadsheetView::increaseDecimals() {
    applyStyleChange([](CellStyle& s) {
        if (s.decimalPlaces < 10) s.decimalPlaces++;
        if (s.numberFormat == "General") s.numberFormat = "Number";
    }, {Qt::DisplayRole});
}

void SpreadsheetView::decreaseDecimals() {
    applyStyleChange([](CellStyle& s) {
        if (s.decimalPlaces > 0) s.decimalPlaces--;
        if (s.numberFormat == "General") s.numberFormat = "Number";
    }, {Qt::DisplayRole});
}

// ============== Alignment ==============

void SpreadsheetView::applyHAlign(HorizontalAlignment align) {
    applyStyleChange([align](CellStyle& s) { s.hAlign = align; }, {Qt::TextAlignmentRole});
}

void SpreadsheetView::applyVAlign(VerticalAlignment align) {
    applyStyleChange([align](CellStyle& s) { s.vAlign = align; }, {Qt::TextAlignmentRole});
}

// ============== Format Painter ==============

void SpreadsheetView::activateFormatPainter() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    CellAddress addr(logicalRow(current), current.column());
    auto cell = m_spreadsheet->getCell(addr);
    m_copiedStyle = cell->getStyle();
    m_formatPainterActive = true;
    viewport()->setCursor(Qt::CrossCursor);
}

// ============== Sheet Protection ==============

void SpreadsheetView::protectSheet(const QString& password) {
    if (m_spreadsheet) {
        m_spreadsheet->setProtected(true, password);
    }
}

void SpreadsheetView::unprotectSheet(const QString& password) {
    if (!m_spreadsheet) return;
    if (m_spreadsheet->checkProtectionPassword(password)) {
        m_spreadsheet->setProtected(false);
    }
}

bool SpreadsheetView::isSheetProtected() const {
    return m_spreadsheet && m_spreadsheet->isProtected();
}

// ============== Sorting ==============

void SpreadsheetView::sortAscending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();

    // Determine sort range from selection
    QItemSelection sel = selectionModel()->selection();
    CellRange range;
    if (!sel.isEmpty()) {
        int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
        for (const auto& r : sel) {
            minR = qMin(minR, r.top());
            maxR = qMax(maxR, r.bottom());
            minC = qMin(minC, r.left());
            maxC = qMax(maxC, r.right());
        }
        // Check if full column selected (all model rows)
        int modelRows = m_model ? m_model->rowCount() : 100;
        bool isFullColumn = (maxR - minR + 1 >= modelRows - 1);

        if (isFullColumn) {
            // Full column: sort ALL logical rows
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow < 1) return;
            range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
        } else if (minR < maxR) {
            // Partial selection: translate model rows to logical rows
            int logMinR = m_model ? m_model->toLogicalRow(minR) : minR;
            int logMaxR = m_model ? m_model->toLogicalRow(maxR) : maxR;
            range = CellRange(CellAddress(logMinR, minC), CellAddress(logMaxR, maxC));
        }
    }

    if (!range.isValid()) {
        int maxRow = m_spreadsheet->getMaxRow();
        int maxCol = m_spreadsheet->getMaxColumn();
        if (maxRow < 1 && maxCol < 1) return;
        if (maxRow < 1) maxRow = 1;
        range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    }

    // BUG 3 FIX: Check for merged cells in sort range
    const auto& merged = m_spreadsheet->getMergedRegions();
    for (const auto& region : merged) {
        if (range.intersects(region.range)) {
            QMessageBox::warning(this, "Sort Warning",
                "This operation requires identically sized merged cells.\n"
                "Please unmerge cells before sorting.");
            return;
        }
    }

    // BUG 2 FIX: Snapshot cells before sort for undo support
    std::vector<CellSnapshot> beforeSnapshots;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            beforeSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->sortRange(range, col, true);

    // Snapshot cells after sort
    std::vector<CellSnapshot> afterSnapshots;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            afterSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(beforeSnapshots, afterSnapshots, "Sort Ascending"));

    // Full model reset to ensure view refreshes completely
    if (m_model) {
        m_model->resetModel();
    }

    // Restore selection
    int selMinR = 0, selMaxR = std::min(range.getEnd().row, m_model ? m_model->rowCount() - 1 : 0);
    int selMinC = range.getStart().col, selMaxC = range.getEnd().col;
    QModelIndex topLeft = m_model->index(selMinR, selMinC);
    QModelIndex bottomRight = m_model->index(selMaxR, selMaxC);
    selectionModel()->setCurrentIndex(topLeft, QItemSelectionModel::NoUpdate);
    selectionModel()->select(QItemSelection(topLeft, bottomRight), QItemSelectionModel::ClearAndSelect);
}

void SpreadsheetView::sortDescending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();

    // Determine sort range from selection
    QItemSelection sel = selectionModel()->selection();
    CellRange range;
    if (!sel.isEmpty()) {
        int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
        for (const auto& r : sel) {
            minR = qMin(minR, r.top());
            maxR = qMax(maxR, r.bottom());
            minC = qMin(minC, r.left());
            maxC = qMax(maxC, r.right());
        }
        int modelRows = m_model ? m_model->rowCount() : 100;
        bool isFullColumn = (maxR - minR + 1 >= modelRows - 1);

        if (isFullColumn) {
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow < 1) return;
            range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
        } else if (minR < maxR) {
            int logMinR = m_model ? m_model->toLogicalRow(minR) : minR;
            int logMaxR = m_model ? m_model->toLogicalRow(maxR) : maxR;
            range = CellRange(CellAddress(logMinR, minC), CellAddress(logMaxR, maxC));
        }
    }

    if (!range.isValid()) {
        int maxRow = m_spreadsheet->getMaxRow();
        int maxCol = m_spreadsheet->getMaxColumn();
        if (maxRow < 1 && maxCol < 1) return;
        if (maxRow < 1) maxRow = 1;
        range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    }

    // BUG 3 FIX: Check for merged cells in sort range
    const auto& merged = m_spreadsheet->getMergedRegions();
    for (const auto& region : merged) {
        if (range.intersects(region.range)) {
            QMessageBox::warning(this, "Sort Warning",
                "This operation requires identically sized merged cells.\n"
                "Please unmerge cells before sorting.");
            return;
        }
    }

    // BUG 2 FIX: Snapshot cells before sort for undo support
    std::vector<CellSnapshot> beforeSnapshots;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            beforeSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->sortRange(range, col, false);

    // Snapshot cells after sort
    std::vector<CellSnapshot> afterSnapshots;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            afterSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(beforeSnapshots, afterSnapshots, "Sort Descending"));

    if (m_model) {
        m_model->resetModel();
    }

    int selMinR = 0, selMaxR = std::min(range.getEnd().row, m_model ? m_model->rowCount() - 1 : 0);
    int selMinC = range.getStart().col, selMaxC = range.getEnd().col;
    QModelIndex topLeft = m_model->index(selMinR, selMinC);
    QModelIndex bottomRight = m_model->index(selMaxR, selMaxC);
    selectionModel()->setCurrentIndex(topLeft, QItemSelectionModel::NoUpdate);
    selectionModel()->select(QItemSelection(topLeft, bottomRight), QItemSelectionModel::ClearAndSelect);
}

void SpreadsheetView::showSortDialog() {
    if (!m_spreadsheet) return;

    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    // Determine range from selection (same logic as sortAscending)
    QItemSelection sel = selectionModel()->selection();
    CellRange range;
    int minC = 0, maxC = 0;

    if (!sel.isEmpty()) {
        int minR = INT_MAX, maxR = 0;
        minC = INT_MAX; maxC = 0;
        for (const auto& r : sel) {
            minR = qMin(minR, r.top());
            maxR = qMax(maxR, r.bottom());
            minC = qMin(minC, r.left());
            maxC = qMax(maxC, r.right());
        }
        int modelRows = m_model ? m_model->rowCount() : 100;
        bool isFullColumn = (maxR - minR + 1 >= modelRows - 1);

        if (isFullColumn) {
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow < 1) return;
            minC = 0;
            maxC = maxCol;
            range = CellRange(CellAddress(0, 0), CellAddress(maxRow, maxCol));
        } else if (minR < maxR) {
            int logMinR = m_model ? m_model->toLogicalRow(minR) : minR;
            int logMaxR = m_model ? m_model->toLogicalRow(maxR) : maxR;
            range = CellRange(CellAddress(logMinR, minC), CellAddress(logMaxR, maxC));
        }
    }

    if (!range.isValid()) {
        int maxRow = m_spreadsheet->getMaxRow();
        int maxCol = m_spreadsheet->getMaxColumn();
        if (maxRow < 1 && maxCol < 1) return;
        if (maxRow < 1) maxRow = 1;
        minC = 0;
        maxC = maxCol;
        range = CellRange(CellAddress(0, 0), CellAddress(maxRow, maxCol));
    }

    // Show the sort dialog
    SortDialog dialog(minC, maxC, false, this);
    if (dialog.exec() != QDialog::Accepted) return;

    auto levels = dialog.getSortLevels();
    if (levels.empty()) return;

    // If headers, skip the first row
    CellRange sortRange = range;
    if (dialog.hasHeaders()) {
        int newStartRow = range.getStart().row + 1;
        if (newStartRow > range.getEnd().row) return;
        sortRange = CellRange(CellAddress(newStartRow, range.getStart().col),
                              range.getEnd());
    }

    // BUG 3 FIX: Check for merged cells in sort range
    const auto& merged = m_spreadsheet->getMergedRegions();
    for (const auto& region : merged) {
        if (sortRange.intersects(region.range)) {
            QMessageBox::warning(this, "Sort Warning",
                "This operation requires identically sized merged cells.\n"
                "Please unmerge cells before sorting.");
            return;
        }
    }

    // Convert SortLevel to pair<int, bool> for the multi-sort API
    std::vector<std::pair<int, bool>> sortKeys;
    sortKeys.reserve(levels.size());
    for (const auto& lvl : levels) {
        sortKeys.emplace_back(lvl.column, lvl.ascending);
    }

    // BUG 2 FIX: Snapshot cells before sort for undo support
    std::vector<CellSnapshot> beforeSnapshots;
    for (int r = sortRange.getStart().row; r <= sortRange.getEnd().row; ++r) {
        for (int c = sortRange.getStart().col; c <= sortRange.getEnd().col; ++c) {
            beforeSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->sortRangeMulti(sortRange, sortKeys);

    // Snapshot cells after sort
    std::vector<CellSnapshot> afterSnapshots;
    for (int r = sortRange.getStart().row; r <= sortRange.getEnd().row; ++r) {
        for (int c = sortRange.getStart().col; c <= sortRange.getEnd().col; ++c) {
            afterSnapshots.push_back(m_spreadsheet->takeCellSnapshot({r, c}));
        }
    }

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(beforeSnapshots, afterSnapshots, "Sort"));

    if (m_model) {
        m_model->resetModel();
    }

    // Restore selection
    int selMinR = 0, selMaxR = std::min(range.getEnd().row, m_model ? m_model->rowCount() - 1 : 0);
    int selMinC = range.getStart().col, selMaxC = range.getEnd().col;
    QModelIndex topLeft = m_model->index(selMinR, selMinC);
    QModelIndex bottomRight = m_model->index(selMaxR, selMaxC);
    selectionModel()->setCurrentIndex(topLeft, QItemSelectionModel::NoUpdate);
    selectionModel()->select(QItemSelection(topLeft, bottomRight), QItemSelectionModel::ClearAndSelect);
}

// ============== Table Style ==============

CellRange SpreadsheetView::detectDataRegion(int startRow, int startCol) const {
    if (!m_spreadsheet) return CellRange(CellAddress(startRow, startCol), CellAddress(startRow, startCol));

    // Expand outward from the starting cell to find the contiguous data region.
    // Uses the cached navigation index for O(columns × log n) instead of O(rows × columns).

    int maxCol = m_spreadsheet->getMaxColumn();

    // Find the left boundary — check if adjacent column has any occupied rows
    int left = startCol;
    while (left > 0) {
        const auto& occupied = m_spreadsheet->getOccupiedRowsInColumn(left - 1);
        if (occupied.empty()) break;
        left--;
    }

    // Find the right boundary
    int right = startCol;
    while (right < maxCol) {
        const auto& occupied = m_spreadsheet->getOccupiedRowsInColumn(right + 1);
        if (occupied.empty()) break;
        right++;
    }

    // Find top and bottom boundaries using per-column binary search.
    // For each column in [left, right], find the contiguous extent from startRow.
    // A row is in the region if ANY column has data, so we take the widest extent.
    int top = startRow;
    int bottom = startRow;

    for (int c = left; c <= right; ++c) {
        const auto& rows = m_spreadsheet->getOccupiedRowsInColumn(c);
        if (rows.empty()) continue;

        // Binary search for startRow
        auto it = std::lower_bound(rows.begin(), rows.end(), startRow);

        // Upward contiguous extent from startRow
        if (it != rows.end() && *it == startRow) {
            // Walk back to find first contiguous row (binary search)
            size_t idx = static_cast<size_t>(it - rows.begin());
            size_t lo = 0, hi = idx;
            while (lo < hi) {
                size_t mid = lo + (hi - lo) / 2;
                // Check if rows[mid..idx] is contiguous: rows[idx] - rows[mid] == idx - mid
                if (rows[idx] - rows[mid] == static_cast<int>(idx - mid)) {
                    hi = mid;
                } else {
                    lo = mid + 1;
                }
            }
            top = std::min(top, rows[lo]);

            // Downward contiguous extent from startRow (binary search)
            size_t maxIdx = rows.size() - 1;
            lo = idx;
            hi = maxIdx;
            while (lo < hi) {
                size_t mid = lo + (hi - lo + 1) / 2;
                if (rows[mid] - rows[idx] == static_cast<int>(mid - idx)) {
                    lo = mid;
                } else {
                    hi = mid - 1;
                }
            }
            bottom = std::max(bottom, rows[lo]);
        } else if (it != rows.end()) {
            // startRow not in this column, but there's data after it — extend bottom
            // Check if adjacent to startRow+1
            if (*it == startRow + 1) {
                size_t idx = static_cast<size_t>(it - rows.begin());
                size_t maxIdx = rows.size() - 1;
                size_t lo = idx, hi = maxIdx;
                while (lo < hi) {
                    size_t mid = lo + (hi - lo + 1) / 2;
                    if (rows[mid] - rows[idx] == static_cast<int>(mid - idx)) {
                        lo = mid;
                    } else {
                        hi = mid - 1;
                    }
                }
                bottom = std::max(bottom, rows[lo]);
            }
        }

        // Also check upward from startRow for data before it
        if (it != rows.begin()) {
            auto prevIt = std::prev(it);
            if (*prevIt == startRow - 1 || (it != rows.end() && *it == startRow)) {
                // Already handled above via contiguous check
            } else if (*prevIt == startRow - 1) {
                size_t idx = static_cast<size_t>(prevIt - rows.begin());
                size_t lo = 0, hi = idx;
                while (lo < hi) {
                    size_t mid = lo + (hi - lo) / 2;
                    if (rows[idx] - rows[mid] == static_cast<int>(idx - mid)) {
                        hi = mid;
                    } else {
                        lo = mid + 1;
                    }
                }
                top = std::min(top, rows[lo]);
            }
        }
    }

    return CellRange(CellAddress(top, left), CellAddress(bottom, right));
}

void SpreadsheetView::applyTableStyle(int themeIndex) {
    if (!m_spreadsheet) return;

    auto themes = generateTableThemes(m_spreadsheet->getDocumentTheme());
    if (themeIndex < 0 || themeIndex >= static_cast<int>(themes.size())) return;

    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow, maxRow, minCol, maxCol;

    // Auto-detect: if only a single cell is selected, detect the contiguous data region
    if (selected.size() == 1) {
        CellRange region = detectDataRegion(selected.first().row(), selected.first().column());
        minRow = region.getStart().row;
        maxRow = region.getEnd().row;
        minCol = region.getStart().col;
        maxCol = region.getEnd().col;
    } else {
        // Find selection bounding box
        minRow = INT_MAX; maxRow = 0; minCol = INT_MAX; maxCol = 0;
        for (const auto& idx : selected) {
            minRow = qMin(minRow, idx.row());
            maxRow = qMax(maxRow, idx.row());
            minCol = qMin(minCol, idx.column());
            maxCol = qMax(maxCol, idx.column());
        }
    }

    SpreadsheetTable table;
    table.range = CellRange(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    table.theme = themes[themeIndex];
    table.hasHeaderRow = true;
    table.bandedRows = true;

    // Auto-name
    int tableNum = static_cast<int>(m_spreadsheet->getTables().size()) + 1;
    table.name = QString("Table%1").arg(tableNum);

    // Extract column names from header row
    for (int c = minCol; c <= maxCol; ++c) {
        auto val = m_spreadsheet->getCellValue(CellAddress(minRow, c));
        QString name = val.toString();
        if (name.isEmpty()) name = QString("Column%1").arg(c - minCol + 1);
        table.columnNames.append(name);
    }

    // Snapshot tables before change
    std::vector<SpreadsheetTable> tablesBefore = m_spreadsheet->getTables();
    CellRange affectedRange = table.range;

    // Also capture the old table's range for dataChanged (if replacing)
    CellRange oldRange = affectedRange;
    const auto& existing = m_spreadsheet->getTables();
    for (const auto& t : existing) {
        if (t.range.intersects(table.range)) {
            oldRange = t.range;
            m_spreadsheet->removeTable(t.name);
            break;
        }
    }

    m_spreadsheet->addTable(table);

    // Snapshot tables after change
    std::vector<SpreadsheetTable> tablesAfter = m_spreadsheet->getTables();

    // Push undo command (don't execute — we already made the change)
    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<TableChangeCommand>(tablesBefore, tablesAfter, affectedRange, "Apply Table Style"));

    // Use targeted dataChanged for both old and new table ranges
    if (m_model) {
        // Refresh old range (in case table was replaced and old range was different)
        int refreshMinRow = qMin(oldRange.getStart().row, minRow);
        int refreshMaxRow = qMax(oldRange.getEnd().row, maxRow);
        int refreshMinCol = qMin(oldRange.getStart().col, minCol);
        int refreshMaxCol = qMax(oldRange.getEnd().col, maxCol);
        QModelIndex topLeft = m_model->index(refreshMinRow, refreshMinCol);
        QModelIndex bottomRight = m_model->index(refreshMaxRow, refreshMaxCol);
        emit m_model->dataChanged(topLeft, bottomRight);
    }
}

// ============== Auto Filter ==============

void SpreadsheetView::toggleAutoFilter() {
    if (m_filterActive) {
        clearAllFilters();
        return;
    }

    if (!m_spreadsheet) return;

    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    // Detect data region from current cell
    m_filterRange = detectDataRegion(logicalRow(current), current.column());
    m_filterHeaderRow = m_filterRange.getStart().row;
    m_filterActive = true;
    m_columnFilters.clear();

    // Connect horizontal header clicks to show filter dropdown
    // (Disconnect any previous connection first)
    disconnect(horizontalHeader(), &QHeaderView::sectionClicked, this, nullptr);
    connect(horizontalHeader(), &QHeaderView::sectionClicked, this, [this](int logicalIndex) {
        if (!m_filterActive) return;
        int startCol = m_filterRange.getStart().col;
        int endCol = m_filterRange.getEnd().col;
        if (logicalIndex >= startCol && logicalIndex <= endCol) {
            showFilterDropdown(logicalIndex);
        }
    });

    viewport()->update();
}

void SpreadsheetView::showPivotFilterDropdown(int filterIndex) {
    if (!m_spreadsheet || !m_spreadsheet->isPivotSheet()) return;
    const PivotConfig* pc = m_spreadsheet->getPivotConfig();
    if (!pc || filterIndex < 0 || filterIndex >= static_cast<int>(pc->filterFields.size()))
        return;

    const PivotFilterField& ff = pc->filterFields[filterIndex];

    // Get unique values from the SOURCE sheet (not the pivot output)
    // We need access to the source sheet — it's stored by index in the config
    // The caller (MainWindow) will have connected us; for now we just need
    // the source sheet. We'll use a PivotEngine to get unique values.
    // Since we don't have the source sheets vector here, we emit the signal
    // with the filter index and let MainWindow handle the heavy lifting.
    // But we CAN show the dialog here using the stored config info.

    // We need the source sheet — get it via a signal roundtrip would be complex.
    // Instead, store a helper pointer. But the simplest approach: just emit
    // a signal and let MainWindow show the dialog and recompute.
    // However, for better UX let's emit the signal directly.
    emit pivotFilterChanged(filterIndex, {});  // empty list = signal to show picker
}

void SpreadsheetView::enableAutoFilter(const CellRange& range) {
    if (m_filterActive) clearAllFilters();
    m_filterRange = range;
    m_filterHeaderRow = range.getStart().row;
    m_filterActive = true;
    m_columnFilters.clear();

    disconnect(horizontalHeader(), &QHeaderView::sectionClicked, this, nullptr);
    connect(horizontalHeader(), &QHeaderView::sectionClicked, this, [this](int logicalIndex) {
        if (!m_filterActive) return;
        int startCol = m_filterRange.getStart().col;
        int endCol = m_filterRange.getEnd().col;
        if (logicalIndex >= startCol && logicalIndex <= endCol) {
            showFilterDropdown(logicalIndex);
        }
    });

    viewport()->update();
}

void SpreadsheetView::clearAllFilters() {
    m_filterActive = false;
    m_columnFilters.clear();

    // Unhide all rows
    int startRow = m_filterRange.getStart().row;
    int endRow = m_filterRange.getEnd().row;
    for (int r = startRow; r <= endRow; ++r) {
        setRowHidden(r, false);
    }

    disconnect(horizontalHeader(), &QHeaderView::sectionClicked, this, nullptr);
    viewport()->update();
}

void SpreadsheetView::showFilterDropdown(int column) {
    if (!m_spreadsheet) return;

    int dataStartRow = m_filterHeaderRow + 1;
    int dataEndRow = m_filterRange.getEnd().row;

    // Collect unique values — use nav index to skip empty rows (O(occupied) not O(grid))
    QStringList uniqueValues;
    QSet<QString> seen;
    const auto& occupiedRows = m_spreadsheet->getOccupiedRowsInColumn(column);
    auto lo = std::lower_bound(occupiedRows.begin(), occupiedRows.end(), dataStartRow);
    auto hi = std::upper_bound(occupiedRows.begin(), occupiedRows.end(), dataEndRow);
    for (auto it = lo; it != hi; ++it) {
        auto val = m_spreadsheet->getCellValue(CellAddress(*it, column));
        QString text = val.toString();
        if (!seen.contains(text)) {
            seen.insert(text);
            uniqueValues.append(text);
        }
    }
    uniqueValues.sort(Qt::CaseInsensitive);

    // Get current filter for this column (if any)
    QSet<QString> currentFilter;
    bool hasFilter = m_columnFilters.count(column) > 0;
    if (hasFilter) {
        currentFilter = m_columnFilters[column];
    }

    // Create filter dropdown dialog
    QDialog dialog(this);
    dialog.setWindowTitle("Auto Filter");
    dialog.setMinimumSize(240, 340);
    dialog.setMaximumSize(320, 500);

    QVBoxLayout* layout = new QVBoxLayout(&dialog);
    layout->setContentsMargins(6, 6, 6, 6);
    layout->setSpacing(4);

    // Search box for filtering items
    QLineEdit* searchBox = new QLineEdit(&dialog);
    searchBox->setPlaceholderText("Search...");
    searchBox->setFixedHeight(24);
    searchBox->setClearButtonEnabled(true);
    layout->addWidget(searchBox);

    // Select All / Clear All buttons
    QHBoxLayout* btnRow = new QHBoxLayout();
    QPushButton* selectAllBtn = new QPushButton("Select All", &dialog);
    QPushButton* clearAllBtn = new QPushButton("Clear All", &dialog);
    selectAllBtn->setFixedHeight(24);
    clearAllBtn->setFixedHeight(24);
    btnRow->addWidget(selectAllBtn);
    btnRow->addWidget(clearAllBtn);
    layout->addLayout(btnRow);

    // QListWidget with checkable items — virtualized, handles thousands of items instantly
    QListWidget* listWidget = new QListWidget(&dialog);
    listWidget->setStyleSheet(
        "QListWidget::item { padding: 1px 2px; }"
        "QListWidget { font-size: 11px; }");
    listWidget->setUniformItemSizes(true); // enables fast layout for large lists

    // Add "(Blanks)" entry
    auto* blanksItem = new QListWidgetItem("(Blanks)", listWidget);
    blanksItem->setFlags(blanksItem->flags() | Qt::ItemIsUserCheckable);
    blanksItem->setCheckState((!hasFilter || currentFilter.contains("")) ? Qt::Checked : Qt::Unchecked);
    blanksItem->setData(Qt::UserRole, QString("")); // actual value

    for (const QString& val : uniqueValues) {
        if (val.isEmpty()) continue;
        auto* item = new QListWidgetItem(val, listWidget);
        item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
        item->setCheckState((!hasFilter || currentFilter.contains(val)) ? Qt::Checked : Qt::Unchecked);
        item->setData(Qt::UserRole, val);
    }
    layout->addWidget(listWidget);

    // Search filtering — hide non-matching items
    QObject::connect(searchBox, &QLineEdit::textChanged, [listWidget](const QString& text) {
        for (int i = 0; i < listWidget->count(); i++) {
            auto* item = listWidget->item(i);
            bool match = text.isEmpty() || item->text().contains(text, Qt::CaseInsensitive);
            item->setHidden(!match);
        }
    });

    // Select/clear all (only visible items)
    QObject::connect(selectAllBtn, &QPushButton::clicked, [listWidget]() {
        for (int i = 0; i < listWidget->count(); i++) {
            auto* item = listWidget->item(i);
            if (!item->isHidden()) item->setCheckState(Qt::Checked);
        }
    });
    QObject::connect(clearAllBtn, &QPushButton::clicked, [listWidget]() {
        for (int i = 0; i < listWidget->count(); i++) {
            auto* item = listWidget->item(i);
            if (!item->isHidden()) item->setCheckState(Qt::Unchecked);
        }
    });

    // OK / Cancel
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dialog);
    layout->addWidget(buttons);
    QObject::connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    QObject::connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);

    // Position dialog near the column header
    int headerX = horizontalHeader()->sectionViewportPosition(column);
    QPoint globalPos = horizontalHeader()->mapToGlobal(
        QPoint(headerX, horizontalHeader()->height()));
    dialog.move(globalPos);

    if (dialog.exec() == QDialog::Accepted) {
        QSet<QString> selectedValues;
        for (int i = 0; i < listWidget->count(); i++) {
            auto* item = listWidget->item(i);
            if (item->checkState() == Qt::Checked) {
                selectedValues.insert(item->data(Qt::UserRole).toString());
            }
        }

        // Count total possible values (unique non-empty + blank)
        int totalPossible = seen.size();
        if (!seen.contains("")) totalPossible++;

        if (static_cast<int>(selectedValues.size()) >= totalPossible) {
            m_columnFilters.erase(column);
        } else {
            m_columnFilters[column] = selectedValues;
        }

        applyFilters();
    }
}

void SpreadsheetView::applyFilters() {
    if (!m_spreadsheet || !m_filterActive) return;

    int dataStartRow = m_filterHeaderRow + 1;
    int dataEndRow = m_filterRange.getEnd().row;
    int totalRows = dataEndRow - dataStartRow + 1;

    // For large datasets (>50K rows), use bitmap-based filtering
    // Avoids O(n) setRowHidden() calls that freeze the UI
    if (totalRows > 50000) {
        // Build a hidden-rows set from column filters using ColumnStore scanning
        QSet<int> visibleRows;
        bool firstFilter = true;

        for (const auto& [col, allowedValues] : m_columnFilters) {
            QSet<int> colVisible;
            m_spreadsheet->getColumnStore().scanColumnValues(col, dataStartRow, dataEndRow,
                [&](int row, const QVariant& val) {
                    if (allowedValues.contains(val.toString())) {
                        colVisible.insert(row);
                    }
                });
            // Also include rows with empty values if "(Blanks)" is allowed
            if (allowedValues.contains(QString())) {
                for (int r = dataStartRow; r <= dataEndRow; r++) {
                    if (!m_spreadsheet->getColumnStore().hasCell(r, col)) {
                        colVisible.insert(r);
                    }
                }
            }
            if (firstFilter) {
                visibleRows = colVisible;
                firstFilter = false;
            } else {
                visibleRows &= colVisible;  // AND logic
            }
        }

        // Apply visibility in batch — only hide/show rows that changed
        for (int r = dataStartRow; r <= dataEndRow; ++r) {
            bool shouldBeVisible = m_columnFilters.empty() || visibleRows.contains(r);
            if (isRowHidden(r) != !shouldBeVisible) {
                setRowHidden(r, !shouldBeVisible);
            }
        }
    } else {
        // Small dataset: original per-row approach (fast enough)
        for (int r = dataStartRow; r <= dataEndRow; ++r) {
            bool visible = true;
            for (const auto& [col, allowedValues] : m_columnFilters) {
                auto val = m_spreadsheet->getCellValue(CellAddress(r, col));
                QString text = val.toString();
                if (!allowedValues.contains(text)) {
                    visible = false;
                    break;
                }
            }
            setRowHidden(r, !visible);
        }
    }

    viewport()->update();
}

// ============== Clear Operations ==============

void SpreadsheetView::clearAll() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

    // Compute bounding box of selection for validation/conditional formatting cleanup
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;

    std::vector<CellSnapshot> before, after;
    for (const auto& index : selected) {
        CellAddress addr(logicalRow(index), index.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        cell->clear();
        cell->setStyle(CellStyle()); // Reset to default style
        cell->setComment("");
        cell->setHyperlink("");
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));

        minRow = qMin(minRow, addr.row);
        maxRow = qMax(maxRow, addr.row);
        minCol = qMin(minCol, addr.col);
        maxCol = qMax(maxCol, addr.col);
    }

    // Remove validation rules that intersect the cleared selection
    CellRange selRange(minRow, minCol, maxRow, maxCol);
    auto& valRules = m_spreadsheet->getValidationRules();
    valRules.erase(
        std::remove_if(valRules.begin(), valRules.end(),
            [&selRange](const Spreadsheet::DataValidationRule& rule) {
                return rule.range.intersects(selRange);
            }),
        valRules.end());

    // Remove conditional formatting rules that intersect the cleared selection
    auto& cfRules = m_spreadsheet->getConditionalFormatting().getAllRules();
    // Iterate in reverse to safely remove by index
    for (int i = static_cast<int>(cfRules.size()) - 1; i >= 0; --i) {
        if (cfRules[i] && cfRules[i]->getRange().intersects(selRange)) {
            m_spreadsheet->getConditionalFormatting().removeRule(static_cast<size_t>(i));
        }
    }

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Clear All"));

    if (m_model) m_model->resetModel();
}

void SpreadsheetView::clearContent() {
    deleteSelection(); // Already implemented - clears values but keeps formatting
}

void SpreadsheetView::clearFormats() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    std::vector<CellSnapshot> before, after;
    for (const auto& index : selected) {
        CellAddress addr(logicalRow(index), index.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        cell->setStyle(CellStyle()); // Reset style only, keep value
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<StyleChangeCommand>(before, after));

    if (m_model) m_model->resetModel();
}

// ============== Indent ==============

void SpreadsheetView::increaseIndent() {
    applyStyleChange([](CellStyle& s) {
        s.indentLevel = qMin(s.indentLevel + 1, 10);
        if (s.hAlign == HorizontalAlignment::General || s.hAlign == HorizontalAlignment::Right || s.hAlign == HorizontalAlignment::Center) {
            s.hAlign = HorizontalAlignment::Left; // Indent forces left-align like Excel
        }
    }, {Qt::TextAlignmentRole, Qt::UserRole + 10});
}

void SpreadsheetView::decreaseIndent() {
    applyStyleChange([](CellStyle& s) {
        s.indentLevel = qMax(s.indentLevel - 1, 0);
    }, {Qt::UserRole + 10});
}

void SpreadsheetView::applyTextRotation(int degrees) {
    applyStyleChange([degrees](CellStyle& s) {
        s.textRotation = degrees;
    }, {Qt::UserRole + 16});
}

void SpreadsheetView::applyTextOverflow(TextOverflowMode mode) {
    applyStyleChange([mode](CellStyle& s) {
        s.textOverflow = mode;
    }, {Qt::DisplayRole});
}

// ============== Borders ==============

void SpreadsheetView::applyBorderStyle(const QString& borderType, const QColor& color, int width, int penStyle) {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    BorderStyle on;
    on.enabled = true;
    on.color = color.name();
    on.width = width;
    on.penStyle = penStyle;

    BorderStyle off;
    off.enabled = false;

    auto modifier = [&](CellStyle& s, int row, int col) {
        if (borderType == "none") {
            s.borderTop = off;
            s.borderBottom = off;
            s.borderLeft = off;
            s.borderRight = off;
        } else if (borderType == "all") {
            // Only set top/left on outer edges to avoid double-drawing internal borders
            s.borderTop = (row == minRow) ? on : off;
            s.borderLeft = (col == minCol) ? on : off;
            s.borderBottom = on;
            s.borderRight = on;
        } else if (borderType == "outside") {
            if (row == minRow) s.borderTop = on;
            if (row == maxRow) s.borderBottom = on;
            if (col == minCol) s.borderLeft = on;
            if (col == maxCol) s.borderRight = on;
        } else if (borderType == "bottom") {
            if (row == maxRow) s.borderBottom = on;
        } else if (borderType == "top") {
            if (row == minRow) s.borderTop = on;
        } else if (borderType == "thick_outside") {
            BorderStyle thick = on;
            thick.width = 2;
            if (row == minRow) s.borderTop = thick;
            if (row == maxRow) s.borderBottom = thick;
            if (col == minCol) s.borderLeft = thick;
            if (col == maxCol) s.borderRight = thick;
        } else if (borderType == "left") {
            if (col == minCol) s.borderLeft = on;
        } else if (borderType == "right") {
            if (col == maxCol) s.borderRight = on;
        } else if (borderType == "inside_h") {
            // Only set borderBottom to avoid double-drawing at shared edges
            if (row < maxRow) s.borderBottom = on;
        } else if (borderType == "inside_v") {
            // Only set borderRight to avoid double-drawing at shared edges
            if (col < maxCol) s.borderRight = on;
        } else if (borderType == "inside") {
            if (row < maxRow) s.borderBottom = on;
            if (col < maxCol) s.borderRight = on;
        }
    };

    std::vector<CellSnapshot> before, after;
    for (const auto& idx : selected) {
        CellAddress addr(logicalRow(idx), idx.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        CellStyle style = cell->getStyle();
        modifier(style, idx.row(), idx.column());
        cell->setStyle(style);
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }

    if (!before.empty()) {
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<StyleChangeCommand>(before, after));
    }

    // Use targeted dataChanged instead of resetModel to preserve selection
    if (m_model) {
        QModelIndex topLeft = m_model->index(minRow, minCol);
        QModelIndex bottomRight = m_model->index(maxRow, maxCol);
        emit m_model->dataChanged(topLeft, bottomRight);
    }
}

// ============== Merge Cells ==============

void SpreadsheetView::mergeCells() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.size() <= 1) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->mergeCells(range);

    // Set span on the QTableView
    int rowSpan = maxRow - minRow + 1;
    int colSpan = maxCol - minCol + 1;
    setSpan(minRow, minCol, rowSpan, colSpan);

    // Center the content in the merged cell
    auto cell = m_spreadsheet->getCell(CellAddress(minRow, minCol));
    CellStyle style = cell->getStyle();
    style.hAlign = HorizontalAlignment::Center;
    style.vAlign = VerticalAlignment::Middle;
    cell->setStyle(style);

    if (m_model) m_model->resetModel();

    // Keep focus on the merged cell
    QModelIndex mergedIdx = m_model->index(minRow, minCol);
    setCurrentIndex(mergedIdx);
    selectionModel()->select(mergedIdx, QItemSelectionModel::ClearAndSelect);
}

void SpreadsheetView::unmergeCells() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    auto* mr = m_spreadsheet->getMergedRegionAt(logicalRow(current), current.column());
    if (!mr) return;

    int startRow = mr->range.getStart().row;
    int startCol = mr->range.getStart().col;
    int endRow = mr->range.getEnd().row;
    int endCol = mr->range.getEnd().col;

    // Clear span
    setSpan(startRow, startCol, 1, 1);

    m_spreadsheet->unmergeCells(mr->range);

    if (m_model) m_model->resetModel();
}

// ============== Context Menu ==============

void SpreadsheetView::showCellContextMenu(const QPoint& pos) {
    QMenu menu(this);

    menu.addAction("Cu&t", QKeySequence(Qt::CTRL | Qt::Key_X), this, &SpreadsheetView::cut);
    menu.addAction("&Copy", QKeySequence(Qt::CTRL | Qt::Key_C), this, &SpreadsheetView::copy);
    menu.addAction("&Paste", QKeySequence(Qt::CTRL | Qt::Key_V), this, &SpreadsheetView::paste);
    menu.addAction("Paste &Special...", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_V),
                   this, &SpreadsheetView::pasteSpecial);

    menu.addSeparator();

    // Clear submenu
    QMenu* clearMenu = menu.addMenu("Clea&r");
    clearMenu->addAction("Clear &All", this, &SpreadsheetView::clearAll);
    clearMenu->addAction("Clear &Contents", QKeySequence(Qt::Key_Delete),
                         this, &SpreadsheetView::clearContent);
    clearMenu->addAction("Clear &Formats", this, &SpreadsheetView::clearFormats);

    // Pick from Drop-down List (Excel feature: shows unique values from current column)
    menu.addSeparator();
    menu.addAction("Pick from &Drop-down List...", this, [this]() {
        QModelIndex idx = currentIndex();
        if (!idx.isValid() || !m_spreadsheet) return;

        int col = idx.column();
        int maxRow = m_spreadsheet->getMaxRow();

        // Collect unique non-empty values from this column
        QStringList values;
        QSet<QString> seen;
        for (int r = 0; r <= maxRow; ++r) {
            if (r == idx.row()) continue; // skip current cell
            auto cell = m_spreadsheet->getCellIfExists(r, col);
            if (cell && cell->getValue().isValid()) {
                QString val = cell->getValue().toString().trimmed();
                if (!val.isEmpty() && !seen.contains(val)) {
                    seen.insert(val);
                    values.append(val);
                }
            }
        }

        if (values.isEmpty()) return;
        values.sort(Qt::CaseInsensitive);

        // Show popup menu with values
        QMenu popup(this);
        popup.setStyleSheet(ThemeManager::instance().dialogStylesheet());
        for (const QString& val : values) {
            popup.addAction(val, [this, idx, val]() {
                model()->setData(idx, val, Qt::EditRole);
            });
        }
        popup.exec(viewport()->mapToGlobal(visualRect(idx).bottomLeft()));
    });

    menu.addSeparator();

    // Insert submenu
    QMenu* insertMenu = menu.addMenu("&Insert...");
    insertMenu->addAction("Shift cells &right", this, &SpreadsheetView::insertCellsShiftRight);
    insertMenu->addAction("Shift cells &down", this, &SpreadsheetView::insertCellsShiftDown);
    insertMenu->addSeparator();
    insertMenu->addAction("Entire ro&w", this, &SpreadsheetView::insertEntireRow);
    insertMenu->addAction("Entire co&lumn", this, &SpreadsheetView::insertEntireColumn);

    // Delete submenu
    QMenu* deleteMenu = menu.addMenu("De&lete...");
    deleteMenu->addAction("Shift cells &left", this, &SpreadsheetView::deleteCellsShiftLeft);
    deleteMenu->addAction("Shift cells &up", this, &SpreadsheetView::deleteCellsShiftUp);
    deleteMenu->addSeparator();
    deleteMenu->addAction("Entire ro&w", this, &SpreadsheetView::deleteEntireRow);
    deleteMenu->addAction("Entire co&lumn", this, &SpreadsheetView::deleteEntireColumn);

    menu.addSeparator();

    // Merge cells
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.size() > 1) {
        menu.addAction("Merge && Center", this, &SpreadsheetView::mergeCells);
    }
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid() && m_spreadsheet->getMergedRegionAt(logicalRow(cur), cur.column())) {
            menu.addAction("Unmerge Cells", this, &SpreadsheetView::unmergeCells);
        }
    }

    menu.addSeparator();

    // Checkbox / Picklist context actions
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            auto cell = m_spreadsheet->getCell(CellAddress(logicalRow(cur), cur.column()));
            if (cell) {
                QString fmt = cell->getStyle().numberFormat;
                if (fmt == "Checkbox") {
                    menu.addAction("Remove Checkbox", this, [this]() {
                        if (!m_spreadsheet || !m_model) return;
                        auto sel = selectionModel()->selectedIndexes();
                        if (sel.isEmpty()) return;
                        std::vector<CellSnapshot> before, after;
                        for (const auto& idx : sel) {
                            CellAddress addr(logicalRow(idx), idx.column());
                            before.push_back(m_spreadsheet->takeCellSnapshot(addr));
                            auto c = m_spreadsheet->getCell(addr);
                            CellStyle st = c->getStyle();
                            if (st.numberFormat == "Checkbox") {
                                st.numberFormat = "General";
                                c->setStyle(st);
                                c->setValue(QVariant());
                            }
                            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                        }
                        m_spreadsheet->getUndoManager().pushCommand(
                            std::make_unique<MultiCellEditCommand>(before, after, "Remove Checkbox"));
                        QModelIndex tl = m_model->index(sel.first().row(), sel.first().column());
                        QModelIndex br = m_model->index(sel.last().row(), sel.last().column());
                        emit m_model->dataChanged(tl, br);
                    });
                    menu.addSeparator();
                } else if (fmt == "Picklist") {
                    menu.addAction("Manage Picklist...", this, [this, cur]() {
                        showPicklistPopup(cur);
                    });
                    menu.addAction("Remove Picklist", this, [this]() {
                        if (!m_spreadsheet || !m_model) return;
                        auto sel = selectionModel()->selectedIndexes();
                        if (sel.isEmpty()) return;
                        std::vector<CellSnapshot> before, after;
                        for (const auto& idx : sel) {
                            CellAddress addr(logicalRow(idx), idx.column());
                            before.push_back(m_spreadsheet->takeCellSnapshot(addr));
                            auto c = m_spreadsheet->getCell(addr);
                            CellStyle st = c->getStyle();
                            if (st.numberFormat == "Picklist") {
                                st.numberFormat = "General";
                                c->setStyle(st);
                                c->setValue(QVariant());
                            }
                            // Also remove validation rule if exists
                            auto& rules = m_spreadsheet->getValidationRules();
                            for (int ri = (int)rules.size() - 1; ri >= 0; --ri) {
                                if (rules[ri].range.contains(addr)) {
                                    m_spreadsheet->removeValidationRule(ri);
                                    break;
                                }
                            }
                            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                        }
                        m_spreadsheet->getUndoManager().pushCommand(
                            std::make_unique<MultiCellEditCommand>(before, after, "Remove Picklist"));
                        QModelIndex tl = m_model->index(sel.first().row(), sel.first().column());
                        QModelIndex br = m_model->index(sel.last().row(), sel.last().column());
                        emit m_model->dataChanged(tl, br);
                    });
                    menu.addSeparator();
                }
            }
        }
    }

    // Comment actions
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            auto cell = m_spreadsheet->getCellIfExists(logicalRow(cur), cur.column());
            bool hasComment = cell && cell->hasComment();
            menu.addAction(hasComment ? "Edit Co&mment..." : "Insert Co&mment...",
                           QKeySequence(Qt::SHIFT | Qt::Key_F2),
                           this, &SpreadsheetView::insertOrEditComment);
            if (hasComment) {
                menu.addAction("Delete Comment", this, &SpreadsheetView::deleteComment);
            }
            menu.addSeparator();
        }
    }

    // Hyperlink actions
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            auto cell = m_spreadsheet->getCellIfExists(logicalRow(cur), cur.column());
            bool hasLink = cell && cell->hasHyperlink();
            menu.addAction(hasLink ? "Edit Hyperlink..." : "Insert Hyperlink...",
                           this, &SpreadsheetView::insertOrEditHyperlink);
            if (hasLink) {
                menu.addAction("Open Hyperlink", this, [this, cur]() {
                    openHyperlink(cur.row(), cur.column());
                });
                menu.addAction("Remove Hyperlink", this, &SpreadsheetView::removeHyperlink);
            }
            menu.addSeparator();
        }
    }

    menu.addAction("&Format Cells...", QKeySequence(Qt::CTRL | Qt::Key_1), this, [this]() {
        emit formatCellsRequested();
    });

    menu.addSeparator();

    // Sort submenu
    QMenu* sortMenu = menu.addMenu("Sor&t");
    sortMenu->addAction("Sort &A to Z", this, &SpreadsheetView::sortAscending);
    sortMenu->addAction("Sort &Z to A", this, &SpreadsheetView::sortDescending);
    sortMenu->addSeparator();
    sortMenu->addAction("&Custom Sort...", this, &SpreadsheetView::showSortDialog);

    menu.addSeparator();

    // Define Name
    menu.addAction("De&fine Name...", [this]() {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) return;
        QString addr = CellAddress(logicalRow(cur), cur.column()).toString();
        bool ok;
        QString name = QInputDialog::getText(this, "Define Name",
            QString("Name for %1:").arg(addr), QLineEdit::Normal, "", &ok);
        if (ok && !name.isEmpty()) {
            CellRange range(CellAddress(logicalRow(cur), cur.column()),
                           CellAddress(logicalRow(cur), cur.column()));
            m_spreadsheet->addNamedRange(name, range);
        }
    });

    // Hyperlink
    menu.addAction("&Hyperlink...", QKeySequence(Qt::CTRL | Qt::Key_K),
                   this, &SpreadsheetView::insertOrEditHyperlink);

    menu.exec(viewport()->mapToGlobal(pos));
}

// ============== Insert/Delete with shift ==============

void SpreadsheetView::insertCellsShiftRight() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        minRow = qMin(minRow, lr);
        maxRow = qMax(maxRow, lr);
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->insertCellsShiftRight(range);
    refreshView();
}

void SpreadsheetView::insertCellsShiftDown() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        minRow = qMin(minRow, lr);
        maxRow = qMax(maxRow, lr);
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->insertCellsShiftDown(range);
    refreshView();
}

void SpreadsheetView::insertEntireRow() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> rows;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        rows.insert(logicalRow(current));
    } else {
        for (const auto& idx : selected) {
            rows.insert(logicalRow(idx));
        }
    }

    // Sort descending so inserts don't shift subsequent indices
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());

    // Build a single compound undo command for all row inserts
    auto compound = std::make_unique<CompoundUndoCommand>("Insert Row");
    for (int row : sortedRows) {
        compound->addChild(std::make_unique<InsertRowCommand>(row, 1));
    }

    if (m_model && m_model->isVirtualMode()) {
        int totalRows = m_spreadsheet->getRowCount();
        if (totalRows > 500000) {
            qDebug() << "[InsertRow] Skipped: dataset too large (" << totalRows << " rows)";
        }
        m_spreadsheet->getUndoManager().execute(std::move(compound), m_spreadsheet.get());
        m_model->resetModel();
    } else {
        int minRow = *std::min_element(rows.begin(), rows.end());
        int count = static_cast<int>(rows.size());
        verticalHeader()->blockSignals(true);
        m_model->beginRowInsertion(minRow, count);
        m_spreadsheet->getUndoManager().execute(std::move(compound), m_spreadsheet.get());
        m_model->endRowInsertion();
        verticalHeader()->blockSignals(false);
    }
}

void SpreadsheetView::insertEntireColumn() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> cols;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        cols.insert(current.column());
    } else {
        for (const auto& idx : selected) {
            cols.insert(idx.column());
        }
    }

    // Sort descending so inserts don't shift subsequent indices
    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());

    // Build a single compound undo command for all column inserts
    auto compound = std::make_unique<CompoundUndoCommand>("Insert Column");
    for (int col : sortedCols) {
        compound->addChild(std::make_unique<InsertColumnCommand>(col, 1, col));
    }

    int minCol = *std::min_element(cols.begin(), cols.end());
    int count = static_cast<int>(cols.size());

    horizontalHeader()->blockSignals(true);
    m_model->beginColumnInsertion(minCol, count);
    m_spreadsheet->getUndoManager().execute(std::move(compound), m_spreadsheet.get());
    m_model->endColumnInsertion();
    horizontalHeader()->blockSignals(false);
}

void SpreadsheetView::deleteCellsShiftLeft() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        minRow = qMin(minRow, lr);
        maxRow = qMax(maxRow, lr);
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->deleteCellsShiftLeft(range);
    refreshView();
}

void SpreadsheetView::deleteCellsShiftUp() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        minRow = qMin(minRow, lr);
        maxRow = qMax(maxRow, lr);
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->deleteCellsShiftUp(range);
    refreshView();
}

void SpreadsheetView::deleteEntireRow() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> rows;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        rows.insert(logicalRow(current));
    } else {
        for (const auto& idx : selected) {
            rows.insert(logicalRow(idx));
        }
    }

    int focusRow = *std::min_element(rows.begin(), rows.end());
    int focusCol = current.column();

    // Delete rows high-to-low to preserve indices.
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());

    // Build a single compound undo command for all row deletes
    auto compound = std::make_unique<CompoundUndoCommand>("Delete Row");
    for (int row : sortedRows) {
        // Snapshot deleted cells for undo
        std::vector<CellSnapshot> deleted;
        int maxCol = m_spreadsheet->getMaxColumn();
        for (int c = 0; c <= maxCol; ++c) {
            auto cell = m_spreadsheet->getCellIfExists(row, c);
            if (cell) {
                deleted.push_back(m_spreadsheet->takeCellSnapshot(CellAddress(row, c)));
            }
        }
        std::vector<int> savedHeights = { verticalHeader()->defaultSectionSize() };
        compound->addChild(std::make_unique<DeleteRowCommand>(row, 1, deleted, savedHeights));
    }

    // Block header signals so Qt's internal section adjustments don't
    // trigger onVerticalSectionResized (which would corrupt m_rowHeights).
    verticalHeader()->blockSignals(true);

    bool isVirtual = m_model && m_model->isVirtualMode();
    int count = static_cast<int>(rows.size());

    if (!isVirtual) {
        m_model->beginRowRemoval(focusRow, count);
    }

    m_spreadsheet->getUndoManager().execute(std::move(compound), m_spreadsheet.get());

    if (!isVirtual) {
        m_model->endRowRemoval();
    }

    if (isVirtual) {
        m_model->resetModel();
    }

    verticalHeader()->blockSignals(false);

    // Restore focus
    if (m_model) {
        int clampedRow = qMin(focusRow, m_model->rowCount() - 1);
        QModelIndex focusIdx = m_model->index(clampedRow, focusCol);
        selectionModel()->setCurrentIndex(focusIdx, QItemSelectionModel::ClearAndSelect);
        scrollTo(focusIdx);
        emitCellSelected(focusIdx);
    }
}

void SpreadsheetView::deleteEntireColumn() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> cols;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        cols.insert(current.column());
    } else {
        for (const auto& idx : selected) {
            cols.insert(idx.column());
        }
    }

    int focusCol = *std::min_element(cols.begin(), cols.end());
    int focusRow = current.row();

    // Delete columns high-to-low to preserve indices.
    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());

    // Build a single compound undo command for all column deletes
    auto compound = std::make_unique<CompoundUndoCommand>("Delete Column");
    for (int col : sortedCols) {
        // Snapshot deleted cells for undo
        std::vector<CellSnapshot> deleted;
        int maxRow = m_spreadsheet->getMaxRow();
        for (int r = 0; r <= maxRow; ++r) {
            auto cell = m_spreadsheet->getCellIfExists(r, col);
            if (cell) {
                deleted.push_back(m_spreadsheet->takeCellSnapshot(CellAddress(r, col)));
            }
        }
        std::vector<int> savedWidths = { horizontalHeader()->sectionSize(col) };
        compound->addChild(std::make_unique<DeleteColumnCommand>(col, 1, deleted, savedWidths));
    }

    // Block header signals so Qt's internal section adjustments don't
    // trigger onHorizontalSectionResized (which would corrupt m_columnWidths).
    horizontalHeader()->blockSignals(true);

    int count = static_cast<int>(cols.size());
    m_model->beginColumnRemoval(focusCol, count);
    m_spreadsheet->getUndoManager().execute(std::move(compound), m_spreadsheet.get());
    m_model->endColumnRemoval();

    horizontalHeader()->blockSignals(false);

    // Restore focus
    if (m_model) {
        int clampedCol = qMin(focusCol, m_model->columnCount() - 1);
        QModelIndex focusIdx = m_model->index(focusRow, clampedCol);
        selectionModel()->setCurrentIndex(focusIdx, QItemSelectionModel::ClearAndSelect);
        scrollTo(focusIdx);
        emitCellSelected(focusIdx);
    }
}

// ============== Autofit ==============

void SpreadsheetView::autofitColumn(int column) {
    if (!m_model) return;
    int maxWidth = 40;
    for (int row = 0; row < m_model->rowCount(); ++row) {
        QModelIndex idx = m_model->index(row, column);
        QString text = idx.data(Qt::DisplayRole).toString();
        if (text.isEmpty()) continue;

        QFont cellFont = font();
        QVariant fontData = idx.data(Qt::FontRole);
        if (fontData.isValid()) {
            cellFont = fontData.value<QFont>();
        }
        QFontMetrics fm(cellFont);
        int width = fm.horizontalAdvance(text) + 16;
        maxWidth = qMax(maxWidth, width);
    }
    horizontalHeader()->resizeSection(column, maxWidth);
}

void SpreadsheetView::autofitRow(int row) {
    if (!m_model) return;
    int maxHeight = 18;
    for (int col = 0; col < m_model->columnCount(); ++col) {
        QModelIndex idx = m_model->index(row, col);
        QString text = idx.data(Qt::DisplayRole).toString();
        if (text.isEmpty()) continue;

        QFont cellFont = font();
        QVariant fontData = idx.data(Qt::FontRole);
        if (fontData.isValid()) {
            cellFont = fontData.value<QFont>();
        }
        QFontMetrics fm(cellFont);
        int height = fm.height() + 6;
        maxHeight = qMax(maxHeight, height);
    }
    verticalHeader()->resizeSection(row, maxHeight);
}

void SpreadsheetView::autofitSelectedColumns() {
    QModelIndexList selected = selectionModel()->selectedColumns();
    if (selected.isEmpty()) {
        autofitColumn(currentIndex().column());
    } else {
        for (const auto& idx : selected) {
            autofitColumn(idx.column());
        }
    }
}

void SpreadsheetView::autofitSelectedRows() {
    QModelIndexList selected = selectionModel()->selectedRows();
    if (selected.isEmpty()) {
        autofitRow(currentIndex().row());
    } else {
        for (const auto& idx : selected) {
            autofitRow(idx.row());
        }
    }
}

// ============== UI Operations ==============

void SpreadsheetView::setRowHeight(int row, int height) {
    if (row >= 0 && height > 0) {
        verticalHeader()->resizeSection(row, height);
        if (m_spreadsheet) m_spreadsheet->setRowHeight(row, height);
    }
}

void SpreadsheetView::setColumnWidth(int col, int width) {
    if (col >= 0 && width > 0) {
        horizontalHeader()->resizeSection(col, width);
        if (m_spreadsheet) m_spreadsheet->setColumnWidth(col, width);
    }
}

void SpreadsheetView::applyStoredDimensions() {
    if (!m_spreadsheet) return;

    // Reset all sections to default sizes first (so new sheets get clean dimensions)
    int defaultColW = 80;
    int defaultRowH = 25;
    horizontalHeader()->blockSignals(true);
    verticalHeader()->blockSignals(true);
    for (int c = 0; c < model()->columnCount(); ++c)
        horizontalHeader()->resizeSection(c, defaultColW);
    for (int r = 0; r < model()->rowCount(); ++r)
        verticalHeader()->resizeSection(r, defaultRowH);

    // Then apply stored per-sheet dimensions
    for (auto& [col, width] : m_spreadsheet->getColumnWidths()) {
        if (col >= 0 && col < model()->columnCount() && width > 0)
            horizontalHeader()->resizeSection(col, width);
    }
    for (auto& [row, height] : m_spreadsheet->getRowHeights()) {
        if (row >= 0 && row < model()->rowCount() && height > 0)
            verticalHeader()->resizeSection(row, height);
    }
    horizontalHeader()->blockSignals(false);
    verticalHeader()->blockSignals(false);
}

void SpreadsheetView::setGridlinesVisible(bool visible) {
    if (m_delegate) {
        m_delegate->setShowGridlines(visible);
        viewport()->update();
    }
    // Propagate to freeze overlay delegates
    auto syncDelegate = [visible](QTableView* v) {
        if (!v) return;
        auto* d = qobject_cast<CellDelegate*>(v->itemDelegate());
        if (d) { d->setShowGridlines(visible); v->viewport()->update(); }
    };
    syncDelegate(m_frozenRowView);
    syncDelegate(m_frozenColView);
    syncDelegate(m_frozenCornerView);
}

void SpreadsheetView::refreshView() {
    if (m_model) {
        // Full reset: forces the view to re-query ALL data roles (colors, fonts, etc.)
        // This is needed when document theme changes so theme-referenced colors resolve correctly.
        m_model->resetModel();
    }
}

void SpreadsheetView::toggleFormulaView() {
    m_showFormulas = !m_showFormulas;
    if (m_model) {
        m_model->setShowFormulas(m_showFormulas);
        m_model->resetModel();
    }
}

void SpreadsheetView::insertOrEditComment() {
    QModelIndex cur = currentIndex();
    if (!cur.isValid() || !m_spreadsheet) return;

    auto cell = m_spreadsheet->getCell(CellAddress(logicalRow(cur), cur.column()));
    QString existing = cell->getComment();

    bool ok = false;
    QString comment = QInputDialog::getMultiLineText(
        this, cell->hasComment() ? "Edit Comment" : "Insert Comment",
        "Comment:", existing, &ok);

    if (ok) {
        cell->setComment(comment);
        viewport()->update();
    }
}

void SpreadsheetView::deleteComment() {
    QModelIndex cur = currentIndex();
    if (!cur.isValid() || !m_spreadsheet) return;

    auto cell = m_spreadsheet->getCellIfExists(logicalRow(cur), cur.column());
    if (cell && cell->hasComment()) {
        cell->setComment(QString());
        viewport()->update();
    }
}

// ============== Hyperlinks ==============

void SpreadsheetView::insertOrEditHyperlink() {
    QModelIndex cur = currentIndex();
    if (!cur.isValid() || !m_spreadsheet) return;

    auto cell = m_spreadsheet->getCell(CellAddress(logicalRow(cur), cur.column()));
    QString existingUrl = cell->getHyperlink();
    QString existingText = cell->getValue().toString();

    // Create a small dialog for URL and display text
    QDialog dlg(this);
    dlg.setWindowTitle(existingUrl.isEmpty() ? "Insert Hyperlink" : "Edit Hyperlink");
    dlg.setMinimumWidth(400);
    dlg.setStyleSheet(ThemeManager::instance().dialogStylesheet());

    auto* layout = new QVBoxLayout(&dlg);
    auto* form = new QFormLayout();

    auto* urlEdit = new QLineEdit(&dlg);
    urlEdit->setText(existingUrl);
    urlEdit->setPlaceholderText("https://example.com");
    form->addRow("URL:", urlEdit);

    auto* textEdit = new QLineEdit(&dlg);
    textEdit->setText(existingText);
    textEdit->setPlaceholderText("Display text (optional)");
    form->addRow("Display text:", textEdit);

    layout->addLayout(form);

    auto* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    layout->addWidget(buttons);

    connect(buttons, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);

    if (dlg.exec() == QDialog::Accepted) {
        QString url = urlEdit->text().trimmed();
        QString text = textEdit->text().trimmed();

        if (!url.isEmpty()) {
            cell->setHyperlink(url);
            if (!text.isEmpty() && text != existingText) {
                cell->setValue(text);
            } else if (existingText.isEmpty()) {
                cell->setValue(url);
            }
        }
        if (m_model) {
            emit m_model->dataChanged(cur, cur);
        }
        viewport()->update();
    }
}

void SpreadsheetView::removeHyperlink() {
    QModelIndex cur = currentIndex();
    if (!cur.isValid() || !m_spreadsheet) return;

    auto cell = m_spreadsheet->getCellIfExists(logicalRow(cur), cur.column());
    if (cell && cell->hasHyperlink()) {
        cell->setHyperlink(QString());
        if (m_model) {
            emit m_model->dataChanged(cur, cur);
        }
        viewport()->update();
    }
}

void SpreadsheetView::openHyperlink(int row, int col) {
    if (!m_spreadsheet) return;
    auto cell = m_spreadsheet->getCellIfExists(row, col);
    if (cell && cell->hasHyperlink()) {
        QUrl url(cell->getHyperlink());
        if (url.scheme().isEmpty()) {
            url = QUrl("https://" + cell->getHyperlink());
        }
        QDesktopServices::openUrl(url);
    }
}

bool SpreadsheetView::viewportEvent(QEvent* event) {
    if (event->type() == QEvent::ToolTip) {
        QHelpEvent* helpEvent = static_cast<QHelpEvent*>(event);
        QModelIndex index = indexAt(helpEvent->pos());
        if (index.isValid() && m_spreadsheet) {
            auto cell = m_spreadsheet->getCellIfExists(logicalRow(index), index.column());
            if (cell) {
                QStringList tips;
                if (cell->hasComment()) tips << cell->getComment();
                if (cell->hasHyperlink()) tips << "Ctrl+Click to open: " + cell->getHyperlink();
                if (!tips.isEmpty()) {
                    // Excel-style yellow tooltip for comments
                    if (cell->hasComment()) {
                        // Apply yellow Excel-style tooltip palette
                        QPalette ttPalette = QToolTip::palette();
                        ttPalette.setColor(QPalette::ToolTipBase, QColor("#FFFFD0"));
                        ttPalette.setColor(QPalette::ToolTipText, QColor("#000000"));
                        QToolTip::setPalette(ttPalette);
                        QToolTip::showText(helpEvent->globalPos(), tips.join("\n"), this);
                    } else {
                        QToolTip::showText(helpEvent->globalPos(), tips.join("\n"), this);
                    }
                    return true;
                }
            }
        }
        QToolTip::hideText();
        event->ignore();
        return true;
    }
    return QTableView::viewportEvent(event);
}

void SpreadsheetView::setFrozenRow(int row) {
    m_frozenRow = row;
    if (m_frozenRow > 0 || m_frozenCol > 0)
        setupFreezeViews();
    else
        destroyFreezeViews();
}

void SpreadsheetView::setFrozenColumn(int col) {
    m_frozenCol = col;
    if (m_frozenRow > 0 || m_frozenCol > 0)
        setupFreezeViews();
    else
        destroyFreezeViews();
}

QTableView* SpreadsheetView::createFreezeOverlay() {
    auto* v = new QTableView(this);
    v->setModel(model());
    auto* delegate = new CellDelegate(v);
    if (m_delegate) delegate->setShowGridlines(m_delegate->showGridlines());
    v->setItemDelegate(delegate);
    v->setShowGrid(false);
    v->horizontalHeader()->hide();
    v->verticalHeader()->hide();
    v->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    v->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    v->setHorizontalScrollMode(horizontalScrollMode());
    v->setVerticalScrollMode(verticalScrollMode());
    v->setSelectionMode(QAbstractItemView::NoSelection);
    v->setFocusPolicy(Qt::NoFocus);
    v->setAttribute(Qt::WA_TransparentForMouseEvents);
    v->setFont(font());
    v->setStyleSheet(
        "QTableView { background: white; border: none; }"
        "QTableView::item { padding: 0; border: none; background: transparent; }"
        "QTableView::item:selected { background: transparent; }"
    );

    // Sync dimensions from main view
    v->horizontalHeader()->setDefaultSectionSize(horizontalHeader()->defaultSectionSize());
    v->verticalHeader()->setDefaultSectionSize(verticalHeader()->defaultSectionSize());
    for (int c = 0; c < model()->columnCount(); c++)
        v->setColumnWidth(c, columnWidth(c));
    for (int r = 0; r < model()->rowCount(); r++)
        v->setRowHeight(r, rowHeight(r));

    return v;
}

void SpreadsheetView::setupFreezeViews() {
    destroyFreezeViews();
    if (m_frozenRow <= 0 && m_frozenCol <= 0) return;

    // Frozen row view (top strip, scrolls horizontally with main)
    if (m_frozenRow > 0) {
        m_frozenRowView = createFreezeOverlay();
        m_freezeConnections.append(
            connect(horizontalScrollBar(), &QScrollBar::valueChanged,
                    m_frozenRowView->horizontalScrollBar(), &QScrollBar::setValue));
    }

    // Frozen column view (left strip, scrolls vertically with main)
    if (m_frozenCol > 0) {
        m_frozenColView = createFreezeOverlay();
        m_freezeConnections.append(
            connect(verticalScrollBar(), &QScrollBar::valueChanged,
                    m_frozenColView->verticalScrollBar(), &QScrollBar::setValue));
    }

    // Corner view (no scrolling, sits on top of both)
    if (m_frozenRow > 0 && m_frozenCol > 0) {
        m_frozenCornerView = createFreezeOverlay();
    }

    // Freeze divider lines
    if (m_frozenRow > 0) {
        m_freezeHLine = new QWidget(this);
        m_freezeHLine->setFixedHeight(2);
        m_freezeHLine->setStyleSheet(QString("background: %1;").arg(ThemeManager::instance().currentTheme().freezeLineColor.name()));
        m_freezeHLine->setAttribute(Qt::WA_TransparentForMouseEvents);
    }
    if (m_frozenCol > 0) {
        m_freezeVLine = new QWidget(this);
        m_freezeVLine->setFixedWidth(2);
        m_freezeVLine->setStyleSheet(QString("background: %1;").arg(ThemeManager::instance().currentTheme().freezeLineColor.name()));
        m_freezeVLine->setAttribute(Qt::WA_TransparentForMouseEvents);
    }

    // Sync column width changes from main to overlays
    m_freezeConnections.append(
        connect(horizontalHeader(), &QHeaderView::sectionResized,
                this, [this](int idx, int, int newSize) {
            if (m_frozenRowView) m_frozenRowView->setColumnWidth(idx, newSize);
            if (m_frozenColView) m_frozenColView->setColumnWidth(idx, newSize);
            if (m_frozenCornerView) m_frozenCornerView->setColumnWidth(idx, newSize);
            updateFreezeGeometry();
        }));

    // Sync row height changes from main to overlays
    m_freezeConnections.append(
        connect(verticalHeader(), &QHeaderView::sectionResized,
                this, [this](int idx, int, int newSize) {
            if (m_frozenRowView) m_frozenRowView->setRowHeight(idx, newSize);
            if (m_frozenColView) m_frozenColView->setRowHeight(idx, newSize);
            if (m_frozenCornerView) m_frozenCornerView->setRowHeight(idx, newSize);
            updateFreezeGeometry();
        }));

    updateFreezeGeometry();
}

void SpreadsheetView::destroyFreezeViews() {
    for (auto& conn : m_freezeConnections)
        disconnect(conn);
    m_freezeConnections.clear();

    delete m_frozenRowView; m_frozenRowView = nullptr;
    delete m_frozenColView; m_frozenColView = nullptr;
    delete m_frozenCornerView; m_frozenCornerView = nullptr;
    delete m_freezeHLine; m_freezeHLine = nullptr;
    delete m_freezeVLine; m_freezeVLine = nullptr;
}

void SpreadsheetView::updateFreezeGeometry() {
    if (m_frozenRow <= 0 && m_frozenCol <= 0) return;

    int fw = frameWidth();
    int hdrH = horizontalHeader()->height();
    int hdrW = verticalHeader()->width();
    int vpW = viewport()->width();
    int vpH = viewport()->height();

    int frozenH = 0;
    for (int r = 0; r < m_frozenRow && r < model()->rowCount(); r++) {
        if (!isRowHidden(r))
            frozenH += rowHeight(r);
    }

    int frozenW = 0;
    for (int c = 0; c < m_frozenCol && c < model()->columnCount(); c++) {
        if (!isColumnHidden(c))
            frozenW += columnWidth(c);
    }

    if (m_frozenRowView) {
        m_frozenRowView->setGeometry(hdrW + fw, hdrH + fw, vpW, frozenH);
        m_frozenRowView->show();
    }
    if (m_frozenColView) {
        m_frozenColView->setGeometry(hdrW + fw, hdrH + fw, frozenW, vpH);
        m_frozenColView->show();
    }
    if (m_frozenCornerView) {
        m_frozenCornerView->setGeometry(hdrW + fw, hdrH + fw, frozenW, frozenH);
        m_frozenCornerView->show();
    }

    // Divider lines at the freeze boundary
    if (m_freezeHLine) {
        m_freezeHLine->setGeometry(hdrW + fw, hdrH + fw + frozenH - 1, vpW, 2);
        m_freezeHLine->show();
        m_freezeHLine->raise();
    }
    if (m_freezeVLine) {
        m_freezeVLine->setGeometry(hdrW + fw + frozenW - 1, hdrH + fw, 2, vpH);
        m_freezeVLine->show();
        m_freezeVLine->raise();
    }

    // Z-order: divider lines on top, then corner, then column, then row
    if (m_frozenRowView) m_frozenRowView->raise();
    if (m_frozenColView) m_frozenColView->raise();
    if (m_frozenCornerView) m_frozenCornerView->raise();
    if (m_freezeHLine) m_freezeHLine->raise();
    if (m_freezeVLine) m_freezeVLine->raise();
}

void SpreadsheetView::resizeEvent(QResizeEvent* event) {
    QTableView::resizeEvent(event);
    updateFreezeGeometry();

    // Update visible row tracking and reposition virtual scrollbar
    if (m_model && m_model->isVirtualMode()) {
        int rh = verticalHeader()->defaultSectionSize();
        if (rh <= 0) rh = 25;
        int visible = viewport()->height() / rh + 2;
        m_model->setVisibleRows(std::max(20, visible));
        // Reposition virtual scrollbar on resize
        if (m_virtualScrollBar) {
            int sbWidth = m_virtualScrollBar->sizeHint().width();
            m_virtualScrollBar->setGeometry(width() - sbWidth, 0, sbWidth, height());
        }
    }
}

void SpreadsheetView::scrollTo(const QModelIndex& index, ScrollHint hint) {
    // With the buffer window approach, Qt's native scrollTo works within the buffer.
    QTableView::scrollTo(index, hint);
}

void SpreadsheetView::scrollContentsBy(int dx, int dy) {
    // Let Qt do native smooth pixel-level scrolling
    QTableView::scrollContentsBy(dx, dy);

    if (m_recentering) return;

    // In virtual mode, check if we need to recenter the buffer window
    if (m_model && m_model->isVirtualMode() && dy != 0) {
        int topModelRow = rowAt(0);
        if (topModelRow < 0) topModelRow = 0;
        int windowRows = m_model->rowCount();
        int margin = SpreadsheetModel::RECENTER_MARGIN;

        bool needRecenter = false;
        if (topModelRow > windowRows - margin && m_model->windowBase() + windowRows < m_model->totalLogicalRows()) {
            needRecenter = true;  // approaching bottom of window
        }
        if (topModelRow < margin && m_model->windowBase() > 0) {
            needRecenter = true;  // approaching top of window
        }

        if (needRecenter) {
            int logicalTopRow = m_model->toLogicalRow(topModelRow);
            m_recentering = true;

            int shift = m_model->recenterWindow(logicalTopRow);
            if (shift != 0) {
                // Compensate Qt's scroll position so the view doesn't jump
                int rh = verticalHeader()->defaultSectionSize();
                if (rh <= 0) rh = 25;
                auto* vsb = verticalScrollBar();
                vsb->setValue(vsb->value() - shift * rh);
            }

            m_recentering = false;
        }

        // Update virtual scrollbar position to reflect current logical position
        if (m_virtualScrollBar) {
            int logicalTop = m_model->toLogicalRow(rowAt(0));
            m_virtualScrollBar->blockSignals(true);
            m_virtualScrollBar->setValue(logicalTop);
            m_virtualScrollBar->blockSignals(false);
        }
    }
}

void SpreadsheetView::wheelEvent(QWheelEvent* event) {
    // Ctrl+Wheel = zoom in/out
    if (event->modifiers() & Qt::ControlModifier) {
        int delta = event->angleDelta().y();
        if (delta > 0) zoomIn();
        else if (delta < 0) zoomOut();
        event->accept();
        return;
    }

    if (m_model && m_model->isVirtualMode()) {
        // At dataset boundaries, block momentum to prevent bounce
        auto phase = event->phase();
        if (phase == Qt::ScrollMomentum) {
            int topModelRow = rowAt(0);
            int logicalTop = m_model->toLogicalRow(topModelRow < 0 ? 0 : topModelRow);
            if (logicalTop <= 0 || logicalTop >= m_model->totalLogicalRows() - m_model->visibleRows()) {
                event->accept();
                return;
            }
        }
    }
    // Let Qt handle scrolling natively — smooth pixel blitting
    QTableView::wheelEvent(event);
}

void SpreadsheetView::navigateToLogicalRow(int logicalRow, int col) {
    if (!m_model) return;

    if (m_model->isVirtualMode()) {
        // Check if target row is within current buffer window
        int modelRow = m_model->toModelRow(logicalRow);
        if (modelRow < 0 || modelRow >= m_model->rowCount()) {
            // Row not in current window — recenter window (uses dataChanged, not reset)
            int rh = verticalHeader()->defaultSectionSize();
            if (rh <= 0) rh = 25;
            m_recentering = true;
            int shift = m_model->recenterWindow(logicalRow);
            if (shift != 0) {
                verticalScrollBar()->setValue(verticalScrollBar()->value() - shift * rh);
            }
            m_recentering = false;
            modelRow = m_model->toModelRow(logicalRow);
            // Update virtual scrollbar
            if (m_virtualScrollBar) {
                m_virtualScrollBar->blockSignals(true);
                m_virtualScrollBar->setValue(logicalRow);
                m_virtualScrollBar->blockSignals(false);
            }
        }

        QModelIndex target = model()->index(modelRow, col);
        if (target.isValid()) {
            selectionModel()->setCurrentIndex(target, QItemSelectionModel::ClearAndSelect);
            scrollTo(target);
        }
    } else {
        QModelIndex target = model()->index(logicalRow, col);
        if (target.isValid()) {
            selectionModel()->setCurrentIndex(target, QItemSelectionModel::ClearAndSelect);
            scrollTo(target);
        }
    }
}

void SpreadsheetView::zoomIn() {
    setZoomLevel(m_zoomLevel + 10);
}

void SpreadsheetView::zoomOut() {
    setZoomLevel(m_zoomLevel - 10);
}

void SpreadsheetView::resetZoom() {
    setZoomLevel(100);
}

void SpreadsheetView::setZoomLevel(int percent) {
    percent = qBound(25, percent, 400);
    if (percent == m_zoomLevel) return;
    m_zoomLevel = percent;

    double scale = m_zoomLevel / 100.0;

    // Scale font
    QFont f = this->font();
    f.setPointSize(qMax(6, static_cast<int>(m_baseFontSize * scale)));
    setFont(f);

    // Scale row heights
    int defaultRowHeight = qMax(12, static_cast<int>(22 * scale));
    verticalHeader()->setDefaultSectionSize(defaultRowHeight);

    // Scale column widths
    int defaultColWidth = qMax(30, static_cast<int>(80 * scale));
    horizontalHeader()->setDefaultSectionSize(defaultColWidth);

    // Scale custom row/col dimensions if spreadsheet has them
    if (m_spreadsheet) {
        for (const auto& [row, height] : m_spreadsheet->getRowHeights()) {
            setRowHeight(row, qMax(12, static_cast<int>(height * scale)));
        }
        for (const auto& [col, width] : m_spreadsheet->getColumnWidths()) {
            setColumnWidth(col, qMax(30, static_cast<int>(width * scale)));
        }
    }

    // Scale header font
    QFont hf = horizontalHeader()->font();
    hf.setPointSize(qMax(6, static_cast<int>(m_baseFontSize * scale)));
    horizontalHeader()->setFont(hf);
    verticalHeader()->setFont(hf);

    viewport()->update();
    emit zoomChanged(m_zoomLevel);
}

// ============== Event handlers ==============

void SpreadsheetView::keyPressEvent(QKeyEvent* event) {
    bool ctrl = event->modifiers() & Qt::ControlModifier;
    bool shift = event->modifiers() & Qt::ShiftModifier;

    // Sheet protection helper: returns true if the current cell is locked on a protected sheet
    auto isCellProtected = [this]() -> bool {
        if (!m_spreadsheet || !m_spreadsheet->isProtected()) return false;
        QModelIndex cur = currentIndex();
        if (!cur.isValid()) return false;
        int row = m_model ? m_model->toLogicalRow(cur.row()) : cur.row();
        int col = cur.column();
        auto cell = m_spreadsheet->getCell(row, col);
        return cell->getStyle().locked;
    };

    // Delete / Backspace: clear selection (on Mac, "Delete" key = Backspace)
    if (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace) {
        if (state() != QAbstractItemView::EditingState) {
            if (isCellProtected()) {
                QMessageBox::warning(this, "Protected Sheet",
                    "The cell or chart you're trying to change is on a protected sheet.\n"
                    "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
                event->accept();
                return;
            }
            deleteSelection();
            event->accept();
            return;
        }
        // If editing, let the editor handle the key
        QTableView::keyPressEvent(event);
        return;
    }

    // Alt+Enter (Win) / Option+Enter / Ctrl+Enter (Mac): insert line break in cell
    if ((event->key() == Qt::Key_Return || event->key() == Qt::Key_Enter)
        && (event->modifiers() & (Qt::AltModifier | Qt::ControlModifier))
        && !(event->modifiers() & Qt::ShiftModifier)) {
        QModelIndex idx = currentIndex();
        if (!idx.isValid() || !m_spreadsheet) { event->accept(); return; }

        // If already in multiline editor, insert newline directly
        if (m_multilineEditor && m_multilineEditor->isVisible()) {
            m_multilineEditor->textCursor().insertText("\n");
            event->accept();
            return;
        }

        // Get current text (from QLineEdit editor if editing, or from cell)
        QString text;
        int insertPos = -1;
        if (state() == QAbstractItemView::EditingState) {
            QWidget* editor = indexWidget(idx);
            if (!editor) editor = viewport()->findChild<QLineEdit*>();
            if (auto* lineEdit = qobject_cast<QLineEdit*>(editor)) {
                text = lineEdit->text();
                insertPos = lineEdit->cursorPosition();
            }
            setState(QAbstractItemView::NoState);
        } else {
            auto cell = m_spreadsheet->getCellIfExists(idx.row(), idx.column());
            text = cell ? cell->getValue().toString() : QString();
            insertPos = text.length();
        }

        // Insert newline and open multiline editor
        text.insert(insertPos, '\n');
        openMultilineEditor(idx, text, insertPos + 1);

        event->accept();
        return;
    }

    // If multiline editor is open, route keys to it
    if (m_multilineEditor && m_multilineEditor->isVisible()) {
        if (event->key() == Qt::Key_Escape) {
            cancelMultilineEditor();
            event->accept();
            return;
        }
        if ((event->key() == Qt::Key_Return || event->key() == Qt::Key_Enter)
            && !(event->modifiers() & (Qt::AltModifier | Qt::ControlModifier))) {
            commitMultilineEditor();
            event->accept();
            return;
        }
        // Let the multiline editor handle all other keys
        return;
    }

    // Enter/Return: commit and move down (within selection if multi-cell)
    if (event->key() == Qt::Key_Return || event->key() == Qt::Key_Enter) {
        m_formulaEditMode = false;  // Must be set before commit/close so overrides allow it
        if (state() == QAbstractItemView::EditingState) {
            QWidget* editor = indexWidget(currentIndex());
            if (!editor) editor = viewport()->findChild<QLineEdit*>();
            if (editor) {
                commitData(editor);
                closeEditor(editor, QAbstractItemDelegate::NoHint);
            }
        }

        // Navigate within selection bounds (Excel behavior)
        QItemSelection sel = selectionModel()->selection();
        if (!sel.isEmpty()) {
            QItemSelectionRange range = sel.first();
            if (range.width() > 1 || range.height() > 1) {
                int curRow = currentIndex().row();
                int curCol = currentIndex().column();
                if (shift) {
                    // Shift+Enter: move up within selection, wrap to bottom of previous column
                    int nextRow = curRow - 1;
                    int nextCol = curCol;
                    if (nextRow < range.top()) {
                        nextRow = range.bottom();
                        nextCol = curCol - 1;
                        if (nextCol < range.left()) nextCol = range.right();
                    }
                    QModelIndex idx = model()->index(nextRow, nextCol);
                    selectionModel()->setCurrentIndex(idx, QItemSelectionModel::NoUpdate);
                    scrollTo(idx);
                } else {
                    // Enter: move down within selection, wrap to top of next column
                    int nextRow = curRow + 1;
                    int nextCol = curCol;
                    if (nextRow > range.bottom()) {
                        nextRow = range.top();
                        nextCol = curCol + 1;
                        if (nextCol > range.right()) nextCol = range.left();
                    }
                    QModelIndex idx = model()->index(nextRow, nextCol);
                    selectionModel()->setCurrentIndex(idx, QItemSelectionModel::NoUpdate);
                    scrollTo(idx);
                }
                viewport()->update();
                event->accept();
                return;
            }
        }

        int newRow = currentIndex().row() + (shift ? -1 : 1);
        newRow = qBound(0, newRow, model()->rowCount() - 1);
        QModelIndex next = model()->index(newRow, currentIndex().column());
        if (next.isValid()) {
            setCurrentIndex(next);
            scrollTo(next);
            viewport()->update();
        }
        event->accept();
        return;
    }

    // Tab: commit and move right; Shift+Tab: move left (within selection if multi-cell)
    if (event->key() == Qt::Key_Tab || event->key() == Qt::Key_Backtab) {
        m_formulaEditMode = false;
        if (state() == QAbstractItemView::EditingState) {
            QWidget* editor = indexWidget(currentIndex());
            if (!editor) editor = viewport()->findChild<QLineEdit*>();
            if (editor) {
                commitData(editor);
                closeEditor(editor, QAbstractItemDelegate::NoHint);
            }
        }

        // Navigate within selection bounds (Excel behavior)
        QItemSelection sel = selectionModel()->selection();
        if (!sel.isEmpty()) {
            QItemSelectionRange range = sel.first();
            if (range.width() > 1 || range.height() > 1) {
                int curRow = currentIndex().row();
                int curCol = currentIndex().column();
                if (event->key() == Qt::Key_Backtab) {
                    // Shift+Tab: move left within selection, wrap to last column of previous row
                    int nextCol = curCol - 1;
                    int nextRow = curRow;
                    if (nextCol < range.left()) {
                        nextCol = range.right();
                        nextRow = curRow - 1;
                        if (nextRow < range.top()) nextRow = range.bottom();
                    }
                    QModelIndex idx = model()->index(nextRow, nextCol);
                    selectionModel()->setCurrentIndex(idx, QItemSelectionModel::NoUpdate);
                    scrollTo(idx);
                } else {
                    // Tab: move right within selection, wrap to first column of next row
                    int nextCol = curCol + 1;
                    int nextRow = curRow;
                    if (nextCol > range.right()) {
                        nextCol = range.left();
                        nextRow = curRow + 1;
                        if (nextRow > range.bottom()) nextRow = range.top();
                    }
                    QModelIndex idx = model()->index(nextRow, nextCol);
                    selectionModel()->setCurrentIndex(idx, QItemSelectionModel::NoUpdate);
                    scrollTo(idx);
                }
                viewport()->update();
                event->accept();
                return;
            }
        }

        int newCol = currentIndex().column() + (event->key() == Qt::Key_Backtab ? -1 : 1);
        newCol = qBound(0, newCol, model()->columnCount() - 1);
        QModelIndex next = model()->index(currentIndex().row(), newCol);
        if (next.isValid()) {
            setCurrentIndex(next);
            scrollTo(next);
            viewport()->update();
        }
        event->accept();
        return;
    }

    // Arrow keys in multi-cell selection: move active cell within selection, don't change selection
    // This matches Excel behavior where arrows navigate within the selected range.
    if (!ctrl && !shift && !m_formulaEditMode
        && state() != QAbstractItemView::EditingState
        && (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down
            || event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QItemSelection sel = selectionModel()->selection();
        if (!sel.isEmpty()) {
            QItemSelectionRange range = sel.first();
            if (range.width() > 1 || range.height() > 1) {
                int row = currentIndex().row();
                int col = currentIndex().column();
                switch (event->key()) {
                    case Qt::Key_Up:    row = qMax(range.top(), row - 1); break;
                    case Qt::Key_Down:  row = qMin(range.bottom(), row + 1); break;
                    case Qt::Key_Left:  col = qMax(range.left(), col - 1); break;
                    case Qt::Key_Right: col = qMin(range.right(), col + 1); break;
                }
                QModelIndex idx = model()->index(row, col);
                selectionModel()->setCurrentIndex(idx, QItemSelectionModel::NoUpdate);
                scrollTo(idx);
                event->accept();
                return;
            }
        }
    }

    // F2: Edit current cell / toggle Edit mode vs Enter mode (like Excel)
    if (event->key() == Qt::Key_F2) {
        if (state() == QAbstractItemView::EditingState) {
            // Toggle between Edit mode (arrows navigate within text) and
            // Enter mode (arrows commit and move to next cell) — Excel F2 behavior
            if (m_delegate) {
                bool newMode = !m_delegate->isFormulaEditMode();
                m_delegate->setFormulaEditMode(newMode);
                emit editModeChanged(newMode ? "Edit" : "Enter");
            }
        } else {
            if (isCellProtected()) {
                QMessageBox::warning(this, "Protected Sheet",
                    "The cell or chart you're trying to change is on a protected sheet.\n"
                    "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
                event->accept();
                return;
            }
            QModelIndex current = currentIndex();
            if (current.isValid()) {
                edit(current);
                // Enter Edit mode by default when F2 is pressed (arrows stay in text)
                if (m_delegate) {
                    m_delegate->setFormulaEditMode(true);
                }
                emit editModeChanged("Edit");
            }
        }
        event->accept();
        return;
    }

    // Escape: cancel editing / format painter / formula edit mode
    if (event->key() == Qt::Key_Escape) {
        // Clear marching ants on Escape (Excel behavior)
        if (m_hasClipboardRange) {
            clearClipboardRange();
        }
        // Cancel any pending cut operation
        m_isCutOperation = false;
        if (m_formulaEditMode) {
            m_formulaEditMode = false;
            if (state() == QAbstractItemView::EditingState) {
                QModelIndex idx = currentIndex();
                if (idx.isValid())
                    closeEditor(indexWidget(idx), QAbstractItemDelegate::RevertModelCache);
            }
            event->accept();
            return;
        }
        if (m_formatPainterActive) {
            m_formatPainterActive = false;
            viewport()->setCursor(Qt::ArrowCursor);
            event->accept();
            return;
        }
    }

    // ===== Ctrl/Cmd+Shift+V: Paste Special =====
    if (ctrl && shift && event->key() == Qt::Key_V) {
        pasteSpecial();
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+Shift+L: Toggle AutoFilter (Excel shortcut) =====
    if (ctrl && shift && event->key() == Qt::Key_L) {
        toggleAutoFilter();
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+E: Flash Fill (pattern detection from examples) =====
    if (ctrl && event->key() == Qt::Key_E) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndex cur = currentIndex();
        if (!cur.isValid()) { event->accept(); return; }

        int col = cur.column();
        int row = logicalRow(cur);

        // Find the first example (the cell with data in this column)
        // Look upward from current row to find the first non-empty cell
        int exampleRow = -1;
        QString exampleOutput;
        for (int r = row - 1; r >= 0; --r) {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, col));
            if (val.isValid() && !val.toString().isEmpty()) {
                exampleRow = r;
                exampleOutput = val.toString();
                break;
            }
        }

        if (exampleRow < 0) { event->accept(); return; }

        // Find the source column (typically the column to the left)
        int srcCol = col - 1;
        if (srcCol < 0) { event->accept(); return; }

        QString srcValue = m_spreadsheet->getCellValue(CellAddress(exampleRow, srcCol)).toString();
        if (srcValue.isEmpty()) { event->accept(); return; }

        // Detect pattern: what transformation was applied from srcValue → exampleOutput?
        // Common patterns: extract part, change case, concatenate

        // Pattern 1: Extract substring
        int startPos = srcValue.indexOf(exampleOutput, 0, Qt::CaseInsensitive);
        bool isExtract = (startPos >= 0 && exampleOutput.length() < srcValue.length());

        // Pattern 2: Case change
        bool isUpper = (exampleOutput == srcValue.toUpper());
        bool isLower = (exampleOutput == srcValue.toLower());
        bool isProper = false;
        {
            QString proper = srcValue.toLower();
            bool capitalize = true;
            for (int i = 0; i < proper.length(); ++i) {
                if (capitalize && proper[i].isLetter()) {
                    proper[i] = proper[i].toUpper();
                    capitalize = false;
                }
                if (proper[i] == ' ' || proper[i] == '-') capitalize = true;
            }
            isProper = (exampleOutput == proper);
        }

        // Pattern 3: Take first N characters or last N characters
        bool isLeft = srcValue.startsWith(exampleOutput, Qt::CaseInsensitive);
        bool isRight = srcValue.endsWith(exampleOutput, Qt::CaseInsensitive);

        // Apply pattern to remaining rows with undo support
        int maxRow = m_spreadsheet->getMaxRow();
        int applied = 0;
        std::vector<CellSnapshot> ffBefore, ffAfter;
        for (int r = row; r <= maxRow; ++r) {
            QString src = m_spreadsheet->getCellValue(CellAddress(r, srcCol)).toString();
            if (src.isEmpty()) continue;

            QString result;
            if (isUpper) result = src.toUpper();
            else if (isLower) result = src.toLower();
            else if (isProper) {
                result = src.toLower();
                bool cap = true;
                for (int i = 0; i < result.length(); ++i) {
                    if (cap && result[i].isLetter()) { result[i] = result[i].toUpper(); cap = false; }
                    if (result[i] == ' ' || result[i] == '-') cap = true;
                }
            } else if (isLeft) result = src.left(exampleOutput.length());
            else if (isRight) result = src.right(exampleOutput.length());
            else if (isExtract && startPos >= 0) {
                result = src.mid(startPos, exampleOutput.length());
            } else {
                continue; // Can't detect pattern
            }

            CellAddress fillAddr(r, col);
            ffBefore.push_back(m_spreadsheet->takeCellSnapshot(fillAddr));
            m_spreadsheet->setCellValue(fillAddr, QVariant(result));
            ffAfter.push_back(m_spreadsheet->takeCellSnapshot(fillAddr));
            applied++;
            if (applied > 100000) break; // Safety limit
        }

        // Push undo command for flash fill
        if (!ffBefore.empty()) {
            m_spreadsheet->getUndoManager().pushCommand(
                std::make_unique<MultiCellEditCommand>(ffBefore, ffAfter, "Flash Fill"));
        }

        if (m_model) m_model->resetModel();
        qDebug() << "[Flash Fill]" << applied << "cells filled";
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+D: Fill Down (copy value from cell above into selection) =====
    if (ctrl && event->key() == Qt::Key_D) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndexList selected = selectionModel()->selectedIndexes();
        if (selected.size() <= 1) {
            // Single cell: copy from cell above
            QModelIndex cur = currentIndex();
            if (cur.isValid() && cur.row() > 0) {
                auto valAbove = m_spreadsheet->getCellValue(CellAddress(logicalRow(cur) - 1, cur.column()));
                auto cellAbove = m_spreadsheet->getCell(CellAddress(logicalRow(cur) - 1, cur.column()));

                std::vector<CellSnapshot> before, after;
                CellAddress addr(logicalRow(cur), cur.column());
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                if (cellAbove->getType() == CellType::Formula) {
                    m_model->setData(cur, cellAbove->getFormula());
                } else {
                    m_model->setData(cur, valAbove);
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        } else {
            // Multi-cell selection: for each column, copy the topmost selected cell value down
            // Find bounds
            int minRow = INT_MAX;
            std::map<int, int> colMinRow; // col -> min row in selection
            for (const auto& idx : selected) {
                if (colMinRow.find(idx.column()) == colMinRow.end() || idx.row() < colMinRow[idx.column()]) {
                    colMinRow[idx.column()] = idx.row();
                }
                minRow = qMin(minRow, idx.row());
            }

            std::vector<CellSnapshot> before, after;
            m_model->setSuppressUndo(true);

            for (const auto& idx : selected) {
                int sourceRow = colMinRow[idx.column()];
                if (idx.row() == sourceRow) continue; // Skip source cells

                CellAddress addr(logicalRow(idx), idx.column());
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                auto srcCell = m_spreadsheet->getCell(CellAddress(sourceRow, idx.column()));
                if (srcCell->getType() == CellType::Formula) {
                    m_model->setData(idx, srcCell->getFormula());
                } else {
                    m_model->setData(idx, srcCell->getValue());
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }

            m_model->setSuppressUndo(false);
            if (!before.empty()) {
                m_spreadsheet->getUndoManager().pushCommand(
                    std::make_unique<MultiCellEditCommand>(before, after, "Fill Down"));
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+Arrow: Jump to edge of data region =====
    // Uses ColumnStore direct queries — O(1) per column check, no full cell scan.
    // Skip during bulk loading — background thread is mutating data concurrently.
    if (ctrl && !shift && !m_bulkLoading &&
        (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        // Convert model row to logical row for virtual mode
        int logicalRow = m_model ? m_model->toLogicalRow(cur.row()) : cur.row();
        int col = cur.column();
        int maxRow = m_spreadsheet->getRowCount() - 1;
        int maxCol = m_spreadsheet->getColumnCount() - 1;

        if (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down) {
            // Direct chunk bitmap queries — O(chunks) not O(cells), no index building
            auto& cs = m_spreadsheet->getColumnStore();
            if (event->key() == Qt::Key_Down) {
                if (logicalRow < maxRow) {
                    bool nextHasData = cs.hasCell(logicalRow + 1, col);
                    if (nextHasData) {
                        // Jump to end of contiguous filled region
                        int emptyRow = cs.nextEmptyRow(col, logicalRow + 1);
                        logicalRow = std::min(emptyRow - 1, maxRow);
                    } else {
                        // Jump to next occupied cell
                        int next = cs.nextOccupiedRow(col, logicalRow + 1);
                        logicalRow = (next >= 0) ? next : maxRow;
                    }
                }
            } else { // Key_Up
                if (logicalRow > 0) {
                    bool prevHasData = cs.hasCell(logicalRow - 1, col);
                    if (prevHasData) {
                        // Jump to start of contiguous filled region
                        int emptyRow = cs.prevEmptyRow(col, logicalRow - 1);
                        logicalRow = std::max(emptyRow + 1, 0);
                    } else {
                        // Jump to previous occupied cell
                        int prev = cs.prevOccupiedRow(col, logicalRow - 1);
                        logicalRow = (prev >= 0) ? prev : 0;
                    }
                }
            }
        } else { // Key_Left or Key_Right
            // Horizontal navigation — only ~20 columns, always fast
            auto occupied = m_spreadsheet->getOccupiedColsInRow(logicalRow);
            if (event->key() == Qt::Key_Right) {
                if (col < maxCol) {
                    auto it = std::upper_bound(occupied.begin(), occupied.end(), col);
                    bool nextHasData = (it != occupied.end() && *it == col + 1);
                    if (nextHasData) {
                        int last = col + 1;
                        ++it;
                        while (it != occupied.end() && *it == last + 1) { last = *it; ++it; }
                        col = last;
                    } else {
                        col = (it != occupied.end()) ? *it : maxCol;
                    }
                }
            } else { // Key_Left
                if (col > 0) {
                    auto prevIt = std::lower_bound(occupied.begin(), occupied.end(), col - 1);
                    bool prevHasData = (prevIt != occupied.end() && *prevIt == col - 1);
                    if (prevHasData) {
                        int first = col - 1;
                        while (prevIt != occupied.begin()) {
                            auto before = std::prev(prevIt);
                            if (*before != first - 1) break;
                            first = *before;
                            prevIt = before;
                        }
                        col = first;
                    } else {
                        auto it = std::lower_bound(occupied.begin(), occupied.end(), col);
                        col = (it != occupied.begin()) ? *std::prev(it) : 0;
                    }
                }
            }
        }

        navigateToLogicalRow(logicalRow, col);
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+Home: Extend selection from current cell to A1 =====
    if (ctrl && shift && event->key() == Qt::Key_Home) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            QModelIndex topLeft = model()->index(0, 0);
            QItemSelection sel(topLeft, cur);
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
            selectionModel()->setCurrentIndex(topLeft, QItemSelectionModel::Current);
            scrollTo(topLeft);
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+End: Extend selection from current cell to last used cell =====
    if (ctrl && shift && !m_bulkLoading && event->key() == Qt::Key_End) {
        QModelIndex cur = currentIndex();
        if (cur.isValid() && m_spreadsheet) {
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow >= 0 && maxCol >= 0) {
                QModelIndex bottomRight = model()->index(maxRow, maxCol);
                QItemSelection sel(cur, bottomRight);
                selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
                selectionModel()->setCurrentIndex(bottomRight, QItemSelectionModel::Current);
                scrollTo(bottomRight);
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+Home: Go to cell A1 =====
    if (ctrl && event->key() == Qt::Key_Home) {
        QModelIndex first = model()->index(0, 0);
        setCurrentIndex(first);
        scrollTo(first);
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+End: Go to last used cell =====
    if (ctrl && !m_bulkLoading && event->key() == Qt::Key_End) {
        if (m_spreadsheet) {
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow >= 0 && maxCol >= 0) {
                QModelIndex last = model()->index(maxRow, maxCol);
                setCurrentIndex(last);
                scrollTo(last);
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+Arrow: Extend selection to data edge =====
    // Direct chunk bitmap queries — same as Ctrl+Arrow but extends selection.
    if (ctrl && shift && !m_bulkLoading &&
        (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        int logicalRow = m_model ? m_model->toLogicalRow(cur.row()) : cur.row();
        int col = cur.column();
        int maxRowIdx = m_spreadsheet->getRowCount() - 1;
        int maxColIdx = m_spreadsheet->getColumnCount() - 1;
        auto& cs = m_spreadsheet->getColumnStore();

        if (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down) {
            if (event->key() == Qt::Key_Down) {
                if (logicalRow < maxRowIdx) {
                    bool nextHasData = cs.hasCell(logicalRow + 1, col);
                    if (nextHasData) {
                        int emptyRow = cs.nextEmptyRow(col, logicalRow + 1);
                        logicalRow = std::min(emptyRow - 1, maxRowIdx);
                    } else {
                        int next = cs.nextOccupiedRow(col, logicalRow + 1);
                        logicalRow = (next >= 0) ? next : maxRowIdx;
                    }
                }
            } else {
                if (logicalRow > 0) {
                    bool prevHasData = cs.hasCell(logicalRow - 1, col);
                    if (prevHasData) {
                        int emptyRow = cs.prevEmptyRow(col, logicalRow - 1);
                        logicalRow = std::max(emptyRow + 1, 0);
                    } else {
                        int prev = cs.prevOccupiedRow(col, logicalRow - 1);
                        logicalRow = (prev >= 0) ? prev : 0;
                    }
                }
            }
        } else {
            auto occupied = m_spreadsheet->getOccupiedColsInRow(logicalRow);
            if (event->key() == Qt::Key_Right) {
                if (col < maxColIdx) {
                    auto it = std::upper_bound(occupied.begin(), occupied.end(), col);
                    bool nextHasData = (it != occupied.end() && *it == col + 1);
                    if (nextHasData) {
                        int last = col + 1; ++it;
                        while (it != occupied.end() && *it == last + 1) { last = *it; ++it; }
                        col = last;
                    } else {
                        col = (it != occupied.end()) ? *it : maxColIdx;
                    }
                }
            } else {
                if (col > 0) {
                    auto prevIt = std::lower_bound(occupied.begin(), occupied.end(), col - 1);
                    bool prevHasData = (prevIt != occupied.end() && *prevIt == col - 1);
                    if (prevHasData) {
                        int first = col - 1;
                        while (prevIt != occupied.begin()) {
                            auto before = std::prev(prevIt);
                            if (*before != first - 1) break;
                            first = *before; prevIt = before;
                        }
                        col = first;
                    } else {
                        auto it = std::lower_bound(occupied.begin(), occupied.end(), col);
                        col = (it != occupied.begin()) ? *std::prev(it) : 0;
                    }
                }
            }
        }

        // Navigate viewport then extend selection
        navigateToLogicalRow(logicalRow, col);
        int modelRow = m_model ? m_model->toModelRow(logicalRow) : logicalRow;
        QModelIndex target = model()->index(modelRow, col);
        if (target.isValid()) {
            QItemSelection sel(cur, target);
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
            setCurrentIndex(target);
        }
        event->accept();
        return;
    }

    // ===== Ctrl+R: Fill Right (copy from left cell) =====
    if (ctrl && event->key() == Qt::Key_R) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndexList selected = selectionModel()->selectedIndexes();
        if (selected.size() <= 1) {
            QModelIndex cur = currentIndex();
            if (cur.isValid() && cur.column() > 0) {
                auto valLeft = m_spreadsheet->getCellValue(CellAddress(logicalRow(cur), cur.column() - 1));
                auto cellLeft = m_spreadsheet->getCell(CellAddress(logicalRow(cur), cur.column() - 1));

                CellAddress addr(logicalRow(cur), cur.column());
                std::vector<CellSnapshot> before, after;
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                if (cellLeft->getType() == CellType::Formula) {
                    m_model->setData(cur, cellLeft->getFormula());
                } else {
                    m_model->setData(cur, valLeft);
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+; : Insert current time =====
    if (ctrl && shift && event->key() == Qt::Key_Semicolon) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            QString timeStr = QTime::currentTime().toString("hh:mm:ss");
            m_model->setData(cur, timeStr);
        }
        event->accept();
        return;
    }

    // ===== Ctrl+; : Insert current date =====
    if (ctrl && event->key() == Qt::Key_Semicolon) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            QString today = QDate::currentDate().toString("MM/dd/yyyy");
            m_model->setData(cur, today);
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+~ : General number format =====
    // (Must be checked before Ctrl+` since ~ is Shift+` on most keyboards)
    if (ctrl && shift && (event->key() == Qt::Key_AsciiTilde || event->key() == Qt::Key_QuoteLeft)) {
        applyNumberFormat("General");
        event->accept();
        return;
    }

    // ===== Ctrl+` (backtick): Toggle formula view =====
    if (ctrl && !shift && event->key() == Qt::Key_QuoteLeft) {
        toggleFormulaView();
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+! : Number format (2 decimal places, thousands separator) =====
    if (ctrl && shift && event->key() == Qt::Key_Exclam) {
        applyNumberFormat("Number");
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+@ : Time format =====
    if (ctrl && shift && event->key() == Qt::Key_At) {
        applyNumberFormat("hh:mm:ss");
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+# : Date format =====
    if (ctrl && shift && event->key() == Qt::Key_NumberSign) {
        applyNumberFormat("mm/dd/yyyy");
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+$ : Currency format =====
    if (ctrl && shift && event->key() == Qt::Key_Dollar) {
        applyNumberFormat("Currency");
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+% : Percentage format =====
    if (ctrl && shift && event->key() == Qt::Key_Percent) {
        applyNumberFormat("Percentage");
        event->accept();
        return;
    }

    // ===== F5: Go To dialog (same as Ctrl+G) =====
    if (event->key() == Qt::Key_F5) {
        emit goToRequested();
        event->accept();
        return;
    }

    // ===== Ctrl+Plus: Insert cells dialog =====
    if (ctrl && (event->key() == Qt::Key_Plus || (shift && event->key() == Qt::Key_Equal))) {
        QDialog dlg(this);
        dlg.setWindowTitle("Insert");
        auto* layout = new QVBoxLayout(&dlg);
        auto* rbRight = new QRadioButton("Shift cells right", &dlg);
        auto* rbDown = new QRadioButton("Shift cells down", &dlg);
        auto* rbRow = new QRadioButton("Entire row", &dlg);
        auto* rbCol = new QRadioButton("Entire column", &dlg);
        rbDown->setChecked(true);
        layout->addWidget(rbRight);
        layout->addWidget(rbDown);
        layout->addWidget(rbRow);
        layout->addWidget(rbCol);
        auto* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
        layout->addWidget(buttons);
        connect(buttons, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
        connect(buttons, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);
        if (dlg.exec() == QDialog::Accepted) {
            if (rbRight->isChecked()) insertCellsShiftRight();
            else if (rbDown->isChecked()) insertCellsShiftDown();
            else if (rbRow->isChecked()) insertEntireRow();
            else if (rbCol->isChecked()) insertEntireColumn();
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Minus: Delete cells dialog =====
    if (ctrl && event->key() == Qt::Key_Minus) {
        QDialog dlg(this);
        dlg.setWindowTitle("Delete");
        auto* layout = new QVBoxLayout(&dlg);
        auto* rbLeft = new QRadioButton("Shift cells left", &dlg);
        auto* rbUp = new QRadioButton("Shift cells up", &dlg);
        auto* rbRow = new QRadioButton("Entire row", &dlg);
        auto* rbCol = new QRadioButton("Entire column", &dlg);
        rbUp->setChecked(true);
        layout->addWidget(rbLeft);
        layout->addWidget(rbUp);
        layout->addWidget(rbRow);
        layout->addWidget(rbCol);
        auto* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
        layout->addWidget(buttons);
        connect(buttons, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
        connect(buttons, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);
        if (dlg.exec() == QDialog::Accepted) {
            if (rbLeft->isChecked()) deleteCellsShiftLeft();
            else if (rbUp->isChecked()) deleteCellsShiftUp();
            else if (rbRow->isChecked()) deleteEntireRow();
            else if (rbCol->isChecked()) deleteEntireColumn();
        }
        event->accept();
        return;
    }

    // ===== F4: Absolute reference toggle (only while editing a formula) =====
    if (event->key() == Qt::Key_F4 && state() == QAbstractItemView::EditingState) {
        QWidget* editor = indexWidget(currentIndex());
        if (!editor) editor = viewport()->findChild<QLineEdit*>();
        auto* lineEdit = qobject_cast<QLineEdit*>(editor);
        if (lineEdit) {
            QString text = lineEdit->text();
            if (text.startsWith("=")) {
                int cursor = lineEdit->cursorPosition();
                // Find cell reference token around cursor
                static QRegularExpression refRe("\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}");
                QRegularExpressionMatchIterator it = refRe.globalMatch(text);
                while (it.hasNext()) {
                    QRegularExpressionMatch match = it.next();
                    int start = match.capturedStart();
                    int end = match.capturedEnd();
                    if (cursor >= start && cursor <= end) {
                        QString token = match.captured();
                        // Parse current $ state
                        static QRegularExpression parseRe("^(\\$?)([A-Za-z]{1,3})(\\$?)([0-9]{1,7})$");
                        auto parsed = parseRe.match(token);
                        if (parsed.hasMatch()) {
                            bool colAbs = !parsed.captured(1).isEmpty();
                            QString col = parsed.captured(2);
                            bool rowAbs = !parsed.captured(3).isEmpty();
                            QString row = parsed.captured(4);
                            // Cycle: A1 -> $A$1 -> A$1 -> $A1 -> A1
                            QString newToken;
                            if (!colAbs && !rowAbs) {
                                newToken = "$" + col + "$" + row;       // A1 -> $A$1
                            } else if (colAbs && rowAbs) {
                                newToken = col + "$" + row;             // $A$1 -> A$1
                            } else if (!colAbs && rowAbs) {
                                newToken = "$" + col + row;             // A$1 -> $A1
                            } else {
                                newToken = col + row;                   // $A1 -> A1
                            }
                            QString newText = text.left(start) + newToken + text.mid(end);
                            lineEdit->setText(newText);
                            lineEdit->setCursorPosition(start + newToken.length());
                        }
                        break;
                    }
                }
            }
        }
        event->accept();
        return;
    }

    // ===== Shift+Space: Select entire row =====
    if (shift && !ctrl && event->key() == Qt::Key_Space) {
        if (state() != QAbstractItemView::EditingState) {
            QModelIndex cur = currentIndex();
            if (cur.isValid()) {
                selectRow(cur.row());
            }
            event->accept();
            return;
        }
    }

    // ===== Ctrl+Space: Select entire column =====
    if (ctrl && !shift && event->key() == Qt::Key_Space) {
        if (state() != QAbstractItemView::EditingState) {
            QModelIndex cur = currentIndex();
            if (cur.isValid()) {
                selectColumn(cur.column());
            }
            event->accept();
            return;
        }
    }

    // ===== Home key: go to beginning of row (column A) =====
    if (event->key() == Qt::Key_Home && !ctrl && !shift) {
        if (state() != QAbstractItemView::EditingState) {
            QModelIndex idx = model()->index(currentIndex().row(), 0);
            setCurrentIndex(idx);
            scrollTo(idx);
            event->accept();
            return;
        }
    }

    // ===== Shift+Home: select from current cell to column A =====
    if (event->key() == Qt::Key_Home && shift && !ctrl) {
        if (state() != QAbstractItemView::EditingState) {
            QModelIndex cur = currentIndex();
            QModelIndex first = model()->index(cur.row(), 0);
            QItemSelection sel(first, cur);
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
            selectionModel()->setCurrentIndex(first, QItemSelectionModel::NoUpdate);
            scrollTo(first);
            event->accept();
            return;
        }
    }

    // ===== Alt+= : AutoSum (insert =SUM() with auto-detected range above) =====
    if (event->key() == Qt::Key_Equal && (event->modifiers() & Qt::AltModifier)) {
        if (m_spreadsheet) {
            QModelIndex cur = currentIndex();
            if (cur.isValid()) {
                int row = logicalRow(cur);
                int col = cur.column();
                // Walk up from row-1 while cells have numeric data
                int startRow = row - 1;
                while (startRow >= 0) {
                    auto val = m_spreadsheet->getCellValue(CellAddress(startRow, col));
                    if (!val.isValid() || val.toString().isEmpty()) break;
                    // Check if it's numeric
                    bool ok = false;
                    val.toDouble(&ok);
                    if (!ok) break;
                    startRow--;
                }
                startRow++; // first numeric row
                if (startRow < row) {
                    // Build cell references using CellAddress::toString()
                    QString startRef = CellAddress(startRow, col).toString();
                    QString endRef = CellAddress(row - 1, col).toString();
                    QString formula = QString("=SUM(%1:%2)").arg(startRef, endRef);
                    m_model->setData(cur, formula);
                }
            }
        }
        event->accept();
        return;
    }

    // Block typing-to-edit if cell is protected
    if (state() != QAbstractItemView::EditingState && !ctrl && !event->text().isEmpty()) {
        QChar ch = event->text().at(0);
        if (ch.isPrint() && isCellProtected()) {
            QMessageBox::warning(this, "Protected Sheet",
                "The cell or chart you're trying to change is on a protected sheet.\n"
                "To make changes, unprotect the sheet (Data menu > Protect Sheet).");
            event->accept();
            return;
        }
    }

    // Detect transition from non-editing to editing (typing starts the editor)
    bool wasEditing = (state() == QAbstractItemView::EditingState);
    QTableView::keyPressEvent(event);
    if (!wasEditing && state() == QAbstractItemView::EditingState) {
        // Check if typing "=" starts formula (Enter mode), otherwise Edit mode
        QString typed = event->text();
        if (typed == "=") {
            emit editModeChanged("Enter");
        } else {
            emit editModeChanged("Edit");
        }
    }
}

void SpreadsheetView::setFormulaEditMode(bool active) {
    m_formulaEditMode = active;
    if (m_delegate) {
        m_delegate->setFormulaEditMode(active);
    }
    if (active) {
        m_formulaEditCell = currentIndex();
        emit editModeChanged("Enter");
    } else {
        m_formulaRangeDragging = false;
    }
}

void SpreadsheetView::closeEditor(QWidget* editor, QAbstractItemDelegate::EndEditHint hint) {
    // During formula edit mode, block editor closing (e.g. from FocusOut
    // when user clicks on grid to select a cell reference)
    if (m_formulaEditMode) {
        return;
    }
    QTableView::closeEditor(editor, hint);
    emit editModeChanged("Ready");
}

void SpreadsheetView::commitData(QWidget* editor) {
    // During formula edit mode, block data commit (user is still building the formula)
    if (m_formulaEditMode) {
        return;
    }
    QTableView::commitData(editor);
}

void SpreadsheetView::setChartRangeHighlight(const CellRange& range,
                                              const QVector<QPair<int, QColor>>& seriesColumns,
                                              const QColor& categoryColor) {
    m_chartHighlight.fullRange = range;
    m_chartHighlight.seriesColumns = seriesColumns;
    m_chartHighlight.categoryColor = categoryColor;
    m_chartHighlightActive = true;
    viewport()->update();
}

void SpreadsheetView::clearChartRangeHighlight() {
    m_chartHighlightActive = false;
    viewport()->update();
}

void SpreadsheetView::insertCellReference(const QString& ref) {
    // Insert into the active cell editor if editing inline
    if (state() == QAbstractItemView::EditingState) {
        QLineEdit* lineEdit = viewport()->findChild<QLineEdit*>();
        if (lineEdit) {
            lineEdit->insert(ref);
            return;
        }
    }
    // Otherwise signal for the formula bar to handle
    emit cellReferenceInserted(ref);
}

void SpreadsheetView::mouseDoubleClickEvent(QMouseEvent* event) {
    m_lastDoubleClickPos = event->pos();
    m_editTriggeredByDoubleClick = true;
    QTableView::mouseDoubleClickEvent(event);
    if (state() == QAbstractItemView::EditingState) {
        emit editModeChanged("Edit");
    }
}

void SpreadsheetView::mousePressEvent(QMouseEvent* event) {
    // Outline gutter click: check if clicking on the outline gutter area
    if (event->button() == Qt::LeftButton && m_spreadsheet && outlineGutterTotalWidth() > 0) {
        handleOutlineGutterClick(event->pos());
        // If click was in gutter area, it was already handled
        QPoint vpPos = event->pos();
        // The gutter is painted on the widget (not viewport), so we check against
        // the vertical header area. The gutter is drawn to the left of the row header.
        // But since we paint on the viewport, clicks on the gutter are actually in
        // the negative x area relative to viewport, which maps to the header area.
        // We handle this in handleOutlineGutterClick which uses mapToParent.
    }

    // Filter button click: check if clicking on a filter dropdown button
    if (m_filterActive && event->button() == Qt::LeftButton && m_model) {
        QModelIndex clickedIdx = indexAt(event->pos());
        if (clickedIdx.isValid() && clickedIdx.row() == m_filterHeaderRow) {
            int col = clickedIdx.column();
            int startCol = m_filterRange.getStart().col;
            int endCol = m_filterRange.getEnd().col;
            if (col >= startCol && col <= endCol) {
                QRect cellRect = visualRect(clickedIdx);
                int btnSize = 16;
                int margin = 2;
                QRect btnRect(cellRect.right() - btnSize - margin,
                              cellRect.top() + (cellRect.height() - btnSize) / 2,
                              btnSize, btnSize);
                // Expand hit area slightly for easier clicking
                if (btnRect.adjusted(-3, -3, 3, 3).contains(event->pos())) {
                    showFilterDropdown(col);
                    event->accept();
                    return;
                }
            }
        }
    }

    // Pivot report filter: clicking col 1 of a filter row opens value picker
    if (m_spreadsheet && m_spreadsheet->isPivotSheet() && event->button() == Qt::LeftButton) {
        const PivotConfig* pc = m_spreadsheet->getPivotConfig();
        if (pc && !pc->filterFields.empty()) {
            QModelIndex clickedIdx = indexAt(event->pos());
            if (clickedIdx.isValid() && clickedIdx.column() == 1
                && clickedIdx.row() < static_cast<int>(pc->filterFields.size())) {
                showPivotFilterDropdown(clickedIdx.row());
                event->accept();
                return;
            }
        }
    }

    // Formula edit mode: clicking a cell inserts its reference, drag selects range
    // The editor stays open — user is still building the formula (like Excel)
    if (m_formulaEditMode && event->button() == Qt::LeftButton) {
        QModelIndex clickedIdx = indexAt(event->pos());
        if (clickedIdx.isValid() && clickedIdx != m_formulaEditCell) {
            CellAddress addr(clickedIdx.row(), clickedIdx.column());
            QString ref = addr.toString();

            // Try to insert into the in-cell editor (delegate editors aren't
            // accessible via indexWidget, so search viewport children)
            bool insertedInCell = false;
            if (state() == QAbstractItemView::EditingState) {
                QLineEdit* lineEdit = viewport()->findChild<QLineEdit*>();
                if (lineEdit) {
                    m_formulaRefInsertPos = lineEdit->cursorPosition();
                    lineEdit->insert(ref);
                    m_formulaRefInsertLen = ref.length();
                    insertedInCell = true;
                    // Keep focus on the editor so the formula stays active
                    lineEdit->setFocus();
                }
            }

            // Otherwise insert into formula bar
            if (!insertedInCell) {
                m_formulaRefInsertPos = -1;
                m_formulaRefInsertLen = ref.length();
                emit cellReferenceInserted(ref);
            }

            m_formulaRangeDragging = true;
            m_formulaRangeStart = clickedIdx;
            m_formulaRangeEnd = clickedIdx;
            event->accept();
            return;
        }
        // Clicking on the formula cell itself — let it through to position cursor
        // but don't close the editor
        if (clickedIdx == m_formulaEditCell) {
            event->accept();
            return;
        }
    }

    // Format painter: apply copied style
    if (m_formatPainterActive && event->button() == Qt::LeftButton) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && m_spreadsheet) {
            std::vector<CellSnapshot> before, after;
            CellAddress addr(logicalRow(idx), idx.column());
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            cell->setStyle(m_copiedStyle);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            m_spreadsheet->getUndoManager().execute(
                std::make_unique<StyleChangeCommand>(before, after), m_spreadsheet.get());

            if (m_model) {
                emit m_model->dataChanged(idx, idx);
            }
        }
        m_formatPainterActive = false;
        viewport()->setCursor(Qt::ArrowCursor);
        return;
    }

    // Ctrl+Click on hyperlink: open URL
    if (event->button() == Qt::LeftButton && (event->modifiers() & Qt::ControlModifier) && m_spreadsheet) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid()) {
            auto cell = m_spreadsheet->getCellIfExists(logicalRow(idx), idx.column());
            if (cell && cell->hasHyperlink()) {
                openHyperlink(logicalRow(idx), idx.column());
                event->accept();
                return;
            }
        }
    }

    // Validation dropdown button click handling
    if (event->button() == Qt::LeftButton && m_spreadsheet) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && idx == currentIndex()) {
            auto* rule = m_spreadsheet->getValidationAt(logicalRow(idx), idx.column());
            if (rule && rule->type == Spreadsheet::DataValidationRule::List) {
                QRect cellRect = visualRect(idx);
                int btnWidth = 18;
                QRect btnRect(cellRect.right() - btnWidth, cellRect.top(), btnWidth, cellRect.height());
                if (btnRect.contains(event->pos())) {
                    showValidationDropdown(idx, rule);
                    event->accept();
                    return;
                }
            }
        }
    }

    // Checkbox & Picklist click handling
    if (event->button() == Qt::LeftButton) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && m_spreadsheet) {
            auto cell = m_spreadsheet->getCellIfExists(logicalRow(idx), idx.column());
            if (cell) {
                const QString& fmt = cell->getStyle().numberFormat;
                if (fmt == "Checkbox") {
                    QRect cellRect = visualRect(idx);
                    int boxSize = 14;
                    int cx = cellRect.left() + (cellRect.width() - boxSize) / 2;
                    int cy = cellRect.top() + (cellRect.height() - boxSize) / 2;
                    QRect hitArea(cx - 4, cy - 4, boxSize + 8, boxSize + 8);
                    if (hitArea.contains(event->pos())) {
                        toggleCheckbox(logicalRow(idx), idx.column());
                        setCurrentIndex(idx);
                        event->accept();
                        return;
                    }
                } else if (fmt == "Picklist" && !m_picklistPopupOpen) {
                    // Open picklist popup on any click within the cell
                    m_picklistPopupOpen = true;
                    setCurrentIndex(idx);
                    QTimer::singleShot(0, this, [this, idx]() {
                        showPicklistPopup(idx);
                    });
                    event->accept();
                    return;
                }
            }
        }
    }

    // Fill handle drag start
    if (event->button() == Qt::LeftButton && isOverFillHandle(event->pos())) {
        m_fillDragging = true;
        m_fillDragStart = currentIndex();
        m_fillDragCurrent = event->pos();
        event->accept();
        return;
    }

    // Store anchor for range selection (Excel: first cell = active cell)
    if (event->button() == Qt::LeftButton) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && !(event->modifiers() & Qt::ControlModifier)) {
            m_selectionAnchor = idx;
            m_isDragSelecting = true;
        }
    }

    QTableView::mousePressEvent(event);

    // Immediately set anchor as current index so it stays white during drag
    if (m_isDragSelecting && m_selectionAnchor.isValid()) {
        selectionModel()->setCurrentIndex(m_selectionAnchor, QItemSelectionModel::NoUpdate);
    }
}

void SpreadsheetView::mouseMoveEvent(QMouseEvent* event) {
    // Formula range drag: extend selection as user drags
    if (m_formulaRangeDragging && m_formulaEditMode) {
        QModelIndex hoverIdx = indexAt(event->pos());
        if (hoverIdx.isValid() && hoverIdx != m_formulaRangeEnd) {
            m_formulaRangeEnd = hoverIdx;

            // Build range string
            CellAddress startAddr(m_formulaRangeStart.row(), m_formulaRangeStart.column());
            CellAddress endAddr(m_formulaRangeEnd.row(), m_formulaRangeEnd.column());
            QString newRef;
            if (m_formulaRangeStart == m_formulaRangeEnd) {
                newRef = startAddr.toString();
            } else {
                newRef = startAddr.toString() + ":" + endAddr.toString();
            }

            // Replace previously inserted reference
            bool replacedInCell = false;
            if (state() == QAbstractItemView::EditingState && m_formulaRefInsertPos >= 0) {
                QLineEdit* lineEdit = viewport()->findChild<QLineEdit*>();
                if (lineEdit) {
                    lineEdit->setSelection(m_formulaRefInsertPos, m_formulaRefInsertLen);
                    lineEdit->insert(newRef);
                    m_formulaRefInsertLen = newRef.length();
                    lineEdit->setFocus();
                    replacedInCell = true;
                }
            }
            if (!replacedInCell) {
                emit cellReferenceReplaced(newRef);
                m_formulaRefInsertLen = newRef.length();
            }
            viewport()->update();
        }
        event->accept();
        return;
    }

    if (m_fillDragging) {
        m_fillDragCurrent = event->pos();
        viewport()->update();
        return;
    }

    // Change cursor when hovering over fill handle
    if (isOverFillHandle(event->pos())) {
        viewport()->setCursor(Qt::CrossCursor);
    } else if (!m_formatPainterActive) {
        viewport()->setCursor(Qt::ArrowCursor);
    }

    QTableView::mouseMoveEvent(event);

    // Keep anchor as active cell during drag selection (Excel: anchor stays white)
    if (m_isDragSelecting && m_selectionAnchor.isValid()) {
        selectionModel()->setCurrentIndex(m_selectionAnchor, QItemSelectionModel::NoUpdate);
    }
}

void SpreadsheetView::mouseReleaseEvent(QMouseEvent* event) {
    hideResizeTooltip();

    if (m_formulaRangeDragging) {
        m_formulaRangeDragging = false;
        viewport()->update();
        return;
    }

    if (m_fillDragging) {
        m_fillDragging = false;
        performFillSeries();
        viewport()->update();
        return;
    }

    QTableView::mouseReleaseEvent(event);

    // Excel behavior: after drag selection, the anchor cell (where mouse-down started)
    // is the active cell, not the cell where mouse was released
    if (m_isDragSelecting && m_selectionAnchor.isValid()) {
        m_isDragSelecting = false;
        QModelIndex current = currentIndex();
        if (current != m_selectionAnchor) {
            // Set anchor as current without changing the selection
            selectionModel()->setCurrentIndex(m_selectionAnchor, QItemSelectionModel::NoUpdate);
        }
    }
}

void SpreadsheetView::paintEvent(QPaintEvent* event) {
    QTableView::paintEvent(event);

    // --- Selection range outer border (Excel-style) ---
    // Draw a separate rounded rectangle border around EACH selection range.
    {
        QItemSelection sel = selectionModel()->selection();
        if (!sel.isEmpty()) {
            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, true);
            QPen borderPen(ThemeManager::instance().currentTheme().focusBorderColor, 2, Qt::SolidLine);
            painter.setPen(borderPen);
            painter.setBrush(Qt::NoBrush);

            for (const auto& range : sel) {
                // Skip single-cell ranges (current cell already has focus border)
                if (range.width() == 1 && range.height() == 1 && sel.size() == 1) continue;

                QModelIndex topLeft = model()->index(range.top(), range.left());
                QModelIndex bottomRight = model()->index(range.bottom(), range.right());
                QRect tlRect = visualRect(topLeft);
                QRect brRect = visualRect(bottomRight);

                if (tlRect.isValid() && brRect.isValid()) {
                    QRect rangePixelRect = tlRect.united(brRect);
                    painter.drawRoundedRect(rangePixelRect, 3, 3);
                }
            }
        }
    }

    // --- Marching ants animation around clipboard range (Ctrl+C) ---
    if (m_hasClipboardRange) {
        // m_clipboardRange stores: x()=minCol, y()=minRow, width()=maxCol, height()=maxRow
        int minCol = m_clipboardRange.x();
        int minRow = m_clipboardRange.y();
        int maxCol = m_clipboardRange.width();
        int maxRow = m_clipboardRange.height();

        QModelIndex topLeft = model()->index(minRow, minCol);
        QModelIndex bottomRight = model()->index(maxRow, maxCol);
        QRect tlRect = visualRect(topLeft);
        QRect brRect = visualRect(bottomRight);

        if (tlRect.isValid() && brRect.isValid()) {
            QRect clipPixelRect = tlRect.united(brRect);
            QPainter antsPainter(viewport());
            antsPainter.setRenderHint(QPainter::Antialiasing, false);

            // White background line (so dashes are visible on any background)
            QPen bgPen(Qt::white, 2, Qt::SolidLine);
            antsPainter.setPen(bgPen);
            antsPainter.setBrush(Qt::NoBrush);
            antsPainter.drawRect(clipPixelRect);

            // Animated dashed foreground line
            QPen dashPen(QColor("#1a73e8"), 2, Qt::CustomDashLine);
            QVector<qreal> pattern = {6, 4};
            dashPen.setDashPattern(pattern);
            dashPen.setDashOffset(m_marchingAntsOffset);
            antsPainter.setPen(dashPen);
            antsPainter.drawRect(clipPixelRect);
        }
    }

    // --- Formula reference highlighting (Excel-style colored borders) ---
    if (m_formulaEditMode && m_spreadsheet && state() == QAbstractItemView::EditingState) {
        // Parse cell references from the formula being edited
        QWidget* editor = indexWidget(currentIndex());
        if (!editor) editor = viewport()->findChild<QLineEdit*>();
        if (editor) {
            QString text = qobject_cast<QLineEdit*>(editor) ?
                qobject_cast<QLineEdit*>(editor)->text() : QString();
            if (text.startsWith("=")) {
                // Colors for referenced ranges (Excel uses these colors)
                static const QColor refColors[] = {
                    QColor(0, 0, 255),      // Blue
                    QColor(255, 0, 0),      // Red
                    QColor(128, 0, 128),    // Purple
                    QColor(0, 128, 0),      // Green
                    QColor(255, 128, 0),    // Orange
                    QColor(0, 128, 128),    // Teal
                    QColor(128, 0, 0),      // Maroon
                    QColor(0, 0, 128),      // Navy
                };
                int colorIdx = 0;

                // Find all cell references in the formula
                static QRegularExpression refRe(
                    "\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?",
                    QRegularExpression::CaseInsensitiveOption);
                auto it = refRe.globalMatch(text);
                QPainter refPainter(viewport());
                refPainter.setRenderHint(QPainter::Antialiasing, false);

                while (it.hasNext()) {
                    auto match = it.next();
                    QString ref = match.captured();
                    QColor color = refColors[colorIdx % 8];
                    colorIdx++;

                    // Parse the reference
                    CellRange range(ref);
                    if (!range.isValid()) continue;

                    int startRow = range.getStart().row;
                    int startCol = range.getStart().col;
                    int endRow = range.getEnd().row;
                    int endCol = range.getEnd().col;

                    // Convert to model rows
                    int modelStartRow = m_model ? m_model->toModelRow(startRow) : startRow;
                    int modelEndRow = m_model ? m_model->toModelRow(endRow) : endRow;
                    if (modelStartRow < 0 || modelEndRow < 0) continue;

                    QModelIndex tl = model()->index(modelStartRow, startCol);
                    QModelIndex br = model()->index(modelEndRow, endCol);
                    QRect tlRect = visualRect(tl);
                    QRect brRect = visualRect(br);
                    if (!tlRect.isValid() || !brRect.isValid()) continue;

                    QRect refRect = tlRect.united(brRect);

                    // Draw colored border + light fill
                    QPen pen(color, 2);
                    refPainter.setPen(pen);
                    refPainter.setBrush(QColor(color.red(), color.green(), color.blue(), 30));
                    refPainter.drawRect(refRect);
                }
            }
        }
    }

    // Draw fill handle on current selection (Excel-style solid square at bottom-right corner)
    QModelIndex current = currentIndex();
    if (current.isValid() && !m_fillDragging) {
        QRect selRect = getSelectionBoundingRect();
        if (!selRect.isNull()) {
            int handleSize = 8;
            m_fillHandleRect = QRect(
                selRect.right() - handleSize / 2,
                selRect.bottom() - handleSize / 2,
                handleSize, handleSize);

            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, false);
            // Solid filled square with the focus border color (Excel green)
            painter.fillRect(m_fillHandleRect, ThemeManager::instance().currentTheme().focusBorderColor);
            // Thin white outline for visibility against any background
            painter.setPen(QPen(Qt::white, 1));
            painter.drawRect(m_fillHandleRect);
        }
    }

    // Draw filter dropdown buttons on header row cells (Excel-style)
    if (m_filterActive && m_model) {
        QPainter filterPainter(viewport());
        filterPainter.setRenderHint(QPainter::Antialiasing, true);
        int startCol = m_filterRange.getStart().col;
        int endCol = m_filterRange.getEnd().col;
        for (int c = startCol; c <= endCol; ++c) {
            QModelIndex headerIdx = m_model->index(m_filterHeaderRow, c);
            QRect cellRect = visualRect(headerIdx);
            if (cellRect.isNull() || !viewport()->rect().intersects(cellRect)) continue;

            // Draw a dropdown button in the right side of the cell
            int btnSize = 16;
            int margin = 2;
            QRect btnRect(cellRect.right() - btnSize - margin,
                          cellRect.top() + (cellRect.height() - btnSize) / 2,
                          btnSize, btnSize);

            bool hasActiveFilter = m_columnFilters.count(c) > 0;

            // Button background
            filterPainter.setPen(QPen(QColor("#C0C0C0"), 0.5));
            filterPainter.setBrush(hasActiveFilter ? QColor("#D6E4F0") : QColor("#F0F0F0"));
            filterPainter.drawRoundedRect(btnRect, 2, 2);

            // Draw small dropdown arrow
            QColor arrowColor = hasActiveFilter ? ThemeManager::instance().currentTheme().accentDarker : QColor("#555555");
            filterPainter.setPen(Qt::NoPen);
            filterPainter.setBrush(arrowColor);
            int ax = btnRect.center().x();
            int ay = btnRect.center().y();
            QPolygonF arrow;
            arrow << QPointF(ax - 3, ay - 1) << QPointF(ax + 3, ay - 1) << QPointF(ax, ay + 2.5);
            filterPainter.drawPolygon(arrow);
        }
    }

    // Draw chart data range highlights when a chart is selected
    if (m_chartHighlightActive && m_model) {
        QPainter chartPainter(viewport());
        chartPainter.setRenderHint(QPainter::Antialiasing, false);

        int startRow = m_chartHighlight.fullRange.getStart().row;
        int endRow = m_chartHighlight.fullRange.getEnd().row;
        int startCol = m_chartHighlight.fullRange.getStart().col;
        int endCol = m_chartHighlight.fullRange.getEnd().col;

        // Draw each column with its color
        for (int col = startCol; col <= endCol; ++col) {
            QColor color = m_chartHighlight.categoryColor; // default for category col
            for (const auto& sc : m_chartHighlight.seriesColumns) {
                if (sc.first == col) {
                    color = sc.second;
                    break;
                }
            }

            // Compute bounding rect for the column's cells in this range
            QRect colBounds;
            for (int row = startRow; row <= endRow; ++row) {
                // Convert logical row to model row for virtual mode
                int modelRow = m_model->toModelRow(row);
                if (modelRow < 0 || modelRow >= m_model->rowCount()) continue;
                QModelIndex idx = m_model->index(modelRow, col);
                QRect cellRect = visualRect(idx);
                if (cellRect.isNull() || !viewport()->rect().intersects(cellRect)) continue;

                // Fill individual cells with semi-transparent color
                QColor fillColor(color.red(), color.green(), color.blue(), 35);
                chartPainter.fillRect(cellRect, fillColor);

                if (colBounds.isNull()) colBounds = cellRect;
                else colBounds = colBounds.united(cellRect);
            }

            // Draw outer border around the column range
            if (!colBounds.isNull()) {
                chartPainter.setPen(QPen(color, 2));
                chartPainter.setBrush(Qt::NoBrush);
                chartPainter.drawRect(colBounds.adjusted(0, 0, -1, -1));
            }
        }
    }

    // Draw formula range selection preview (blue highlight during drag)
    if (m_formulaRangeDragging && m_formulaEditMode && m_formulaRangeStart.isValid() && m_model) {
        int r1 = qMin(m_formulaRangeStart.row(), m_formulaRangeEnd.row());
        int r2 = qMax(m_formulaRangeStart.row(), m_formulaRangeEnd.row());
        int c1 = qMin(m_formulaRangeStart.column(), m_formulaRangeEnd.column());
        int c2 = qMax(m_formulaRangeStart.column(), m_formulaRangeEnd.column());

        QRect rangeBounds;
        QPainter fPainter(viewport());
        fPainter.setRenderHint(QPainter::Antialiasing, false);
        QColor rangeColor(68, 114, 196); // Excel-like blue

        for (int row = r1; row <= r2; ++row) {
            for (int col = c1; col <= c2; ++col) {
                QRect cellRect = visualRect(m_model->index(row, col));
                if (cellRect.isNull()) continue;
                fPainter.fillRect(cellRect, QColor(rangeColor.red(), rangeColor.green(), rangeColor.blue(), 40));
                if (rangeBounds.isNull()) rangeBounds = cellRect;
                else rangeBounds = rangeBounds.united(cellRect);
            }
        }
        if (!rangeBounds.isNull()) {
            fPainter.setPen(QPen(rangeColor, 2));
            fPainter.setBrush(Qt::NoBrush);
            fPainter.drawRect(rangeBounds.adjusted(0, 0, -1, -1));
        }
    }

    // Draw fill drag preview
    if (m_fillDragging && m_fillDragStart.isValid()) {
        QModelIndex dragTarget = indexAt(m_fillDragCurrent);
        if (dragTarget.isValid()) {
            QRect selRect = getSelectionBoundingRect();
            QRect targetRect = visualRect(dragTarget);

            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, false);

            // Determine fill direction and draw dashed border
            QRect fillRect;
            if (dragTarget.row() > m_fillDragStart.row()) {
                fillRect = QRect(selRect.left(), selRect.bottom() + 1,
                                 selRect.width(), targetRect.bottom() - selRect.bottom());
            } else if (dragTarget.column() > m_fillDragStart.column()) {
                fillRect = QRect(selRect.right() + 1, selRect.top(),
                                 targetRect.right() - selRect.right(), selRect.height());
            }

            if (!fillRect.isNull()) {
                QPen dashPen(ThemeManager::instance().currentTheme().focusBorderColor, 1, Qt::DashLine);
                painter.setPen(dashPen);
                painter.setBrush(QColor(198, 217, 240, 40));
                painter.drawRect(fillRect);
            }
        }
    }

    // Draw trace precedent/dependent arrows
    if ((m_showPrecedents || m_showDependents) && !m_tracedCells.empty() && m_model) {
        QPainter tracePainter(viewport());
        drawTraceArrows(tracePainter);
    }

    // Draw outline gutter (row grouping indicators)
    if (m_spreadsheet && m_spreadsheet->getMaxRowOutlineLevel() > 0) {
        QPainter gutterPainter(this);
        paintOutlineGutter(gutterPainter);
    }
}

// ============== Marching ants (clipboard range animation) ==============

void SpreadsheetView::onMarchingAntsTick() {
    m_marchingAntsOffset = (m_marchingAntsOffset + 1) % 10;
    viewport()->update();
}

void SpreadsheetView::clearClipboardRange() {
    m_hasClipboardRange = false;
    m_marchingAntsOffset = 0;
    if (m_marchingAntsTimer) m_marchingAntsTimer->stop();
    viewport()->update();
}

// ============== Fill series helpers ==============

QRect SpreadsheetView::getSelectionBoundingRect() const {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        QModelIndex current = currentIndex();
        if (current.isValid()) return visualRect(current);
        return QRect();
    }

    QRect result;
    for (const auto& idx : selected) {
        QRect r = visualRect(idx);
        if (result.isNull()) result = r;
        else result = result.united(r);
    }
    return result;
}

bool SpreadsheetView::isOverFillHandle(const QPoint& pos) const {
    if (m_fillHandleRect.isNull()) return false;
    // Generous hit area (5px padding) so it's easy to grab the fill handle
    QRect hitRect = m_fillHandleRect.adjusted(-5, -5, 5, 5);
    return hitRect.contains(pos);
}

// Adjust cell references in a formula by row/col offsets.
// Respects absolute references ($A$1 = no shift, $A1 = shift row only, A$1 = shift col only).
static QString adjustFormulaReferences(const QString& formula, int rowOffset, int colOffset) {
    static QRegularExpression cellRefRe("(\\$?)([A-Za-z]+)(\\$?)(\\d+)");
    QString result;
    int lastEnd = 0;
    auto it = cellRefRe.globalMatch(formula);
    while (it.hasNext()) {
        auto match = it.next();
        // Skip if letters look like a function name (preceded by nothing or operator, not letters)
        // Cell refs have 1-3 letters; function names are longer but also letter-only before '('
        // We rely on the digit requirement to distinguish: SUM( has no digits, A1 does
        result += formula.mid(lastEnd, match.capturedStart() - lastEnd);

        bool colAbsolute = !match.captured(1).isEmpty();
        QString colLetters = match.captured(2);
        bool rowAbsolute = !match.captured(3).isEmpty();
        int rowNum = match.captured(4).toInt();

        // Adjust column if not absolute
        if (!colAbsolute && colOffset != 0) {
            int colIdx = 0;
            for (QChar ch : colLetters)
                colIdx = colIdx * 26 + (ch.toUpper().toLatin1() - 'A' + 1);
            colIdx += colOffset;
            if (colIdx < 1) colIdx = 1;
            bool wasUpper = !colLetters.isEmpty() && colLetters[0].isUpper();
            colLetters.clear();
            int c = colIdx;
            while (c > 0) {
                colLetters = QChar((wasUpper ? 'A' : 'a') + (c - 1) % 26) + colLetters;
                c = (c - 1) / 26;
            }
        }

        // Adjust row if not absolute
        if (!rowAbsolute && rowOffset != 0) {
            rowNum += rowOffset;
            if (rowNum < 1) rowNum = 1;
        }

        result += (colAbsolute ? "$" : "") + colLetters + (rowAbsolute ? "$" : "") + QString::number(rowNum);
        lastEnd = match.capturedEnd();
    }
    result += formula.mid(lastEnd);
    return result;
}

void SpreadsheetView::performFillSeries() {
    if (!m_spreadsheet || !m_fillDragStart.isValid()) return;

    QModelIndex dragTarget = indexAt(m_fillDragCurrent);
    if (!dragTarget.isValid()) return;

    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    // Find selection bounds
    int selMinRow = selected.first().row(), selMaxRow = selMinRow;
    int selMinCol = selected.first().column(), selMaxCol = selMinCol;
    for (const auto& idx : selected) {
        selMinRow = qMin(selMinRow, idx.row());
        selMaxRow = qMax(selMaxRow, idx.row());
        selMinCol = qMin(selMinCol, idx.column());
        selMaxCol = qMax(selMaxCol, idx.column());
    }

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);

    // Fill down
    if (dragTarget.row() > selMaxRow) {
        int fillCount = dragTarget.row() - selMaxRow;
        int seedCount = selMaxRow - selMinRow + 1;

        for (int col = selMinCol; col <= selMaxCol; ++col) {
            // Check if any seed is a formula
            bool hasFormula = false;
            for (int row = selMinRow; row <= selMaxRow; ++row) {
                auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                if (cell->getType() == CellType::Formula) { hasFormula = true; break; }
            }

            if (hasFormula) {
                // Formula fill: cycle through seeds, adjust row references
                for (int i = 0; i < fillCount; ++i) {
                    int targetRow = selMaxRow + 1 + i;
                    int seedIdx = i % seedCount;
                    int seedRow = selMinRow + seedIdx;
                    auto seedCell = m_spreadsheet->getCell(CellAddress(seedRow, col));

                    CellAddress addr(targetRow, col);
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                    if (seedCell->getType() == CellType::Formula) {
                        int rowOffset = targetRow - seedRow;
                        QString adjusted = adjustFormulaReferences(seedCell->getFormula(), rowOffset, 0);
                        m_spreadsheet->setCellFormula(addr, adjusted);
                    } else {
                        QModelIndex idx = m_model->index(targetRow, col);
                        m_model->setData(idx, seedCell->getValue().toString());
                    }

                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            } else {
                // Original behavior: numeric/text series
                QStringList seeds;
                for (int row = selMinRow; row <= selMaxRow; ++row) {
                    auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                    seeds.append(cell->getValue().toString());
                }
                int totalCount = seeds.size() + fillCount;
                QStringList series = FillSeries::generateSeries(seeds, totalCount);

                for (int i = 0; i < fillCount; ++i) {
                    int targetRow = selMaxRow + 1 + i;
                    CellAddress addr(targetRow, col);
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                    QModelIndex idx = m_model->index(targetRow, col);
                    m_model->setData(idx, series[seeds.size() + i]);

                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            }
        }
    }
    // Fill right
    else if (dragTarget.column() > selMaxCol) {
        int fillCount = dragTarget.column() - selMaxCol;
        int seedCount = selMaxCol - selMinCol + 1;

        for (int row = selMinRow; row <= selMaxRow; ++row) {
            // Check if any seed is a formula
            bool hasFormula = false;
            for (int col = selMinCol; col <= selMaxCol; ++col) {
                auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                if (cell->getType() == CellType::Formula) { hasFormula = true; break; }
            }

            if (hasFormula) {
                // Formula fill: cycle through seeds, adjust column references
                for (int i = 0; i < fillCount; ++i) {
                    int targetCol = selMaxCol + 1 + i;
                    int seedIdx = i % seedCount;
                    int seedCol = selMinCol + seedIdx;
                    auto seedCell = m_spreadsheet->getCell(CellAddress(row, seedCol));

                    CellAddress addr(row, targetCol);
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                    if (seedCell->getType() == CellType::Formula) {
                        int colOffset = targetCol - seedCol;
                        QString adjusted = adjustFormulaReferences(seedCell->getFormula(), 0, colOffset);
                        m_spreadsheet->setCellFormula(addr, adjusted);
                    } else {
                        QModelIndex idx = m_model->index(row, targetCol);
                        m_model->setData(idx, seedCell->getValue().toString());
                    }

                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            } else {
                // Original behavior: numeric/text series
                QStringList seeds;
                for (int col = selMinCol; col <= selMaxCol; ++col) {
                    auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                    seeds.append(cell->getValue().toString());
                }
                int totalCount = seeds.size() + fillCount;
                QStringList series = FillSeries::generateSeries(seeds, totalCount);

                for (int i = 0; i < fillCount; ++i) {
                    int targetCol = selMaxCol + 1 + i;
                    CellAddress addr(row, targetCol);
                    before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                    QModelIndex idx = m_model->index(row, targetCol);
                    m_model->setData(idx, series[seeds.size() + i]);

                    after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                }
            }
        }
    }

    m_model->setSuppressUndo(false);

    if (!before.empty()) {
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<MultiCellEditCommand>(before, after, "Fill Series"));
    }
}

// ============== Multi-select resize ==============

void SpreadsheetView::showResizeTooltip(const QPoint& globalPos, const QString& text) {
    if (!m_resizeTooltip) {
        m_resizeTooltip = new QLabel(nullptr, Qt::ToolTip | Qt::FramelessWindowHint);
        m_resizeTooltip->setStyleSheet(
            "QLabel { background: #333; color: white; padding: 4px 8px; "
            "border-radius: 3px; font-size: 12px; }");
    }
    if (!m_resizeTooltipTimer) {
        m_resizeTooltipTimer = new QTimer(this);
        m_resizeTooltipTimer->setSingleShot(true);
        connect(m_resizeTooltipTimer, &QTimer::timeout, this, &SpreadsheetView::hideResizeTooltip);
    }
    m_resizeTooltip->setText(text);
    m_resizeTooltip->adjustSize();
    m_resizeTooltip->move(globalPos.x() + 15, globalPos.y() + 15);
    m_resizeTooltip->show();
    // Auto-hide 500ms after the last resize event (when user releases mouse)
    m_resizeTooltipTimer->start(500);
}

void SpreadsheetView::hideResizeTooltip() {
    if (m_resizeTooltip) {
        m_resizeTooltip->hide();
    }
}

void SpreadsheetView::onHorizontalSectionResized(int logicalIndex, int /*oldSize*/, int newSize) {
    if (m_resizingMultiple) return;
    m_resizingMultiple = true;

    // Persist dimension to the current sheet
    if (m_spreadsheet) m_spreadsheet->setColumnWidth(logicalIndex, newSize);

    // Show resize tooltip
    showResizeTooltip(QCursor::pos(), QString("Width: %1 px").arg(newSize));

    QModelIndexList selected = selectionModel()->selectedColumns();
    if (selected.size() > 1) {
        for (const auto& idx : selected) {
            if (idx.column() != logicalIndex) {
                horizontalHeader()->resizeSection(idx.column(), newSize);
                if (m_spreadsheet) m_spreadsheet->setColumnWidth(idx.column(), newSize);
            }
        }
    }

    m_resizingMultiple = false;
}

void SpreadsheetView::onVerticalSectionResized(int logicalIndex, int /*oldSize*/, int newSize) {
    if (m_resizingMultiple) return;
    m_resizingMultiple = true;

    // Persist dimension to the current sheet
    if (m_spreadsheet) m_spreadsheet->setRowHeight(logicalIndex, newSize);

    // Show resize tooltip
    showResizeTooltip(QCursor::pos(), QString("Height: %1 px").arg(newSize));

    QModelIndexList selected = selectionModel()->selectedRows();
    if (selected.size() > 1) {
        for (const auto& idx : selected) {
            if (idx.row() != logicalIndex) {
                verticalHeader()->resizeSection(idx.row(), newSize);
                if (m_spreadsheet) m_spreadsheet->setRowHeight(logicalRow(idx), newSize);
            }
        }
    }

    m_resizingMultiple = false;
}

// ============== Slots ==============

void SpreadsheetView::onCellClicked(const QModelIndex& index) {
    emitCellSelected(index);
}

void SpreadsheetView::onCellDoubleClicked(const QModelIndex& index) {
    // Picklist/Checkbox: block normal editor on double-click
    if (index.isValid() && m_spreadsheet) {
        int logRow = m_model ? m_model->toLogicalRow(index.row()) : index.row();
        auto cell = m_spreadsheet->getCellIfExists(logRow, index.column());
        if (cell) {
            const auto& style = cell->getStyle();
            if (style.numberFormat == "Picklist") {
                // Already opened by single-click; just block editor
                return;
            }
            if (style.numberFormat == "Checkbox") {
                return;
            }
        }
    }
    edit(index);
}

void SpreadsheetView::onDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight) {
    update(topLeft);
    update(bottomRight);
}

// --- Formula cell flash animation ---
// Simple timer-driven approach: hold yellow for 2s, then fade out over 0.5s.
// A single QTimer ticks every 30ms and updates all active cell flashes.

double SpreadsheetView::cellAnimationProgress(int row, int col) const {
    auto it = m_cellAnimations.find({row, col});
    if (it == m_cellAnimations.end()) return 0.0;
    return it->progress;
}

void SpreadsheetView::startCellFlashAnimation(int row, int col) {
    auto key = QPair<int,int>(row, col);

    CellAnim ca;
    ca.progress = 1.0;
    ca.elapsedMs = 0;
    m_cellAnimations[key] = ca;

    // Force immediate repaint so yellow shows right away
    viewport()->update();

    // Start the shared timer if not already running
    if (!m_flashTimer) {
        m_flashTimer = new QTimer(this);
        m_flashTimer->setInterval(FLASH_TICK_MS);
        connect(m_flashTimer, &QTimer::timeout, this, &SpreadsheetView::onFlashTimerTick);
    }
    if (!m_flashTimer->isActive()) {
        m_flashTimer->start();
    }
}

void SpreadsheetView::onFlashTimerTick() {
    if (m_cellAnimations.isEmpty()) {
        m_flashTimer->stop();
        return;
    }

    QVector<QPair<int,int>> toRemove;

    for (auto it = m_cellAnimations.begin(); it != m_cellAnimations.end(); ++it) {
        it->elapsedMs += FLASH_TICK_MS;

        if (it->elapsedMs <= FLASH_HOLD_MS) {
            it->progress = 1.0;
        } else {
            int fadeElapsed = it->elapsedMs - FLASH_HOLD_MS;
            if (fadeElapsed >= FLASH_FADE_MS) {
                it->progress = 0.0;
                toRemove.append(it.key());
            } else {
                it->progress = 1.0 - static_cast<double>(fadeElapsed) / FLASH_FADE_MS;
            }
        }
    }

    for (const auto& key : toRemove) {
        m_cellAnimations.remove(key);
    }

    // Force full viewport repaint so delegate redraws with updated progress
    viewport()->update();

    if (m_cellAnimations.isEmpty()) {
        m_flashTimer->stop();
    }
}

// ===== Checkbox toggle with undo =====
void SpreadsheetView::toggleCheckbox(int row, int col) {
    if (!m_spreadsheet) return;
    CellAddress addr(row, col);
    CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);

    auto cell = m_spreadsheet->getCell(addr);
    bool current = false;
    auto val = cell->getValue();
    if (val.typeId() == QMetaType::Bool) current = val.toBool();
    else {
        QString s = val.toString().toLower();
        current = (s == "true" || s == "1");
    }
    cell->setValue(QVariant(!current));

    CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<CellEditCommand>(before, after));

    if (m_model) {
        QModelIndex idx = m_model->index(row, col);
        emit m_model->dataChanged(idx, idx);
    }
}

// ===== Insert Checkbox on selected cells =====
void SpreadsheetView::insertCheckbox() {
    if (!m_spreadsheet || !m_model) return;
    auto selection = selectionModel()->selectedIndexes();
    if (selection.isEmpty()) return;

    std::vector<CellSnapshot> before, after;
    for (const auto& idx : selection) {
        CellAddress addr(logicalRow(idx), idx.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        CellStyle style = cell->getStyle();
        style.numberFormat = "Checkbox";
        cell->setStyle(style);
        if (cell->getValue().isNull() || cell->getValue().toString().isEmpty()) {
            cell->setValue(QVariant(false));
        }
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }
    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Insert Checkbox"));

    QModelIndex topLeft = m_model->index(selection.first().row(), selection.first().column());
    QModelIndex bottomRight = m_model->index(selection.last().row(), selection.last().column());
    emit m_model->dataChanged(topLeft, bottomRight);
}

// ===== Create Picklist Dialog (same as Edit dialog) =====
void SpreadsheetView::showCreatePicklistDialog() {
    if (!m_spreadsheet || !m_model) return;
    auto selection = selectionModel()->selectedIndexes();
    if (selection.isEmpty()) return;

    static const QColor defaultTagBg[] = {
        QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
        QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
        QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
    };

    // Save original target — the picklist will be applied to these cells
    int origMinRow = INT_MAX, origMaxRow = 0, origMinCol = INT_MAX, origMaxCol = 0;
    for (const auto& idx : selection) {
        origMinRow = qMin(origMinRow, idx.row());
        origMaxRow = qMax(origMaxRow, idx.row());
        origMinCol = qMin(origMinCol, idx.column());
        origMaxCol = qMax(origMaxCol, idx.column());
    }
    std::shared_ptr<Spreadsheet> originalSheet = m_spreadsheet;

    // State preserved across dialog re-opens (for picker round-trip)
    QString prefilledRange;
    QString prefilledSheet;
    bool startInRangeMode = false;
    static constexpr int PickerRequested = 2;

    for (;;) {

    QDialog dlg(this);
    dlg.setWindowTitle("Create Picklist");
    dlg.setFixedWidth(400);
    dlg.setStyleSheet(
        "QDialog { background: white; }"
        "QLineEdit { border: 1px solid #D1D5DB; border-radius: 6px; padding: 6px 8px; "
        "font-size: 13px; color: #1F2937; }"
        "QLineEdit:focus { border-color: #2563EB; }");

    QVBoxLayout* lo = new QVBoxLayout(&dlg);
    lo->setContentsMargins(16, 12, 16, 12);
    lo->setSpacing(8);

    // === Source toggle: Manual vs Cell Range ===
    QRadioButton* manualRadio = new QRadioButton("Enter manually", &dlg);
    QRadioButton* rangeRadio = new QRadioButton("From cell range", &dlg);
    if (startInRangeMode) rangeRadio->setChecked(true);
    else manualRadio->setChecked(true);
    manualRadio->setStyleSheet("QRadioButton { font-size: 12px; color: #374151; }");
    rangeRadio->setStyleSheet("QRadioButton { font-size: 12px; color: #374151; }");
    QHBoxLayout* sourceRow = new QHBoxLayout();
    sourceRow->setSpacing(12);
    sourceRow->addWidget(manualRadio);
    sourceRow->addWidget(rangeRadio);
    sourceRow->addStretch();
    lo->addLayout(sourceRow);

    // --- Range source panel ---
    QWidget* rangePanel = new QWidget(&dlg);
    QVBoxLayout* rangePanelLo = new QVBoxLayout(rangePanel);
    rangePanelLo->setContentsMargins(0, 0, 0, 0);
    rangePanelLo->setSpacing(4);

    // Row 1: Sheet combo + Range edit + Pick button
    QHBoxLayout* rangeLo = new QHBoxLayout();
    rangeLo->setContentsMargins(0, 0, 0, 0);
    rangeLo->setSpacing(6);

    QComboBox* sheetCombo = new QComboBox(rangePanel);
    if (m_allSheets.empty() && m_spreadsheet) {
        sheetCombo->addItem(m_spreadsheet->getSheetName());
    } else {
        for (const auto& s : m_allSheets)
            sheetCombo->addItem(s->getSheetName());
    }
    if (!prefilledSheet.isEmpty()) {
        int idx = sheetCombo->findText(prefilledSheet);
        if (idx >= 0) sheetCombo->setCurrentIndex(idx);
    } else if (m_spreadsheet) {
        int idx = sheetCombo->findText(m_spreadsheet->getSheetName());
        if (idx >= 0) sheetCombo->setCurrentIndex(idx);
    }
    sheetCombo->setFixedWidth(100);
    sheetCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D1D5DB; border-radius: 6px; padding: 4px 8px; font-size: 12px; }"
        "QComboBox:focus { border-color: #2563EB; }");
    rangeLo->addWidget(sheetCombo);

    QLineEdit* rangeEdit = new QLineEdit(rangePanel);
    rangeEdit->setPlaceholderText("e.g. A1:A10");
    if (!prefilledRange.isEmpty()) rangeEdit->setText(prefilledRange);
    rangeLo->addWidget(rangeEdit, 1);

    // Pick button: closes dialog, opens picker bar, then re-opens dialog
    QPushButton* pickBtn = new QPushButton(QString::fromUtf8("\xe2\xac\x92"), rangePanel);
    pickBtn->setFixedSize(28, 28);
    pickBtn->setToolTip("Select range from spreadsheet");
    pickBtn->setCursor(Qt::PointingHandCursor);
    pickBtn->setStyleSheet(
        "QPushButton { background: #EFF6FF; border: 1px solid #93C5FD; border-radius: 6px; "
        "font-size: 14px; color: #2563EB; }"
        "QPushButton:hover { background: #DBEAFE; }");
    rangeLo->addWidget(pickBtn);
    rangePanelLo->addLayout(rangeLo);

    // Row 2: Ignore blanks + Sort
    QHBoxLayout* rangeOptsLo = new QHBoxLayout();
    rangeOptsLo->setContentsMargins(0, 2, 0, 0);
    rangeOptsLo->setSpacing(10);

    QCheckBox* ignoreBlanks = new QCheckBox("Ignore blank cells", rangePanel);
    ignoreBlanks->setChecked(true);
    ignoreBlanks->setStyleSheet("QCheckBox { font-size: 11px; color: #6B7280; }");
    rangeOptsLo->addWidget(ignoreBlanks);

    QLabel* sortLabel = new QLabel("Sort:", rangePanel);
    sortLabel->setStyleSheet("QLabel { font-size: 11px; color: #6B7280; }");
    rangeOptsLo->addWidget(sortLabel);

    QComboBox* sortCombo = new QComboBox(rangePanel);
    sortCombo->addItems({"None", "A \u2192 Z", "Z \u2192 A"});
    sortCombo->setFixedWidth(80);
    sortCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D1D5DB; border-radius: 4px; padding: 2px 6px; font-size: 11px; }"
        "QComboBox:focus { border-color: #2563EB; }");
    rangeOptsLo->addWidget(sortCombo);
    rangeOptsLo->addStretch();
    rangePanelLo->addLayout(rangeOptsLo);

    rangePanel->setVisible(startInRangeMode);
    lo->addWidget(rangePanel);

    // --- Manual source panel ---
    QWidget* manualPanel = new QWidget(&dlg);
    QVBoxLayout* manualLo = new QVBoxLayout(manualPanel);
    manualLo->setContentsMargins(0, 0, 0, 0);
    manualLo->setSpacing(6);
    manualPanel->setVisible(!startInRangeMode);

    // Toggle panel visibility
    QObject::connect(manualRadio, &QRadioButton::toggled, &dlg, [manualPanel, rangePanel](bool checked) {
        manualPanel->setVisible(checked);
        rangePanel->setVisible(!checked);
    });

    // Pick button: close dialog with custom result code
    QObject::connect(pickBtn, &QPushButton::clicked, &dlg, [&dlg]() {
        dlg.done(PickerRequested);
    });

    QLabel* sub = new QLabel("Set label and color for each option:", manualPanel);
    sub->setStyleSheet("QLabel { font-size: 12px; color: #6B7280; }");
    manualLo->addWidget(sub);

    // Scrollable list of option rows
    QScrollArea* scroll = new QScrollArea(manualPanel);
    scroll->setWidgetResizable(true);
    scroll->setFrameShape(QFrame::NoFrame);
    scroll->setMinimumHeight(140);
    scroll->setStyleSheet("QScrollArea { background: transparent; border: none; }");

    QWidget* listWidget = new QWidget(scroll);
    QVBoxLayout* listLo = new QVBoxLayout(listWidget);
    listLo->setContentsMargins(0, 0, 0, 0);
    listLo->setSpacing(6);

    struct OptionRow { QLineEdit* text; QPushButton* colorBtn; QString color; };
    auto optionRows = std::make_shared<QList<OptionRow>>();

    auto addOptionRow = [&](const QString& text, const QString& colorStr, int idx) {
        QHBoxLayout* rowLo = new QHBoxLayout();
        rowLo->setSpacing(8);

        QLineEdit* lineEdit = new QLineEdit(text, listWidget);
        lineEdit->setPlaceholderText("Option name...");
        rowLo->addWidget(lineEdit, 1);

        // Color button
        QPushButton* colorBtn = new QPushButton(listWidget);
        colorBtn->setFixedSize(28, 28);
        colorBtn->setCursor(Qt::PointingHandCursor);
        QColor displayColor = colorStr.isEmpty() ? defaultTagBg[idx % 12] : QColor(colorStr);
        colorBtn->setStyleSheet(QString(
            "QPushButton { background: %1; border: 1px solid #D1D5DB; border-radius: 6px; }"
            "QPushButton:hover { border-color: #2563EB; }").arg(displayColor.name()));
        rowLo->addWidget(colorBtn);

        // Remove button
        QPushButton* removeBtn = new QPushButton(QString::fromUtf8("\u2715"), listWidget);
        removeBtn->setFixedSize(24, 24);
        removeBtn->setCursor(Qt::PointingHandCursor);
        removeBtn->setStyleSheet(
            "QPushButton { background: transparent; border: none; color: #9CA3AF; font-size: 13px; }"
            "QPushButton:hover { color: #EF4444; }");
        rowLo->addWidget(removeBtn);

        listLo->addLayout(rowLo);

        OptionRow opt{lineEdit, colorBtn, colorStr};
        optionRows->append(opt);
        int rowIdx = optionRows->size() - 1;

        // Color picker click
        QObject::connect(colorBtn, &QPushButton::clicked, &dlg, [optionRows, colorBtn, rowIdx, &dlg]() {
            QColor cur = (*optionRows)[rowIdx].color.isEmpty()
                ? colorBtn->palette().button().color() : QColor((*optionRows)[rowIdx].color);
            QColor picked = QColorDialog::getColor(cur, &dlg, "Option Color");
            if (picked.isValid()) {
                (*optionRows)[rowIdx].color = picked.name();
                colorBtn->setStyleSheet(QString(
                    "QPushButton { background: %1; border: 1px solid #D1D5DB; border-radius: 6px; }"
                    "QPushButton:hover { border-color: #2563EB; }").arg(picked.name()));
            }
        });

        // Remove click
        QObject::connect(removeBtn, &QPushButton::clicked, &dlg, [optionRows, rowIdx, rowLo, lineEdit, colorBtn, removeBtn]() {
            lineEdit->hide(); colorBtn->hide(); removeBtn->hide();
            (*optionRows)[rowIdx].text = nullptr;
        });
    };

    // Start with 3 empty option rows
    for (int i = 0; i < 3; ++i) {
        addOptionRow(QString("Option %1").arg(i + 1), "", i);
    }

    listLo->addStretch();
    scroll->setWidget(listWidget);
    manualLo->addWidget(scroll, 1);

    // Add option button
    QPushButton* addBtn = new QPushButton("+ Add Option", manualPanel);
    addBtn->setCursor(Qt::PointingHandCursor);
    addBtn->setFixedHeight(30);
    addBtn->setStyleSheet(
        "QPushButton { background: transparent; color: #2563EB; border: 1px dashed #93C5FD; "
        "border-radius: 6px; font-size: 12px; font-weight: 500; }"
        "QPushButton:hover { background: #EFF6FF; }");
    manualLo->addWidget(addBtn);
    connect(addBtn, &QPushButton::clicked, &dlg, [&, addOptionRow]() {
        int idx = optionRows->size();
        addOptionRow("", "", idx);
        listWidget->adjustSize();
    });

    lo->addWidget(manualPanel, 1);
    lo->addSpacing(4);
    QHBoxLayout* btnRow = new QHBoxLayout();
    btnRow->addStretch();

    QPushButton* cancelBtn = new QPushButton("Cancel", &dlg);
    cancelBtn->setFixedHeight(34);
    cancelBtn->setCursor(Qt::PointingHandCursor);
    cancelBtn->setStyleSheet(
        "QPushButton { background: white; color: #374151; border: 1px solid #D1D5DB; "
        "border-radius: 6px; padding: 0 20px; font-size: 13px; font-weight: 500; }"
        "QPushButton:hover { background: #F9FAFB; border-color: #9CA3AF; }");
    btnRow->addWidget(cancelBtn);

    QPushButton* saveBtn = new QPushButton("Create", &dlg);
    saveBtn->setFixedHeight(34);
    saveBtn->setCursor(Qt::PointingHandCursor);
    saveBtn->setStyleSheet(
        "QPushButton { background: #2563EB; color: white; border: none; "
        "border-radius: 6px; padding: 0 24px; font-size: 13px; font-weight: 600; }"
        "QPushButton:hover { background: #1D4ED8; }");
    btnRow->addWidget(saveBtn);
    lo->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &dlg, &QDialog::reject);
    connect(saveBtn, &QPushButton::clicked, &dlg, &QDialog::accept);

    int result = dlg.exec();

    // --- Picker mode: dialog closed, show picker bar with its own event loop ---
    if (result == PickerRequested) {
        // Use the SpreadsheetView as parent (not viewport) so the bar survives
        // sheet switches which call setModel(nullptr) and repaint the viewport.
        QFrame* bar = new QFrame(this);
        bar->setObjectName("plRangeBar");
        bar->setStyleSheet(
            "QFrame#plRangeBar { background: #EFF6FF; border: 1px solid #93C5FD; border-radius: 8px; }");
        QHBoxLayout* barLo = new QHBoxLayout(bar);
        barLo->setContentsMargins(10, 6, 10, 6);
        barLo->setSpacing(8);

        QLabel* barLabel = new QLabel("Select range:", bar);
        barLabel->setStyleSheet("font-size: 12px; color: #1E40AF; font-weight: 500;");
        barLo->addWidget(barLabel);

        QLineEdit* barRange = new QLineEdit(prefilledRange, bar);
        barRange->setReadOnly(true);
        barRange->setPlaceholderText("Click and drag to select cells...");
        barRange->setStyleSheet(
            "QLineEdit { border: 1px solid #93C5FD; border-radius: 4px; padding: 3px 6px; "
            "font-size: 12px; background: white; }");
        barLo->addWidget(barRange, 1);

        QPushButton* barOk = new QPushButton("Done", bar);
        barOk->setFixedHeight(26);
        barOk->setCursor(Qt::PointingHandCursor);
        barOk->setStyleSheet(
            "QPushButton { background: #2563EB; color: white; border: none; border-radius: 4px; "
            "padding: 0 12px; font-size: 12px; font-weight: 500; }"
            "QPushButton:hover { background: #1D4ED8; }");
        barLo->addWidget(barOk);

        QPushButton* barCancel = new QPushButton("Cancel", bar);
        barCancel->setFixedHeight(26);
        barCancel->setCursor(Qt::PointingHandCursor);
        barCancel->setStyleSheet(
            "QPushButton { background: white; color: #6B7280; border: 1px solid #D1D5DB; "
            "border-radius: 4px; padding: 0 10px; font-size: 12px; }"
            "QPushButton:hover { background: #F3F4F6; }");
        barLo->addWidget(barCancel);

        // Position over the viewport area, offset by headers
        int hdrH = horizontalHeader()->height();
        int hdrW = verticalHeader()->width();
        bar->setGeometry(hdrW + 8, hdrH + 8, viewport()->width() - 16, 38);
        bar->show();
        bar->raise();

        // Poll selection with a timer — robust across sheet switches
        QTimer* pollTimer = new QTimer(bar);
        pollTimer->setInterval(80);
        connect(pollTimer, &QTimer::timeout, bar, [barRange, bar, this]() {
            // Keep bar raised above the viewport after sheet switches
            bar->raise();
            if (!selectionModel()) return;
            auto indexes = selectionModel()->selectedIndexes();
            if (indexes.isEmpty()) return;
            int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
            for (const auto& idx : indexes) {
                minR = qMin(minR, idx.row());
                maxR = qMax(maxR, idx.row());
                minC = qMin(minC, idx.column());
                maxC = qMax(maxC, idx.column());
            }
            CellRange cr(CellAddress(minR, minC), CellAddress(maxR, maxC));
            QString newText = cr.toString();
            if (barRange->text() != newText)
                barRange->setText(newText);
        });
        pollTimer->start();

        // Local event loop — blocks this function but allows full app interaction
        QEventLoop pickerLoop;
        bool pickerAccepted = false;

        connect(barOk, &QPushButton::clicked, bar, [&]() {
            prefilledRange = barRange->text();
            prefilledSheet = m_spreadsheet ? m_spreadsheet->getSheetName() : QString();
            pickerAccepted = true;
            pickerLoop.quit();
        });

        connect(barCancel, &QPushButton::clicked, bar, [&]() {
            pickerLoop.quit();
        });

        pickerLoop.exec();
        bar->deleteLater();

        startInRangeMode = true;
        continue;  // Re-open dialog with range pre-filled
    }

    if (result == QDialog::Accepted) {
        QStringList newOpts;
        QStringList newColors;
        QString sourceRef;

        if (rangeRadio->isChecked()) {
            QString sheetName = sheetCombo->currentText();
            QString rangeText = rangeEdit->text().trimmed();
            if (rangeText.isEmpty()) return;

            sourceRef = sheetName + "!" + rangeText;
            newOpts = resolvePicklistFromRange(sourceRef);
            if (newOpts.isEmpty()) return;

            // Apply ignore-blanks filter
            if (ignoreBlanks->isChecked()) {
                newOpts.removeAll("");
            }
            // Apply sorting
            int sortMode = sortCombo->currentIndex(); // 0=None, 1=A-Z, 2=Z-A
            if (sortMode == 1) {
                std::sort(newOpts.begin(), newOpts.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) < 0;
                });
            } else if (sortMode == 2) {
                std::sort(newOpts.begin(), newOpts.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) > 0;
                });
            }
        } else {
            // Manual entry
            for (const auto& opt : *optionRows) {
                if (!opt.text) continue;
                QString t = opt.text->text().trimmed();
                if (!t.isEmpty()) {
                    newOpts.append(t);
                    newColors.append(opt.color);
                }
            }
        }

        if (!newOpts.isEmpty()) {
            // Apply picklist to the ORIGINAL selection (not current, which may have changed during picker)
            CellRange targetRange(CellAddress(origMinRow, origMinCol), CellAddress(origMaxRow, origMaxCol));
            Spreadsheet::DataValidationRule rule;
            rule.range = targetRange;
            rule.type = Spreadsheet::DataValidationRule::List;
            rule.listItems = newOpts;
            rule.listItemColors = newColors;
            rule.showErrorAlert = false;
            if (!sourceRef.isEmpty()) {
                rule.listSourceRange = sourceRef;
                rule.listSortMode = sortCombo->currentIndex();
                rule.listIgnoreBlanks = ignoreBlanks->isChecked();
            }
            originalSheet->addValidationRule(rule);

            // Set Picklist format on original target cells
            std::vector<CellSnapshot> before, after;
            for (int r = origMinRow; r <= origMaxRow; ++r) {
                for (int c = origMinCol; c <= origMaxCol; ++c) {
                    CellAddress addr(r, c);
                    before.push_back(originalSheet->takeCellSnapshot(addr));
                    auto cell = originalSheet->getCell(addr);
                    CellStyle style = cell->getStyle();
                    style.numberFormat = "Picklist";
                    cell->setStyle(style);
                    after.push_back(originalSheet->takeCellSnapshot(addr));
                }
            }
            originalSheet->getUndoManager().pushCommand(
                std::make_unique<MultiCellEditCommand>(before, after, "Insert Picklist"));

            // Switch back to original sheet if we navigated away
            if (m_spreadsheet != originalSheet) {
                for (int si = 0; si < (int)m_allSheets.size(); ++si) {
                    if (m_allSheets[si] == originalSheet) {
                        emit requestSwitchToSheet(si);
                        break;
                    }
                }
            }
            // Refresh view and select the target range
            if (m_spreadsheet == originalSheet && m_model) {
                QModelIndex topLeft = m_model->index(origMinRow, origMinCol);
                QModelIndex bottomRight = m_model->index(origMaxRow, origMaxCol);
                emit m_model->dataChanged(topLeft, bottomRight);
                // Select the original target range
                QItemSelection sel(topLeft, bottomRight);
                selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
                setCurrentIndex(topLeft);
            }
        }
    }

    break;  // Exit loop (accepted or rejected)
    } // end for(;;)
}

// ===== Insert Picklist on selected cells =====
void SpreadsheetView::insertPicklist(const QStringList& options, const QStringList& colors) {
    if (!m_spreadsheet || !m_model) return;
    auto selection = selectionModel()->selectedIndexes();
    if (selection.isEmpty()) return;

    // Determine range for validation rule
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selection) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    // Create validation rule with list items
    Spreadsheet::DataValidationRule rule;
    rule.range = CellRange(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    rule.type = Spreadsheet::DataValidationRule::List;
    rule.listItems = options;
    rule.listItemColors = colors;
    rule.showErrorAlert = false; // picklist allows free typing
    m_spreadsheet->addValidationRule(rule);

    // Set numberFormat to Picklist on all cells
    std::vector<CellSnapshot> before, after;
    for (const auto& idx : selection) {
        CellAddress addr(logicalRow(idx), idx.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        CellStyle style = cell->getStyle();
        style.numberFormat = "Picklist";
        cell->setStyle(style);
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }
    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Insert Picklist"));

    QModelIndex topLeft = m_model->index(minRow, minCol);
    QModelIndex bottomRight = m_model->index(maxRow, maxCol);
    emit m_model->dataChanged(topLeft, bottomRight);
}

// ===== Picklist multi-select popup =====
QStringList SpreadsheetView::resolvePicklistFromRange(const QString& listSourceRange) const {
    if (listSourceRange.isEmpty()) return {};
    int excl = listSourceRange.indexOf('!');
    if (excl < 0) return {};

    QString sheetName = listSourceRange.left(excl);
    QString rangeText = listSourceRange.mid(excl + 1);

    std::shared_ptr<Spreadsheet> targetSheet;
    for (const auto& s : m_allSheets) {
        if (s->getSheetName() == sheetName) { targetSheet = s; break; }
    }
    if (!targetSheet && m_spreadsheet && m_spreadsheet->getSheetName() == sheetName)
        targetSheet = m_spreadsheet;
    if (!targetSheet) return {};

    CellRange cr(rangeText);
    if (!cr.isValid()) return {};

    QStringList result;
    for (int r = cr.getStart().row; r <= cr.getEnd().row; ++r) {
        for (int c = cr.getStart().col; c <= cr.getEnd().col; ++c) {
            QVariant val = targetSheet->getCellValue(CellAddress(r, c));
            QString s = val.toString().trimmed();
            if (!s.isEmpty()) result.append(s);
        }
    }
    return result;
}

void SpreadsheetView::showValidationDropdown(const QModelIndex& idx, const Spreadsheet::DataValidationRule* rule) {
    if (!rule || !m_spreadsheet) return;

    QMenu popup(this);
    popup.setStyleSheet(
        "QMenu { background: white; border: 1px solid #D0D5DD; padding: 2px 0; }"
        "QMenu::item { padding: 4px 16px; font-size: 12px; color: #344054; }"
        "QMenu::item:selected { background: #EFF6FF; color: #1D2939; }");

    QStringList items = rule->listItems;
    if (items.isEmpty() && !rule->listSourceRange.isEmpty()) {
        items = resolvePicklistFromRange(rule->listSourceRange);
    }
    for (const QString& item : items) {
        popup.addAction(item, [this, idx, item]() {
            model()->setData(idx, item, Qt::EditRole);
        });
    }
    QRect cellRect = visualRect(idx);
    popup.exec(viewport()->mapToGlobal(cellRect.bottomLeft()));
}

void SpreadsheetView::showPicklistPopup(const QModelIndex& index) {
    if (!m_spreadsheet || !index.isValid()) return;

    auto cell = m_spreadsheet->getCellIfExists(logicalRow(index), index.column());
    if (!cell) return;

    auto* rule = const_cast<Spreadsheet::DataValidationRule*>(
        m_spreadsheet->getValidationAt(logicalRow(index), index.column()));
    QStringList options;
    if (rule) {
        if (!rule->listSourceRange.isEmpty()) {
            // Dynamic resolution from cell range (Excel-like live binding)
            options = resolvePicklistFromRange(rule->listSourceRange);
            // Apply ignore-blanks
            if (rule->listIgnoreBlanks)
                options.removeAll("");
            // Apply sort preference
            if (rule->listSortMode == 1) {
                std::sort(options.begin(), options.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) < 0;
                });
            } else if (rule->listSortMode == 2) {
                std::sort(options.begin(), options.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) > 0;
                });
            }
            rule->listItems = options;  // Update cache for tag rendering & validation
            rule->listItemColors.clear();
        } else {
            options = rule->listItems;
        }
    }
    if (options.isEmpty()) return;

    // Current selected items
    QString currentVal = cell->getValue().toString();
    QStringList selected = currentVal.split('|', Qt::SkipEmptyParts);
    QSet<QString> selectedSet;
    for (const auto& s : selected) selectedSet.insert(s.trimmed());

    static const QColor tagBg[] = {
        QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
        QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
        QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
    };
    static const QColor tagFg[] = {
        QColor("#1E40AF"), QColor("#9D174D"), QColor("#5B21B6"), QColor("#065F46"),
        QColor("#92400E"), QColor("#9F1239"), QColor("#155E75"), QColor("#991B1B"),
        QColor("#374151"), QColor("#3F6212"), QColor("#3730A3"), QColor("#831843")
    };
    QStringList optionColors = rule ? rule->listItemColors : QStringList();

    int row = index.row(), col = index.column();

    // Save initial value for ESC undo
    QString initialVal = currentVal;

    // --- Popup ---
    QFrame* popup = new QFrame(this, Qt::Popup | Qt::FramelessWindowHint);
    popup->setAttribute(Qt::WA_DeleteOnClose);
    connect(popup, &QObject::destroyed, this, [this]() { m_picklistPopupOpen = false; });
    popup->setFixedWidth(220);
    popup->setObjectName("plPopup");
    popup->setStyleSheet(
        "QFrame#plPopup { background: white; border: 1px solid #D1D5DB; border-radius: 8px; }");

    QVBoxLayout* layout = new QVBoxLayout(popup);
    layout->setContentsMargins(6, 6, 6, 6);
    layout->setSpacing(0);

    // --- Search box ---
    QLineEdit* searchBox = new QLineEdit(popup);
    searchBox->setPlaceholderText("Search");
    searchBox->setFixedHeight(28);
    searchBox->setStyleSheet(
        "QLineEdit { border: none; border-bottom: 2px solid #4ADE80; border-radius: 0; "
        "padding: 4px 6px; font-size: 12px; color: #1F2937; background: transparent; }"
        "QLineEdit:focus { border-bottom-color: #22C55E; }");
    layout->addWidget(searchBox);
    layout->addSpacing(6);

    // --- Scrollable option area ---
    QWidget* optionsContainer = new QWidget(popup);
    QVBoxLayout* optLay = new QVBoxLayout(optionsContainer);
    optLay->setContentsMargins(0, 0, 0, 0);
    optLay->setSpacing(2);

    // --- Option rows: checkbox + colored tag pill ---
    struct OptionRowWidgets { QWidget* row; QCheckBox* cb; };
    auto optionWidgets = std::make_shared<QList<OptionRowWidgets>>();

    for (int i = 0; i < options.size(); ++i) {
        QColor bg = (i < optionColors.size() && !optionColors[i].isEmpty())
                     ? QColor(optionColors[i]) : tagBg[i % 12];
        QColor fg;
        if (i < optionColors.size() && !optionColors[i].isEmpty()) {
            fg = (bg.lightness() > 140) ? bg.darker(300) : QColor(Qt::white);
        } else {
            fg = tagFg[i % 12];
        }
        bool isChecked = selectedSet.contains(options[i]);

        QWidget* rowWidget = new QWidget(optionsContainer);
        rowWidget->setCursor(Qt::PointingHandCursor);
        rowWidget->setFixedHeight(30);
        rowWidget->setStyleSheet(
            "QWidget { background: transparent; border-radius: 4px; }");
        QHBoxLayout* rowLo = new QHBoxLayout(rowWidget);
        rowLo->setContentsMargins(4, 2, 6, 2);
        rowLo->setSpacing(8);

        // Checkbox
        QCheckBox* cb = new QCheckBox(rowWidget);
        cb->setChecked(isChecked);
        cb->setFixedSize(16, 16);
        // Create checkmark image once (static)
        static QString checkImgPath;
        if (checkImgPath.isEmpty()) {
            QPixmap pix(14, 14);
            pix.fill(Qt::transparent);
            QPainter pp(&pix);
            pp.setRenderHint(QPainter::Antialiasing);
            pp.setPen(QPen(QColor("#2563EB"), 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
            pp.drawLine(QPointF(3, 7), QPointF(5.5, 10));
            pp.drawLine(QPointF(5.5, 10), QPointF(11, 3.5));
            pp.end();
            checkImgPath = QDir::temp().filePath("nexel_picklist_check.png");
            pix.save(checkImgPath);
        }
        cb->setStyleSheet(QString(
            "QCheckBox { spacing: 0; }"
            "QCheckBox::indicator { width: 14px; height: 14px; border: 1.5px solid #D1D5DB; "
            "border-radius: 3px; background: white; }"
            "QCheckBox::indicator:checked { background: white; border-color: #2563EB; "
            "image: url(%1); }"
            "QCheckBox::indicator:hover { border-color: #9CA3AF; }").arg(checkImgPath));
        rowLo->addWidget(cb);

        // Tag pill (colored rounded rect with text)
        QLabel* tagLabel = new QLabel(options[i], rowWidget);
        tagLabel->setFixedHeight(22);
        tagLabel->setStyleSheet(QString(
            "QLabel { background: %1; color: %2; border-radius: 11px; "
            "padding: 2px 10px; font-size: 11px; font-weight: 500; border: none; }")
            .arg(bg.name(), fg.name()));
        tagLabel->setAttribute(Qt::WA_TransparentForMouseEvents);
        rowLo->addWidget(tagLabel);
        rowLo->addStretch();

        optLay->addWidget(rowWidget);
        optionWidgets->append({rowWidget, cb});

        // Click anywhere on row to toggle checkbox
        connect(cb, &QCheckBox::toggled, rowWidget, [](bool) {});
    }

    optLay->addStretch();
    layout->addWidget(optionsContainer, 1);

    // Lambda to commit current checked state to the cell
    auto commitToCell = [this, optionWidgets, options, row, col]() {
        QStringList sel;
        for (int j = 0; j < optionWidgets->size(); ++j) {
            if ((*optionWidgets)[j].cb->isChecked()) sel.append(options[j]);
        }
        CellAddress addr(row, col);
        m_spreadsheet->getCell(addr)->setValue(QVariant(sel.join('|')));
        if (m_model) {
            QModelIndex idx = m_model->index(row, col);
            emit m_model->dataChanged(idx, idx);
        }
    };

    // Connect each checkbox to commit on the fly
    for (int i = 0; i < optionWidgets->size(); ++i) {
        connect((*optionWidgets)[i].cb, &QCheckBox::toggled, this, [commitToCell](bool) {
            commitToCell();
        });
    }

    // Search filtering
    connect(searchBox, &QLineEdit::textChanged, this,
        [optionWidgets, options](const QString& text) {
        QString filter = text.trimmed().toLower();
        for (int i = 0; i < optionWidgets->size(); ++i) {
            bool visible = filter.isEmpty() || options[i].toLower().contains(filter);
            (*optionWidgets)[i].row->setVisible(visible);
        }
    });

    // --- Separator ---
    QFrame* sep = new QFrame(popup);
    sep->setFixedHeight(1);
    sep->setStyleSheet("QFrame { background: #E5E7EB; border: none; }");
    layout->addSpacing(4);
    layout->addWidget(sep);
    layout->addSpacing(2);

    // --- Bottom row: Edit + OK ---
    QHBoxLayout* bottomRow = new QHBoxLayout();
    bottomRow->setContentsMargins(4, 0, 4, 2);

    QPushButton* editBtn = new QPushButton(QString::fromUtf8("\u270E Edit"), popup);
    editBtn->setCursor(Qt::PointingHandCursor);
    editBtn->setFixedHeight(22);
    editBtn->setStyleSheet(
        "QPushButton { background: transparent; color: #6B7280; border: none; "
        "font-size: 11px; border-radius: 4px; padding: 0 6px; }"
        "QPushButton:hover { color: #2563EB; background: #F3F4F6; }");
    bottomRow->addWidget(editBtn);
    bottomRow->addStretch();

    QPushButton* okBtn = new QPushButton("OK", popup);
    okBtn->setCursor(Qt::PointingHandCursor);
    okBtn->setFixedHeight(22);
    okBtn->setStyleSheet(
        "QPushButton { background: #2563EB; color: white; border: none; "
        "border-radius: 4px; padding: 0 14px; font-size: 11px; font-weight: 500; }"
        "QPushButton:hover { background: #1D4ED8; }");
    bottomRow->addWidget(okBtn);
    layout->addLayout(bottomRow);

    connect(okBtn, &QPushButton::clicked, popup, &QFrame::close);

    connect(editBtn, &QPushButton::clicked, this, [this, popup, row, col]() {
        popup->close();
        openPicklistManageDialog(row, col);
    });

    // --- ESC key: clear cell value and close ---
    QShortcut* escShortcut = new QShortcut(QKeySequence(Qt::Key_Escape), popup);
    connect(escShortcut, &QShortcut::activated, this,
        [this, popup, row, col, initialVal]() {
        // Clear the cell value
        CellAddress addr(row, col);
        CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);
        m_spreadsheet->getCell(addr)->setValue(QVariant(QString()));
        CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<CellEditCommand>(before, after));
        if (m_model) {
            QModelIndex idx = m_model->index(row, col);
            emit m_model->dataChanged(idx, idx);
        }
        popup->close();
    });

    // When popup closes (click outside), push undo for any changes
    connect(popup, &QObject::destroyed, this,
        [this, row, col, initialVal]() {
        // Check if value changed from initial
        auto cell = m_spreadsheet->getCellIfExists(row, col);
        if (!cell) return;
        QString finalVal = cell->getValue().toString();
        if (finalVal != initialVal) {
            CellAddress addr(row, col);
            CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
            // Temporarily set back to capture "before"
            cell->setValue(QVariant(initialVal));
            CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);
            cell->setValue(QVariant(finalVal));
            m_spreadsheet->getUndoManager().pushCommand(
                std::make_unique<CellEditCommand>(before, after));
        }
    });

    // --- Position ---
    QRect cellRect = visualRect(index);
    QPoint pos = viewport()->mapToGlobal(QPoint(cellRect.left(), cellRect.bottom() + 2));
    popup->adjustSize();
    QScreen* screen = this->screen();
    if (screen) {
        QRect sr = screen->availableGeometry();
        if (pos.y() + popup->height() > sr.bottom())
            pos.setY(viewport()->mapToGlobal(QPoint(0, cellRect.top())).y() - popup->height() - 2);
        if (pos.x() + popup->width() > sr.right())
            pos.setX(sr.right() - popup->width());
    }
    popup->move(pos);
    popup->show();
}

// ===== Manage Picklist Dialog =====
void SpreadsheetView::openPicklistManageDialog(int row, int col) {
    if (!m_spreadsheet) return;
    auto* rule = const_cast<Spreadsheet::DataValidationRule*>(
        m_spreadsheet->getValidationAt(row, col));
    if (!rule) return;

    static const QColor defaultTagBg[] = {
        QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
        QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
        QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
    };

    bool isRangeMode = !rule->listSourceRange.isEmpty();

    // If range mode, resolve fresh values with sort/blank preferences
    if (isRangeMode) {
        QStringList resolved = resolvePicklistFromRange(rule->listSourceRange);
        if (rule->listIgnoreBlanks) resolved.removeAll("");
        if (rule->listSortMode == 1)
            std::sort(resolved.begin(), resolved.end(), [](const QString& a, const QString& b) {
                return a.compare(b, Qt::CaseInsensitive) < 0; });
        else if (rule->listSortMode == 2)
            std::sort(resolved.begin(), resolved.end(), [](const QString& a, const QString& b) {
                return a.compare(b, Qt::CaseInsensitive) > 0; });
        rule->listItems = resolved;
        rule->listItemColors.clear();
    }

    QDialog dlg(this);
    dlg.setWindowTitle("Manage Picklist");
    dlg.setFixedWidth(420);
    dlg.setStyleSheet(
        "QDialog { background: white; }"
        "QLineEdit { border: 1px solid #D1D5DB; border-radius: 6px; padding: 6px 8px; "
        "font-size: 13px; color: #1F2937; }"
        "QLineEdit:focus { border-color: #2563EB; }");

    QVBoxLayout* lo = new QVBoxLayout(&dlg);
    lo->setContentsMargins(16, 12, 16, 12);
    lo->setSpacing(8);

    // === Mode toggle: Manual vs Cell Range ===
    QRadioButton* manualRadio = new QRadioButton("Manual options", &dlg);
    QRadioButton* rangeRadio = new QRadioButton("From cell range", &dlg);
    if (isRangeMode) rangeRadio->setChecked(true);
    else manualRadio->setChecked(true);
    manualRadio->setStyleSheet("QRadioButton { font-size: 12px; color: #374151; }");
    rangeRadio->setStyleSheet("QRadioButton { font-size: 12px; color: #374151; }");
    QHBoxLayout* modeRow = new QHBoxLayout();
    modeRow->setSpacing(12);
    modeRow->addWidget(manualRadio);
    modeRow->addWidget(rangeRadio);
    modeRow->addStretch();
    lo->addLayout(modeRow);

    // === Range panel ===
    QWidget* rangePanel = new QWidget(&dlg);
    QVBoxLayout* rangePanelLo = new QVBoxLayout(rangePanel);
    rangePanelLo->setContentsMargins(0, 0, 0, 0);
    rangePanelLo->setSpacing(6);

    // Sheet combo + Range edit row
    QHBoxLayout* rangeLo = new QHBoxLayout();
    rangeLo->setSpacing(6);

    QComboBox* sheetCombo = new QComboBox(rangePanel);
    if (m_allSheets.empty() && m_spreadsheet) {
        sheetCombo->addItem(m_spreadsheet->getSheetName());
    } else {
        for (const auto& s : m_allSheets)
            sheetCombo->addItem(s->getSheetName());
    }
    sheetCombo->setFixedWidth(100);
    sheetCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D1D5DB; border-radius: 6px; padding: 4px 8px; font-size: 12px; }"
        "QComboBox:focus { border-color: #2563EB; }");
    rangeLo->addWidget(sheetCombo);

    QLineEdit* rangeEdit = new QLineEdit(rangePanel);
    rangeEdit->setPlaceholderText("e.g. A1:A10");
    rangeLo->addWidget(rangeEdit, 1);
    rangePanelLo->addLayout(rangeLo);

    // Pre-fill from existing listSourceRange
    if (isRangeMode) {
        int excl = rule->listSourceRange.indexOf('!');
        if (excl >= 0) {
            int idx = sheetCombo->findText(rule->listSourceRange.left(excl));
            if (idx >= 0) sheetCombo->setCurrentIndex(idx);
            rangeEdit->setText(rule->listSourceRange.mid(excl + 1));
        }
    }

    // Options row: Ignore blanks + Sort
    QHBoxLayout* mgRangeOptsLo = new QHBoxLayout();
    mgRangeOptsLo->setContentsMargins(0, 2, 0, 0);
    mgRangeOptsLo->setSpacing(10);

    QCheckBox* mgIgnoreBlanks = new QCheckBox("Ignore blank cells", rangePanel);
    mgIgnoreBlanks->setChecked(rule->listIgnoreBlanks);
    mgIgnoreBlanks->setStyleSheet("QCheckBox { font-size: 11px; color: #6B7280; }");
    mgRangeOptsLo->addWidget(mgIgnoreBlanks);

    QLabel* mgSortLabel = new QLabel("Sort:", rangePanel);
    mgSortLabel->setStyleSheet("QLabel { font-size: 11px; color: #6B7280; }");
    mgRangeOptsLo->addWidget(mgSortLabel);

    QComboBox* mgSortCombo = new QComboBox(rangePanel);
    mgSortCombo->addItems({"None", "A \u2192 Z", "Z \u2192 A"});
    mgSortCombo->setCurrentIndex(rule->listSortMode);
    mgSortCombo->setFixedWidth(80);
    mgSortCombo->setStyleSheet(
        "QComboBox { border: 1px solid #D1D5DB; border-radius: 4px; padding: 2px 6px; font-size: 11px; }"
        "QComboBox:focus { border-color: #2563EB; }");
    mgRangeOptsLo->addWidget(mgSortCombo);
    mgRangeOptsLo->addStretch();
    rangePanelLo->addLayout(mgRangeOptsLo);

    // Preview of resolved values
    QLabel* previewLabel = new QLabel(rangePanel);
    previewLabel->setWordWrap(true);
    previewLabel->setStyleSheet("QLabel { font-size: 11px; color: #6B7280; padding: 4px 0; }");
    auto updatePreview = [&]() {
        QString src = sheetCombo->currentText() + "!" + rangeEdit->text().trimmed();
        QStringList vals = resolvePicklistFromRange(src);
        if (mgIgnoreBlanks->isChecked()) vals.removeAll("");
        int sm = mgSortCombo->currentIndex();
        if (sm == 1) std::sort(vals.begin(), vals.end(), [](const QString& a, const QString& b) {
            return a.compare(b, Qt::CaseInsensitive) < 0; });
        else if (sm == 2) std::sort(vals.begin(), vals.end(), [](const QString& a, const QString& b) {
            return a.compare(b, Qt::CaseInsensitive) > 0; });
        if (vals.isEmpty()) {
            previewLabel->setText("No values found in range");
        } else {
            QString preview = vals.mid(0, 10).join(", ");
            if (vals.size() > 10) preview += QString(" ... (%1 total)").arg(vals.size());
            previewLabel->setText(QString("Values: %1").arg(preview));
        }
    };
    if (isRangeMode) updatePreview();
    else previewLabel->setText("Enter a range to see values");
    rangePanelLo->addWidget(previewLabel);

    // Update preview when range/sort/blanks change
    QObject::connect(rangeEdit, &QLineEdit::textChanged, &dlg, [&]() { updatePreview(); });
    QObject::connect(sheetCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
        &dlg, [&]() { updatePreview(); });
    QObject::connect(mgSortCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
        &dlg, [&]() { updatePreview(); });
    QObject::connect(mgIgnoreBlanks, &QCheckBox::toggled, &dlg, [&]() { updatePreview(); });

    rangePanel->setVisible(isRangeMode);
    lo->addWidget(rangePanel);

    // === Manual panel ===
    QWidget* manualPanel = new QWidget(&dlg);
    QVBoxLayout* manualLo = new QVBoxLayout(manualPanel);
    manualLo->setContentsMargins(0, 0, 0, 0);
    manualLo->setSpacing(6);

    QLabel* sub = new QLabel("Set label and color for each option:", manualPanel);
    sub->setStyleSheet("QLabel { font-size: 12px; color: #6B7280; }");
    manualLo->addWidget(sub);

    QScrollArea* scroll = new QScrollArea(manualPanel);
    scroll->setWidgetResizable(true);
    scroll->setFrameShape(QFrame::NoFrame);
    scroll->setMinimumHeight(180);
    scroll->setStyleSheet("QScrollArea { background: transparent; border: none; }");

    QWidget* listWidget = new QWidget(scroll);
    QVBoxLayout* listLo = new QVBoxLayout(listWidget);
    listLo->setContentsMargins(0, 0, 0, 0);
    listLo->setSpacing(6);

    struct OptionRow { QLineEdit* text; QPushButton* colorBtn; QString color; };
    auto optionRows = std::make_shared<QList<OptionRow>>();

    while (rule->listItemColors.size() < rule->listItems.size())
        rule->listItemColors.append("");

    auto addOptionRow = [&](const QString& text, const QString& colorStr, int idx) {
        QHBoxLayout* rowLo = new QHBoxLayout();
        rowLo->setSpacing(8);

        QLineEdit* lineEdit = new QLineEdit(text, listWidget);
        lineEdit->setPlaceholderText("Option name...");
        rowLo->addWidget(lineEdit, 1);

        QPushButton* colorBtn = new QPushButton(listWidget);
        colorBtn->setFixedSize(28, 28);
        colorBtn->setCursor(Qt::PointingHandCursor);
        QColor displayColor = colorStr.isEmpty() ? defaultTagBg[idx % 12] : QColor(colorStr);
        colorBtn->setStyleSheet(QString(
            "QPushButton { background: %1; border: 1px solid #D1D5DB; border-radius: 6px; }"
            "QPushButton:hover { border-color: #2563EB; }").arg(displayColor.name()));
        rowLo->addWidget(colorBtn);

        QPushButton* removeBtn = new QPushButton(QString::fromUtf8("\u2715"), listWidget);
        removeBtn->setFixedSize(24, 24);
        removeBtn->setCursor(Qt::PointingHandCursor);
        removeBtn->setStyleSheet(
            "QPushButton { background: transparent; border: none; color: #9CA3AF; font-size: 13px; }"
            "QPushButton:hover { color: #EF4444; }");
        rowLo->addWidget(removeBtn);

        listLo->addLayout(rowLo);

        OptionRow opt{lineEdit, colorBtn, colorStr};
        optionRows->append(opt);
        int rowIdx = optionRows->size() - 1;

        QObject::connect(colorBtn, &QPushButton::clicked, &dlg, [optionRows, colorBtn, rowIdx, &dlg]() {
            QColor cur = (*optionRows)[rowIdx].color.isEmpty()
                ? colorBtn->palette().button().color() : QColor((*optionRows)[rowIdx].color);
            QColor picked = QColorDialog::getColor(cur, &dlg, "Option Color");
            if (picked.isValid()) {
                (*optionRows)[rowIdx].color = picked.name();
                colorBtn->setStyleSheet(QString(
                    "QPushButton { background: %1; border: 1px solid #D1D5DB; border-radius: 6px; }"
                    "QPushButton:hover { border-color: #2563EB; }").arg(picked.name()));
            }
        });

        QObject::connect(removeBtn, &QPushButton::clicked, &dlg, [optionRows, rowIdx, rowLo, lineEdit, colorBtn, removeBtn]() {
            lineEdit->hide(); colorBtn->hide(); removeBtn->hide();
            (*optionRows)[rowIdx].text = nullptr;
        });
    };

    // Only populate manual options if NOT in range mode
    if (!isRangeMode) {
        for (int i = 0; i < rule->listItems.size(); ++i) {
            QString colorStr = (i < rule->listItemColors.size()) ? rule->listItemColors[i] : "";
            addOptionRow(rule->listItems[i], colorStr, i);
        }
    } else {
        // Start with 3 empty rows for manual mode (shown if user switches)
        for (int i = 0; i < 3; ++i)
            addOptionRow(QString("Option %1").arg(i + 1), "", i);
    }

    listLo->addStretch();
    scroll->setWidget(listWidget);
    manualLo->addWidget(scroll, 1);

    QPushButton* addBtn = new QPushButton("+ Add Option", manualPanel);
    addBtn->setCursor(Qt::PointingHandCursor);
    addBtn->setFixedHeight(30);
    addBtn->setStyleSheet(
        "QPushButton { background: transparent; color: #2563EB; border: 1px dashed #93C5FD; "
        "border-radius: 6px; font-size: 12px; font-weight: 500; }"
        "QPushButton:hover { background: #EFF6FF; }");
    manualLo->addWidget(addBtn);
    connect(addBtn, &QPushButton::clicked, &dlg, [&, addOptionRow]() {
        int idx = optionRows->size();
        addOptionRow("", "", idx);
        listWidget->adjustSize();
    });

    manualPanel->setVisible(!isRangeMode);
    lo->addWidget(manualPanel, 1);

    // Toggle panels
    QObject::connect(manualRadio, &QRadioButton::toggled, &dlg,
        [manualPanel, rangePanel](bool checked) {
        manualPanel->setVisible(checked);
        rangePanel->setVisible(!checked);
    });

    lo->addSpacing(4);
    QHBoxLayout* btnRow = new QHBoxLayout();
    btnRow->addStretch();

    QPushButton* cancelBtn = new QPushButton("Cancel", &dlg);
    cancelBtn->setFixedHeight(34);
    cancelBtn->setCursor(Qt::PointingHandCursor);
    cancelBtn->setStyleSheet(
        "QPushButton { background: white; color: #374151; border: 1px solid #D1D5DB; "
        "border-radius: 6px; padding: 0 20px; font-size: 13px; font-weight: 500; }"
        "QPushButton:hover { background: #F9FAFB; border-color: #9CA3AF; }");
    btnRow->addWidget(cancelBtn);

    QPushButton* saveBtn = new QPushButton("Save", &dlg);
    saveBtn->setFixedHeight(34);
    saveBtn->setCursor(Qt::PointingHandCursor);
    saveBtn->setStyleSheet(
        "QPushButton { background: #2563EB; color: white; border: none; "
        "border-radius: 6px; padding: 0 24px; font-size: 13px; font-weight: 600; }"
        "QPushButton:hover { background: #1D4ED8; }");
    btnRow->addWidget(saveBtn);
    lo->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &dlg, &QDialog::reject);
    connect(saveBtn, &QPushButton::clicked, &dlg, &QDialog::accept);

    int result = dlg.exec();

    if (result == QDialog::Accepted) {
        if (rangeRadio->isChecked()) {
            // Range mode: update source reference, sort, ignore-blanks
            QString newSource = sheetCombo->currentText() + "!" + rangeEdit->text().trimmed();
            rule->listSourceRange = newSource;
            rule->listSortMode = mgSortCombo->currentIndex();
            rule->listIgnoreBlanks = mgIgnoreBlanks->isChecked();
            QStringList resolved = resolvePicklistFromRange(newSource);
            if (rule->listIgnoreBlanks) resolved.removeAll("");
            if (rule->listSortMode == 1)
                std::sort(resolved.begin(), resolved.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) < 0; });
            else if (rule->listSortMode == 2)
                std::sort(resolved.begin(), resolved.end(), [](const QString& a, const QString& b) {
                    return a.compare(b, Qt::CaseInsensitive) > 0; });
            rule->listItems = resolved;
            rule->listItemColors.clear();
        } else {
            // Manual mode: clear source reference, use manual entries
            rule->listSourceRange.clear();
            QStringList newOpts;
            QStringList newColors;
            for (const auto& opt : *optionRows) {
                if (!opt.text) continue;
                QString t = opt.text->text().trimmed();
                if (!t.isEmpty()) {
                    newOpts.append(t);
                    newColors.append(opt.color);
                }
            }
            rule->listItems = newOpts;
            rule->listItemColors = newColors;
        }
        if (m_model) {
            emit m_model->dataChanged(
                m_model->index(rule->range.getStart().row, rule->range.getStart().col),
                m_model->index(rule->range.getEnd().row, rule->range.getEnd().col));
        }
    }
}

// ===== Sheet-wide Picklist Manager Dialog =====
void SpreadsheetView::openPicklistManagerDialog() {
    if (!m_spreadsheet) return;

    // Collect all picklist validation rules
    const auto& allRules = m_spreadsheet->getValidationRules();
    std::vector<int> picklistIndices;
    for (int i = 0; i < (int)allRules.size(); ++i) {
        if (allRules[i].type == Spreadsheet::DataValidationRule::List) {
            picklistIndices.push_back(i);
        }
    }

    QDialog dlg(this);
    dlg.setWindowTitle("Manage Picklists");
    dlg.setMinimumSize(480, 360);
    dlg.setStyleSheet(
        "QDialog { background: white; }"
        "QLabel { background: transparent; border: none; }");

    QVBoxLayout* mainLo = new QVBoxLayout(&dlg);
    mainLo->setContentsMargins(20, 20, 20, 16);
    mainLo->setSpacing(14);

    QLabel* title = new QLabel("Picklists in this Sheet", &dlg);
    title->setStyleSheet("QLabel { font-size: 16px; font-weight: 600; color: #111827; }");
    mainLo->addWidget(title);

    if (picklistIndices.empty()) {
        QLabel* emptyLbl = new QLabel("No picklists found in this sheet.", &dlg);
        emptyLbl->setStyleSheet("QLabel { font-size: 13px; color: #6B7280; padding: 20px 0; }");
        emptyLbl->setAlignment(Qt::AlignCenter);
        mainLo->addWidget(emptyLbl, 1);
    } else {
        // Scrollable list of picklist cards
        QScrollArea* scroll = new QScrollArea(&dlg);
        scroll->setWidgetResizable(true);
        scroll->setFrameShape(QFrame::NoFrame);
        scroll->setStyleSheet("QScrollArea { background: transparent; border: none; }");

        QWidget* listWidget = new QWidget(scroll);
        QVBoxLayout* listLo = new QVBoxLayout(listWidget);
        listLo->setContentsMargins(0, 0, 0, 0);
        listLo->setSpacing(8);

        static const QColor tagBg[] = {
            QColor("#DBEAFE"), QColor("#FCE7F3"), QColor("#EDE9FE"), QColor("#D1FAE5"),
            QColor("#FEF3C7"), QColor("#FFE4E6"), QColor("#CFFAFE"), QColor("#FEE2E2"),
            QColor("#F3F4F6"), QColor("#ECFCCB"), QColor("#E0E7FF"), QColor("#FDF2F8")
        };

        for (int pi : picklistIndices) {
            const auto& rule = allRules[pi];
            QString rangeStr = rule.range.toString();
            int ruleIndex = pi; // capture for lambda

            // Card frame
            QFrame* card = new QFrame(listWidget);
            card->setStyleSheet(
                "QFrame { background: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 8px; }");

            QVBoxLayout* cardLo = new QVBoxLayout(card);
            cardLo->setContentsMargins(14, 10, 14, 10);
            cardLo->setSpacing(6);

            // Row 1: Range + action buttons
            QHBoxLayout* headerRow = new QHBoxLayout();
            headerRow->setSpacing(8);

            QLabel* rangeLbl = new QLabel(rangeStr, card);
            rangeLbl->setStyleSheet(
                "QLabel { font-size: 13px; font-weight: 600; color: #1F2937; "
                "background: #E0E7FF; border-radius: 4px; padding: 2px 8px; }");
            headerRow->addWidget(rangeLbl);

            QLabel* countLbl = new QLabel(
                QString("%1 option%2").arg(rule.listItems.size()).arg(rule.listItems.size() != 1 ? "s" : ""), card);
            countLbl->setStyleSheet("QLabel { font-size: 11px; color: #6B7280; }");
            headerRow->addWidget(countLbl);

            headerRow->addStretch();

            QPushButton* editBtn = new QPushButton("Edit", card);
            editBtn->setCursor(Qt::PointingHandCursor);
            editBtn->setFixedHeight(26);
            editBtn->setStyleSheet(
                "QPushButton { background: white; color: #374151; border: 1px solid #D1D5DB; "
                "border-radius: 5px; padding: 0 12px; font-size: 11px; font-weight: 500; }"
                "QPushButton:hover { background: #F3F4F6; border-color: #9CA3AF; }");
            headerRow->addWidget(editBtn);

            QPushButton* delBtn = new QPushButton("Delete", card);
            delBtn->setCursor(Qt::PointingHandCursor);
            delBtn->setFixedHeight(26);
            delBtn->setStyleSheet(
                "QPushButton { background: white; color: #DC2626; border: 1px solid #FCA5A5; "
                "border-radius: 5px; padding: 0 12px; font-size: 11px; font-weight: 500; }"
                "QPushButton:hover { background: #FEF2F2; border-color: #DC2626; }");
            headerRow->addWidget(delBtn);

            cardLo->addLayout(headerRow);

            // Row 2: Tag pills showing the options
            QHBoxLayout* tagsRow = new QHBoxLayout();
            tagsRow->setSpacing(4);
            for (int t = 0; t < rule.listItems.size() && t < 8; ++t) {
                QLabel* tagLbl = new QLabel(rule.listItems[t], card);
                QColor bg = tagBg[t % 12];
                tagLbl->setStyleSheet(QString(
                    "QLabel { background: %1; color: #374151; font-size: 10px; font-weight: 500; "
                    "border-radius: 8px; padding: 1px 8px; border: none; }").arg(bg.name()));
                tagsRow->addWidget(tagLbl);
            }
            if (rule.listItems.size() > 8) {
                QLabel* moreLbl = new QLabel(QString("+%1").arg(rule.listItems.size() - 8), card);
                moreLbl->setStyleSheet("QLabel { font-size: 10px; color: #6B7280; }");
                tagsRow->addWidget(moreLbl);
            }
            tagsRow->addStretch();
            cardLo->addLayout(tagsRow);

            listLo->addWidget(card);

            // Edit button: open per-cell manage dialog using first cell in range
            connect(editBtn, &QPushButton::clicked, &dlg, [this, &dlg, ruleIndex]() {
                auto& rules = m_spreadsheet->getValidationRules();
                if (ruleIndex >= 0 && ruleIndex < (int)rules.size()) {
                    auto& r = rules[ruleIndex];
                    dlg.close();
                    openPicklistManageDialog(r.range.getStart().row, r.range.getStart().col);
                }
            });

            // Delete button: remove rule + clear picklist format from cells
            connect(delBtn, &QPushButton::clicked, &dlg, [this, &dlg, ruleIndex]() {
                auto& rules = m_spreadsheet->getValidationRules();
                if (ruleIndex < 0 || ruleIndex >= (int)rules.size()) return;

                CellRange range = rules[ruleIndex].range;

                // Clear Picklist format + value from all cells in range
                std::vector<CellSnapshot> before, after;
                auto cells = range.getCells();
                for (const auto& addr : cells) {
                    auto cell = m_spreadsheet->getCellIfExists(addr.row, addr.col);
                    if (cell && cell->getStyle().numberFormat == "Picklist") {
                        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
                        CellStyle st = cell->getStyle();
                        st.numberFormat = "General";
                        cell->setStyle(st);
                        cell->setValue(QVariant());
                        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
                    }
                }
                if (!before.empty()) {
                    m_spreadsheet->getUndoManager().pushCommand(
                        std::make_unique<MultiCellEditCommand>(before, after, "Delete Picklist"));
                }

                m_spreadsheet->removeValidationRule(ruleIndex);

                if (m_model) m_model->resetModel();
                dlg.close();
            });
        }

        listLo->addStretch();
        scroll->setWidget(listWidget);
        mainLo->addWidget(scroll, 1);
    }

    // Close button
    QHBoxLayout* bottomRow = new QHBoxLayout();
    bottomRow->addStretch();
    QPushButton* closeBtn = new QPushButton("Close", &dlg);
    closeBtn->setFixedHeight(34);
    closeBtn->setCursor(Qt::PointingHandCursor);
    closeBtn->setStyleSheet(
        "QPushButton { background: #2563EB; color: white; border: none; "
        "border-radius: 6px; padding: 0 24px; font-size: 13px; font-weight: 600; }"
        "QPushButton:hover { background: #1D4ED8; }");
    connect(closeBtn, &QPushButton::clicked, &dlg, &QDialog::accept);
    bottomRow->addWidget(closeBtn);
    mainLo->addLayout(bottomRow);

    dlg.exec();
}

// ============== Trace Precedents/Dependents ==============

void SpreadsheetView::tracePrecedents() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    m_traceCell = CellAddress(logicalRow(current), current.column());
    const auto& depGraph = m_spreadsheet->getDependencyGraph();
    m_tracedCells = depGraph.getDependencies(m_traceCell);
    m_showPrecedents = true;
    m_showDependents = false;
    viewport()->update();
}

void SpreadsheetView::traceDependents() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    m_traceCell = CellAddress(logicalRow(current), current.column());
    const auto& depGraph = m_spreadsheet->getDependencyGraph();
    m_tracedCells = depGraph.getDependents(m_traceCell);
    m_showDependents = true;
    m_showPrecedents = false;
    viewport()->update();
}

void SpreadsheetView::clearTraceArrows() {
    m_showPrecedents = false;
    m_showDependents = false;
    m_tracedCells.clear();
    viewport()->update();
}

void SpreadsheetView::drawTraceArrows(QPainter& painter) {
    painter.setRenderHint(QPainter::Antialiasing, true);

    // Blue for precedents (arrows point FROM precedent TO traced cell)
    // Red for dependents (arrows point FROM traced cell TO dependent)
    QColor arrowColor = m_showPrecedents ? QColor(50, 100, 220, 180) : QColor(220, 50, 50, 180);

    QModelIndex traceCellIdx = m_model->index(m_traceCell.row, m_traceCell.col);
    QRect traceCellRect = visualRect(traceCellIdx);
    if (traceCellRect.isNull()) return;

    QPointF traceCellCenter = traceCellRect.center();

    for (const auto& addr : m_tracedCells) {
        QModelIndex idx = m_model->index(addr.row, addr.col);
        QRect cellRect = visualRect(idx);
        if (cellRect.isNull()) continue;

        QPointF cellCenter = cellRect.center();

        QPointF from, to;
        if (m_showPrecedents) {
            from = cellCenter;
            to = traceCellCenter;
        } else {
            from = traceCellCenter;
            to = cellCenter;
        }

        // Draw the line
        QPen linePen(arrowColor, 2.0);
        painter.setPen(linePen);
        painter.setBrush(Qt::NoBrush);
        painter.drawLine(from, to);

        // Draw arrowhead at destination
        double angle = std::atan2(to.y() - from.y(), to.x() - from.x());
        double arrowLen = 10.0;
        double arrowAngle = 0.45; // ~25 degrees

        QPointF p1(to.x() - arrowLen * std::cos(angle - arrowAngle),
                   to.y() - arrowLen * std::sin(angle - arrowAngle));
        QPointF p2(to.x() - arrowLen * std::cos(angle + arrowAngle),
                   to.y() - arrowLen * std::sin(angle + arrowAngle));

        painter.setBrush(arrowColor);
        painter.setPen(Qt::NoPen);
        QPolygonF arrowHead;
        arrowHead << to << p1 << p2;
        painter.drawPolygon(arrowHead);

        // Draw a small dot at the source cell center
        painter.setBrush(arrowColor);
        painter.drawEllipse(from, 3.0, 3.0);
    }
}

// ============== Row/Column Grouping (Outline) ==============

void SpreadsheetView::groupSelectedRows() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        if (lr < minRow) minRow = lr;
        if (lr > maxRow) maxRow = lr;
    }
    m_spreadsheet->groupRows(minRow, maxRow);
    viewport()->update();
    update();
}

void SpreadsheetView::ungroupSelectedRows() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0;
    for (const auto& idx : selected) {
        int lr = logicalRow(idx);
        if (lr < minRow) minRow = lr;
        if (lr > maxRow) maxRow = lr;
    }
    m_spreadsheet->ungroupRows(minRow, maxRow);
    viewport()->update();
    update();
}

void SpreadsheetView::groupSelectedColumns() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int c = idx.column();
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
    }
    m_spreadsheet->groupColumns(minCol, maxCol);
    viewport()->update();
    update();
}

void SpreadsheetView::ungroupSelectedColumns() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int c = idx.column();
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
    }
    m_spreadsheet->ungroupColumns(minCol, maxCol);
    viewport()->update();
    update();
}

int SpreadsheetView::outlineGutterTotalWidth() const {
    if (!m_spreadsheet) return 0;
    int maxLevel = m_spreadsheet->getMaxRowOutlineLevel();
    if (maxLevel <= 0) return 0;
    return (maxLevel + 1) * OUTLINE_GUTTER_WIDTH;
}

void SpreadsheetView::paintOutlineGutter(QPainter& painter) {
    if (!m_spreadsheet || !m_model) return;

    int maxLevel = m_spreadsheet->getMaxRowOutlineLevel();
    if (maxLevel <= 0) return;

    const auto& outlineLevels = m_spreadsheet->getRowOutlineLevels();
    if (outlineLevels.empty()) return;

    painter.save();
    painter.setRenderHint(QPainter::Antialiasing, true);

    int headerWidth = verticalHeader()->width();
    int headerHeight = horizontalHeader()->height();
    int gutterWidth = outlineGutterTotalWidth();
    int gutterX = headerWidth - gutterWidth;
    if (gutterX < 0) gutterX = 0;

    QColor lineColor(0x98, 0xA2, 0xB3);
    QColor buttonColor(0x42, 0x85, 0xF4);
    QColor bgColor(0xF5, 0xF5, 0xF5);

    int totalH = height();
    painter.fillRect(QRect(gutterX, headerHeight, gutterWidth, totalH - headerHeight), bgColor);

    // Draw level buttons at the top
    for (int lvl = 1; lvl <= maxLevel + 1; ++lvl) {
        int btnX = gutterX + (lvl - 1) * OUTLINE_GUTTER_WIDTH;
        int btnY = 2;
        int btnSize = OUTLINE_GUTTER_WIDTH - 2;
        QRect btnRect(btnX + 1, btnY, btnSize, btnSize);
        painter.setPen(QPen(buttonColor, 1));
        painter.setBrush(QColor(255, 255, 255));
        painter.drawRoundedRect(btnRect, 2, 2);
        painter.setPen(buttonColor);
        QFont font = painter.font();
        font.setPixelSize(10);
        font.setBold(true);
        painter.setFont(font);
        painter.drawText(btnRect, Qt::AlignCenter, QString::number(lvl));
    }

    // Draw tree lines and collapse/expand buttons for each group at each level
    for (int lvl = 1; lvl <= maxLevel; ++lvl) {
        int colX = gutterX + (lvl - 1) * OUTLINE_GUTTER_WIDTH + OUTLINE_GUTTER_WIDTH / 2;

        struct GroupInfo { int startRow; int endRow; bool collapsed; };
        std::vector<GroupInfo> groups;
        {
            int gStart = -1;
            for (auto it = outlineLevels.begin(); it != outlineLevels.end(); ++it) {
                int row = it->first;
                int rowLevel = it->second;
                if (rowLevel >= lvl) {
                    if (gStart < 0) gStart = row;
                } else {
                    if (gStart >= 0) {
                        auto prevIt = std::prev(it);
                        int gEnd = prevIt->first;
                        groups.push_back({gStart, gEnd, m_spreadsheet->isRowOutlineCollapsed(gEnd)});
                        gStart = -1;
                    }
                }
            }
            if (gStart >= 0) {
                int gEnd = std::prev(outlineLevels.end())->first;
                groups.push_back({gStart, gEnd, m_spreadsheet->isRowOutlineCollapsed(gEnd)});
            }
        }

        for (const auto& grp : groups) {
            int modelStart = m_model->toModelRow(grp.startRow);
            int modelEnd = m_model->toModelRow(grp.endRow);
            if (modelStart < 0) modelStart = 0;
            if (modelEnd < 0) continue;

            QRect startRect = visualRect(m_model->index(modelStart, 0));
            QRect endRect = visualRect(m_model->index(modelEnd, 0));
            int vpOffsetY = viewport()->mapTo(this, QPoint(0, 0)).y();
            int yStart = startRect.top() + vpOffsetY;
            int yEnd = endRect.bottom() + vpOffsetY;

            if (yEnd < headerHeight || yStart > totalH) continue;
            if (yStart < headerHeight) yStart = headerHeight;

            // Vertical tree line
            painter.setPen(QPen(lineColor, 1));
            painter.drawLine(colX, yStart, colX, yEnd - 6);
            // Horizontal tick
            painter.drawLine(colX, yEnd - 6, colX + 5, yEnd - 6);

            // Collapse/expand button
            int bsz = 11;
            QRect bRect(colX - bsz / 2, yEnd - bsz - 1, bsz, bsz);
            painter.setPen(QPen(lineColor, 1));
            painter.setBrush(QColor(255, 255, 255));
            painter.drawRect(bRect);

            painter.setPen(QPen(QColor(0x33, 0x33, 0x33), 1.5));
            int cx = bRect.center().x();
            int cy = bRect.center().y();
            painter.drawLine(cx - 3, cy, cx + 3, cy);
            if (grp.collapsed) {
                painter.drawLine(cx, cy - 3, cx, cy + 3);
            }
        }
    }

    // Separator line
    painter.setPen(QPen(lineColor, 1));
    painter.drawLine(gutterX + gutterWidth - 1, headerHeight,
                     gutterX + gutterWidth - 1, totalH);
    painter.restore();
}

void SpreadsheetView::handleOutlineGutterClick(const QPoint& pos) {
    if (!m_spreadsheet || !m_model) return;

    int maxLevel = m_spreadsheet->getMaxRowOutlineLevel();
    if (maxLevel <= 0) return;

    int headerWidth = verticalHeader()->width();
    int headerHeight = horizontalHeader()->height();
    int gutterWidth = outlineGutterTotalWidth();
    int gutterX = headerWidth - gutterWidth;
    if (gutterX < 0) gutterX = 0;

    int clickX = pos.x();
    int clickY = pos.y();

    if (clickX < gutterX || clickX >= gutterX + gutterWidth) return;
    if (clickY < 0) return;

    // Level button click
    if (clickY < headerHeight) {
        int levelClicked = (clickX - gutterX) / OUTLINE_GUTTER_WIDTH + 1;
        if (levelClicked >= 1 && levelClicked <= maxLevel + 1) {
            const auto& outlineLevels = m_spreadsheet->getRowOutlineLevels();
            for (const auto& [row, level] : outlineLevels) {
                int modelRow = m_model->toModelRow(row);
                if (modelRow < 0) continue;
                if (level >= levelClicked) {
                    m_spreadsheet->setRowHeight(row, 0);
                    setRowHidden(modelRow, true);
                } else {
                    if (m_spreadsheet->getRowHeight(row) == 0) {
                        m_spreadsheet->getRowHeights().erase(row);
                        setRowHidden(modelRow, false);
                    }
                }
            }
            viewport()->update();
            update();
            return;
        }
    }

    // Collapse/expand button click
    const auto& outlineLevels = m_spreadsheet->getRowOutlineLevels();
    for (int lvl = 1; lvl <= maxLevel; ++lvl) {
        int colX = gutterX + (lvl - 1) * OUTLINE_GUTTER_WIDTH + OUTLINE_GUTTER_WIDTH / 2;

        std::vector<std::pair<int, int>> groups;
        int gStart = -1;
        for (auto it = outlineLevels.begin(); it != outlineLevels.end(); ++it) {
            int row = it->first;
            int rowLevel = it->second;
            if (rowLevel >= lvl) {
                if (gStart < 0) gStart = row;
            } else {
                if (gStart >= 0) {
                    groups.push_back({gStart, std::prev(it)->first});
                    gStart = -1;
                }
            }
        }
        if (gStart >= 0) {
            groups.push_back({gStart, std::prev(outlineLevels.end())->first});
        }

        for (const auto& [grpStart, grpEnd] : groups) {
            int modelEnd = m_model->toModelRow(grpEnd);
            if (modelEnd < 0) continue;

            QRect endRect = visualRect(m_model->index(modelEnd, 0));
            int vpOffsetY = viewport()->mapTo(this, QPoint(0, 0)).y();
            int yEnd = endRect.bottom() + vpOffsetY;

            int bsz = 11;
            QRect bRect(colX - bsz / 2, yEnd - bsz - 1, bsz, bsz);

            if (bRect.adjusted(-3, -3, 3, 3).contains(QPoint(clickX, clickY))) {
                m_spreadsheet->toggleRowGroup(grpEnd, lvl);
                bool collapsed = m_spreadsheet->isRowOutlineCollapsed(grpEnd);
                int scanStart = grpEnd;
                while (scanStart > 0 && m_spreadsheet->getRowOutlineLevel(scanStart - 1) >= lvl) {
                    scanStart--;
                }
                for (int r = scanStart; r <= grpEnd; ++r) {
                    if (m_spreadsheet->getRowOutlineLevel(r) >= lvl) {
                        int modelRow = m_model->toModelRow(r);
                        if (modelRow >= 0) setRowHidden(modelRow, collapsed);
                    }
                }
                viewport()->update();
                update();
                return;
            }
        }
    }
}

// ============================================================================
// Multiline cell editor (QPlainTextEdit overlay for Alt+Enter)
// ============================================================================

void SpreadsheetView::openMultilineEditor(const QModelIndex& idx, const QString& text, int cursorPos) {
    // Close any existing multiline editor
    if (m_multilineEditor) {
        commitMultilineEditor();
    }

    m_multilineIndex = idx;

    // Create QPlainTextEdit overlay on the cell
    m_multilineEditor = new QPlainTextEdit(viewport());
    m_multilineEditor->setPlainText(text);

    // Style to match cell editor look
    const auto& t = ThemeManager::instance().currentTheme();
    m_multilineEditor->setStyleSheet(QString(
        "QPlainTextEdit { background: white; color: black; padding: 2px 4px; "
        "border: 2px solid %1; font-size: %2px; font-family: %3; }")
        .arg(t.editorBorderColor.name())
        .arg(font().pointSize() > 0 ? font().pointSize() : 11)
        .arg(font().family()));
    m_multilineEditor->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_multilineEditor->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_multilineEditor->setTabChangesFocus(true);

    // Position over the cell, expanding downward
    QRect cellRect = visualRect(idx);
    int minHeight = qMax(cellRect.height() * 3, 80);  // at least 3 rows tall
    int editorWidth = qMax(cellRect.width(), 200);
    m_multilineEditor->setGeometry(cellRect.x(), cellRect.y(), editorWidth, minHeight);
    m_multilineEditor->show();
    m_multilineEditor->setFocus();

    // Position cursor
    QTextCursor tc = m_multilineEditor->textCursor();
    tc.setPosition(qMin(cursorPos, text.length()));
    m_multilineEditor->setTextCursor(tc);

    // Auto-resize as user types
    connect(m_multilineEditor, &QPlainTextEdit::textChanged, this, [this]() {
        if (!m_multilineEditor) return;
        QRect cellRect = visualRect(m_multilineIndex);
        int docHeight = static_cast<int>(m_multilineEditor->document()->size().height()) + 16;
        int minHeight = qMax(cellRect.height() * 3, 80);
        int h = qMax(docHeight, minHeight);
        int w = qMax(cellRect.width(), 200);
        m_multilineEditor->setGeometry(cellRect.x(), cellRect.y(), w, h);
    });
}

void SpreadsheetView::commitMultilineEditor() {
    if (!m_multilineEditor || !m_spreadsheet) return;

    QString text = m_multilineEditor->toPlainText();
    CellAddress addr{m_multilineIndex.row(), m_multilineIndex.column()};

    // Set cell value
    m_spreadsheet->setCellValue(addr, text);

    // Auto-enable wrap text if there are newlines
    if (text.contains('\n')) {
        auto cell = m_spreadsheet->getCell(addr);
        CellStyle style = cell->getStyle();
        if (style.textOverflow != TextOverflowMode::Wrap) {
            style.textOverflow = TextOverflowMode::Wrap;
            cell->setStyle(style);
        }
    }

    // Clean up
    m_multilineEditor->deleteLater();
    m_multilineEditor = nullptr;
    setFocus();
    refreshView();
}

void SpreadsheetView::cancelMultilineEditor() {
    if (!m_multilineEditor) return;
    m_multilineEditor->deleteLater();
    m_multilineEditor = nullptr;
    setFocus();
}

// =============================================================================
// Go To Special — select cells by type (blanks, constants, formulas, comments)
// =============================================================================
void SpreadsheetView::goToSpecial() {
    if (!m_spreadsheet) return;

    GoToSpecialDialog dialog(this);
    if (dialog.exec() != QDialog::Accepted) return;

    auto selType = dialog.getSelectionType();
    bool wantNumbers = dialog.includeNumbers();
    bool wantText = dialog.includeText();
    bool wantLogicals = dialog.includeLogicals();
    bool wantErrors = dialog.includeErrors();

    int maxRow = m_spreadsheet->getMaxRow();
    int maxCol = m_spreadsheet->getMaxColumn();

    // For CurrentRegion, use detectDataRegion from current cell
    if (selType == GoToSpecialDialog::CurrentRegion) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid()) return;
        CellRange region = detectDataRegion(logicalRow(cur), cur.column());
        if (region.isValid() && m_model) {
            QItemSelection sel;
            sel.select(m_model->index(region.getStart().row, region.getStart().col),
                       m_model->index(region.getEnd().row, region.getEnd().col));
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
        }
        return;
    }

    if (maxRow < 0 || maxCol < 0) {
        QMessageBox::information(this, "Go To Special", "No cells found matching the criteria.");
        return;
    }

    QItemSelection matchingSel;
    int matchCount = 0;

    for (int r = 0; r <= maxRow; ++r) {
        for (int c = 0; c <= maxCol; ++c) {
            auto cell = m_spreadsheet->getCellIfExists(r, c);
            bool match = false;

            switch (selType) {
            case GoToSpecialDialog::Blanks:
                match = !cell || cell.isEmpty() || cell.getType() == CellType::Empty;
                break;

            case GoToSpecialDialog::Constants: {
                if (!cell || cell.isEmpty()) break;
                if (cell.hasFormula()) break; // constants only = no formulas
                CellType ct = cell.getType();
                if (wantNumbers && (ct == CellType::Number || ct == CellType::Date)) match = true;
                if (wantText && ct == CellType::Text) match = true;
                if (wantLogicals && ct == CellType::Boolean) match = true;
                if (wantErrors && ct == CellType::Error) match = true;
                break;
            }

            case GoToSpecialDialog::Formulas: {
                if (!cell || !cell.hasFormula()) break;
                // Check the computed value type
                QVariant val = cell.getComputedValue();
                if (cell.hasError()) {
                    if (wantErrors) match = true;
                } else if (val.typeId() == QMetaType::Bool) {
                    if (wantLogicals) match = true;
                } else if (val.typeId() == QMetaType::Double || val.typeId() == QMetaType::Int ||
                           val.typeId() == QMetaType::LongLong || val.typeId() == QMetaType::Float) {
                    if (wantNumbers) match = true;
                } else if (val.typeId() == QMetaType::QString) {
                    if (wantText) match = true;
                } else {
                    if (wantNumbers) match = true; // default: treat as number
                }
                break;
            }

            case GoToSpecialDialog::Comments:
                if (cell && cell.hasComment()) match = true;
                break;

            case GoToSpecialDialog::ConditionalFormats: {
                const auto& cf = m_spreadsheet->getConditionalFormatting();
                const auto& rules = cf.getAllRules();
                for (const auto& rule : rules) {
                    if (rule->getRange().contains(r, c)) {
                        match = true;
                        break;
                    }
                }
                break;
            }

            case GoToSpecialDialog::DataValidation: {
                if (m_spreadsheet->getValidationAt(r, c)) match = true;
                break;
            }

            case GoToSpecialDialog::VisibleCells:
                if (!isRowHidden(r) && !isColumnHidden(c)) match = true;
                break;

            case GoToSpecialDialog::CurrentRegion:
                break; // handled above
            }

            if (match) {
                QModelIndex idx = m_model ? m_model->index(r, c) : QModelIndex();
                if (idx.isValid()) {
                    matchingSel.select(idx, idx);
                    matchCount++;
                }
            }
        }
    }

    if (matchCount > 0) {
        selectionModel()->select(matchingSel, QItemSelectionModel::ClearAndSelect);
        // Scroll to the first match
        if (!matchingSel.isEmpty() && !matchingSel.first().indexes().isEmpty()) {
            scrollTo(matchingSel.first().indexes().first());
        }
    } else {
        QMessageBox::information(this, "Go To Special", "No cells found matching the criteria.");
    }
}

// =============================================================================
// Remove Duplicates — find and remove duplicate rows based on selected columns
// =============================================================================
void SpreadsheetView::removeDuplicates() {
    if (!m_spreadsheet) return;

    // Determine the data range from selection or auto-detect
    QModelIndex cur = currentIndex();
    if (!cur.isValid()) return;

    CellRange dataRange;
    QItemSelection sel = selectionModel()->selection();
    if (!sel.isEmpty()) {
        int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
        for (const auto& r : sel) {
            minR = qMin(minR, r.top());
            maxR = qMax(maxR, r.bottom());
            minC = qMin(minC, r.left());
            maxC = qMax(maxC, r.right());
        }
        // If selection is more than 1 cell, use it; otherwise auto-detect
        if (maxR > minR || maxC > minC) {
            int logMinR = m_model ? m_model->toLogicalRow(minR) : minR;
            int logMaxR = m_model ? m_model->toLogicalRow(maxR) : maxR;
            dataRange = CellRange(CellAddress(logMinR, minC), CellAddress(logMaxR, maxC));
        }
    }

    if (!dataRange.isValid()) {
        dataRange = detectDataRegion(logicalRow(cur), cur.column());
    }
    if (!dataRange.isValid()) {
        QMessageBox::information(this, "Remove Duplicates", "No data range found.");
        return;
    }

    int startRow = dataRange.getStart().row;
    int endRow = dataRange.getEnd().row;
    int startCol = dataRange.getStart().col;
    int endCol = dataRange.getEnd().col;

    // Build column headers from first row
    QStringList headers;
    for (int c = startCol; c <= endCol; ++c) {
        auto cell = m_spreadsheet->getCellIfExists(startRow, c);
        QString val = cell ? cell.getValue().toString().trimmed() : QString();
        if (val.isEmpty()) {
            // Generate column letter
            QString colLetter;
            int cc = c;
            do {
                colLetter.prepend(QChar('A' + cc % 26));
                cc = cc / 26 - 1;
            } while (cc >= 0);
            val = QString("Column %1").arg(colLetter);
        }
        headers.append(val);
    }

    RemoveDuplicatesDialog dialog(headers, this);
    if (dialog.exec() != QDialog::Accepted) return;

    QVector<int> selectedCols = dialog.getSelectedColumns();
    if (selectedCols.isEmpty()) {
        QMessageBox::information(this, "Remove Duplicates", "No columns selected.");
        return;
    }

    bool hasHeaders = dialog.hasHeaders();
    int dataStartRow = hasHeaders ? startRow + 1 : startRow;

    if (dataStartRow > endRow) {
        QMessageBox::information(this, "Remove Duplicates",
            "0 duplicate values found and removed; 0 unique values remain.");
        return;
    }

    // Build set of seen row-tuples, keep track of duplicate row indices
    QSet<QString> seen;
    QVector<int> duplicateRows;

    for (int r = dataStartRow; r <= endRow; ++r) {
        // Build a key string from the selected columns
        QString key;
        for (int ci : selectedCols) {
            int actualCol = startCol + ci;
            if (actualCol > endCol) continue;
            auto cell = m_spreadsheet->getCellIfExists(r, actualCol);
            QString val = cell ? cell.getValue().toString() : QString();
            key += val + QChar('\x1F'); // unit separator as delimiter
        }

        if (seen.contains(key)) {
            duplicateRows.append(r);
        } else {
            seen.insert(key);
        }
    }

    if (duplicateRows.isEmpty()) {
        int uniqueCount = endRow - dataStartRow + 1;
        QMessageBox::information(this, "Remove Duplicates",
            QString("0 duplicate values found and removed; %1 unique values remain.")
                .arg(uniqueCount));
        return;
    }

    // Delete duplicate rows from bottom to top to preserve indices
    std::sort(duplicateRows.begin(), duplicateRows.end(), std::greater<int>());
    for (int r : duplicateRows) {
        m_spreadsheet->deleteRow(r);
    }

    int removedCount = duplicateRows.size();
    int uniqueCount = (endRow - dataStartRow + 1) - removedCount;

    // Refresh view
    if (m_model) {
        m_model->resetModel();
    }

    QMessageBox::information(this, "Remove Duplicates",
        QString("%1 duplicate values found and removed; %2 unique values remain.")
            .arg(removedCount).arg(uniqueCount));
}
