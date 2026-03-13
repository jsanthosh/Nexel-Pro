#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "CellDelegate.h"
#include "PasteSpecialDialog.h"
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
#include <QCheckBox>
#include <QRegularExpression>
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
    setModel(m_model);

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

void SpreadsheetView::initializeView() {
    m_model = new SpreadsheetModel(m_spreadsheet, this);
    setModel(m_model);

    m_delegate = new CellDelegate(this);
    m_delegate->setSpreadsheetView(this);
    setItemDelegate(m_delegate);

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

    // Enable mouse tracking for fill handle cursor changes
    viewport()->setMouseTracking(true);

    // Cell context menu (right-click)
    setContextMenuPolicy(Qt::CustomContextMenu);
    connect(this, &QTableView::customContextMenuRequested,
            this, &SpreadsheetView::showCellContextMenu);

    // Setup header context menus
    setupHeaderContextMenus();
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
        t.headerText.name()
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
        menu.addAction("Autofit Column Width", this, &SpreadsheetView::autofitSelectedColumns);
        menu.addSeparator();
        int clickedCol = horizontalHeader()->logicalIndexAt(pos);

        // Determine selected column range
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
        int minCol = *std::min_element(selCols.begin(), selCols.end());

        QString insertLabel = colCount > 1 ? QString("Insert %1 Columns").arg(colCount) : "Insert Column";
        QString deleteLabel = colCount > 1 ? QString("Delete %1 Columns").arg(colCount) : "Delete Column";

        menu.addAction(insertLabel, [this]() {
            insertEntireColumn();
        });
        menu.addAction(deleteLabel, [this]() {
            deleteEntireColumn();
        });
        menu.exec(horizontalHeader()->mapToGlobal(pos));
    });

    verticalHeader()->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(verticalHeader(), &QHeaderView::customContextMenuRequested,
            this, [this](const QPoint& pos) {
        QMenu menu;
        menu.addAction("Autofit Row Height", this, &SpreadsheetView::autofitSelectedRows);
        menu.addSeparator();
        int clickedRow = verticalHeader()->logicalIndexAt(pos);

        // Determine selected row range
        QSet<int> selRows;
        QModelIndexList selIdx = selectionModel()->selectedRows();
        if (selIdx.size() > 1) {
            for (const auto& idx : selIdx) selRows.insert(idx.row());
        } else {
            QModelIndexList allSel = selectionModel()->selectedIndexes();
            for (const auto& idx : allSel) selRows.insert(idx.row());
        }
        if (selRows.isEmpty()) selRows.insert(clickedRow);
        int rowCount = selRows.size();
        int minRow = *std::min_element(selRows.begin(), selRows.end());

        QString insertLabel = rowCount > 1 ? QString("Insert %1 Rows").arg(rowCount) : "Insert Row";
        QString deleteLabel = rowCount > 1 ? QString("Delete %1 Rows").arg(rowCount) : "Delete Row";

        menu.addAction(insertLabel, [this]() {
            insertEntireRow();
        });
        menu.addAction(deleteLabel, [this]() {
            deleteEntireRow();
        });
        menu.exec(verticalHeader()->mapToGlobal(pos));
    });
}

void SpreadsheetView::emitCellSelected(const QModelIndex& index) {
    if (!index.isValid() || !m_spreadsheet) return;

    CellAddress addr(index.row(), index.column());
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

    // Force repaint of previous cell to clear its focus border and fill handle
    if (previous.isValid()) {
        QRect prevRect = visualRect(previous);
        // Expand to cover 2px focus border + fill handle (7px square at corner)
        viewport()->update(prevRect.adjusted(-2, -2, 6, 6));
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
}

// ============== Clipboard operations ==============

void SpreadsheetView::cut() {
    copy();
    deleteSelection();
}

void SpreadsheetView::copy() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

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
        CellAddress addr(idx.row(), idx.column());
        auto cell = m_spreadsheet->getCell(addr);
        m_internalClipboard[r][c].value = cell->getValue();
        m_internalClipboard[r][c].style = cell->getStyle();
        m_internalClipboard[r][c].type = cell->getType();
        m_internalClipboard[r][c].formula = cell->getFormula();
        m_internalClipboard[r][c].sourceAddr = addr;
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
}

void SpreadsheetView::paste() {
    QString data = QApplication::clipboard()->text();
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int startRow = current.row();
    int startCol = current.column();

    std::vector<CellSnapshot> before, after;
    m_model->setSuppressUndo(true);

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

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
    } else {
        // External paste: plain text only
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
    m_model->setSuppressUndo(false);

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Paste"));

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
    int startRow = current.row();
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

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);
    for (const auto& index : selected) {
        CellAddress addr(index.row(), index.column());
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
    QTableView::selectAll();
}

// ============== Style operations ==============

// Efficient style application: for large selections (select all), only iterate occupied cells
void SpreadsheetView::applyStyleChange(std::function<void(CellStyle&)> modifier, const QList<int>& roles) {
    if (!m_spreadsheet) return;

    // Use selection ranges (compact) instead of selectedIndexes() (one per cell — very slow for Select All)
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

    static constexpr int LARGE_SELECTION_THRESHOLD = 5000;
    bool isLargeSelection = totalCells > LARGE_SELECTION_THRESHOLD;

    std::vector<CellSnapshot> before, after;

    if (isLargeSelection) {
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
    } else {
        QModelIndexList selected = selectionModel()->selectedIndexes();
        for (const auto& index : selected) {
            CellAddress addr(index.row(), index.column());
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
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
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return false;
    for (const auto& idx : selected) {
        auto cell = m_spreadsheet->getCellIfExists(idx.row(), idx.column());
        CellStyle style = cell ? cell->getStyle() : CellStyle();
        if (!predicate(style)) return false;
    }
    return true;
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
    applyStyleChange([&family](CellStyle& s) { s.fontName = family; }, {Qt::FontRole});
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
    applyStyleChange([&colorStr](CellStyle& s) { s.foregroundColor = colorStr; }, {Qt::ForegroundRole});
}

void SpreadsheetView::applyBackgroundColor(const QString& colorStr) {
    if (m_macroEngine && m_macroEngine->isRecording()) {
        m_macroEngine->recordAction(QString("sheet.setBackgroundColor(\"%1\", \"%2\");").arg(selectionRangeStr(selectionModel()), colorStr));
    }
    applyStyleChange([&colorStr](CellStyle& s) { s.backgroundColor = colorStr; }, {Qt::BackgroundRole});
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

    CellAddress addr(current.row(), current.column());
    auto cell = m_spreadsheet->getCell(addr);
    m_copiedStyle = cell->getStyle();
    m_formatPainterActive = true;
    viewport()->setCursor(Qt::CrossCursor);
}

// ============== Sorting ==============

void SpreadsheetView::sortAscending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();

    // Use the user's selection if it spans multiple rows; otherwise sort the whole data region
    QModelIndexList selected = selectionModel()->selectedIndexes();
    CellRange range;
    if (selected.size() > 1) {
        int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
        for (const auto& idx : selected) {
            minR = qMin(minR, idx.row());
            maxR = qMax(maxR, idx.row());
            minC = qMin(minC, idx.column());
            maxC = qMax(maxC, idx.column());
        }
        if (minR < maxR) {
            range = CellRange(CellAddress(minR, minC), CellAddress(maxR, maxC));
        }
    }

    if (!range.isValid()) {
        int maxRow = m_spreadsheet->getMaxRow();
        int maxCol = m_spreadsheet->getMaxColumn();
        if (maxRow < 1 && maxCol < 1) return;
        if (maxRow < 1) maxRow = 1;
        range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    }

    m_spreadsheet->sortRange(range, col, true);

    // Save selection bounds before reset clears them
    int selMinR = range.getStart().row, selMaxR = range.getEnd().row;
    int selMinC = range.getStart().col, selMaxC = range.getEnd().col;

    // Full model reset to ensure view refreshes completely
    if (m_model) {
        m_model->resetModel();
    }

    // Restore the selection range — set current index first, then select,
    // so setCurrentIndex doesn't clear the restored selection.
    QModelIndex topLeft = m_model->index(selMinR, selMinC);
    QModelIndex bottomRight = m_model->index(selMaxR, selMaxC);
    selectionModel()->setCurrentIndex(topLeft, QItemSelectionModel::NoUpdate);
    selectionModel()->select(QItemSelection(topLeft, bottomRight), QItemSelectionModel::ClearAndSelect);
}

void SpreadsheetView::sortDescending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();

    // Use the user's selection if it spans multiple rows; otherwise sort the whole data region
    QModelIndexList selected = selectionModel()->selectedIndexes();
    CellRange range;
    if (selected.size() > 1) {
        int minR = INT_MAX, maxR = 0, minC = INT_MAX, maxC = 0;
        for (const auto& idx : selected) {
            minR = qMin(minR, idx.row());
            maxR = qMax(maxR, idx.row());
            minC = qMin(minC, idx.column());
            maxC = qMax(maxC, idx.column());
        }
        if (minR < maxR) {
            range = CellRange(CellAddress(minR, minC), CellAddress(maxR, maxC));
        }
    }

    if (!range.isValid()) {
        int maxRow = m_spreadsheet->getMaxRow();
        int maxCol = m_spreadsheet->getMaxColumn();
        if (maxRow < 1 && maxCol < 1) return;
        if (maxRow < 1) maxRow = 1;
        range = CellRange(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    }

    m_spreadsheet->sortRange(range, col, false);

    // Save selection bounds before reset clears them
    int selMinR = range.getStart().row, selMaxR = range.getEnd().row;
    int selMinC = range.getStart().col, selMaxC = range.getEnd().col;

    if (m_model) {
        m_model->resetModel();
    }

    // Restore the selection range — set current index first, then select,
    // so setCurrentIndex doesn't clear the restored selection.
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
    m_filterRange = detectDataRegion(current.row(), current.column());
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

    viewport()->update();
}

// ============== Clear Operations ==============

void SpreadsheetView::clearAll() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

    std::vector<CellSnapshot> before, after;
    for (const auto& index : selected) {
        CellAddress addr(index.row(), index.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        auto cell = m_spreadsheet->getCell(addr);
        cell->clear();
        cell->setStyle(CellStyle()); // Reset to default style
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
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
        CellAddress addr(index.row(), index.column());
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
            s.borderTop = on;
            s.borderBottom = on;
            s.borderLeft = on;
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
        CellAddress addr(idx.row(), idx.column());
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

    auto* mr = m_spreadsheet->getMergedRegionAt(current.row(), current.column());
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

    menu.addAction("Cut", this, &SpreadsheetView::cut, QKeySequence::Cut);
    menu.addAction("Copy", this, &SpreadsheetView::copy, QKeySequence::Copy);
    menu.addAction("Paste", this, &SpreadsheetView::paste, QKeySequence::Paste);
    menu.addAction("Paste Special...", this, &SpreadsheetView::pasteSpecial,
                   QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_V));

    menu.addSeparator();

    // Clear submenu
    QMenu* clearMenu = menu.addMenu("Clear");
    clearMenu->addAction("Clear All", this, &SpreadsheetView::clearAll);
    clearMenu->addAction("Clear Contents", this, &SpreadsheetView::clearContent);
    clearMenu->addAction("Clear Formats", this, &SpreadsheetView::clearFormats);

    menu.addSeparator();

    // Insert submenu
    QMenu* insertMenu = menu.addMenu("Insert...");
    insertMenu->addAction("Shift cells right", this, &SpreadsheetView::insertCellsShiftRight);
    insertMenu->addAction("Shift cells down", this, &SpreadsheetView::insertCellsShiftDown);
    insertMenu->addSeparator();
    insertMenu->addAction("Entire row", this, &SpreadsheetView::insertEntireRow);
    insertMenu->addAction("Entire column", this, &SpreadsheetView::insertEntireColumn);

    // Delete submenu
    QMenu* deleteMenu = menu.addMenu("Delete...");
    deleteMenu->addAction("Shift cells left", this, &SpreadsheetView::deleteCellsShiftLeft);
    deleteMenu->addAction("Shift cells up", this, &SpreadsheetView::deleteCellsShiftUp);
    deleteMenu->addSeparator();
    deleteMenu->addAction("Entire row", this, &SpreadsheetView::deleteEntireRow);
    deleteMenu->addAction("Entire column", this, &SpreadsheetView::deleteEntireColumn);

    menu.addSeparator();

    // Merge cells
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.size() > 1) {
        menu.addAction("Merge && Center", this, &SpreadsheetView::mergeCells);
    }
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid() && m_spreadsheet->getMergedRegionAt(cur.row(), cur.column())) {
            menu.addAction("Unmerge Cells", this, &SpreadsheetView::unmergeCells);
        }
    }

    menu.addSeparator();

    // Checkbox / Picklist context actions
    if (m_spreadsheet) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            auto cell = m_spreadsheet->getCell(CellAddress(cur.row(), cur.column()));
            if (cell) {
                QString fmt = cell->getStyle().numberFormat;
                if (fmt == "Checkbox") {
                    menu.addAction("Remove Checkbox", this, [this]() {
                        if (!m_spreadsheet || !m_model) return;
                        auto sel = selectionModel()->selectedIndexes();
                        if (sel.isEmpty()) return;
                        std::vector<CellSnapshot> before, after;
                        for (const auto& idx : sel) {
                            CellAddress addr(idx.row(), idx.column());
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
                            CellAddress addr(idx.row(), idx.column());
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
            auto cell = m_spreadsheet->getCellIfExists(cur.row(), cur.column());
            bool hasComment = cell && cell->hasComment();
            menu.addAction(hasComment ? "Edit Comment..." : "Insert Comment...",
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
            auto cell = m_spreadsheet->getCellIfExists(cur.row(), cur.column());
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

    menu.addAction("Format Cells...", this, [this]() {
        emit formatCellsRequested();
    }, QKeySequence(Qt::CTRL | Qt::Key_1));

    menu.addSeparator();

    menu.addAction("Sort Ascending", this, &SpreadsheetView::sortAscending);
    menu.addAction("Sort Descending", this, &SpreadsheetView::sortDescending);

    menu.exec(viewport()->mapToGlobal(pos));
}

// ============== Insert/Delete with shift ==============

void SpreadsheetView::insertCellsShiftRight() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
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

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
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
        rows.insert(current.row());
    } else {
        for (const auto& idx : selected) {
            rows.insert(idx.row());
        }
    }

    // Use beginInsertRows/endInsertRows so Qt automatically shifts header heights.
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());

    verticalHeader()->blockSignals(true);

    for (int row : sortedRows) {
        m_model->beginRowInsertion(row, 1);
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<InsertRowCommand>(row, 1), m_spreadsheet.get());
        m_model->endRowInsertion();
    }

    verticalHeader()->blockSignals(false);
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

    // Use beginInsertColumns/endInsertColumns so Qt automatically shifts header widths.
    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());

    horizontalHeader()->blockSignals(true);

    for (int col : sortedCols) {
        m_model->beginColumnInsertion(col, 1);
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<InsertColumnCommand>(col, 1, col), m_spreadsheet.get());
        m_model->endColumnInsertion();
    }

    horizontalHeader()->blockSignals(false);
}

void SpreadsheetView::deleteCellsShiftLeft() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
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

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
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
        rows.insert(current.row());
    } else {
        for (const auto& idx : selected) {
            rows.insert(idx.row());
        }
    }

    int focusRow = *std::min_element(rows.begin(), rows.end());
    int focusCol = current.column();

    // Delete rows high-to-low to preserve indices.
    // Use beginRemoveRows/endRemoveRows so Qt automatically shifts header heights.
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());

    // Block header signals so Qt's internal section adjustments don't
    // trigger onVerticalSectionResized (which would corrupt m_rowHeights).
    verticalHeader()->blockSignals(true);

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

        // Save the height of the row being deleted (for undo restore)
        std::vector<int> savedHeights = { verticalHeader()->sectionSize(row) };

        // Tell Qt model: rows are about to be removed (BEFORE data change)
        m_model->beginRowRemoval(row, 1);

        // Execute the delete (changes data + shifts m_rowHeights in Spreadsheet)
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<DeleteRowCommand>(row, 1, deleted, savedHeights), m_spreadsheet.get());

        // Tell Qt model: removal complete (AFTER data change)
        // Qt will now automatically shift header section sizes
        m_model->endRowRemoval();
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
    // Use beginRemoveColumns/endRemoveColumns so Qt automatically shifts header widths.
    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());

    // Block header signals so Qt's internal section adjustments don't
    // trigger onHorizontalSectionResized (which would corrupt m_columnWidths).
    horizontalHeader()->blockSignals(true);

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

        // Save the width of the column being deleted (for undo restore)
        std::vector<int> savedWidths = { horizontalHeader()->sectionSize(col) };

        // Tell Qt model: columns are about to be removed (BEFORE data change)
        m_model->beginColumnRemoval(col, 1);

        // Execute the delete (changes data + shifts m_columnWidths in Spreadsheet)
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<DeleteColumnCommand>(col, 1, deleted, savedWidths), m_spreadsheet.get());

        // Tell Qt model: removal complete (AFTER data change)
        // Qt will now automatically shift header section sizes
        m_model->endColumnRemoval();
    }

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

    auto cell = m_spreadsheet->getCell(CellAddress(cur.row(), cur.column()));
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

    auto cell = m_spreadsheet->getCellIfExists(cur.row(), cur.column());
    if (cell && cell->hasComment()) {
        cell->setComment(QString());
        viewport()->update();
    }
}

// ============== Hyperlinks ==============

void SpreadsheetView::insertOrEditHyperlink() {
    QModelIndex cur = currentIndex();
    if (!cur.isValid() || !m_spreadsheet) return;

    auto cell = m_spreadsheet->getCell(CellAddress(cur.row(), cur.column()));
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

    auto cell = m_spreadsheet->getCellIfExists(cur.row(), cur.column());
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
            auto cell = m_spreadsheet->getCellIfExists(index.row(), index.column());
            if (cell) {
                QStringList tips;
                if (cell->hasComment()) tips << cell->getComment();
                if (cell->hasHyperlink()) tips << "Ctrl+Click to open: " + cell->getHyperlink();
                if (!tips.isEmpty()) {
                    QToolTip::showText(helpEvent->globalPos(), tips.join("\n"), this);
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
}

void SpreadsheetView::zoomIn() {
    m_zoomLevel += 10;
    if (m_zoomLevel > 200) m_zoomLevel = 200;

    QFont f = this->font();
    f.setPointSize(m_baseFontSize * m_zoomLevel / 100);
    setFont(f);
}

void SpreadsheetView::zoomOut() {
    m_zoomLevel -= 10;
    if (m_zoomLevel < 50) m_zoomLevel = 50;

    QFont f = this->font();
    f.setPointSize(m_baseFontSize * m_zoomLevel / 100);
    setFont(f);
}

void SpreadsheetView::resetZoom() {
    m_zoomLevel = 100;
    setFont(QFont("Arial", m_baseFontSize));
}

// ============== Event handlers ==============

void SpreadsheetView::keyPressEvent(QKeyEvent* event) {
    bool ctrl = event->modifiers() & Qt::ControlModifier;
    bool shift = event->modifiers() & Qt::ShiftModifier;

    // Delete / Backspace: clear selection (on Mac, "Delete" key = Backspace)
    if (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace) {
        if (state() != QAbstractItemView::EditingState) {
            deleteSelection();
            event->accept();
            return;
        }
        // If editing, let the editor handle the key
        QTableView::keyPressEvent(event);
        return;
    }

    // Enter/Return: commit and move down
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

    // Tab: commit and move right; Shift+Tab: move left
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

    // F2: Edit current cell (like Excel)
    if (event->key() == Qt::Key_F2) {
        QModelIndex current = currentIndex();
        if (current.isValid() && state() != QAbstractItemView::EditingState) {
            edit(current);
        }
        event->accept();
        return;
    }

    // Escape: cancel editing / format painter / formula edit mode
    if (event->key() == Qt::Key_Escape) {
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

    // ===== Ctrl/Cmd+D: Fill Down (copy value from cell above into selection) =====
    if (ctrl && event->key() == Qt::Key_D) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndexList selected = selectionModel()->selectedIndexes();
        if (selected.size() <= 1) {
            // Single cell: copy from cell above
            QModelIndex cur = currentIndex();
            if (cur.isValid() && cur.row() > 0) {
                auto valAbove = m_spreadsheet->getCellValue(CellAddress(cur.row() - 1, cur.column()));
                auto cellAbove = m_spreadsheet->getCell(CellAddress(cur.row() - 1, cur.column()));

                std::vector<CellSnapshot> before, after;
                CellAddress addr(cur.row(), cur.column());
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

                CellAddress addr(idx.row(), idx.column());
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
    // Uses sparse cell map scan + binary search instead of row-by-row iteration.
    // This makes navigation O(total_cells) instead of O(grid_rows), critical for 1M+ row sheets.
    // Skip during bulk loading — background thread is mutating m_cells concurrently.
    if (ctrl && !shift && !m_bulkLoading &&
        (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        int row = cur.row();
        int col = cur.column();
        int maxRow = m_spreadsheet->getRowCount() - 1;
        int maxCol = m_spreadsheet->getColumnCount() - 1;

        if (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down) {
            auto occupied = m_spreadsheet->getOccupiedRowsInColumn(col);
            if (event->key() == Qt::Key_Down) {
                if (row < maxRow) {
                    auto it = std::upper_bound(occupied.begin(), occupied.end(), row);
                    bool nextHasData = (it != occupied.end() && *it == row + 1);
                    if (nextHasData) {
                        // Walk contiguous data block downward
                        int last = row + 1;
                        ++it;
                        while (it != occupied.end() && *it == last + 1) { last = *it; ++it; }
                        row = last;
                    } else {
                        // Jump to next non-empty row, or grid bottom
                        row = (it != occupied.end()) ? *it : maxRow;
                    }
                }
            } else { // Key_Up
                if (row > 0) {
                    auto prevIt = std::lower_bound(occupied.begin(), occupied.end(), row - 1);
                    bool prevHasData = (prevIt != occupied.end() && *prevIt == row - 1);
                    if (prevHasData) {
                        // Walk contiguous data block upward
                        int first = row - 1;
                        while (prevIt != occupied.begin()) {
                            auto before = std::prev(prevIt);
                            if (*before != first - 1) break;
                            first = *before;
                            prevIt = before;
                        }
                        row = first;
                    } else {
                        // Jump to prev non-empty row, or grid top
                        auto it = std::lower_bound(occupied.begin(), occupied.end(), row);
                        row = (it != occupied.begin()) ? *std::prev(it) : 0;
                    }
                }
            }
        } else { // Key_Left or Key_Right
            auto occupied = m_spreadsheet->getOccupiedColsInRow(row);
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

        QModelIndex target = model()->index(row, col);
        if (target.isValid()) {
            selectionModel()->setCurrentIndex(target,
                QItemSelectionModel::ClearAndSelect);
            scrollTo(target);
        }
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
    // Same sparse-map navigation as Ctrl+Arrow above, but extends selection.
    // Skip during bulk loading — background thread is mutating m_cells concurrently.
    if (ctrl && shift && !m_bulkLoading &&
        (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        int row = cur.row();
        int col = cur.column();
        int maxRowIdx = m_spreadsheet->getRowCount() - 1;
        int maxColIdx = m_spreadsheet->getColumnCount() - 1;

        if (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down) {
            auto occupied = m_spreadsheet->getOccupiedRowsInColumn(col);
            if (event->key() == Qt::Key_Down) {
                if (row < maxRowIdx) {
                    auto it = std::upper_bound(occupied.begin(), occupied.end(), row);
                    bool nextHasData = (it != occupied.end() && *it == row + 1);
                    if (nextHasData) {
                        int last = row + 1;
                        ++it;
                        while (it != occupied.end() && *it == last + 1) { last = *it; ++it; }
                        row = last;
                    } else {
                        row = (it != occupied.end()) ? *it : maxRowIdx;
                    }
                }
            } else {
                if (row > 0) {
                    auto prevIt = std::lower_bound(occupied.begin(), occupied.end(), row - 1);
                    bool prevHasData = (prevIt != occupied.end() && *prevIt == row - 1);
                    if (prevHasData) {
                        int first = row - 1;
                        while (prevIt != occupied.begin()) {
                            auto before = std::prev(prevIt);
                            if (*before != first - 1) break;
                            first = *before;
                            prevIt = before;
                        }
                        row = first;
                    } else {
                        auto it = std::lower_bound(occupied.begin(), occupied.end(), row);
                        row = (it != occupied.begin()) ? *std::prev(it) : 0;
                    }
                }
            }
        } else {
            auto occupied = m_spreadsheet->getOccupiedColsInRow(row);
            if (event->key() == Qt::Key_Right) {
                if (col < maxColIdx) {
                    auto it = std::upper_bound(occupied.begin(), occupied.end(), col);
                    bool nextHasData = (it != occupied.end() && *it == col + 1);
                    if (nextHasData) {
                        int last = col + 1;
                        ++it;
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

        QModelIndex target = model()->index(row, col);
        if (target.isValid()) {
            // Extend selection from current to target
            QItemSelection sel(cur, target);
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
            setCurrentIndex(target);
            scrollTo(target);
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
                auto valLeft = m_spreadsheet->getCellValue(CellAddress(cur.row(), cur.column() - 1));
                auto cellLeft = m_spreadsheet->getCell(CellAddress(cur.row(), cur.column() - 1));

                CellAddress addr(cur.row(), cur.column());
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

    // ===== Ctrl+` (backtick): Toggle formula view =====
    if (ctrl && event->key() == Qt::Key_QuoteLeft) {
        toggleFormulaView();
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

    QTableView::keyPressEvent(event);
}

void SpreadsheetView::setFormulaEditMode(bool active) {
    m_formulaEditMode = active;
    if (m_delegate) {
        m_delegate->setFormulaEditMode(active);
    }
    if (active) {
        m_formulaEditCell = currentIndex();
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

void SpreadsheetView::mousePressEvent(QMouseEvent* event) {
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
            CellAddress addr(idx.row(), idx.column());
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
            auto cell = m_spreadsheet->getCellIfExists(idx.row(), idx.column());
            if (cell && cell->hasHyperlink()) {
                openHyperlink(idx.row(), idx.column());
                event->accept();
                return;
            }
        }
    }

    // Checkbox & Picklist click handling
    if (event->button() == Qt::LeftButton) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && m_spreadsheet) {
            auto cell = m_spreadsheet->getCellIfExists(idx.row(), idx.column());
            if (cell) {
                const QString& fmt = cell->getStyle().numberFormat;
                if (fmt == "Checkbox") {
                    QRect cellRect = visualRect(idx);
                    int boxSize = 14;
                    int cx = cellRect.left() + (cellRect.width() - boxSize) / 2;
                    int cy = cellRect.top() + (cellRect.height() - boxSize) / 2;
                    QRect hitArea(cx - 4, cy - 4, boxSize + 8, boxSize + 8);
                    if (hitArea.contains(event->pos())) {
                        toggleCheckbox(idx.row(), idx.column());
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

    QTableView::mousePressEvent(event);
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
}

void SpreadsheetView::paintEvent(QPaintEvent* event) {
    QTableView::paintEvent(event);

    // Draw fill handle on current selection
    QModelIndex current = currentIndex();
    if (current.isValid() && !m_fillDragging) {
        QRect selRect = getSelectionBoundingRect();
        if (!selRect.isNull()) {
            int handleSize = 7;
            m_fillHandleRect = QRect(
                selRect.right() - handleSize / 2,
                selRect.bottom() - handleSize / 2,
                handleSize, handleSize);

            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, false);
            painter.fillRect(m_fillHandleRect, ThemeManager::instance().currentTheme().focusBorderColor);
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
                QModelIndex idx = m_model->index(row, col);
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
    QRect hitRect = m_fillHandleRect.adjusted(-4, -4, 4, 4);
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
                if (m_spreadsheet) m_spreadsheet->setRowHeight(idx.row(), newSize);
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
        auto cell = m_spreadsheet->getCellIfExists(index.row(), index.column());
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
        CellAddress addr(idx.row(), idx.column());
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
        CellAddress addr(idx.row(), idx.column());
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

void SpreadsheetView::showPicklistPopup(const QModelIndex& index) {
    if (!m_spreadsheet || !index.isValid()) return;

    auto cell = m_spreadsheet->getCellIfExists(index.row(), index.column());
    if (!cell) return;

    auto* rule = const_cast<Spreadsheet::DataValidationRule*>(
        m_spreadsheet->getValidationAt(index.row(), index.column()));
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

    m_traceCell = CellAddress(current.row(), current.column());
    const auto& depGraph = m_spreadsheet->getDependencyGraph();
    m_tracedCells = depGraph.getDependencies(m_traceCell);
    m_showPrecedents = true;
    m_showDependents = false;
    viewport()->update();
}

void SpreadsheetView::traceDependents() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    m_traceCell = CellAddress(current.row(), current.column());
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
