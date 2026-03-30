#include "MainWindow.h"
#ifdef _WIN32
#include <process.h>
#else
#include <unistd.h>
#endif
#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "FormulaBar.h"
#include "Toolbar.h"
#include "FormatCellsDialog.h"
#include "FindReplaceDialog.h"
#include "GoToDialog.h"
#include "ConditionalFormatDialog.h"
#include "DataValidationDialog.h"
#include "ChatPanel.h"
#include "ChartWidget.h"
#ifdef HAS_DATA2APP
#include "NativeChartWidget.h"
#endif
#include "ShapeWidget.h"
#include "ChartDialog.h"
#include "ShapePropertiesDialog.h"
#include "ChartPropertiesPanel.h"
#include "../core/Spreadsheet.h"
#include "../core/Cell.h"
#include "../core/UndoManager.h"
#include "../core/CellRange.h"
#include "../services/DocumentService.h"
#include "../services/CsvService.h"
#include "../services/XlsxService.h"
#include "../services/AutoSaveService.h"
#include "../core/PivotEngine.h"
#include "PivotTableDialog.h"
#include "TemplateGallery.h"
#include "ImageWidget.h"
#include "SparklineDialog.h"
#include "../core/MacroEngine.h"
#include "MacroEditorDialog.h"
#include "GoalSeekDialog.h"
#include "Theme.h"
#include <cmath>
#include "../core/SparklineConfig.h"
#include "../core/DocumentTheme.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QMenuBar>
#include <QMenu>
#include <QFileDialog>
#include <QMessageBox>
#include <QStatusBar>
#include <QMimeData>
#include <QUrl>
#include <QFileInfo>
#include <QInputDialog>
#include <QLineEdit>
#include <QLabel>
#include <QDialogButtonBox>
#include <QTabBar>
#include <QToolButton>
#include <QScrollBar>
#include <QHeaderView>
#include <QDockWidget>
#include <QJsonObject>
#include <QTimer>
#include <QJsonArray>
#include <QProgressDialog>
#include <QtConcurrent>
#include <QFutureWatcher>
#include <QElapsedTimer>
#include <QApplication>
#include <QActionGroup>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent) {
    setWindowTitle("Nexel Pro");
    setGeometry(100, 100, 1280, 800);

#ifdef HAS_DATA2APP
    // Load chart backend preference (Data2App available)
    QSettings settings("Nexel", "Nexel");
    int backendVal = settings.value("chartBackend", static_cast<int>(ChartBackend::Data2App)).toInt();
    m_chartBackend = (backendVal == static_cast<int>(ChartBackend::QtPainter))
        ? ChartBackend::QtPainter : ChartBackend::Data2App;
#else
    QSettings settings("Nexel", "Nexel");
    m_chartBackend = ChartBackend::QtPainter;
#endif
    m_lazyLoadCharts = settings.value("lazyLoadCharts", false).toBool();

    // Initialize with one default sheet
    auto defaultSheet = std::make_shared<Spreadsheet>();
    defaultSheet->setSheetName("Sheet1");
    m_sheets.push_back(defaultSheet);
    m_activeSheetIndex = 0;

    QWidget* centralWidget = new QWidget(this);
    QVBoxLayout* layout = new QVBoxLayout(centralWidget);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(0);

    m_toolbar = new Toolbar(this);
    addToolBar(m_toolbar);
    addToolBarBreak(Qt::TopToolBarArea);
    QToolBar* toolbar2 = m_toolbar->createSecondaryToolbar(this);
    addToolBar(toolbar2);

    m_formulaBar = new FormulaBar(this);
    layout->addWidget(m_formulaBar);

    m_spreadsheetView = new SpreadsheetView(this);
    m_spreadsheetView->setSpreadsheet(m_sheets[0]);
    m_toolbar->setDocumentTheme(&m_sheets[0]->getDocumentTheme());
    layout->addWidget(m_spreadsheetView);

    // Sheet tab bar at bottom
    createSheetTabBar();
    layout->addWidget(m_bottomBar);

    setCentralWidget(centralWidget);

    // Chat assistant panel (dock widget on the right) — created before menu bar so menu can connect
    m_chatPanel = new ChatPanel(this);
    m_chatPanel->setSpreadsheet(m_sheets[0]);
    m_chatDock = new QDockWidget("Claude Assistant", this);
    m_chatDock->setWidget(m_chatPanel);
    m_chatDock->setFeatures(QDockWidget::DockWidgetClosable | QDockWidget::DockWidgetMovable);
    m_chatDock->setMinimumWidth(300);
    // Dock stylesheet is set by onThemeChanged()
    addDockWidget(Qt::RightDockWidgetArea, m_chatDock);
    m_chatDock->hide(); // Hidden by default, toggled from View menu

    // Chart properties panel (dock widget on the right)
    m_chartPropsPanel = new ChartPropertiesPanel(this);
    m_chartPropsDock = new QDockWidget(this);
    m_chartPropsDock->setTitleBarWidget(new QWidget()); // hide default title bar
    m_chartPropsDock->setWidget(m_chartPropsPanel);
    m_chartPropsDock->setFeatures(QDockWidget::DockWidgetClosable | QDockWidget::DockWidgetMovable);
    m_chartPropsDock->setMinimumWidth(260);
    m_chartPropsDock->setMaximumWidth(320);
    m_chartPropsDock->setStyleSheet("QDockWidget { border: none; }");
    addDockWidget(Qt::RightDockWidgetArea, m_chartPropsDock);
    m_chartPropsDock->hide();

    // Tabify docks so they don't overlap — they share the right area
    tabifyDockWidget(m_chatDock, m_chartPropsDock);

    connect(m_chartPropsPanel, &ChartPropertiesPanel::closeRequested, this, [this]() {
        m_chartPropsDock->hide();
    });

    // Macro engine
    m_macroEngine = new MacroEngine(this);
    m_macroEngine->setSpreadsheet(m_sheets[0]);
    connect(m_macroEngine, &MacroEngine::logMessage, this, [this](const QString& msg) {
        statusBar()->showMessage("Macro: " + msg, 3000);
    });

    // Connect macro engine to model and view for action recording
    m_spreadsheetView->setMacroEngine(m_macroEngine);
    if (m_spreadsheetView->getModel()) {
        m_spreadsheetView->getModel()->setMacroEngine(m_macroEngine);
    }

    createMenuBar();
    createStatusBar();
    connectSignals();

    // Deselect charts/shapes when clicking on the spreadsheet
    m_spreadsheetView->viewport()->installEventFilter(this);

    // Reposition all overlays from their cell anchors on every scroll (Excel-like anchoring).
    auto moveOverlays = [this]() {
        repositionAllOverlays();
        for (auto* chart : m_charts) chart->updateOverlayPosition();
        if (m_lazyLoadCharts) loadVisibleLazyCharts();
    };
    connect(m_spreadsheetView->verticalScrollBar(), &QScrollBar::valueChanged, this, moveOverlays);
    connect(m_spreadsheetView->horizontalScrollBar(), &QScrollBar::valueChanged, this, moveOverlays);
    // Also reposition on virtual scroll (scrollbar hidden in virtual mode)
    connect(m_spreadsheetView, &SpreadsheetView::virtualViewportChanged, this, moveOverlays);

    setAcceptDrops(true);

    // Global stylesheet is applied by ThemeManager::applyTheme() from main.cpp
    // Apply theme to sub-widgets
    onThemeChanged();

    // Auto-save with crash recovery
    m_autoSave = new AutoSaveService(this);
    m_autoSave->setMainWindow(this);
    m_autoSave->setInterval(60);
    m_autoSave->setEnabled(true);

    // Ensure cell A1 is focused on startup (deferred so widget has keyboard focus)
    QTimer::singleShot(0, this, [this]() {
        if (m_spreadsheetView && m_spreadsheetView->model()) {
            m_spreadsheetView->setCurrentIndex(m_spreadsheetView->model()->index(0, 0));
            m_spreadsheetView->setFocus();
        }
        // Check for crash recovery after UI is ready
        checkAutoSaveRecovery();
    });
}

void MainWindow::createSheetTabBar() {
    m_bottomBar = new QWidget(this);
    m_bottomBar->setFixedHeight(28);

    QHBoxLayout* bottomLayout = new QHBoxLayout(m_bottomBar);
    bottomLayout->setContentsMargins(4, 0, 0, 0);
    bottomLayout->setSpacing(2);

    // Add sheet button
    m_addSheetBtn = new QToolButton(m_bottomBar);
    m_addSheetBtn->setText("+");
    m_addSheetBtn->setFixedSize(24, 22);
    m_addSheetBtn->setToolTip("Add New Sheet");
    // Add sheet button stylesheet set by onThemeChanged()
    bottomLayout->addWidget(m_addSheetBtn);
    connect(m_addSheetBtn, &QToolButton::clicked, this, &MainWindow::onAddSheet);

    // Tab bar
    m_sheetTabBar = new QTabBar(m_bottomBar);
    m_sheetTabBar->setExpanding(false);
    m_sheetTabBar->setMovable(true);
    m_sheetTabBar->setTabsClosable(false);
    m_sheetTabBar->setDocumentMode(true);
    // Tab bar stylesheet set by onThemeChanged()

    // Populate tabs from m_sheets
    for (const auto& sheet : m_sheets) {
        m_sheetTabBar->addTab(sheet->getSheetName());
    }

    bottomLayout->addWidget(m_sheetTabBar);
    bottomLayout->addStretch();

    // Connect signals
    connect(m_sheetTabBar, &QTabBar::currentChanged, this, &MainWindow::onSheetTabChanged);
    connect(m_sheetTabBar, &QTabBar::tabBarDoubleClicked, this, &MainWindow::onSheetTabDoubleClicked);

    // Reorder underlying m_sheets when user drags tabs
    connect(m_sheetTabBar, &QTabBar::tabMoved, this, [this](int from, int to) {
        if (from < 0 || from >= static_cast<int>(m_sheets.size())) return;
        if (to < 0 || to >= static_cast<int>(m_sheets.size())) return;

        // Move the sheet in the vector
        auto sheet = m_sheets[from];
        m_sheets.erase(m_sheets.begin() + from);
        m_sheets.insert(m_sheets.begin() + to, sheet);

        // Update sheetIndex property on all overlays to match new positions
        for (auto* c : m_charts) {
            int si = c->property("sheetIndex").toInt();
            if (si == from) {
                c->setProperty("sheetIndex", to);
            } else if (from < to && si > from && si <= to) {
                c->setProperty("sheetIndex", si - 1);
            } else if (from > to && si >= to && si < from) {
                c->setProperty("sheetIndex", si + 1);
            }
        }
        for (auto* s : m_shapes) {
            int si = s->property("sheetIndex").toInt();
            if (si == from) {
                s->setProperty("sheetIndex", to);
            } else if (from < to && si > from && si <= to) {
                s->setProperty("sheetIndex", si - 1);
            } else if (from > to && si >= to && si < from) {
                s->setProperty("sheetIndex", si + 1);
            }
        }
        for (auto* img : m_images) {
            int si = img->property("sheetIndex").toInt();
            if (si == from) {
                img->setProperty("sheetIndex", to);
            } else if (from < to && si > from && si <= to) {
                img->setProperty("sheetIndex", si - 1);
            } else if (from > to && si >= to && si < from) {
                img->setProperty("sheetIndex", si + 1);
            }
        }

        // Update active sheet index to follow the current tab
        m_activeSheetIndex = m_sheetTabBar->currentIndex();
        setDirty();
    });

    // Right-click context menu on tab bar
    m_sheetTabBar->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_sheetTabBar, &QTabBar::customContextMenuRequested, this, &MainWindow::showSheetContextMenu);
}

void MainWindow::onSheetTabChanged(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    switchToSheet(index);
}

void MainWindow::switchToSheet(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    m_activeSheetIndex = index;
    m_spreadsheetView->clearChartRangeHighlight();
    m_spreadsheetView->setAllSheets(m_sheets);
    m_spreadsheetView->setSpreadsheet(m_sheets[index]);
    m_sheets[index]->getFormulaEngine().setAllSheets(&m_sheets);
    m_toolbar->setDocumentTheme(&m_sheets[index]->getDocumentTheme());
    m_spreadsheetView->refreshView();
    m_spreadsheetView->applyStoredDimensions();

    // Sync gridline visibility with sheet setting
    bool gridlines = m_sheets[index]->showGridlines();
    m_spreadsheetView->setGridlinesVisible(gridlines);
    if (m_gridlinesAction) m_gridlinesAction->setChecked(gridlines);

    if (m_chatPanel) m_chatPanel->setSpreadsheet(m_sheets[index]);

    // Show/hide charts, shapes, and images per sheet
    for (auto* c : m_charts) {
        c->setVisible(c->property("sheetIndex").toInt() == index);
    }
    for (auto* s : m_shapes) {
        s->setVisible(s->property("sheetIndex").toInt() == index);
    }
    for (auto* img : m_images) {
        img->setVisible(img->property("sheetIndex").toInt() == index);
    }
    applyZOrder();
    repositionAllOverlays();

    // Load any lazy-pending charts that are now visible on this sheet
    if (m_lazyLoadCharts) loadVisibleLazyCharts();

    // Update macro engine's spreadsheet reference
    if (m_macroEngine) {
        m_macroEngine->setSpreadsheet(m_sheets[index]);
        if (m_spreadsheetView->getModel()) {
            m_spreadsheetView->getModel()->setMacroEngine(m_macroEngine);
        }
    }

    // Reconnect dataChanged for live chart updates on the new model
    reconnectDataChanged();

    // Reset scroll position and focus to A1
    QModelIndex first = m_spreadsheetView->getModel()->index(0, 0);
    m_spreadsheetView->setCurrentIndex(first);
    m_spreadsheetView->scrollTo(first, QAbstractItemView::PositionAtTop);
    // Also reset horizontal scroll
    m_spreadsheetView->horizontalScrollBar()->setValue(0);
    m_spreadsheetView->verticalScrollBar()->setValue(0);
}

void MainWindow::onSheetTabDoubleClicked(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    QString currentName = m_sheetTabBar->tabText(index);
    bool ok;
    QString newName = QInputDialog::getText(this, "Rename Sheet",
        "Sheet name:", QLineEdit::Normal, currentName, &ok);
    if (ok && !newName.isEmpty()) {
        m_sheetTabBar->setTabText(index, newName);
        m_sheets[index]->setSheetName(newName);
        setDirty();
    }
}

void MainWindow::onAddSheet() {
    int num = nextSheetNumber();
    QString name = QString("Sheet%1").arg(num);
    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setSheetName(name);
    m_sheets.push_back(sheet);
    sheet->getFormulaEngine().setAllSheets(&m_sheets);
    m_sheetTabBar->addTab(name);
    m_sheetTabBar->setCurrentIndex(m_sheetTabBar->count() - 1);
    setDirty();
    statusBar()->showMessage("Added: " + name);
}

void MainWindow::onDeleteSheet() {
    if (m_sheets.size() <= 1) {
        QMessageBox::information(this, "Delete Sheet", "Cannot delete the only sheet.");
        return;
    }

    int idx = m_sheetTabBar->currentIndex();
    if (idx < 0) return;

    QString name = m_sheetTabBar->tabText(idx);
    if (QMessageBox::question(this, "Delete Sheet",
            QString("Delete sheet \"%1\"?").arg(name)) != QMessageBox::Yes) {
        return;
    }

    // Delete charts/shapes/images belonging to the deleted sheet
    for (int i = m_charts.size() - 1; i >= 0; --i) {
        if (m_charts[i]->property("sheetIndex").toInt() == idx) {
            removeOverlay(m_charts[i]);
            m_charts[i]->hide();
            m_charts[i]->deleteLater();
            m_charts.removeAt(i);
        }
    }
    for (int i = m_shapes.size() - 1; i >= 0; --i) {
        if (m_shapes[i]->property("sheetIndex").toInt() == idx) {
            removeOverlay(m_shapes[i]);
            m_shapes[i]->hide();
            m_shapes[i]->deleteLater();
            m_shapes.removeAt(i);
        }
    }
    for (int i = m_images.size() - 1; i >= 0; --i) {
        if (m_images[i]->property("sheetIndex").toInt() == idx) {
            removeOverlay(m_images[i]);
            m_images[i]->hide();
            m_images[i]->deleteLater();
            m_images.removeAt(i);
        }
    }
    // Shift sheetIndex down for charts/shapes/images on sheets after the deleted one
    for (auto* c : m_charts) {
        int si = c->property("sheetIndex").toInt();
        if (si > idx) c->setProperty("sheetIndex", si - 1);
    }
    for (auto* s : m_shapes) {
        int si = s->property("sheetIndex").toInt();
        if (si > idx) s->setProperty("sheetIndex", si - 1);
    }
    for (auto* img : m_images) {
        int si = img->property("sheetIndex").toInt();
        if (si > idx) img->setProperty("sheetIndex", si - 1);
    }

    // Block signals during removal to avoid triggering onSheetTabChanged prematurely
    m_sheetTabBar->blockSignals(true);
    m_sheetTabBar->removeTab(idx);
    m_sheets.erase(m_sheets.begin() + idx);
    m_sheetTabBar->blockSignals(false);

    int newIdx = qMin(idx, static_cast<int>(m_sheets.size()) - 1);
    m_sheetTabBar->setCurrentIndex(newIdx);
    switchToSheet(newIdx);
    setDirty();
    statusBar()->showMessage("Deleted: " + name);
}

void MainWindow::onDuplicateSheet() {
    int idx = m_sheetTabBar->currentIndex();
    if (idx < 0 || idx >= static_cast<int>(m_sheets.size())) return;

    auto source = m_sheets[idx];
    auto copy = std::make_shared<Spreadsheet>();
    copy->setSheetName(source->getSheetName() + " (Copy)");
    copy->setAutoRecalculate(false);

    // Copy all cells
    source->forEachCell([&](int row, int col, const Cell& cell) {
        CellAddress addr(row, col);
        auto val = source->getCellValue(addr);
        if (val.isValid() && !val.toString().isEmpty()) {
            copy->setCellValue(addr, val);
        }
        auto srcCell = source->getCell(addr);
        auto dstCell = copy->getCell(addr);
        dstCell->setStyle(srcCell->getStyle());
    });

    copy->setRowCount(source->getRowCount());
    copy->setColumnCount(source->getColumnCount());
    copy->setAutoRecalculate(true);

    m_sheets.insert(m_sheets.begin() + idx + 1, copy);
    copy->getFormulaEngine().setAllSheets(&m_sheets);
    m_sheetTabBar->insertTab(idx + 1, copy->getSheetName());
    m_sheetTabBar->setCurrentIndex(idx + 1);
    setDirty();
    statusBar()->showMessage("Duplicated sheet");
}

void MainWindow::showSheetContextMenu(const QPoint& pos) {
    int tabIdx = m_sheetTabBar->tabAt(pos);
    if (tabIdx < 0) tabIdx = m_sheetTabBar->currentIndex();

    QMenu menu(this);
    menu.addAction("Insert New Sheet", this, &MainWindow::onAddSheet);
    menu.addAction("Duplicate Sheet", this, &MainWindow::onDuplicateSheet);
    menu.addSeparator();
    menu.addAction("Rename Sheet", this, [this, tabIdx]() {
        onSheetTabDoubleClicked(tabIdx);
    });
    menu.addAction("Delete Sheet", this, &MainWindow::onDeleteSheet);
    menu.exec(m_sheetTabBar->mapToGlobal(pos));
}

int MainWindow::nextSheetNumber() const {
    int maxNum = 0;
    for (const auto& sheet : m_sheets) {
        QString name = sheet->getSheetName();
        if (name.startsWith("Sheet")) {
            bool ok;
            int num = name.mid(5).toInt(&ok);
            if (ok && num > maxNum) maxNum = num;
        }
    }
    return maxNum + 1;
}

void MainWindow::setSheets(const std::vector<std::shared_ptr<Spreadsheet>>& sheets) {
    // Clear existing charts, shapes, and images from the viewport
    for (auto* c : m_charts) { c->hide(); c->deleteLater(); }
    m_charts.clear();
    for (auto* s : m_shapes) { s->hide(); s->deleteLater(); }
    m_shapes.clear();
    for (auto* img : m_images) { img->hide(); img->deleteLater(); }
    m_images.clear();
    m_overlayZOrder.clear();
    m_overlayGroups.clear();

    // Move old sheets to background thread for deferred destruction
    // (destroying 9M+ Cell objects synchronously would freeze the UI)
    if (!m_sheets.empty()) {
        auto oldSheets = std::make_shared<std::vector<std::shared_ptr<Spreadsheet>>>(
            std::move(m_sheets));
        QtConcurrent::run([oldSheets]() { /* destructor runs here */ });
    }

    m_sheets = sheets;
    m_activeSheetIndex = 0;

    // Wire up cross-sheet reference support for all sheets
    for (auto& sheet : m_sheets) {
        sheet->getFormulaEngine().setAllSheets(&m_sheets);
    }

    // Rebuild tab bar
    m_sheetTabBar->blockSignals(true);
    while (m_sheetTabBar->count() > 0) {
        m_sheetTabBar->removeTab(0);
    }
    for (const auto& sheet : m_sheets) {
        m_sheetTabBar->addTab(sheet->getSheetName());
    }
    m_sheetTabBar->setCurrentIndex(0);
    m_sheetTabBar->blockSignals(false);

    switchToSheet(0);
}

void MainWindow::createMenuBar() {
    QMenuBar* menuBar = new QMenuBar(this);
    setMenuBar(menuBar);

    QMenu* fileMenu = menuBar->addMenu("&File");
    fileMenu->addAction("&New", this, &MainWindow::onNewDocument, QKeySequence::New);
    fileMenu->addAction("New from &Template...", this, &MainWindow::onTemplateGallery);
    fileMenu->addAction("&Open", this, &MainWindow::onOpenDocument, QKeySequence::Open);
    fileMenu->addAction("&Save", this, &MainWindow::onSaveDocument, QKeySequence::Save);
    fileMenu->addAction("Save &As", this, &MainWindow::onSaveAs, QKeySequence::SaveAs);
    fileMenu->addAction("&Rename Document...", this, [this]() {
        QString baseName = "Untitled";
        if (!m_currentFilePath.isEmpty()) {
            baseName = QFileInfo(m_currentFilePath).completeBaseName();
        } else {
            QString title = windowTitle();
            if (title.contains(" - "))
                baseName = title.section(" - ", 1);
        }
        bool ok;
        QString newName = QInputDialog::getText(this, "Rename Document",
            "Document name:", QLineEdit::Normal, baseName, &ok);
        if (ok && !newName.isEmpty()) {
            setWindowTitle("Nexel Pro - " + newName);
            statusBar()->showMessage("Renamed to: " + newName);
        }
    });
    fileMenu->addSeparator();
    fileMenu->addAction("&Import CSV...", this, &MainWindow::onImportCsv);
    fileMenu->addAction("&Export CSV...", this, &MainWindow::onExportCsv);
    fileMenu->addSeparator();
    fileMenu->addAction("E&xit", this, &QWidget::close, QKeySequence::Quit);

    QMenu* editMenu = menuBar->addMenu("&Edit");
    editMenu->addAction("&Undo", QKeySequence::Undo, this, &MainWindow::onUndo);
    auto* redoAction = editMenu->addAction("&Redo", QKeySequence::Redo, this, &MainWindow::onRedo);
    // Add Ctrl+Y as additional redo shortcut (Cmd+Y on Mac)
    redoAction->setShortcuts({QKeySequence::Redo, QKeySequence(Qt::CTRL | Qt::Key_Y)});
    editMenu->addSeparator();
    editMenu->addAction("Cu&t", this, &MainWindow::onCut, QKeySequence::Cut);
    editMenu->addAction("&Copy", this, &MainWindow::onCopy, QKeySequence::Copy);
    editMenu->addAction("&Paste", this, &MainWindow::onPaste, QKeySequence::Paste);
    editMenu->addAction("Paste &Special...", this, [this]() {
        if (m_spreadsheetView) m_spreadsheetView->pasteSpecial();
    }, QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_V));
    editMenu->addAction("&Delete", this, &MainWindow::onDelete, QKeySequence::Delete);
    editMenu->addSeparator();
    editMenu->addAction("Select &All", this, &MainWindow::onSelectAll, QKeySequence::SelectAll);
    editMenu->addSeparator();
    editMenu->addAction("&Find and Replace...", this, &MainWindow::onFindReplace, QKeySequence::Find);
    editMenu->addAction("&Go To...", this, &MainWindow::onGoTo,
                         QKeySequence(Qt::CTRL | Qt::Key_G));

    QMenu* formatMenu = menuBar->addMenu("F&ormat");
    formatMenu->addAction("Format &Cells...", this, &MainWindow::onFormatCells,
                          QKeySequence(Qt::CTRL | Qt::Key_1));
    formatMenu->addSeparator();
    formatMenu->addAction("&Bold", m_spreadsheetView, &SpreadsheetView::applyBold,
                          QKeySequence(Qt::CTRL | Qt::Key_B));
    formatMenu->addAction("&Italic", m_spreadsheetView, &SpreadsheetView::applyItalic,
                          QKeySequence(Qt::CTRL | Qt::Key_I));
    formatMenu->addAction("&Underline", m_spreadsheetView, &SpreadsheetView::applyUnderline,
                          QKeySequence(Qt::CTRL | Qt::Key_U));
    formatMenu->addSeparator();
    formatMenu->addAction("&Conditional Formatting...", this, &MainWindow::onConditionalFormat);
    formatMenu->addSeparator();
    formatMenu->addAction("Autofit Column Width", m_spreadsheetView,
                          &SpreadsheetView::autofitSelectedColumns);
    formatMenu->addAction("Autofit Row Height", m_spreadsheetView,
                          &SpreadsheetView::autofitSelectedRows);

    // ===== Insert Menu =====
    QMenu* insertMenu = menuBar->addMenu("&Insert");
    insertMenu->addAction("&Chart...", this, &MainWindow::onInsertChart,
                          QKeySequence(Qt::ALT | Qt::Key_F1));
    insertMenu->addAction("&Shape...", this, &MainWindow::onInsertShape);
    insertMenu->addAction("&Image...", this, &MainWindow::onInsertImage);
    insertMenu->addAction("Spark&line...", this, &MainWindow::onInsertSparkline);
    insertMenu->addSeparator();
    insertMenu->addAction("&Hyperlink...", m_spreadsheetView, &SpreadsheetView::insertOrEditHyperlink,
                          QKeySequence(Qt::CTRL | Qt::Key_K));
    insertMenu->addSeparator();
    insertMenu->addAction("&Row Above", m_spreadsheetView, &SpreadsheetView::insertEntireRow);
    insertMenu->addAction("&Column Left", m_spreadsheetView, &SpreadsheetView::insertEntireColumn);

    // ===== Arrange Menu =====
    QMenu* arrangeMenu = menuBar->addMenu("&Arrange");
    arrangeMenu->addAction("Bring to &Front", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_BracketRight),
        this, [this]() {
            auto sel = selectedOverlays();
            for (auto* w : sel) bringToFront(w);
        });
    arrangeMenu->addAction("Send to &Back", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_BracketLeft),
        this, [this]() {
            auto sel = selectedOverlays();
            for (auto* w : sel) sendToBack(w);
        });
    arrangeMenu->addAction("Bring For&ward", QKeySequence(Qt::CTRL | Qt::Key_BracketRight),
        this, [this]() {
            auto sel = selectedOverlays();
            for (auto* w : sel) bringForward(w);
        });
    arrangeMenu->addAction("Send Back&ward", QKeySequence(Qt::CTRL | Qt::Key_BracketLeft),
        this, [this]() {
            auto sel = selectedOverlays();
            for (auto* w : sel) sendBackward(w);
        });
    arrangeMenu->addSeparator();
    arrangeMenu->addAction("&Group", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_G),
        this, &MainWindow::groupSelectedOverlays);
    arrangeMenu->addAction("&Ungroup", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::ALT | Qt::Key_G),
        this, &MainWindow::ungroupSelectedOverlays);

    QMenu* dataMenu = menuBar->addMenu("&Data");
    dataMenu->addAction("Sort &Ascending", m_spreadsheetView, &SpreadsheetView::sortAscending);
    dataMenu->addAction("Sort &Descending", m_spreadsheetView, &SpreadsheetView::sortDescending);
    dataMenu->addAction("Custom &Sort...", m_spreadsheetView, &SpreadsheetView::showSortDialog);
    dataMenu->addSeparator();
    dataMenu->addAction("&Data Validation...", this, &MainWindow::onDataValidation);
    dataMenu->addSeparator();
    dataMenu->addAction("Create &Pivot Table...", this, &MainWindow::onCreatePivotTable);
    dataMenu->addAction("&Edit Pivot Table...", this, &MainWindow::onEditPivotTable);
    dataMenu->addAction("&Refresh Pivot Table", this, &MainWindow::onRefreshPivotTable,
                        QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_R));
    dataMenu->addSeparator();
    dataMenu->addAction("&Recalculate All", this, [this]() {
        if (!m_sheets.empty() && m_activeSheetIndex < (int)m_sheets.size()) {
            m_sheets[m_activeSheetIndex]->recalculateAll();
            m_spreadsheetView->refreshView();
            statusBar()->showMessage("Recalculated all formulas");
        }
    }, QKeySequence(Qt::Key_F9));
    dataMenu->addSeparator();
    QAction* highlightAction = dataMenu->addAction("&Circle Invalid Data", this, &MainWindow::onHighlightInvalidCells);
    highlightAction->setCheckable(true);
    dataMenu->addSeparator();
    dataMenu->addAction("&Goal Seek...", this, [this]() {
        GoalSeekDialog dialog(this);
        if (dialog.exec() != QDialog::Accepted) return;

        CellAddress setCell = CellAddress::fromString(dialog.getSetCell());
        CellAddress byChangingCell = CellAddress::fromString(dialog.getByChangingCell());
        double targetValue = dialog.getToValue();

        if (setCell.row < 0 || setCell.col < 0 || byChangingCell.row < 0 || byChangingCell.col < 0) {
            QMessageBox::warning(this, "Goal Seek", "Invalid cell reference.");
            return;
        }

        auto& sheet = m_sheets[m_activeSheetIndex];

        // Save original value for undo
        QVariant originalValue = sheet->getCellValue(byChangingCell);

        // Secant method with bisection fallback
        auto evaluate = [&](double x) -> double {
            sheet->setCellValue(byChangingCell, x);
            return sheet->getCellValue(setCell).toDouble() - targetValue;
        };

        double x0 = originalValue.toDouble();
        double x1 = (x0 == 0.0) ? 1.0 : x0 * 1.1;
        double f0 = evaluate(x0);
        double f1 = evaluate(x1);

        bool converged = false;
        const int maxIter = 100;
        const double tol = 1e-7;

        // Secant method
        for (int i = 0; i < maxIter && !converged; ++i) {
            if (std::abs(f1) < tol) {
                converged = true;
                break;
            }
            if (std::abs(f1 - f0) < 1e-15) break; // avoid division by zero

            double x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
            x0 = x1;
            f0 = f1;
            x1 = x2;
            f1 = evaluate(x1);
        }

        // Bisection fallback if secant didn't converge
        if (!converged && std::abs(f1) >= tol) {
            // Try to find a bracket
            double lo = x0 - 1000, hi = x0 + 1000;
            double flo = evaluate(lo), fhi = evaluate(hi);
            if (flo * fhi < 0) {
                for (int i = 0; i < maxIter && !converged; ++i) {
                    double mid = (lo + hi) / 2.0;
                    double fmid = evaluate(mid);
                    if (std::abs(fmid) < tol) {
                        x1 = mid;
                        converged = true;
                    } else if (flo * fmid < 0) {
                        hi = mid;
                        fhi = fmid;
                    } else {
                        lo = mid;
                        flo = fmid;
                    }
                }
                if (!converged) x1 = (lo + hi) / 2.0;
            }
        }

        if (converged || std::abs(evaluate(x1)) < tol * 100) {
            sheet->setCellValue(byChangingCell, x1);
            m_spreadsheetView->refreshView();
            QMessageBox::information(this, "Goal Seek",
                QString("Goal Seek found a solution.\n%1 = %2")
                    .arg(dialog.getByChangingCell())
                    .arg(x1, 0, 'g', 10));
        } else {
            // Restore original value
            sheet->setCellValue(byChangingCell, originalValue);
            m_spreadsheetView->refreshView();
            QMessageBox::warning(this, "Goal Seek",
                "Goal Seek could not find a solution.");
        }
    });
    dataMenu->addSeparator();
    QMenu* groupMenu = dataMenu->addMenu("&Group && Outline");
    groupMenu->addAction("&Group Rows", QKeySequence(Qt::ALT | Qt::SHIFT | Qt::Key_Right),
        m_spreadsheetView, &SpreadsheetView::groupSelectedRows);
    groupMenu->addAction("&Ungroup Rows", QKeySequence(Qt::ALT | Qt::SHIFT | Qt::Key_Left),
        m_spreadsheetView, &SpreadsheetView::ungroupSelectedRows);
    groupMenu->addSeparator();
    groupMenu->addAction("Group &Columns", m_spreadsheetView, &SpreadsheetView::groupSelectedColumns);
    groupMenu->addAction("Ungroup C&olumns", m_spreadsheetView, &SpreadsheetView::ungroupSelectedColumns);

    dataMenu->addSeparator();
    m_protectSheetAction = dataMenu->addAction("&Protect Sheet...", this, &MainWindow::onProtectSheet);

    // ===== Formulas Menu =====
    QMenu* formulasMenu = menuBar->addMenu("&Formulas");
    formulasMenu->addAction("Trace &Precedents", m_spreadsheetView,
                            &SpreadsheetView::tracePrecedents,
                            QKeySequence(Qt::ALT | Qt::SHIFT | Qt::Key_BracketLeft));
    formulasMenu->addAction("Trace &Dependents", m_spreadsheetView,
                            &SpreadsheetView::traceDependents,
                            QKeySequence(Qt::ALT | Qt::SHIFT | Qt::Key_BracketRight));
    formulasMenu->addAction("&Remove Arrows", m_spreadsheetView,
                            &SpreadsheetView::clearTraceArrows,
                            QKeySequence(Qt::ALT | Qt::SHIFT | Qt::Key_Backslash));

    // === Page Layout menu (document themes) ===
    QMenu* layoutMenu = menuBar->addMenu("Page &Layout");
    QMenu* docThemeMenu = layoutMenu->addMenu("&Themes");
    QActionGroup* docThemeGroup = new QActionGroup(this);
    docThemeGroup->setExclusive(true);
    auto docThemes = getBuiltinDocumentThemes();
    for (const auto& dt : docThemes) {
        QAction* action = docThemeMenu->addAction(dt.displayName);
        action->setCheckable(true);
        action->setChecked(dt.id == m_sheets[0]->getDocumentTheme().id);
        action->setData(dt.id);
        docThemeGroup->addAction(action);
        connect(action, &QAction::triggered, this, [this, dt]() {
            for (auto& sheet : m_sheets) {
                sheet->setDocumentTheme(dt);
            }
            m_toolbar->setDocumentTheme(&m_sheets[m_activeSheetIndex]->getDocumentTheme());
            m_spreadsheetView->refreshView();
            refreshActiveCharts();
            setDirty();
        });
    }

    QMenu* viewMenu = menuBar->addMenu("&View");
    m_gridlinesAction = viewMenu->addAction("Show &Gridlines");
    m_gridlinesAction->setCheckable(true);
    m_gridlinesAction->setChecked(true);
    connect(m_gridlinesAction, &QAction::toggled, this, [this](bool checked) {
        if (!m_sheets.empty() && m_activeSheetIndex < (int)m_sheets.size()) {
            m_sheets[m_activeSheetIndex]->setShowGridlines(checked);
        }
        m_spreadsheetView->setGridlinesVisible(checked);
    });
    viewMenu->addAction("&Freeze Panes", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_F),
                         this, &MainWindow::onFreezePane);
    viewMenu->addSeparator();
    QAction* chatAction = viewMenu->addAction("&Claude Assistant");
    chatAction->setCheckable(true);
    chatAction->setChecked(false);
    connect(chatAction, &QAction::toggled, this, [this](bool checked) {
        if (checked) {
            m_chatDock->show();
        } else {
            m_chatDock->hide();
        }
    });
    connect(m_chatDock, &QDockWidget::visibilityChanged, chatAction, &QAction::setChecked);

    // ===== Theme Submenu =====
    viewMenu->addSeparator();
    QMenu* themeMenu = viewMenu->addMenu("&Theme");
    QActionGroup* themeGroup = new QActionGroup(this);
    themeGroup->setExclusive(true);
    auto themes = ThemeManager::instance().availableThemes();
    for (const auto& theme : themes) {
        QAction* action = themeMenu->addAction(theme.displayName);
        action->setCheckable(true);
        action->setChecked(theme.id == ThemeManager::instance().currentTheme().id);
        action->setData(theme.id);
        themeGroup->addAction(action);
        connect(action, &QAction::triggered, this, [this, id = theme.id, themeGroup]() {
            ThemeManager::instance().setTheme(id);
            ThemeManager::instance().applyTheme(this);
            onThemeChanged();
        });
    }

    // Dark Mode quick toggle
    QAction* darkModeAction = viewMenu->addAction("&Dark Mode");
    darkModeAction->setCheckable(true);
    darkModeAction->setChecked(ThemeManager::instance().isDarkTheme());
    connect(darkModeAction, &QAction::toggled, this, [this, themeGroup, darkModeAction](bool checked) {
        QString targetId = checked ? "dark_mode" : "nexel_green";
        ThemeManager::instance().setTheme(targetId);
        ThemeManager::instance().applyTheme(this);
        onThemeChanged();
        // Update theme submenu checkmarks
        for (auto* action : themeGroup->actions()) {
            action->setChecked(action->data().toString() == targetId);
        }
    });
    // Keep dark mode toggle in sync when theme submenu changes
    connect(&ThemeManager::instance(), &ThemeManager::themeChanged, darkModeAction,
            [darkModeAction](const NexelTheme& t) {
        darkModeAction->setChecked(t.id == "dark_mode");
    });

    // ===== Tools Menu =====
    QMenu* toolsMenu = menuBar->addMenu("&Tools");
    toolsMenu->addAction("Macro &Editor...", this, &MainWindow::onMacroEditor,
                         QKeySequence(Qt::CTRL | Qt::ALT | Qt::Key_M));
    toolsMenu->addAction("Run &Last Macro", this, &MainWindow::onRunLastMacro,
                         QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_M));

#ifdef HAS_DATA2APP
    toolsMenu->addSeparator();
    m_nativeChartAction = toolsMenu->addAction("Use Qt Charts (Legacy)");
    m_nativeChartAction->setCheckable(true);
    m_nativeChartAction->setChecked(m_chartBackend == ChartBackend::QtPainter);
    connect(m_nativeChartAction, &QAction::toggled, this, [this](bool checked) {
        m_chartBackend = checked ? ChartBackend::QtPainter : ChartBackend::Data2App;
        QSettings settings("Nexel", "Nexel");
        settings.setValue("chartBackend", static_cast<int>(m_chartBackend));
    });
#endif

    toolsMenu->addSeparator();
    auto* lazyAction = toolsMenu->addAction("Lazy Load Charts");
    lazyAction->setCheckable(true);
    lazyAction->setChecked(m_lazyLoadCharts);
    connect(lazyAction, &QAction::toggled, this, [this](bool checked) {
        m_lazyLoadCharts = checked;
        QSettings settings("Nexel", "Nexel");
        settings.setValue("lazyLoadCharts", checked);
        if (!checked) loadVisibleLazyCharts();  // load any pending charts immediately
    });
}

void MainWindow::createToolBar() {}

void MainWindow::createStatusBar() {
    statusBar()->showMessage("Ready");

    // Permanent stats widgets on the right side (Excel-style)
    QString statStyle = "QLabel { color: #555; font-size: 11px; padding: 0 8px; }";

    m_statusAvgLabel = new QLabel("");
    m_statusAvgLabel->setStyleSheet(statStyle);
    statusBar()->addPermanentWidget(m_statusAvgLabel);

    m_statusCountLabel = new QLabel("");
    m_statusCountLabel->setStyleSheet(statStyle);
    statusBar()->addPermanentWidget(m_statusCountLabel);

    m_statusSumLabel = new QLabel("");
    m_statusSumLabel->setStyleSheet(statStyle);
    statusBar()->addPermanentWidget(m_statusSumLabel);
}

void MainWindow::onThemeChanged() {
    const auto& tm = ThemeManager::instance();
    const auto& t = tm.currentTheme();

    // Bottom bar
    m_bottomBar->setStyleSheet(tm.buildBottomBarStylesheet());

    // Tab bar
    m_sheetTabBar->setStyleSheet(tm.buildTabBarStylesheet());

    // Add sheet button
    m_addSheetBtn->setStyleSheet(tm.buildAddSheetBtnStylesheet());

    // Chat dock
    if (m_chatDock) {
        m_chatDock->setStyleSheet(QString(
            "QDockWidget { border: none; }"
            "QDockWidget::title { background: %1; color: %2; padding: 6px; font-weight: bold; text-align: center; }"
            "QDockWidget::close-button { background: transparent; }"
        ).arg(t.dockTitleBackground.name(), t.dockTitleText.name()));
    }

    // Chart properties dock
    if (m_chartPropsDock) {
        m_chartPropsDock->setStyleSheet("QDockWidget { border: none; }");
    }

    // Cascade to child widgets
    if (m_toolbar) m_toolbar->onThemeChanged();
    if (m_formulaBar) m_formulaBar->onThemeChanged();
    if (m_spreadsheetView) m_spreadsheetView->onThemeChanged();
    if (m_chatPanel) m_chatPanel->onThemeChanged();
}

void MainWindow::connectSignals() {
    connect(m_toolbar, &Toolbar::newDocument, this, &MainWindow::onNewDocument);
    connect(m_toolbar, &Toolbar::saveDocument, this, &MainWindow::onSaveDocument);
    connect(m_toolbar, &Toolbar::undo, this, &MainWindow::onUndo);
    connect(m_toolbar, &Toolbar::redo, this, &MainWindow::onRedo);

    connect(m_toolbar, &Toolbar::bold, this, [this]() {
        if (m_selectedChart) {
            auto cfg = m_selectedChart->config();
            cfg.titleBold = !cfg.titleBold;
            m_selectedChart->setConfig(cfg);
            setDirty();
        } else {
            m_spreadsheetView->applyBold();
        }
    });
    connect(m_toolbar, &Toolbar::italic, this, [this]() {
        if (m_selectedChart) {
            auto cfg = m_selectedChart->config();
            cfg.titleItalic = !cfg.titleItalic;
            m_selectedChart->setConfig(cfg);
            setDirty();
        } else {
            m_spreadsheetView->applyItalic();
        }
    });
    connect(m_toolbar, &Toolbar::underline, m_spreadsheetView, &SpreadsheetView::applyUnderline);
    connect(m_toolbar, &Toolbar::strikethrough, m_spreadsheetView, &SpreadsheetView::applyStrikethrough);
    connect(m_toolbar, &Toolbar::fontFamilyChanged, m_spreadsheetView, &SpreadsheetView::applyFontFamily);
    connect(m_toolbar, &Toolbar::fontSizeChanged, m_spreadsheetView, &SpreadsheetView::applyFontSize);

    connect(m_toolbar, &Toolbar::foregroundColorChanged, this, [this](const QString& colorStr, const QColor& displayColor) {
        if (m_selectedChart) {
            auto cfg = m_selectedChart->config();
            cfg.titleColor = displayColor;
            m_selectedChart->setConfig(cfg);
            setDirty();
        } else {
            m_spreadsheetView->applyForegroundColor(colorStr);
        }
    });
    connect(m_toolbar, &Toolbar::backgroundColorChanged, this, [this](const QString& colorStr, const QColor& displayColor) {
        if (m_selectedChart) {
            auto cfg = m_selectedChart->config();
            cfg.backgroundColor = displayColor;
            m_selectedChart->setConfig(cfg);
            setDirty();
        } else {
            m_spreadsheetView->applyBackgroundColor(colorStr);
        }
    });

    connect(m_toolbar, &Toolbar::hAlignChanged, m_spreadsheetView, &SpreadsheetView::applyHAlign);
    connect(m_toolbar, &Toolbar::vAlignChanged, m_spreadsheetView, &SpreadsheetView::applyVAlign);

    connect(m_toolbar, &Toolbar::thousandSeparatorToggled, m_spreadsheetView, &SpreadsheetView::applyThousandSeparator);
    connect(m_toolbar, &Toolbar::numberFormatChanged, m_spreadsheetView, &SpreadsheetView::applyNumberFormat);
    connect(m_toolbar, &Toolbar::dateFormatSelected, m_spreadsheetView, &SpreadsheetView::applyDateFormat);
    connect(m_toolbar, &Toolbar::currencyFormatSelected, m_spreadsheetView, &SpreadsheetView::applyCurrencyFormat);
    connect(m_toolbar, &Toolbar::accountingFormatSelected, m_spreadsheetView, &SpreadsheetView::applyAccountingFormat);
    connect(m_toolbar, &Toolbar::increaseDecimals, m_spreadsheetView, &SpreadsheetView::increaseDecimals);
    connect(m_toolbar, &Toolbar::decreaseDecimals, m_spreadsheetView, &SpreadsheetView::decreaseDecimals);
    connect(m_toolbar, &Toolbar::formatCellsRequested, this, &MainWindow::onFormatCells);

    connect(m_toolbar, &Toolbar::formatPainterToggled, m_spreadsheetView, &SpreadsheetView::activateFormatPainter);

    connect(m_toolbar, &Toolbar::sortAscending, m_spreadsheetView, &SpreadsheetView::sortAscending);
    connect(m_toolbar, &Toolbar::sortDescending, m_spreadsheetView, &SpreadsheetView::sortDescending);
    connect(m_toolbar, &Toolbar::filterToggled, m_spreadsheetView, &SpreadsheetView::toggleAutoFilter);

    connect(m_toolbar, &Toolbar::tableStyleSelected, m_spreadsheetView, &SpreadsheetView::applyTableStyle);

    connect(m_toolbar, &Toolbar::borderStyleSelected, this, [this](const QString& type, const QColor& color, int width, int penStyle) {
        m_spreadsheetView->applyBorderStyle(type, color, width, penStyle);
    });
    connect(m_toolbar, &Toolbar::mergeCellsRequested, m_spreadsheetView, &SpreadsheetView::mergeCells);
    connect(m_toolbar, &Toolbar::unmergeCellsRequested, m_spreadsheetView, &SpreadsheetView::unmergeCells);
    connect(m_toolbar, &Toolbar::increaseIndent, m_spreadsheetView, &SpreadsheetView::increaseIndent);
    connect(m_toolbar, &Toolbar::decreaseIndent, m_spreadsheetView, &SpreadsheetView::decreaseIndent);
    connect(m_toolbar, &Toolbar::textRotationChanged, m_spreadsheetView, &SpreadsheetView::applyTextRotation);
    connect(m_toolbar, &Toolbar::textOverflowChanged, m_spreadsheetView, &SpreadsheetView::applyTextOverflow);

    connect(m_toolbar, &Toolbar::conditionalFormatRequested, this, &MainWindow::onConditionalFormat);
    connect(m_toolbar, &Toolbar::dataValidationRequested, this, &MainWindow::onDataValidation);

    // Chart and shape insertion from toolbar
    connect(m_toolbar, &Toolbar::insertChartRequested, this, &MainWindow::onInsertChart);
    connect(m_toolbar, &Toolbar::insertShapeRequested, this, &MainWindow::onInsertShape);

    // Checkbox insertion
    connect(m_toolbar, &Toolbar::insertCheckboxRequested, this, [this]() {
        m_spreadsheetView->insertCheckbox();
        setDirty();
    });

    // Picklist manager
    connect(m_toolbar, &Toolbar::managePicklistsRequested, this, [this]() {
        m_spreadsheetView->openPicklistManagerDialog();
    });

    // Picklist insertion — open create dialog (same as edit dialog)
    connect(m_toolbar, &Toolbar::insertPicklistRequested, this, [this]() {
        m_spreadsheetView->showCreatePicklistDialog();
        setDirty();
    });

    // Chat assistant toggle
    connect(m_toolbar, &Toolbar::chatToggleRequested, this, [this]() {
        if (m_chatDock->isVisible()) {
            m_chatDock->hide();
        } else {
            m_chatDock->show();
            m_chatDock->raise();
        }
    });

    // Chat NLP actions
    connect(m_chatPanel, &ChatPanel::executeActions, this, &MainWindow::onChatActions);

    connect(m_spreadsheetView, &SpreadsheetView::formatCellsRequested, this, &MainWindow::onFormatCells);
    connect(m_spreadsheetView, &SpreadsheetView::pivotFilterChanged, this, &MainWindow::onPivotFilterChanged);
    connect(m_spreadsheetView, &SpreadsheetView::requestSwitchToSheet, this, [this](int index) {
        if (index >= 0 && index < static_cast<int>(m_sheets.size())) {
            m_sheetTabBar->setCurrentIndex(index);
            switchToSheet(index);
        }
    });

    connect(m_spreadsheetView, &SpreadsheetView::cellSelected,
            this, [this](int row, int col, const QString& content, const QString& address) {
        // Don't update formula bar if we're in formula editing mode (would overwrite the formula)
        if (m_spreadsheetView->isFormulaEditMode() || m_formulaBar->isFormulaEditing()) {
            return;
        }
        m_formulaBar->setCellAddress(address);
        m_formulaBar->setCellContent(content);

        // Sync toolbar to reflect the focused cell's formatting
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) {
            auto cell = spreadsheet->getCellIfExists(row, col);
            if (cell) {
                m_toolbar->syncToStyle(cell->getStyle());
            } else if (spreadsheet->hasDefaultCellStyle()) {
                m_toolbar->syncToStyle(spreadsheet->getDefaultCellStyle());
            } else {
                m_toolbar->syncToStyle(Cell::defaultStyle());
            }
        }

        // Update status bar with selection summary (SUM, AVERAGE, COUNT like Excel)
        updateStatusBarSummary();
    });

    // Selection change also updates status bar summary
    connect(m_spreadsheetView->selectionModel(), &QItemSelectionModel::selectionChanged,
            this, [this]() { updateStatusBarSummary(); });

    // Formula bar -> cell reference insertion
    connect(m_formulaBar, &FormulaBar::formulaEditModeChanged,
            m_spreadsheetView, &SpreadsheetView::setFormulaEditMode);

    // When SpreadsheetView inserts a cell reference via click, insert it into formula bar
    connect(m_spreadsheetView, &SpreadsheetView::cellReferenceInserted,
            m_formulaBar, &FormulaBar::insertText);
    // When SpreadsheetView replaces a reference during range drag, update formula bar
    connect(m_spreadsheetView, &SpreadsheetView::cellReferenceReplaced,
            m_formulaBar, &FormulaBar::replaceLastInsertedText);

    connect(m_formulaBar, &FormulaBar::contentEdited,
            this, [this](const QString& content) {
        auto index = m_spreadsheetView->currentIndex();
        if (index.isValid()) {
            auto model = m_spreadsheetView->getModel();
            if (model) {
                model->setData(index, content);
            }
        }
    });

    // Enter in formula bar: commit the value and move focus back to grid, move down
    connect(m_formulaBar, &FormulaBar::returnPressed, this, [this]() {
        auto index = m_spreadsheetView->currentIndex();
        if (index.isValid()) {
            auto model = m_spreadsheetView->getModel();
            if (model) {
                model->setData(index, m_formulaBar->getContent());
            }
        }
        // Turn off formula edit mode
        m_spreadsheetView->setFormulaEditMode(false);
        // Move to next row (like pressing Enter in a cell)
        int newRow = index.row() + 1;
        if (newRow < m_spreadsheetView->model()->rowCount()) {
            QModelIndex next = m_spreadsheetView->model()->index(newRow, index.column());
            m_spreadsheetView->setCurrentIndex(next);
            m_spreadsheetView->scrollTo(next);
        }
        // Return focus to the grid
        m_spreadsheetView->setFocus();
    });

    // Live chart updates: refresh charts on the active sheet when data changes
    reconnectDataChanged();
}

void MainWindow::refreshActiveCharts() {
    for (auto* chart : m_charts) {
        if (chart->isVisible() && chart->property("sheetIndex").toInt() == m_activeSheetIndex) {
            chart->refreshData();
        }
    }
}

void MainWindow::reconnectDataChanged() {
    if (m_dataChangedConnection)
        disconnect(m_dataChangedConnection);
    if (m_modelResetConnection)
        disconnect(m_modelResetConnection);

    auto* model = m_spreadsheetView->getModel();
    if (model) {
        m_dataChangedConnection = connect(model, &QAbstractItemModel::dataChanged,
            this, [this]() { refreshActiveCharts(); setDirty(); });
        m_modelResetConnection = connect(model, &QAbstractItemModel::modelReset,
            this, &MainWindow::refreshActiveCharts);
    }
}

void MainWindow::onFormatCells() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    CellAddress addr(current.row(), current.column());
    auto cell = sheet->getCell(addr);
    CellStyle currentStyle = cell->getStyle();

    FormatCellsDialog dialog(currentStyle, this);
    dialog.setDocumentTheme(&sheet->getDocumentTheme());
    if (dialog.exec() == QDialog::Accepted) {
        CellStyle newStyle = dialog.getStyle();

        QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
        if (selected.isEmpty()) selected.append(current);

        std::vector<CellSnapshot> before, after;
        for (const auto& idx : selected) {
            CellAddress a(idx.row(), idx.column());
            before.push_back(sheet->takeCellSnapshot(a));
            auto c = sheet->getCell(a);
            c->setStyle(newStyle);
            after.push_back(sheet->takeCellSnapshot(a));
        }

        sheet->getUndoManager().execute(
            std::make_unique<StyleChangeCommand>(before, after), sheet.get());

        m_spreadsheetView->refreshView();
        statusBar()->showMessage("Format applied");
    }
}

void MainWindow::openFile(const QString& fileName) {
    if (fileName.isEmpty()) return;

    // Cancel any in-progress progressive loading
    if (m_isProgressiveLoading) {
        m_isProgressiveLoading = false;
        if (m_loadingOverlay) {
            m_loadingOverlay->hide();
            m_loadingOverlay->deleteLater();
            m_loadingOverlay = nullptr;
            m_loadingLabel = nullptr;
        }
        if (m_spreadsheetView) {
            m_spreadsheetView->setEditTriggers(
                QAbstractItemView::DoubleClicked | QAbstractItemView::EditKeyPressed |
                QAbstractItemView::AnyKeyPressed);
            m_spreadsheetView->setBulkLoading(false);
        }
        if (m_formulaBar) m_formulaBar->setEnabled(true);
        m_csvLoadState.reset();
    }

    m_currentFilePath = fileName;
    QString ext = QFileInfo(fileName).suffix().toLower();

    if (ext == "xlsx" || ext == "xls") {
        // XLSX: background thread with progress dialog (no progressive loading yet)
        auto* progress = new QProgressDialog(
            "Opening " + QFileInfo(fileName).fileName() + "...",
            QString(), 0, 0, this);
        progress->setWindowModality(Qt::WindowModal);
        progress->setMinimumDuration(400);
        progress->setCancelButton(nullptr);

        QElapsedTimer timer;
        timer.start();

        auto* watcher = new QFutureWatcher<XlsxImportResult>(this);
        connect(watcher, &QFutureWatcher<XlsxImportResult>::finished, this,
            [this, watcher, progress, fileName, timer]() {
                progress->close();
                progress->deleteLater();
                auto result = watcher->result();
                watcher->deleteLater();
                finishXlsxOpen(result, fileName, timer.elapsed());
            });
        // Use streaming import for better memory efficiency on large XLSX files
        watcher->setFuture(QtConcurrent::run([fileName]() {
            return XlsxService::importFromFileStreaming(fileName, [](int sheetIdx, int rowsParsed) {
                // Progress callback (runs on worker thread, can't update UI directly)
                Q_UNUSED(sheetIdx);
                Q_UNUSED(rowsParsed);
            });
        }));
    } else {
        // CSV/TXT: progressive loading — show first 100K rows instantly, load rest in background
        startProgressiveCsvLoad(fileName);
    }
}

// Internal state for progressive CSV loading
struct MainWindow::CsvLoadState {
    CsvProgressiveResult progress;
    QElapsedTimer timer;
};

MainWindow::~MainWindow() = default;

void MainWindow::startProgressiveCsvLoad(const QString& fileName) {
    m_csvLoadState = std::make_unique<CsvLoadState>();
    m_csvLoadState->timer.start();

    // Load first 100K rows on background thread, then show immediately
    auto* watcher = new QFutureWatcher<CsvProgressiveResult>(this);
    connect(watcher, &QFutureWatcher<CsvProgressiveResult>::finished, this,
        [this, watcher, fileName]() {
            m_csvLoadState->progress = watcher->result();
            watcher->deleteLater();

            auto& prog = m_csvLoadState->progress;
            if (!prog.spreadsheet) {
                QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
                m_csvLoadState.reset();
                return;
            }

            prog.spreadsheet->setSheetName(QFileInfo(fileName).baseName());
            std::vector<std::shared_ptr<Spreadsheet>> sheets = { prog.spreadsheet };
            setSheets(sheets);
            m_currentFilePath = fileName;
            updateWindowTitle();

            if (prog.resumeOffset >= 0) {
                // More data to load — show loading indicator and continue in background
                m_isProgressiveLoading = true;

                // Read-only mode during loading: disable editing and cell-scanning nav
                m_spreadsheetView->setEditTriggers(QAbstractItemView::NoEditTriggers);
                m_spreadsheetView->setBulkLoading(true);
                m_formulaBar->setEnabled(false);

                // Create loading badge in top-right corner
                if (!m_loadingOverlay) {
                    m_loadingOverlay = new QWidget(m_spreadsheetView->viewport());
                    m_loadingOverlay->setStyleSheet("background: transparent;");
                    m_loadingLabel = new QLabel(m_loadingOverlay);
                    m_loadingLabel->setStyleSheet(
                        "QLabel { background: rgba(0,120,212,200); color: white; "
                        "padding: 6px 16px; border-radius: 4px; font-size: 12px; font-weight: bold; }");
                    m_loadingLabel->setAlignment(Qt::AlignCenter);
                }
                m_loadingLabel->setText(QString("Loading... %1 rows").arg(prog.currentRow));
                m_loadingLabel->adjustSize();
                auto* vp = m_spreadsheetView->viewport();
                m_loadingOverlay->setGeometry(0, 0, vp->width(), 40);
                m_loadingLabel->move(vp->width() - m_loadingLabel->width() - 10, 8);
                m_loadingOverlay->show();
                m_loadingOverlay->raise();

                statusBar()->showMessage(QString("Loading %1 — %2 rows loaded, continuing...")
                    .arg(QFileInfo(fileName).fileName()).arg(prog.currentRow));

                QTimer::singleShot(0, this, &MainWindow::continueCsvLoadChunk);
            } else {
                // Small file — fully loaded
                m_isProgressiveLoading = false;
                m_dirty = false;
                m_toolbar->setSaveEnabled(false);
                double secs = m_csvLoadState->timer.elapsed() / 1000.0;
                statusBar()->showMessage(QString("Opened: %1 (%2 rows in %3s)")
                    .arg(QFileInfo(fileName).fileName())
                    .arg(prog.currentRow)
                    .arg(QString::number(secs, 'f', 1)));
                m_csvLoadState.reset();
            }
        });

    statusBar()->showMessage("Opening " + QFileInfo(fileName).fileName() + "...");
    watcher->setFuture(QtConcurrent::run(&CsvService::importProgressive, fileName, 100000));
}

void MainWindow::continueCsvLoadChunk() {
    if (!m_isProgressiveLoading || !m_csvLoadState || m_csvLoadState->progress.resumeOffset < 0) {
        finishProgressiveCsvLoad();
        return;
    }

    // Time-budgeted parsing: parse micro-chunks of 2K rows, yield after ~12ms
    // so scrolling/clicks stay fluid at 60fps
    QElapsedTimer timer;
    timer.start();
    auto& prog = m_csvLoadState->progress;

    while (prog.resumeOffset >= 0 && timer.elapsed() < 12) {
        CsvService::continueImport(prog, 2000);
    }

    if (prog.resumeOffset < 0) {
        finishProgressiveCsvLoad();
    } else {
        if (m_loadingLabel) {
            m_loadingLabel->setText(QString("Loading... %1 rows").arg(prog.currentRow));
            m_loadingLabel->adjustSize();
            auto* vp = m_spreadsheetView->viewport();
            m_loadingLabel->move(vp->width() - m_loadingLabel->width() - 10, 8);
        }
        statusBar()->showMessage(QString("Loading — %1 rows...").arg(prog.currentRow));
        // Yield to event loop for UI responsiveness
        QTimer::singleShot(0, this, &MainWindow::continueCsvLoadChunk);
    }
}

void MainWindow::finishProgressiveCsvLoad() {
    m_isProgressiveLoading = false;

    int totalRows = m_csvLoadState ? m_csvLoadState->progress.currentRow : 0;
    qint64 elapsedMs = m_csvLoadState ? m_csvLoadState->timer.elapsed() : 0;

    if (m_loadingOverlay) {
        m_loadingOverlay->hide();
        m_loadingOverlay->deleteLater();
        m_loadingOverlay = nullptr;
        m_loadingLabel = nullptr;
    }

    // Re-enable editing and navigation now that loading is complete
    if (m_spreadsheetView) {
        m_spreadsheetView->setEditTriggers(
            QAbstractItemView::DoubleClicked | QAbstractItemView::EditKeyPressed |
            QAbstractItemView::AnyKeyPressed);
        m_spreadsheetView->setBulkLoading(false);
    }
    if (m_formulaBar) {
        m_formulaBar->setEnabled(true);
    }

    // Final model refresh + scrollbar range update for actual row count
    if (m_spreadsheetView && m_spreadsheetView->model()) {
        QTimer::singleShot(0, this, [this]() {
            if (!m_spreadsheetView) return;
            auto* mdl = m_spreadsheetView->getModel();
            if (!mdl) return;
            // Reset the view so model picks up the final row count
            m_spreadsheetView->reset();
            // Re-setup virtual scrollbar with actual data size
            m_spreadsheetView->refreshVirtualScrollBar();
        });
    }

    m_csvLoadState.reset();
    m_dirty = false;
    m_toolbar->setSaveEnabled(false);
    double secs = elapsedMs / 1000.0;
    statusBar()->showMessage(QString("Opened: %1 (%2 rows in %3s)")
        .arg(QFileInfo(m_currentFilePath).fileName())
        .arg(totalRows)
        .arg(QString::number(secs, 'f', 1)));
}

void MainWindow::finishXlsxOpen(const XlsxImportResult& result, const QString& fileName, qint64 elapsedMs) {
    if (!result.sheets.empty()) {
        setSheets(result.sheets);
        setWindowTitle("Nexel Pro - " + QFileInfo(fileName).fileName());

        // Create chart widgets from imported charts
        static const QVector<QColor> excelColors = {
            QColor("#4472C4"), QColor("#ED7D31"), QColor("#A5A5A5"),
            QColor("#FFC000"), QColor("#5B9BD5"), QColor("#70AD47"),
            QColor("#264478"), QColor("#9E480E"), QColor("#636363")
        };

        for (const auto& imported : result.charts) {
            ChartConfig config;

            if (imported.chartType == "line") config.type = ChartType::Line;
            else if (imported.chartType == "bar") config.type = ChartType::Bar;
            else if (imported.chartType == "scatter") config.type = ChartType::Scatter;
            else if (imported.chartType == "pie") config.type = ChartType::Pie;
            else if (imported.chartType == "area") config.type = ChartType::Area;
            else if (imported.chartType == "donut") config.type = ChartType::Donut;
            else if (imported.chartType == "histogram") config.type = ChartType::Histogram;
            else config.type = ChartType::Column;

            config.title = imported.title;
            config.xAxisTitle = imported.xAxisTitle;
            config.yAxisTitle = imported.yAxisTitle;

            int si = imported.sheetIndex;
            if (si < 0 || si >= static_cast<int>(m_sheets.size())) continue;

            auto* chart = createChartWidget(m_spreadsheetView->viewport());
            chart->setSpreadsheet(m_sheets[si]);
            chart->setLazyLoad(m_lazyLoadCharts);

            if (imported.isNexelNative && !imported.dataRange.isEmpty()) {
                // Nexel-native chart: restore from dataRange
                config.dataRange = imported.dataRange;
                config.themeIndex = imported.themeIndex;
                config.showLegend = imported.showLegend;
                config.showGridLines = imported.showGridLines;
                chart->setConfig(config);
                chart->loadDataFromRange(imported.dataRange);
            } else if (!imported.dataRange.isEmpty()) {
                // Excel chart with parsed cell references: load live data from spreadsheet
                config.dataRange = imported.dataRange;
                chart->setConfig(config);
                chart->loadDataFromRange(imported.dataRange);
            } else {
                // Excel-imported chart without cell refs: use inline series data (fallback)
                for (int i = 0; i < imported.series.size(); ++i) {
                    ChartSeries s;
                    s.name = imported.series[i].name;
                    s.yValues = imported.series[i].values;

                    if (!imported.series[i].xNumeric.isEmpty()) {
                        s.xValues = imported.series[i].xNumeric;
                    } else {
                        s.xValues.resize(s.yValues.size());
                        for (int j = 0; j < s.yValues.size(); ++j) {
                            s.xValues[j] = j;
                        }
                    }
                    s.color = excelColors[i % excelColors.size()];
                    config.series.append(s);
                }
                chart->setConfig(config);
            }

            chart->setGeometry(imported.x, imported.y, imported.width, imported.height);

            connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
            connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
            connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
            connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
            connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
            connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

            chart->setProperty("sheetIndex", si);
            chart->setVisible(si == m_activeSheetIndex);
            if (si == m_activeSheetIndex) {
                chart->show();
                chart->startEntryAnimation();
            }
            m_charts.append(chart);
            addOverlay(chart);
        }
        applyZOrder();

        // Immediately load any lazy charts that are already visible
        if (m_lazyLoadCharts) loadVisibleLazyCharts();

        m_currentFilePath = fileName;
        m_dirty = false;
        m_toolbar->setSaveEnabled(false);
        updateWindowTitle();

        int chartCount = static_cast<int>(result.charts.size());
        double secs = elapsedMs / 1000.0;
        if (chartCount > 0) {
            statusBar()->showMessage(QString("Opened: %1 (%2 chart(s)) in %3s")
                .arg(QFileInfo(fileName).fileName()).arg(chartCount).arg(secs, 0, 'f', 1));
        } else {
            statusBar()->showMessage(QString("Opened: %1 in %2s")
                .arg(QFileInfo(fileName).fileName()).arg(secs, 0, 'f', 1));
        }
    } else {
        QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
    }
}

void MainWindow::finishCsvOpen(const std::shared_ptr<Spreadsheet>& spreadsheet,
                                const QString& fileName, qint64 elapsedMs) {
    if (spreadsheet) {
        spreadsheet->setSheetName(QFileInfo(fileName).baseName());
        std::vector<std::shared_ptr<Spreadsheet>> sheets = { spreadsheet };
        setSheets(sheets);
        m_currentFilePath = fileName;
        m_dirty = false;
        m_toolbar->setSaveEnabled(false);
        updateWindowTitle();
        double secs = elapsedMs / 1000.0;
        statusBar()->showMessage(QString("Opened: %1 in %2s")
            .arg(QFileInfo(fileName).fileName()).arg(secs, 0, 'f', 1));
    } else {
        QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
    }
}

void MainWindow::onNewDocument() {
    // Open a new window instead of replacing the current document
    MainWindow* newWindow = new MainWindow();
    newWindow->setAttribute(Qt::WA_DeleteOnClose);
    ThemeManager::instance().applyTheme(newWindow);

    // Offset the new window slightly right and down from the current one (like Excel)
    QPoint pos = this->pos();
    newWindow->move(pos.x() + 30, pos.y() + 30);
    newWindow->resize(this->size());

    newWindow->show();

    // Ensure cell A1 is focused in the new window (deferred so widget has keyboard focus)
    QTimer::singleShot(0, newWindow, [newWindow]() {
        if (newWindow->m_spreadsheetView && newWindow->m_spreadsheetView->model()) {
            QModelIndex a1 = newWindow->m_spreadsheetView->model()->index(0, 0);
            newWindow->m_spreadsheetView->setCurrentIndex(a1);
            newWindow->m_spreadsheetView->setFocus();
        }
    });
}

void MainWindow::onOpenDocument() {
    QString fileName = QFileDialog::getOpenFileName(this, "Open Document", "",
        "All Spreadsheet Files (*.xlsx *.csv *.txt);;Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    openFile(fileName);
}

static QString chartTypeToString(ChartType type) {
    switch (type) {
        case ChartType::Column: return "column";
        case ChartType::Bar: return "bar";
        case ChartType::Line: return "line";
        case ChartType::Area: return "area";
        case ChartType::Scatter: return "scatter";
        case ChartType::Pie: return "pie";
        case ChartType::Donut: return "donut";
        case ChartType::Histogram: return "histogram";
    }
    return "column";
}

void MainWindow::onSaveDocument() {
    if (m_currentFilePath.isEmpty()) {
        onSaveAs();
        return;
    }

    QString ext = QFileInfo(m_currentFilePath).suffix().toLower();

    if (ext == "xlsx" || ext == "xls") {
        // Warn if >1M rows (XLSX limit)
        int maxRows = 0;
        for (const auto& sheet : m_sheets) {
            if (sheet) maxRows = qMax(maxRows, sheet->getRowCount());
        }
        if (maxRows > 1048576) {
            auto reply = QMessageBox::warning(this, "Large Dataset",
                QString("This document has %1 rows.\n\n"
                        "XLSX format supports max 1,048,576 rows.\n"
                        "Only the first 1M rows will be saved.\n\n"
                        "Use Save As to save as CSV instead?")
                    .arg(QLocale().toString(maxRows)),
                QMessageBox::Ok | QMessageBox::Cancel);
            if (reply == QMessageBox::Cancel) return;
        }

        // Collect chart configs
        std::vector<NexelChartExport> chartExports;
        for (auto* chart : m_charts) {
            NexelChartExport ce;
            auto cfg = chart->config();
            ce.sheetIndex = chart->property("sheetIndex").toInt();
            ce.chartType = chartTypeToString(cfg.type);
            ce.title = cfg.title;
            ce.xAxisTitle = cfg.xAxisTitle;
            ce.yAxisTitle = cfg.yAxisTitle;
            ce.dataRange = cfg.dataRange;
            ce.themeIndex = cfg.themeIndex;
            ce.showLegend = cfg.showLegend;
            ce.showGridLines = cfg.showGridLines;
            ce.x = chart->x();
            ce.y = chart->y();
            ce.width = chart->width();
            ce.height = chart->height();
            chartExports.push_back(ce);
        }

        // Check if large dataset — run async to avoid UI freeze
        bool isLarge = false;
        for (const auto& sheet : m_sheets) {
            if (sheet && sheet->getRowCount() > 100000) { isLarge = true; break; }
        }

        if (isLarge) {
            // Background XLSX save for large files
            statusBar()->showMessage("Saving XLSX (background)...");
            m_toolbar->setSaveEnabled(false);

            auto sheets = m_sheets;
            QString path = m_currentFilePath;
            auto* watcher = new QFutureWatcher<bool>(this);
            connect(watcher, &QFutureWatcher<bool>::finished, this, [this, watcher, path]() {
                bool success = watcher->result();
                watcher->deleteLater();
                if (success) {
                    m_dirty = false;
                    updateWindowTitle();
                    statusBar()->showMessage("Saved: " + path);
                    if (m_autoSave) m_autoSave->onManualSave();
                } else {
                    m_toolbar->setSaveEnabled(true);
                    QMessageBox::warning(this, "Save Failed", "Could not save file.");
                }
            });
            watcher->setFuture(QtConcurrent::run([sheets, path, chartExports]() {
                return XlsxService::exportToFile(sheets, path, chartExports);
            }));
        } else {
            // Small files: save synchronously (fast)
            bool success = XlsxService::exportToFile(m_sheets, m_currentFilePath, chartExports);
            if (success) {
                m_dirty = false;
                updateWindowTitle();
                m_toolbar->setSaveEnabled(false);
                statusBar()->showMessage("Saved: " + m_currentFilePath);
                if (m_autoSave) m_autoSave->onManualSave();
            } else {
                QMessageBox::warning(this, "Save Failed", "Could not save file.");
            }
        }
    } else {
        // CSV save — run in background thread to avoid UI freeze on large files
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (!spreadsheet) return;

        statusBar()->showMessage("Saving...");
        m_toolbar->setSaveEnabled(false);

        QString path = m_currentFilePath;
        auto* watcher = new QFutureWatcher<bool>(this);
        connect(watcher, &QFutureWatcher<bool>::finished, this, [this, watcher, path]() {
            bool success = watcher->result();
            if (success) {
                m_dirty = false;
                updateWindowTitle();
                statusBar()->showMessage("Saved: " + path);
                if (m_autoSave) m_autoSave->onManualSave();
            } else {
                m_toolbar->setSaveEnabled(true);
                QMessageBox::warning(this, "Save Failed", "Could not save file.");
            }
            watcher->deleteLater();
        });

        watcher->setFuture(QtConcurrent::run([spreadsheet, path]() {
            return CsvService::exportToFile(*spreadsheet, path);
        }));
    }
}

void MainWindow::onSaveAs() {
    QString fileName = QFileDialog::getSaveFileName(this, "Save Document As", "",
        "Excel Workbook (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    if (fileName.isEmpty()) return;

    QString ext = QFileInfo(fileName).suffix().toLower();

    // Warn if saving >1M rows as XLSX (Excel's limit is 1,048,576 rows)
    if (ext == "xlsx") {
        int maxRows = 0;
        for (const auto& sheet : m_sheets) {
            if (sheet) maxRows = qMax(maxRows, sheet->getRowCount());
        }
        if (maxRows > 1048576) {
            auto reply = QMessageBox::warning(this, "Large Dataset",
                QString("This document has %1 rows.\n\n"
                        "XLSX format supports max 1,048,576 rows (Excel limit).\n"
                        "Only the first 1M rows will be saved.\n\n"
                        "Save as CSV instead to keep all rows?")
                    .arg(QLocale().toString(maxRows)),
                QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);
            if (reply == QMessageBox::Yes) {
                // Switch to CSV
                fileName = QFileDialog::getSaveFileName(this, "Save as CSV", "",
                    "CSV Files (*.csv);;All Files (*)");
                if (fileName.isEmpty()) return;
                ext = QFileInfo(fileName).suffix().toLower();
            } else if (reply == QMessageBox::Cancel) {
                return;
            }
            // No = proceed with XLSX (truncated to 1M rows)
        }
    }

    bool success = false;

    if (ext == "xlsx") {
        std::vector<NexelChartExport> chartExports;
        for (auto* chart : m_charts) {
            NexelChartExport ce;
            auto cfg = chart->config();
            ce.sheetIndex = chart->property("sheetIndex").toInt();
            ce.chartType = chartTypeToString(cfg.type);
            ce.title = cfg.title;
            ce.xAxisTitle = cfg.xAxisTitle;
            ce.yAxisTitle = cfg.yAxisTitle;
            ce.dataRange = cfg.dataRange;
            ce.themeIndex = cfg.themeIndex;
            ce.showLegend = cfg.showLegend;
            ce.showGridLines = cfg.showGridLines;
            ce.x = chart->x();
            ce.y = chart->y();
            ce.width = chart->width();
            ce.height = chart->height();
            chartExports.push_back(ce);
        }

        // Large files: background save
        bool isLarge = false;
        for (const auto& sheet : m_sheets) {
            if (sheet && sheet->getRowCount() > 100000) { isLarge = true; break; }
        }

        if (isLarge) {
            statusBar()->showMessage("Saving XLSX (background)...");
            auto sheets = m_sheets;
            auto* watcher = new QFutureWatcher<bool>(this);
            connect(watcher, &QFutureWatcher<bool>::finished, this, [this, watcher, fileName]() {
                bool ok = watcher->result();
                watcher->deleteLater();
                if (ok) {
                    if (m_autoSave) m_autoSave->setCurrentFilePath(fileName);
                    m_currentFilePath = fileName;
                    m_dirty = false;
                    updateWindowTitle();
                    m_toolbar->setSaveEnabled(false);
                    statusBar()->showMessage("Saved: " + fileName);
                    if (m_autoSave) m_autoSave->onManualSave();
                } else {
                    QMessageBox::warning(this, "Save Failed", "Could not save file.");
                }
            });
            watcher->setFuture(QtConcurrent::run([sheets, fileName, chartExports]() {
                return XlsxService::exportToFile(sheets, fileName, chartExports);
            }));
            return;
        }

        success = XlsxService::exportToFile(m_sheets, fileName, chartExports);
    } else {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) success = CsvService::exportToFile(*spreadsheet, fileName);
    }

    if (success) {
        if (m_autoSave) m_autoSave->setCurrentFilePath(fileName);
        m_currentFilePath = fileName;
        m_dirty = false;
        updateWindowTitle();
        m_toolbar->setSaveEnabled(false);
        statusBar()->showMessage("Saved: " + fileName);
        if (m_autoSave) m_autoSave->onManualSave();
    } else {
        QMessageBox::warning(this, "Save Failed", "Could not save file.");
    }
}

void MainWindow::onUndo() {
    if (!m_spreadsheetView || m_isProgressiveLoading) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canUndo()) {
        sheet->getUndoManager().undo(sheet.get());
        bool structural = sheet->getUndoManager().lastUndoIsStructural();
        CellAddress target = sheet->getUndoManager().lastUndoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            QModelIndex idx = model->index(target.row, target.col);
            if (structural) {
                model->resetModel();
                QTimer::singleShot(0, m_spreadsheetView, [this]() {
                    m_spreadsheetView->applyStoredDimensions();
                });
            } else {
                emit model->dataChanged(idx, idx);
            }
            // Force full repaint for table style undo/redo
            m_spreadsheetView->viewport()->update();
            m_spreadsheetView->selectionModel()->setCurrentIndex(
                idx, QItemSelectionModel::ClearAndSelect);
            m_spreadsheetView->scrollTo(idx);
        }
        refreshActiveCharts();
        setDirty();
        statusBar()->showMessage("Undo");
    }
}

void MainWindow::onRedo() {
    if (!m_spreadsheetView || m_isProgressiveLoading) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canRedo()) {
        sheet->getUndoManager().redo(sheet.get());
        bool structural = sheet->getUndoManager().lastRedoIsStructural();
        CellAddress target = sheet->getUndoManager().lastRedoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            QModelIndex idx = model->index(target.row, target.col);
            if (structural) {
                model->resetModel();
                QTimer::singleShot(0, m_spreadsheetView, [this]() {
                    m_spreadsheetView->applyStoredDimensions();
                });
            } else {
                emit model->dataChanged(idx, idx);
            }
            // Force full repaint for table style undo/redo
            m_spreadsheetView->viewport()->update();
            m_spreadsheetView->selectionModel()->setCurrentIndex(
                idx, QItemSelectionModel::ClearAndSelect);
            m_spreadsheetView->scrollTo(idx);
        }
        refreshActiveCharts();
        setDirty();
        statusBar()->showMessage("Redo");
    }
}

void MainWindow::onCut() { if (m_spreadsheetView && !m_isProgressiveLoading) m_spreadsheetView->cut(); }
void MainWindow::onCopy() { if (m_spreadsheetView) m_spreadsheetView->copy(); }
void MainWindow::onPaste() { if (m_spreadsheetView && !m_isProgressiveLoading) m_spreadsheetView->paste(); }
void MainWindow::onDelete() { if (m_spreadsheetView && !m_isProgressiveLoading) m_spreadsheetView->deleteSelection(); }
void MainWindow::onSelectAll() { if (m_spreadsheetView) m_spreadsheetView->selectAll(); }

void MainWindow::onImportCsv() {
    QString fileName = QFileDialog::getOpenFileName(this, "Import CSV", "",
        "CSV Files (*.csv);;Text Files (*.txt);;All Files (*)");
    if (!fileName.isEmpty()) openFile(fileName);
}

void MainWindow::onExportCsv() {
    QString fileName = QFileDialog::getSaveFileName(this, "Export CSV", "",
        "CSV Files (*.csv);;All Files (*)");
    if (fileName.isEmpty()) return;

    auto spreadsheet = m_spreadsheetView->getSpreadsheet();
    if (spreadsheet && CsvService::exportToFile(*spreadsheet, fileName)) {
        statusBar()->showMessage("Exported: " + fileName);
    } else {
        QMessageBox::warning(this, "Export Failed", "Could not export CSV file.");
    }
}

void MainWindow::closeEvent(QCloseEvent* event) {
    // Count actual MainWindow instances (not menus, tooltips, dock widgets, etc.)
    auto isLastMainWindow = [&]() {
        int count = 0;
        for (auto* w : QApplication::topLevelWidgets())
            if (qobject_cast<MainWindow*>(w) && w != this) count++;
        return count == 0;
    };

    if (!m_dirty) {
        // Clean exit — remove auto-save files
        if (m_autoSave) AutoSaveService::cleanupAutoSave(m_currentFilePath);
        event->accept();
        if (isLastMainWindow())
            _exit(0);
        return;
    }

    auto result = QMessageBox::question(this, "Unsaved Changes",
        "Do you want to save your changes before closing?",
        QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel);

    if (result == QMessageBox::Save) {
        onSaveDocument();
        if (m_autoSave) AutoSaveService::cleanupAutoSave(m_currentFilePath);
        event->accept();
        if (isLastMainWindow())
            _exit(0);
    } else if (result == QMessageBox::Discard) {
        if (m_autoSave) AutoSaveService::cleanupAutoSave(m_currentFilePath);
        event->accept();
        if (isLastMainWindow())
            _exit(0);
    } else {
        event->ignore();
    }
}

void MainWindow::checkAutoSaveRecovery() {
    // Check for unsaved document recovery first
    QString recoveryPath = AutoSaveService::checkForRecovery(m_currentFilePath);

    // If no recovery for current path, also check for unsaved document recovery
    if (recoveryPath.isEmpty() && m_currentFilePath.isEmpty()) {
        recoveryPath = AutoSaveService::checkForRecovery(QString());
    }

    if (recoveryPath.isEmpty()) return;

    auto result = QMessageBox::question(this, "Recover Auto-Saved Document",
        "An auto-saved version of your document was found.\n"
        "This may be from a previous session that ended unexpectedly.\n\n"
        "Would you like to recover it?",
        QMessageBox::Yes | QMessageBox::No);

    if (result == QMessageBox::Yes) {
        openFile(recoveryPath);
        // Mark as dirty since this is a recovery — user should save explicitly
        setDirty(true);
        statusBar()->showMessage("Recovered auto-saved document", 5000);
    }

    // Clean up the auto-save file regardless of choice
    AutoSaveService::cleanupAutoSave(m_currentFilePath);
    // Also clean up unsaved recovery
    AutoSaveService::cleanupAutoSave(QString());
}

void MainWindow::dragEnterEvent(QDragEnterEvent* event) {
    if (event->mimeData()->hasUrls()) event->acceptProposedAction();
}

void MainWindow::dropEvent(QDropEvent* event) {
    const QMimeData* mimeData = event->mimeData();
    if (mimeData->hasUrls()) {
        openFile(mimeData->urls().first().toLocalFile());
        event->acceptProposedAction();
    }
}

// ============== Find/Replace ==============

void MainWindow::onFindReplace() {
    if (!m_findReplaceDialog) {
        m_findReplaceDialog = new FindReplaceDialog(this);
        connect(m_findReplaceDialog, &FindReplaceDialog::findNext, this, &MainWindow::onFindNext);
        connect(m_findReplaceDialog, &FindReplaceDialog::findPrevious, this, &MainWindow::onFindPrevious);
        connect(m_findReplaceDialog, &FindReplaceDialog::replaceOne, this, &MainWindow::onReplaceOne);
        connect(m_findReplaceDialog, &FindReplaceDialog::replaceAll, this, &MainWindow::onReplaceAll);
    }
    m_findReplaceDialog->show();
    m_findReplaceDialog->raise();
    m_findReplaceDialog->activateWindow();
}

bool MainWindow::cellMatchesSearch(int row, int col, const QString& searchText, bool matchCase, bool wholeCell) const {
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return false;

    auto val = sheet->getCellValue(CellAddress(row, col));
    QString cellText = val.toString();

    if (wholeCell) {
        return matchCase ? (cellText == searchText)
                         : (cellText.compare(searchText, Qt::CaseInsensitive) == 0);
    } else {
        return cellText.contains(searchText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
    }
}

void MainWindow::onFindNext() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    int maxRow = sheet->getMaxRow();
    if (maxRow < 0) {
        m_findReplaceDialog->setStatus("No data to search.");
        return;
    }

    // For large datasets: use cached search results (parallel search, done once)
    bool isLarge = maxRow > 100000;

    if (isLarge && (m_findCache.isEmpty() || m_findCacheQuery != searchText ||
                    m_findCacheCase != matchCase || m_findCacheWhole != wholeCell)) {
        // Run parallel search in background
        m_findReplaceDialog->setStatus("Searching...");
        QApplication::processEvents();

        m_findCache.clear();
        auto results = sheet->searchAllCells(searchText, matchCase, wholeCell);
        for (const auto& addr : results) {
            m_findCache.append(addr);
        }
        m_findCacheQuery = searchText;
        m_findCacheCase = matchCase;
        m_findCacheWhole = wholeCell;
        m_findCacheIdx = -1;
        m_findReplaceDialog->setStatus(QString("%1 matches found").arg(m_findCache.size()));
    }

    if (isLarge) {
        if (m_findCache.isEmpty()) {
            m_findReplaceDialog->setStatus("Not found.");
            return;
        }
        m_findCacheIdx = (m_findCacheIdx + 1) % m_findCache.size();
        CellAddress addr = m_findCache[m_findCacheIdx];
        auto* mdl = m_spreadsheetView->getModel();
        if (mdl) {
            int modelRow = mdl->toModelRow(addr.row);
            if (modelRow < 0) {
                // Row not in current window — recenter
                mdl->jumpToBase(std::max(0, addr.row - SpreadsheetModel::WINDOW_SIZE / 2));
                modelRow = mdl->toModelRow(addr.row);
            }
            if (modelRow >= 0) {
                QModelIndex idx = mdl->index(modelRow, addr.col);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
            }
        }
        m_findReplaceDialog->setStatus(QString("Found %1 of %2: %3")
            .arg(m_findCacheIdx + 1).arg(m_findCache.size()).arg(addr.toString()));
        return;
    }

    // Small datasets: original sequential search
    QModelIndex current = m_spreadsheetView->currentIndex();
    int startRow = current.isValid() ? current.row() : 0;
    int startCol = current.isValid() ? current.column() + 1 : 0;
    int maxCol = sheet->getMaxColumn();

    for (int r = startRow; r <= maxRow; ++r) {
        int cStart = (r == startRow) ? startCol : 0;
        for (int c = cStart; c <= maxCol; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    for (int r = 0; r <= startRow; ++r) {
        int cEnd = (r == startRow) ? startCol - 1 : maxCol;
        for (int c = 0; c <= cEnd; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1 (wrapped)").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    m_findReplaceDialog->setStatus("Not found.");
}

void MainWindow::onFindPrevious() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    int maxRow = sheet->getMaxRow();
    int maxCol = sheet->getMaxColumn();
    if (maxRow < 0 || maxCol < 0) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    int startRow = current.isValid() ? current.row() : maxRow;
    int startCol = current.isValid() ? current.column() - 1 : maxCol;

    // Search backward
    for (int r = startRow; r >= 0; --r) {
        int cStart = (r == startRow) ? startCol : maxCol;
        for (int c = cStart; c >= 0; --c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    // Wrap around from bottom
    for (int r = maxRow; r >= startRow; --r) {
        int cStart = (r == startRow) ? startCol + 1 : 0;
        for (int c = maxCol; c >= cStart; --c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1 (wrapped)").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    m_findReplaceDialog->setStatus("Not found.");
}

void MainWindow::onReplaceOne() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    QString replaceText = m_findReplaceDialog->replaceText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    if (cellMatchesSearch(current.row(), current.column(), searchText, matchCase, wholeCell)) {
        auto model = m_spreadsheetView->getModel();
        if (wholeCell) {
            model->setData(current, replaceText);
        } else {
            auto val = sheet->getCellValue(CellAddress(current.row(), current.column()));
            QString cellText = val.toString();
            cellText.replace(searchText, replaceText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
            model->setData(current, cellText);
        }
        m_findReplaceDialog->setStatus("Replaced. Finding next...");
        onFindNext();
    } else {
        onFindNext();
    }
}

void MainWindow::onReplaceAll() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    QString replaceText = m_findReplaceDialog->replaceText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    auto model = m_spreadsheetView->getModel();
    int maxRow = sheet->getMaxRow();
    int maxCol = sheet->getMaxColumn();
    int count = 0;

    std::vector<CellSnapshot> before, after;
    model->setSuppressUndo(true);

    for (int r = 0; r <= maxRow; ++r) {
        for (int c = 0; c <= maxCol; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                CellAddress addr(r, c);
                before.push_back(sheet->takeCellSnapshot(addr));

                QModelIndex idx = model->index(r, c);
                if (wholeCell) {
                    model->setData(idx, replaceText);
                } else {
                    auto val = sheet->getCellValue(addr);
                    QString cellText = val.toString();
                    cellText.replace(searchText, replaceText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
                    model->setData(idx, cellText);
                }
                after.push_back(sheet->takeCellSnapshot(addr));
                count++;
            }
        }
    }

    model->setSuppressUndo(false);

    if (!before.empty()) {
        sheet->getUndoManager().pushCommand(
            std::make_unique<MultiCellEditCommand>(before, after, "Replace All"));
    }

    m_findReplaceDialog->setStatus(QString("Replaced %1 occurrence(s).").arg(count));
    m_spreadsheetView->refreshView();
}

// ============== Go To ==============

void MainWindow::onGoTo() {
    GoToDialog dialog(this);
    if (dialog.exec() == QDialog::Accepted) {
        CellAddress addr = dialog.getAddress();
        if (addr.row >= 0 && addr.col >= 0) {
            auto model = m_spreadsheetView->getModel();
            if (model) {
                QModelIndex idx = model->index(addr.row, addr.col);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx, QAbstractItemView::PositionAtCenter);
                statusBar()->showMessage(QString("Navigated to %1").arg(addr.toString()));
            }
        } else {
            QMessageBox::warning(this, "Go To", "Invalid cell reference.");
        }
    }
}

// ============== Conditional Formatting ==============

void MainWindow::onConditionalFormat() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    // Use selection as default range for new rules (or A1 if nothing selected)
    CellRange defaultRange(CellAddress(0, 0), CellAddress(0, 0));
    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (!selected.isEmpty()) {
        int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
        for (const auto& idx : selected) {
            minRow = qMin(minRow, idx.row());
            maxRow = qMax(maxRow, idx.row());
            minCol = qMin(minCol, idx.column());
            maxCol = qMax(maxCol, idx.column());
        }
        defaultRange = CellRange(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    }

    ConditionalFormatDialog dialog(defaultRange, sheet->getConditionalFormatting(), this);
    dialog.exec();
    m_spreadsheetView->refreshView();
}

// ============== Data Validation ==============

void MainWindow::onDataValidation() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    DataValidationDialog dialog(range, this);
    dialog.setSheets(m_sheets);

    // Load existing rule if present
    const auto* existing = sheet->getValidationAt(minRow, minCol);
    if (existing) {
        dialog.setRule(*existing);
    }

    if (dialog.exec() == QDialog::Accepted) {
        auto rule = dialog.getRule();
        // Remove old rules for this range
        auto& rules = sheet->getValidationRules();
        for (int i = static_cast<int>(rules.size()) - 1; i >= 0; --i) {
            if (rules[i].range.intersects(range)) {
                sheet->removeValidationRule(i);
            }
        }
        sheet->addValidationRule(rule);
        statusBar()->showMessage("Data validation applied");
    }
}

// ============== Freeze Panes ==============

void MainWindow::onFreezePane() {
    if (!m_spreadsheetView) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    if (m_frozenPanes) {
        // Unfreeze
        m_spreadsheetView->setFrozenRow(-1);
        m_spreadsheetView->setFrozenColumn(-1);
        m_frozenPanes = false;
        statusBar()->showMessage("Panes unfrozen");
    } else {
        // Freeze at current cell position
        m_spreadsheetView->setFrozenRow(current.row());
        m_spreadsheetView->setFrozenColumn(current.column());
        m_frozenPanes = true;
        statusBar()->showMessage(QString("Panes frozen at %1")
            .arg(CellAddress(current.row(), current.column()).toString()));
    }
}

void MainWindow::onProtectSheet() {
    if (m_sheets.empty() || m_activeSheetIndex >= (int)m_sheets.size()) return;
    auto& sheet = m_sheets[m_activeSheetIndex];

    if (sheet->isProtected()) {
        // Unprotect: ask for password if one was set
        if (!sheet->getProtectionPasswordHash().isEmpty()) {
            bool ok = false;
            QString password = QInputDialog::getText(this, "Unprotect Sheet",
                "Enter password to unprotect sheet:", QLineEdit::Password, QString(), &ok);
            if (!ok) return;
            if (!sheet->checkProtectionPassword(password)) {
                QMessageBox::warning(this, "Unprotect Sheet", "Incorrect password.");
                return;
            }
        }
        sheet->setProtected(false);
        if (m_protectSheetAction) m_protectSheetAction->setText("&Protect Sheet...");
        statusBar()->showMessage("Sheet unprotected");
    } else {
        // Protect: show dialog with optional password
        QDialog dlg(this);
        dlg.setWindowTitle("Protect Sheet");
        auto* layout = new QVBoxLayout(&dlg);

        QLabel* label = new QLabel("Optionally enter a password to protect this sheet:", &dlg);
        layout->addWidget(label);

        QLineEdit* pwEdit = new QLineEdit(&dlg);
        pwEdit->setEchoMode(QLineEdit::Password);
        pwEdit->setPlaceholderText("Password (optional)");
        layout->addWidget(pwEdit);

        QLineEdit* pwConfirm = new QLineEdit(&dlg);
        pwConfirm->setEchoMode(QLineEdit::Password);
        pwConfirm->setPlaceholderText("Confirm password");
        layout->addWidget(pwConfirm);

        QLabel* infoLabel = new QLabel(
            "All cells are locked by default. To allow editing specific cells,\n"
            "select them and uncheck 'Locked' in Format Cells > Protection.", &dlg);
        infoLabel->setStyleSheet("color: #666; font-size: 11px;");
        layout->addWidget(infoLabel);

        auto* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
        layout->addWidget(buttons);
        connect(buttons, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
        connect(buttons, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);

        if (dlg.exec() != QDialog::Accepted) return;

        QString pw = pwEdit->text();
        if (!pw.isEmpty() && pw != pwConfirm->text()) {
            QMessageBox::warning(this, "Protect Sheet", "Passwords do not match.");
            return;
        }

        m_spreadsheetView->protectSheet(pw);
        if (m_protectSheetAction) m_protectSheetAction->setText("&Unprotect Sheet...");
        statusBar()->showMessage("Sheet protected");
    }
}

void MainWindow::onHighlightInvalidCells() {
    if (!m_spreadsheetView || !m_spreadsheetView->getModel()) return;
    auto* model = m_spreadsheetView->getModel();
    bool current = model->highlightInvalidCells();
    model->setHighlightInvalidCells(!current);
    model->resetModel();
    statusBar()->showMessage(current ? "Invalid cell highlighting off" : "Invalid cell highlighting on");
}

void MainWindow::updateStatusBarSummary() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    // Use selection ranges (not selectedIndexes) to avoid creating millions of QModelIndex
    // objects when entire rows/columns are selected
    QItemSelection selection = m_spreadsheetView->selectionModel()->selection();
    if (selection.isEmpty()) {
        statusBar()->showMessage("Ready");
        return;
    }

    // Quick check: if only a single cell selected, show "Ready"
    // Count total cells across all selection ranges to handle fragmented ranges
    int totalSelectedCells = 0;
    for (const auto& range : selection) {
        totalSelectedCells += (range.bottom() - range.top() + 1) *
                              (range.right() - range.left() + 1);
        if (totalSelectedCells > 1) break; // early exit
    }
    if (totalSelectedCells <= 1) {
        statusBar()->showMessage("Ready");
        return;
    }

    // Compute SUM, AVERAGE, COUNT for numeric values (like Excel status bar)
    // Iterate only occupied cells within selection ranges for performance
    double sum = 0;
    int numericCount = 0;
    int nonEmptyCount = 0;
    int cellsChecked = 0;
    static constexpr int MAX_CELLS = 50000;

    for (const auto& range : selection) {
        if (cellsChecked >= MAX_CELLS) break;
        int top = range.top(), bottom = range.bottom();
        int left = range.left(), right = range.right();

        // For large ranges (e.g. entire column), only scan occupied cells
        bool isLargeRange = (bottom - top + 1) * (right - left + 1) > 10000;

        if (isLargeRange) {
            for (int c = left; c <= right && cellsChecked < MAX_CELLS; ++c) {
                const auto& occupiedRows = sheet->getOccupiedRowsInColumn(c);
                for (int r : occupiedRows) {
                    if (r < top || r > bottom) continue;
                    if (cellsChecked >= MAX_CELLS) break;
                    auto val = sheet->getCellValue(CellAddress(r, c));
                    QString text = val.toString();
                    if (!text.isEmpty()) {
                        nonEmptyCount++;
                        bool ok;
                        double num = text.toDouble(&ok);
                        if (ok) { sum += num; numericCount++; }
                    }
                    cellsChecked++;
                }
            }
        } else {
            for (int r = top; r <= bottom && cellsChecked < MAX_CELLS; ++r) {
                for (int c = left; c <= right && cellsChecked < MAX_CELLS; ++c) {
                    auto val = sheet->getCellValue(CellAddress(r, c));
                    QString text = val.toString();
                    if (!text.isEmpty()) {
                        nonEmptyCount++;
                        bool ok;
                        double num = text.toDouble(&ok);
                        if (ok) { sum += num; numericCount++; }
                    }
                    cellsChecked++;
                }
            }
        }
    }

    // Update permanent stats widgets (Excel-style, always visible on right)
    if (numericCount > 0) {
        double avg = sum / numericCount;
        if (m_statusAvgLabel) m_statusAvgLabel->setText(QString("Average: %1").arg(QString::number(avg, 'f', 2)));
        if (m_statusCountLabel) m_statusCountLabel->setText(QString("Count: %1").arg(nonEmptyCount));
        if (m_statusSumLabel) m_statusSumLabel->setText(QString("Sum: %1").arg(QString::number(sum, 'f', 2)));
    } else if (nonEmptyCount > 0) {
        if (m_statusAvgLabel) m_statusAvgLabel->setText("");
        if (m_statusCountLabel) m_statusCountLabel->setText(QString("Count: %1").arg(nonEmptyCount));
        if (m_statusSumLabel) m_statusSumLabel->setText("");
    } else {
        if (m_statusAvgLabel) m_statusAvgLabel->setText("");
        if (m_statusCountLabel) m_statusCountLabel->setText("");
        if (m_statusSumLabel) m_statusSumLabel->setText("");
    }
}

// ============== Chat NLP Actions ==============

static CellAddress parseCellRef(const QString& ref) {
    int col = 0;
    int i = 0;
    while (i < ref.length() && ref[i].isLetter()) {
        col = col * 26 + (ref[i].toUpper().unicode() - 'A' + 1);
        i++;
    }
    col--; // 0-indexed
    int row = ref.mid(i).toInt() - 1; // 0-indexed
    return CellAddress(qMax(0, row), qMax(0, col));
}

static CellAddress parseRangeStart(const QString& rangeStr) {
    QStringList parts = rangeStr.split(':');
    return parseCellRef(parts[0]);
}

static CellAddress parseRangeEnd(const QString& rangeStr) {
    QStringList parts = rangeStr.split(':');
    return (parts.size() > 1) ? parseCellRef(parts[1]) : parseCellRef(parts[0]);
}

static int parseColLetter(const QString& col) {
    int result = 0;
    for (int i = 0; i < col.length(); ++i) {
        result = result * 26 + (col[i].toUpper().unicode() - 'A' + 1);
    }
    return result - 1;
}

void MainWindow::onChatActions(const QJsonArray& actions) {
    if (!m_spreadsheetView || m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];
    if (!sheet) return;

    for (const auto& item : actions) {
        QJsonObject action = item.toObject();
        QString type = action["action"].toString();

        if (type == "set_cell") {
            QString cellRef = action["cell"].toString();
            QJsonValue val = action["value"];
            CellAddress addr = parseCellRef(cellRef);
            auto cell = sheet->getCell(addr);
            if (val.isDouble()) {
                cell->setValue(val.toDouble());
            } else {
                cell->setValue(val.toString());
            }

        } else if (type == "set_formula") {
            QString cellRef = action["cell"].toString();
            QString formula = action["formula"].toString();
            CellAddress addr = parseCellRef(cellRef);
            sheet->setCellFormula(addr, formula);

        } else if (type == "format") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());

            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    CellAddress addr(r, c);
                    auto cell = sheet->getCell(addr);
                    CellStyle style = cell->getStyle();

                    if (action.contains("bold")) style.bold = action["bold"].toBool();
                    if (action.contains("italic")) style.italic = action["italic"].toBool();
                    if (action.contains("underline")) style.underline = action["underline"].toBool();
                    if (action.contains("strikethrough")) style.strikethrough = action["strikethrough"].toBool();
                    if (action.contains("bg_color")) style.backgroundColor = action["bg_color"].toString();
                    if (action.contains("fg_color")) style.foregroundColor = action["fg_color"].toString();
                    if (action.contains("font_size")) style.fontSize = action["font_size"].toInt();
                    if (action.contains("font_name")) style.fontName = action["font_name"].toString();
                    if (action.contains("h_align")) {
                        QString align = action["h_align"].toString();
                        if (align == "left") style.hAlign = HorizontalAlignment::Left;
                        else if (align == "center") style.hAlign = HorizontalAlignment::Center;
                        else if (align == "right") style.hAlign = HorizontalAlignment::Right;
                    }
                    if (action.contains("v_align")) {
                        QString align = action["v_align"].toString();
                        if (align == "top") style.vAlign = VerticalAlignment::Top;
                        else if (align == "middle") style.vAlign = VerticalAlignment::Middle;
                        else if (align == "bottom") style.vAlign = VerticalAlignment::Bottom;
                    }

                    cell->setStyle(style);
                }
            }

        } else if (type == "merge") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            sheet->mergeCells(range);
            int rowSpan = end.row - start.row + 1;
            int colSpan = end.col - start.col + 1;
            m_spreadsheetView->setSpan(start.row, start.col, rowSpan, colSpan);
            // Center merged content
            auto cell = sheet->getCell(start);
            CellStyle style = cell->getStyle();
            style.hAlign = HorizontalAlignment::Center;
            style.vAlign = VerticalAlignment::Middle;
            cell->setStyle(style);

        } else if (type == "unmerge") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            m_spreadsheetView->setSpan(start.row, start.col, 1, 1);
            sheet->unmergeCells(range);

        } else if (type == "border") {
            QString borderType = action["type"].toString();
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());

            BorderStyle on;
            on.enabled = true;
            on.color = "#000000";
            on.width = (borderType == "thick_outside") ? 2 : 1;

            BorderStyle off;
            off.enabled = false;

            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    CellAddress addr(r, c);
                    auto cell = sheet->getCell(addr);
                    CellStyle style = cell->getStyle();

                    if (borderType == "none") {
                        style.borderTop = off; style.borderBottom = off;
                        style.borderLeft = off; style.borderRight = off;
                    } else if (borderType == "all") {
                        style.borderTop = on; style.borderBottom = on;
                        style.borderLeft = on; style.borderRight = on;
                    } else if (borderType == "outside" || borderType == "thick_outside") {
                        if (r == start.row) style.borderTop = on;
                        if (r == end.row) style.borderBottom = on;
                        if (c == start.col) style.borderLeft = on;
                        if (c == end.col) style.borderRight = on;
                    } else if (borderType == "bottom") {
                        if (r == end.row) style.borderBottom = on;
                    } else if (borderType == "top") {
                        if (r == start.row) style.borderTop = on;
                    } else if (borderType == "left") {
                        if (c == start.col) style.borderLeft = on;
                    } else if (borderType == "right") {
                        if (c == end.col) style.borderRight = on;
                    }

                    cell->setStyle(style);
                }
            }

        } else if (type == "table") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            int themeIdx = action["theme"].toInt();
            auto themes = getBuiltinTableThemes();
            if (themeIdx >= 0 && themeIdx < static_cast<int>(themes.size())) {
                SpreadsheetTable table;
                table.range = CellRange(start, end);
                table.theme = themes[themeIdx];
                table.hasHeaderRow = true;
                table.bandedRows = true;
                int tableNum = static_cast<int>(sheet->getTables().size()) + 1;
                table.name = QString("Table%1").arg(tableNum);
                for (int c = start.col; c <= end.col; ++c) {
                    auto val = sheet->getCellValue(CellAddress(start.row, c));
                    QString name = val.toString();
                    if (name.isEmpty()) name = QString("Column%1").arg(c - start.col + 1);
                    table.columnNames.append(name);
                }
                sheet->addTable(table);
            }

        } else if (type == "number_format") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            QString fmt = action["format"].toString();
            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    auto cell = sheet->getCell(CellAddress(r, c));
                    CellStyle style = cell->getStyle();
                    style.numberFormat = fmt;
                    cell->setStyle(style);
                }
            }

        } else if (type == "set_row_height") {
            int row = action["row"].toInt() - 1; // 1-based to 0-based
            int height = action["height"].toInt();
            if (row >= 0 && height > 0) {
                m_spreadsheetView->setRowHeight(row, height);
            }

        } else if (type == "set_col_width") {
            QString colStr = action["col"].toString();
            int col = parseColLetter(colStr);
            int width = action["width"].toInt();
            if (col >= 0 && width > 0) {
                m_spreadsheetView->setColumnWidth(col, width);
            }

        } else if (type == "clear") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            sheet->clearRange(range);

        } else if (type == "insert_chart") {
            insertChartFromChat(action);

        } else if (type == "insert_shape") {
            insertShapeFromChat(action);

        } else if (type == "insert_sparkline") {
            QString cellRef = action["cell"].toString();
            QString dataRange = action["data_range"].toString();
            if (!cellRef.isEmpty() && !dataRange.isEmpty()) {
                SparklineConfig config;
                QString typeStr = action["type"].toString().toLower();
                if (typeStr == "column") config.type = SparklineType::Column;
                else if (typeStr == "winloss") config.type = SparklineType::WinLoss;
                else config.type = SparklineType::Line;
                config.dataRange = dataRange;
                if (action.contains("color")) config.lineColor = QColor(action["color"].toString());
                config.showHighPoint = action["show_high"].toBool(false);
                config.showLowPoint = action["show_low"].toBool(false);
                CellAddress addr = parseCellRef(cellRef);
                sheet->setSparkline(addr, config);
            }

        } else if (type == "insert_image") {
            insertImageFromChat(action);

        } else if (type == "run_macro") {
            QString code = action["code"].toString();
            if (!code.isEmpty() && m_macroEngine) {
                auto result = m_macroEngine->execute(code);
                if (!result.success) {
                    statusBar()->showMessage("Macro error: " + result.error, 5000);
                }
            }

        } else if (type == "record_macro") {
            QString macroAction = action["action"].toString().toLower();
            if (m_macroEngine) {
                if (macroAction == "start") m_macroEngine->startRecording();
                else if (macroAction == "stop") m_macroEngine->stopRecording();
            }

        } else if (type == "conditional_format") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);

            // Parse condition type
            QString cond = action["condition"].toString().toLower();
            ConditionType condType = ConditionType::GreaterThan;
            if (cond == "equal") condType = ConditionType::Equal;
            else if (cond == "not_equal") condType = ConditionType::NotEqual;
            else if (cond == "greater_than") condType = ConditionType::GreaterThan;
            else if (cond == "less_than") condType = ConditionType::LessThan;
            else if (cond == "greater_than_or_equal") condType = ConditionType::GreaterThanOrEqual;
            else if (cond == "less_than_or_equal") condType = ConditionType::LessThanOrEqual;
            else if (cond == "between") condType = ConditionType::Between;
            else if (cond == "contains") condType = ConditionType::CellContains;

            auto rule = std::make_shared<ConditionalFormat>(range, condType);
            rule->setValue1(action["value"].toVariant());
            if (action.contains("value2")) {
                rule->setValue2(action["value2"].toVariant());
            }

            // Build style
            CellStyle style;
            style.backgroundColor = action.contains("bg_color")
                ? action["bg_color"].toString() : "#FFEB9C";
            if (action.contains("fg_color"))
                style.foregroundColor = action["fg_color"].toString();
            if (action.contains("bold"))
                style.bold = action["bold"].toBool();
            rule->setStyle(style);

            sheet->getConditionalFormatting().addRule(rule);
        }

        // ===== Insert checkbox =====
        else if (type == "insert_checkbox") {
            QString rangeStr = action["range"].toString();
            CellRange range(rangeStr);
            for (const auto& addr : range.getCells()) {
                auto cell = sheet->getCell(addr);
                CellStyle style = cell->getStyle();
                style.numberFormat = "Checkbox";
                cell->setStyle(style);
                if (cell->getValue().isNull() || cell->getValue().toString().isEmpty()) {
                    bool val = action.contains("checked") ? action["checked"].toBool() : false;
                    cell->setValue(QVariant(val));
                }
            }
        }

        // ===== Insert picklist =====
        else if (type == "insert_picklist") {
            QString rangeStr = action["range"].toString();
            CellRange range(rangeStr);
            QStringList options;
            if (action.contains("options")) {
                for (const auto& v : action["options"].toArray()) {
                    options.append(v.toString());
                }
            }
            // Create validation rule
            Spreadsheet::DataValidationRule rule;
            rule.range = range;
            rule.type = Spreadsheet::DataValidationRule::List;
            rule.listItems = options;
            rule.showErrorAlert = false;
            sheet->addValidationRule(rule);
            // Set numberFormat
            for (const auto& addr : range.getCells()) {
                auto cell = sheet->getCell(addr);
                CellStyle style = cell->getStyle();
                style.numberFormat = "Picklist";
                cell->setStyle(style);
                if (action.contains("value")) {
                    cell->setValue(action["value"].toString());
                }
            }
        }

        // ===== Set picklist value =====
        else if (type == "set_picklist") {
            QString cellRef = action["cell"].toString();
            CellAddress addr = parseCellRef(cellRef);
            auto cell = sheet->getCell(addr);
            cell->setValue(action["value"].toString()); // pipe-separated values
        }

        // ===== Toggle checkbox =====
        else if (type == "toggle_checkbox") {
            QString cellRef = action["cell"].toString();
            CellAddress addr = parseCellRef(cellRef);
            auto cell = sheet->getCell(addr);
            bool current = cell->getValue().toBool();
            cell->setValue(QVariant(!current));
        }
    }

    // Refresh the view
    m_spreadsheetView->refreshView();
    if (m_spreadsheetView->getModel())
        m_spreadsheetView->getModel()->resetModel();
    statusBar()->showMessage(QString("Claude applied %1 action(s)").arg(actions.size()), 5000);
}

bool MainWindow::saveCurrentDocument() {
    auto doc = DocumentService::instance().getCurrentDocument();
    if (doc) return DocumentService::instance().saveDocument();
    return true;
}

void MainWindow::setDirty(bool dirty) {
    if (m_dirty == dirty) return;
    m_dirty = dirty;
    m_toolbar->setSaveEnabled(dirty);
    updateWindowTitle();
    if (dirty && m_autoSave) {
        m_autoSave->markDirty();
    }
}

void MainWindow::updateWindowTitle() {
    QString title = "Nexel Pro";
    if (!m_currentFilePath.isEmpty()) {
        title += " - " + QFileInfo(m_currentFilePath).fileName();
    }
    if (m_dirty) {
        title += " *";
    }
    setWindowTitle(title);
}

// ============== Chart and Shape Insertion ==============

QString MainWindow::getSelectionRange() const {
    if (!m_spreadsheetView) return "";

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return "";

    // In virtual mode, model row != logical row — must translate
    auto* model = m_spreadsheetView->getModel();

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        int logicalRow = model ? model->toLogicalRow(idx.row()) : idx.row();
        minRow = qMin(minRow, logicalRow);
        maxRow = qMax(maxRow, logicalRow);
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    return CellAddress(minRow, minCol).toString() + ":" + CellAddress(maxRow, maxCol).toString();
}

void MainWindow::onInsertChart() {
    ChartDialog dialog(this);
    dialog.setSpreadsheet(m_sheets[m_activeSheetIndex]);

    // Pre-fill with current selection; auto-detect table if single cell
    QString range = getSelectionRange();
    if (!range.isEmpty()) {
        CellRange cr(range);
        if (cr.isSingleCell()) {
            CellRange detected = m_spreadsheetView->detectDataRegion(cr.getStart().row, cr.getStart().col);
            if (!detected.isSingleCell()) {
                range = detected.toString();
            }
        }
        dialog.setDataRange(range);
    }

    if (dialog.exec() == QDialog::Accepted) {
        ChartConfig config = dialog.getConfig();

        // Auto-generate titles from data headers if not specified
        ChartWidget::autoGenerateTitles(config, m_sheets[m_activeSheetIndex]);

        auto* chart = createChartWidget(m_spreadsheetView->viewport());
        chart->setSpreadsheet(m_sheets[m_activeSheetIndex]);
        chart->setConfig(config);

        // Load data from spreadsheet range
        if (!config.dataRange.isEmpty()) {
            chart->loadDataFromRange(config.dataRange);
        }

        // Position in center of visible area
        QRect viewRect = m_spreadsheetView->viewport()->rect();
        int x = (viewRect.width() - 420) / 2;
        int y = (viewRect.height() - 320) / 2;
        chart->setGeometry(qMax(10, x), qMax(10, y), 420, 320);

        connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
        connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
        connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
        connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
        connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
        connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

        chart->setProperty("sheetIndex", m_activeSheetIndex);
        chart->show();
        m_charts.append(chart);
        addOverlay(chart);
        applyZOrder();

        setDirty();
        statusBar()->showMessage("Chart inserted");
    }
}

void MainWindow::onInsertShape() {
    InsertShapeDialog dialog(this);

    if (dialog.exec() == QDialog::Accepted) {
        ShapeConfig config = dialog.getConfig();

        auto* shape = new ShapeWidget(m_spreadsheetView->viewport());
        shape->setConfig(config);

        // Position in center of visible area
        QRect viewRect = m_spreadsheetView->viewport()->rect();
        int x = (viewRect.width() - 160) / 2;
        int y = (viewRect.height() - 120) / 2;
        shape->setGeometry(qMax(10, x), qMax(10, y), 160, 120);

        connect(shape, &ShapeWidget::editRequested, this, &MainWindow::onEditShape);
        connect(shape, &ShapeWidget::deleteRequested, this, &MainWindow::onDeleteShape);
        connect(shape, &ShapeWidget::shapeMoved, this, [this]() { setDirty(); });
        connect(shape, &ShapeWidget::shapeSelected, this, [this](ShapeWidget* s) {
            int si = s->property("sheetIndex").toInt();
            bool multiSelect = QApplication::keyboardModifiers() & Qt::ControlModifier;
            if (!multiSelect) {
                for (auto* other : m_shapes)
                    if (other != s && other->property("sheetIndex").toInt() == si) other->setSelected(false);
                for (auto* c : m_charts)
                    if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
                for (auto* img : m_images)
                    if (img->property("sheetIndex").toInt() == si) img->setSelected(false);
            }
            // Group-aware: select all group members
            OverlayGroup* grp = findGroupContaining(s);
            if (grp) {
                for (auto* w : grp->members) {
                    if (auto* chart = qobject_cast<ChartWidget*>(w)) chart->setSelected(true);
                    else if (auto* shape2 = qobject_cast<ShapeWidget*>(w)) shape2->setSelected(true);
                    else if (auto* img = qobject_cast<ImageWidget*>(w)) img->setSelected(true);
                }
            }
            m_selectedChart = nullptr;
        });

        shape->setProperty("sheetIndex", m_activeSheetIndex);
        shape->show();
        m_shapes.append(shape);
        addOverlay(shape);
        applyZOrder();

        setDirty();
        statusBar()->showMessage("Shape inserted");
    }
}

void MainWindow::highlightChartDataRange(ChartWidget* chart) {
    if (!chart || !m_spreadsheetView) return;

    auto cfg = chart->config();
    if (cfg.dataRange.isEmpty()) return;

    CellRange range(cfg.dataRange);
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;

    auto colors = ChartWidget::themeColors(cfg.themeIndex);
    QColor categoryColor(128, 0, 128); // purple for category/X column

    QVector<QPair<int, QColor>> seriesColumns;
    // First data column after category column gets series colors
    for (int c = startCol + 1; c <= endCol; ++c) {
        QColor sc = colors[(c - startCol - 1) % colors.size()];
        seriesColumns.append({c, sc});
    }

    m_spreadsheetView->setChartRangeHighlight(range, seriesColumns, categoryColor);
}

void MainWindow::onEditChart(ChartWidget* chart) {
    if (!chart) return;
    // Use the side panel for chart editing instead of a dialog
    onChartPropertiesRequested(chart);
}

void MainWindow::onDeleteChart(ChartWidget* chart) {
    if (!chart) return;

    if (chart == m_selectedChart) {
        m_spreadsheetView->clearChartRangeHighlight();
        m_selectedChart = nullptr;
    }
    removeOverlay(chart);
    m_charts.removeOne(chart);
    chart->hide();
    chart->deleteLater();
    setDirty();
    statusBar()->showMessage("Chart deleted");
}

void MainWindow::onEditShape(ShapeWidget* shape) {
    if (!shape) return;

    ShapePropertiesDialog dialog(shape->config(), this);
    if (dialog.exec() == QDialog::Accepted) {
        shape->setConfig(dialog.getConfig());
        statusBar()->showMessage("Shape updated");
    }
}

void MainWindow::onDeleteShape(ShapeWidget* shape) {
    if (!shape) return;

    removeOverlay(shape);
    m_shapes.removeOne(shape);
    shape->hide();
    shape->deleteLater();
    setDirty();
    statusBar()->showMessage("Shape deleted");
}

// ============== Image Insertion ==============

void MainWindow::onInsertImage() {
    QString fileName = QFileDialog::getOpenFileName(this, "Insert Image", "",
        "Image Files (*.png *.jpg *.jpeg *.bmp);;PNG (*.png);;JPEG (*.jpg *.jpeg);;BMP (*.bmp);;All Files (*)");
    if (fileName.isEmpty()) return;

    QPixmap pixmap(fileName);
    if (pixmap.isNull()) {
        QMessageBox::warning(this, "Insert Image", "Could not load image: " + fileName);
        return;
    }

    auto* image = new ImageWidget(m_spreadsheetView->viewport());
    image->setImageFromFile(fileName);

    // Scale to reasonable size while maintaining aspect ratio
    int maxW = 400, maxH = 300;
    int w = pixmap.width(), h = pixmap.height();
    if (w > maxW || h > maxH) {
        double scale = qMin(static_cast<double>(maxW) / w, static_cast<double>(maxH) / h);
        w = static_cast<int>(w * scale);
        h = static_cast<int>(h * scale);
    }

    QRect viewRect = m_spreadsheetView->viewport()->rect();
    int x = (viewRect.width() - w) / 2;
    int y = (viewRect.height() - h) / 2;
    image->setGeometry(qMax(10, x), qMax(10, y), w, h);

    connect(image, &ImageWidget::editRequested, this, &MainWindow::onEditImage);
    connect(image, &ImageWidget::deleteRequested, this, &MainWindow::onDeleteImage);
    connect(image, &ImageWidget::imageMoved, this, [this]() { setDirty(); });
    connect(image, &ImageWidget::imageSelected, this, [this](ImageWidget* img) {
        int si = img->property("sheetIndex").toInt();
        bool multiSelect = QApplication::keyboardModifiers() & Qt::ControlModifier;
        if (!multiSelect) {
            for (auto* other : m_images)
                if (other != img && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            for (auto* c : m_charts)
                if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
            for (auto* s : m_shapes)
                if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
        }
        // Group-aware: select all group members
        OverlayGroup* grp = findGroupContaining(img);
        if (grp) {
            for (auto* w : grp->members) {
                if (auto* chart = qobject_cast<ChartWidget*>(w)) chart->setSelected(true);
                else if (auto* shape = qobject_cast<ShapeWidget*>(w)) shape->setSelected(true);
                else if (auto* img2 = qobject_cast<ImageWidget*>(w)) img2->setSelected(true);
            }
        }
        m_selectedChart = nullptr;
    });

    image->setProperty("sheetIndex", m_activeSheetIndex);
    image->show();
    m_images.append(image);
    addOverlay(image);
    applyZOrder();

    statusBar()->showMessage("Image inserted: " + QFileInfo(fileName).fileName());
}

void MainWindow::onEditImage(ImageWidget* image) {
    if (!image) return;

    QString fileName = QFileDialog::getOpenFileName(this, "Replace Image", "",
        "Image Files (*.png *.jpg *.jpeg *.bmp);;All Files (*)");
    if (fileName.isEmpty()) return;

    image->setImageFromFile(fileName);
    statusBar()->showMessage("Image replaced");
}

void MainWindow::onDeleteImage(ImageWidget* image) {
    if (!image) return;

    removeOverlay(image);
    m_images.removeOne(image);
    image->hide();
    image->deleteLater();
    statusBar()->showMessage("Image deleted");
}

// ============== Sparkline Insertion ==============

void MainWindow::onInsertSparkline() {
    if (m_sheets.empty()) return;

    SparklineDialog dialog(this);

    // Pre-fill with current selection as data range
    QString range = getSelectionRange();
    if (!range.isEmpty()) {
        dialog.setDataRange(range);
    }

    if (dialog.exec() == QDialog::Accepted) {
        SparklineConfig config = dialog.getConfig();
        QString destStr = dialog.getDestinationRange();

        if (destStr.isEmpty()) {
            QMessageBox::warning(this, "Insert Sparkline", "Please specify a destination cell.");
            return;
        }

        // Parse destination — could be a single cell or a range
        CellAddress destStart = parseCellRef(destStr.split(':').first());
        auto sheet = m_sheets[m_activeSheetIndex];
        sheet->setSparkline(destStart, config);

        m_spreadsheetView->refreshView();
        if (m_spreadsheetView->getModel())
            m_spreadsheetView->getModel()->resetModel();

        statusBar()->showMessage("Sparkline inserted");
    }
}

// ============== Macro Editor ==============

void MainWindow::onMacroEditor() {
    if (!m_macroEngine) return;

    MacroEditorDialog dialog(m_macroEngine, this);
    dialog.exec();

    // Refresh view in case macros changed cell values
    m_spreadsheetView->refreshView();
    if (m_spreadsheetView->getModel())
        m_spreadsheetView->getModel()->resetModel();
}

void MainWindow::onRunLastMacro() {
    if (!m_macroEngine) return;

    auto macros = m_macroEngine->getSavedMacros();
    if (macros.isEmpty()) {
        statusBar()->showMessage("No saved macros to run");
        return;
    }

    // Run the most recently saved macro
    auto result = m_macroEngine->execute(macros.last().code);
    if (result.success) {
        statusBar()->showMessage("Macro executed: " + macros.last().name);
    } else {
        statusBar()->showMessage("Macro error: " + result.error, 5000);
    }

    m_spreadsheetView->refreshView();
    if (m_spreadsheetView->getModel())
        m_spreadsheetView->getModel()->resetModel();
}

// ============== Multi-select & Delete key ==============

void MainWindow::onChartSelected(ChartWidget* c) {
    int si = c->property("sheetIndex").toInt();
    bool multiSelect = QApplication::keyboardModifiers() & Qt::ControlModifier;

    if (!multiSelect) {
        for (auto* other : m_charts)
            if (other != c && other->property("sheetIndex").toInt() == si) other->setSelected(false);
        for (auto* s : m_shapes)
            if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
        for (auto* img : m_images)
            if (img->property("sheetIndex").toInt() == si) img->setSelected(false);
    }

    // If widget is in a group, select all group members
    OverlayGroup* grp = findGroupContaining(c);
    if (grp) {
        for (auto* w : grp->members) {
            if (auto* chart = qobject_cast<ChartWidget*>(w)) chart->setSelected(true);
            else if (auto* shape = qobject_cast<ShapeWidget*>(w)) shape->setSelected(true);
            else if (auto* img = qobject_cast<ImageWidget*>(w)) img->setSelected(true);
        }
    }

    m_spreadsheetView->selectionModel()->clearSelection();
    highlightChartDataRange(c);
    m_selectedChart = c;
}

void MainWindow::deselectAllOverlays() {
    for (auto* c : m_charts) {
        if (c->property("sheetIndex").toInt() == m_activeSheetIndex)
            c->setSelected(false);
    }
    for (auto* s : m_shapes) {
        if (s->property("sheetIndex").toInt() == m_activeSheetIndex)
            s->setSelected(false);
    }
    for (auto* img : m_images) {
        if (img->property("sheetIndex").toInt() == m_activeSheetIndex)
            img->setSelected(false);
    }
    // Clear chart data range highlights
    m_spreadsheetView->clearChartRangeHighlight();
    m_selectedChart = nullptr;

    // Hide chart properties panel when nothing is selected
    if (m_chartPropsDock && m_chartPropsDock->isVisible()) {
        m_chartPropsDock->hide();
    }
}

void MainWindow::deleteSelectedOverlays() {
    // Collect selected charts on the active sheet
    QVector<ChartWidget*> chartsToDelete;
    for (auto* c : m_charts) {
        if (c->isSelected() && c->property("sheetIndex").toInt() == m_activeSheetIndex)
            chartsToDelete.append(c);
    }
    for (auto* c : chartsToDelete) {
        removeOverlay(c);
        m_charts.removeOne(c);
        c->hide();
        c->deleteLater();
    }

    // Collect selected shapes on the active sheet
    QVector<ShapeWidget*> shapesToDelete;
    for (auto* s : m_shapes) {
        if (s->isSelected() && s->property("sheetIndex").toInt() == m_activeSheetIndex)
            shapesToDelete.append(s);
    }
    for (auto* s : shapesToDelete) {
        removeOverlay(s);
        m_shapes.removeOne(s);
        s->hide();
        s->deleteLater();
    }

    // Collect selected images on the active sheet
    QVector<ImageWidget*> imagesToDelete;
    for (auto* img : m_images) {
        if (img->isSelected() && img->property("sheetIndex").toInt() == m_activeSheetIndex)
            imagesToDelete.append(img);
    }
    for (auto* img : imagesToDelete) {
        removeOverlay(img);
        m_images.removeOne(img);
        img->hide();
        img->deleteLater();
    }

    int total = chartsToDelete.size() + shapesToDelete.size() + imagesToDelete.size();
    if (total > 0) {
        statusBar()->showMessage(QString("Deleted %1 object(s)").arg(total));
    }
}

// ============== Overlay Z-Order & Grouping ==============

// ---------------------------------------------------------------------------
// Cell-anchor helpers: Excel-like overlay positioning
// Each overlay stores dynamic properties: anchorRow, anchorCol, anchorOffsetX,
// anchorOffsetY — the pixel offset from the top-left of the anchor cell.
// On scroll we recompute the viewport pixel position from these anchors.
// ---------------------------------------------------------------------------

void MainWindow::anchorOverlayToCell(QWidget* overlay) {
    // Convert the overlay's current viewport position to a cell anchor.
    QPoint pos = overlay->pos();  // relative to viewport

    // Find the row whose top edge is at or above pos.y()
    int anchorRow = m_spreadsheetView->rowAt(pos.y());
    int anchorCol = m_spreadsheetView->columnAt(pos.x());

    // If pos is above/left of all visible rows/cols, clamp to 0
    if (anchorRow < 0) anchorRow = 0;
    if (anchorCol < 0) anchorCol = 0;

    // Pixel offset within the anchor cell
    int cellX = m_spreadsheetView->columnViewportPosition(anchorCol);
    int cellY = m_spreadsheetView->rowViewportPosition(anchorRow);
    int offsetX = pos.x() - cellX;
    int offsetY = pos.y() - cellY;

    overlay->setProperty("anchorRow", anchorRow);
    overlay->setProperty("anchorCol", anchorCol);
    overlay->setProperty("anchorOffsetX", offsetX);
    overlay->setProperty("anchorOffsetY", offsetY);
}

void MainWindow::repositionAnchoredOverlay(QWidget* overlay) {
    if (overlay->property("sheetIndex").toInt() != m_activeSheetIndex) return;

    // Skip overlays being dragged — they follow the cursor
    auto* chart = qobject_cast<ChartWidget*>(overlay);
    if (chart && chart->isDragging()) return;
    auto* shape = qobject_cast<ShapeWidget*>(overlay);
    if (shape && shape->isDragging()) return;

    int row = overlay->property("anchorRow").toInt();
    int col = overlay->property("anchorCol").toInt();
    int offX = overlay->property("anchorOffsetX").toInt();
    int offY = overlay->property("anchorOffsetY").toInt();

    // Compute pixel position of anchor cell in viewport
    int cellX = m_spreadsheetView->columnViewportPosition(col);
    int cellY = m_spreadsheetView->rowViewportPosition(row);

    // If anchor cell is scrolled out of view, rowViewportPosition returns -1.
    // Accumulate row/col heights to get the correct off-screen position.
    if (cellY == -1) {
        int firstVisRow = m_spreadsheetView->rowAt(0);
        if (firstVisRow < 0) firstVisRow = 0;
        cellY = 0;
        if (row < firstVisRow) {
            for (int r = row; r < firstVisRow; ++r)
                cellY -= m_spreadsheetView->rowHeight(r);
        } else {
            for (int r = firstVisRow; r < row; ++r)
                cellY += m_spreadsheetView->rowHeight(r);
        }
    }
    if (cellX == -1) {
        int firstVisCol = m_spreadsheetView->columnAt(0);
        if (firstVisCol < 0) firstVisCol = 0;
        cellX = 0;
        if (col < firstVisCol) {
            for (int c = col; c < firstVisCol; ++c)
                cellX -= m_spreadsheetView->columnWidth(c);
        } else {
            for (int c = firstVisCol; c < col; ++c)
                cellX += m_spreadsheetView->columnWidth(c);
        }
    }

    overlay->move(cellX + offX, cellY + offY);
}

void MainWindow::repositionAllOverlays() {
    for (auto* c : m_charts) repositionAnchoredOverlay(c);
    for (auto* s : m_shapes) repositionAnchoredOverlay(s);
    for (auto* i : m_images) repositionAnchoredOverlay(i);
}

void MainWindow::addOverlay(QWidget* w) {
    if (!m_overlayZOrder.contains(w))
        m_overlayZOrder.append(w);

    // Anchor to the cell under the overlay's current position
    anchorOverlayToCell(w);

    // Re-anchor when overlay is moved or resized by the user
    if (auto* chart = qobject_cast<ChartWidget*>(w)) {
        connect(chart, &ChartWidget::chartMoved, this, [this, chart]() { anchorOverlayToCell(chart); });
        connect(chart, &ChartWidget::chartResized, this, [this, chart]() { anchorOverlayToCell(chart); });
    }
    if (auto* shape = qobject_cast<ShapeWidget*>(w)) {
        connect(shape, &ShapeWidget::shapeMoved, this, [this, shape]() { anchorOverlayToCell(shape); });
    }
}

void MainWindow::removeOverlay(QWidget* w) {
    m_overlayZOrder.removeAll(w);
    // Also remove from any group
    for (int g = m_overlayGroups.size() - 1; g >= 0; --g) {
        m_overlayGroups[g].members.removeAll(w);
        if (m_overlayGroups[g].members.size() <= 1)
            m_overlayGroups.removeAt(g);
    }
}

void MainWindow::applyZOrder() {
    // Raise overlays on the active sheet in z-order (index 0 = back, last = front)
    for (auto* w : m_overlayZOrder) {
        if (w->property("sheetIndex").toInt() == m_activeSheetIndex && w->isVisible())
            w->raise();
    }
}

void MainWindow::bringToFront(QWidget* w) {
    m_overlayZOrder.removeAll(w);
    m_overlayZOrder.append(w);
    applyZOrder();
    setDirty();
}

void MainWindow::sendToBack(QWidget* w) {
    m_overlayZOrder.removeAll(w);
    m_overlayZOrder.prepend(w);
    applyZOrder();
    setDirty();
}

void MainWindow::bringForward(QWidget* w) {
    int idx = m_overlayZOrder.indexOf(w);
    if (idx < 0 || idx >= m_overlayZOrder.size() - 1) return;
    m_overlayZOrder.swapItemsAt(idx, idx + 1);
    applyZOrder();
    setDirty();
}

void MainWindow::sendBackward(QWidget* w) {
    int idx = m_overlayZOrder.indexOf(w);
    if (idx <= 0) return;
    m_overlayZOrder.swapItemsAt(idx, idx - 1);
    applyZOrder();
    setDirty();
}

QVector<QWidget*> MainWindow::selectedOverlays() const {
    QVector<QWidget*> sel;
    for (auto* c : m_charts) {
        if (c->isSelected() && c->property("sheetIndex").toInt() == m_activeSheetIndex)
            sel.append(c);
    }
    for (auto* s : m_shapes) {
        if (s->isSelected() && s->property("sheetIndex").toInt() == m_activeSheetIndex)
            sel.append(s);
    }
    for (auto* img : m_images) {
        if (img->isSelected() && img->property("sheetIndex").toInt() == m_activeSheetIndex)
            sel.append(img);
    }
    return sel;
}

OverlayGroup* MainWindow::findGroupContaining(QWidget* w) {
    for (auto& grp : m_overlayGroups) {
        if (grp.members.contains(w))
            return &grp;
    }
    return nullptr;
}

void MainWindow::groupSelectedOverlays() {
    QVector<QWidget*> sel = selectedOverlays();
    if (sel.size() < 2) {
        statusBar()->showMessage("Select at least 2 overlays to group");
        return;
    }

    // Remove selected widgets from any existing groups
    for (auto* w : sel) {
        for (int g = m_overlayGroups.size() - 1; g >= 0; --g) {
            m_overlayGroups[g].members.removeAll(w);
            if (m_overlayGroups[g].members.size() <= 1)
                m_overlayGroups.removeAt(g);
        }
    }

    OverlayGroup grp;
    grp.id = m_nextGroupId++;
    grp.members = sel;
    m_overlayGroups.append(grp);

    // Tag each widget with group id
    for (auto* w : sel)
        w->setProperty("overlayGroupId", grp.id);

    statusBar()->showMessage(QString("Grouped %1 objects").arg(sel.size()));
    setDirty();
}

void MainWindow::ungroupSelectedOverlays() {
    QVector<QWidget*> sel = selectedOverlays();
    QSet<int> groupsToRemove;

    for (auto* w : sel) {
        int gid = w->property("overlayGroupId").toInt();
        if (gid > 0) groupsToRemove.insert(gid);
    }

    for (int g = m_overlayGroups.size() - 1; g >= 0; --g) {
        if (groupsToRemove.contains(m_overlayGroups[g].id)) {
            for (auto* w : m_overlayGroups[g].members)
                w->setProperty("overlayGroupId", 0);
            m_overlayGroups.removeAt(g);
        }
    }

    if (!groupsToRemove.isEmpty()) {
        statusBar()->showMessage("Ungrouped");
        setDirty();
    }
}

void MainWindow::keyPressEvent(QKeyEvent* event) {
    // Check if any chart/shape on the active sheet is selected
    bool hasSelectedOverlay = false;
    for (auto* c : m_charts) {
        if (c->isSelected() && c->property("sheetIndex").toInt() == m_activeSheetIndex)
            { hasSelectedOverlay = true; break; }
    }
    if (!hasSelectedOverlay) {
        for (auto* s : m_shapes) {
            if (s->isSelected() && s->property("sheetIndex").toInt() == m_activeSheetIndex)
                { hasSelectedOverlay = true; break; }
        }
    }
    if (!hasSelectedOverlay) {
        for (auto* img : m_images) {
            if (img->isSelected() && img->property("sheetIndex").toInt() == m_activeSheetIndex)
                { hasSelectedOverlay = true; break; }
        }
    }

    if (hasSelectedOverlay && (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace)) {
        deleteSelectedOverlays();
        return;
    }

    // Escape deselects all overlays
    if (event->key() == Qt::Key_Escape) {
        deselectAllOverlays();
    }

    // Ctrl+PageUp: Switch to previous sheet tab
    bool ctrl = event->modifiers() & Qt::ControlModifier;
    if (ctrl && event->key() == Qt::Key_PageUp) {
        int count = m_sheetTabBar->count();
        if (count > 1) {
            int newIndex = m_activeSheetIndex - 1;
            if (newIndex < 0) newIndex = count - 1;
            m_sheetTabBar->setCurrentIndex(newIndex);
        }
        event->accept();
        return;
    }

    // Ctrl+PageDown: Switch to next sheet tab
    if (ctrl && event->key() == Qt::Key_PageDown) {
        int count = m_sheetTabBar->count();
        if (count > 1) {
            int newIndex = m_activeSheetIndex + 1;
            if (newIndex >= count) newIndex = 0;
            m_sheetTabBar->setCurrentIndex(newIndex);
        }
        event->accept();
        return;
    }

    QMainWindow::keyPressEvent(event);
}

bool MainWindow::eventFilter(QObject* obj, QEvent* event) {
    // When user clicks on the spreadsheet viewport, deselect all chart/shape overlays
    if (obj == m_spreadsheetView->viewport() && event->type() == QEvent::MouseButtonPress) {
        auto* me = static_cast<QMouseEvent*>(event);
        // Check that the click is NOT on a chart or shape widget
        QWidget* child = m_spreadsheetView->viewport()->childAt(me->pos());
        bool clickedOverlay = false;
        for (auto* c : m_charts) {
            if (child == c) { clickedOverlay = true; break; }
        }
        if (!clickedOverlay) {
            for (auto* s : m_shapes) {
                if (child == s) { clickedOverlay = true; break; }
            }
        }
        if (!clickedOverlay) {
            for (auto* img : m_images) {
                if (child == img) { clickedOverlay = true; break; }
            }
        }
        if (!clickedOverlay) {
            deselectAllOverlays();
        }
    }
    return QMainWindow::eventFilter(obj, event);
}

// ============== Chat-driven Chart/Shape insertion ==============

void MainWindow::insertChartFromChat(const QJsonObject& params) {
    ChartConfig config;

    // Parse chart type
    QString typeStr = params["type"].toString().toLower();
    if (typeStr == "line") config.type = ChartType::Line;
    else if (typeStr == "bar") config.type = ChartType::Bar;
    else if (typeStr == "scatter") config.type = ChartType::Scatter;
    else if (typeStr == "pie") config.type = ChartType::Pie;
    else if (typeStr == "area") config.type = ChartType::Area;
    else if (typeStr == "donut") config.type = ChartType::Donut;
    else if (typeStr == "histogram") config.type = ChartType::Histogram;
    else config.type = ChartType::Column;

    config.title = params["title"].toString();
    config.dataRange = params["range"].toString();
    config.xAxisTitle = params["x_axis"].toString();
    config.yAxisTitle = params["y_axis"].toString();
    config.themeIndex = params["theme"].toInt(0);
    config.showLegend = !params.contains("show_legend") || params["show_legend"].toBool(true);
    config.showGridLines = !params.contains("show_grid") || params["show_grid"].toBool(true);

    // Auto-generate titles from data headers if not specified
    ChartWidget::autoGenerateTitles(config, m_sheets[m_activeSheetIndex]);

    auto* chart = createChartWidget(m_spreadsheetView->viewport());
    chart->setSpreadsheet(m_sheets[m_activeSheetIndex]);
    chart->setConfig(config);

    if (!config.dataRange.isEmpty()) {
        chart->loadDataFromRange(config.dataRange);
    }

    QRect viewRect = m_spreadsheetView->viewport()->rect();
    int x = (viewRect.width() - 420) / 2;
    int y = (viewRect.height() - 320) / 2;
    chart->setGeometry(qMax(10, x), qMax(10, y), 420, 320);

    connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
    connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
    connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
    connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
    connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
    connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

    chart->setProperty("sheetIndex", m_activeSheetIndex);
    chart->show();
    m_charts.append(chart);
    addOverlay(chart);
    applyZOrder();
    setDirty();
}

void MainWindow::insertShapeFromChat(const QJsonObject& params) {
    ShapeConfig config;

    QString typeStr = params["type"].toString().toLower();
    if (typeStr == "rectangle" || typeStr == "rect") config.type = ShapeType::Rectangle;
    else if (typeStr == "rounded_rect" || typeStr == "rounded") config.type = ShapeType::RoundedRect;
    else if (typeStr == "circle") config.type = ShapeType::Circle;
    else if (typeStr == "ellipse") config.type = ShapeType::Ellipse;
    else if (typeStr == "triangle") config.type = ShapeType::Triangle;
    else if (typeStr == "star") config.type = ShapeType::Star;
    else if (typeStr == "arrow") config.type = ShapeType::Arrow;
    else if (typeStr == "diamond") config.type = ShapeType::Diamond;
    else if (typeStr == "pentagon") config.type = ShapeType::Pentagon;
    else if (typeStr == "hexagon") config.type = ShapeType::Hexagon;
    else if (typeStr == "callout") config.type = ShapeType::Callout;
    else if (typeStr == "line") config.type = ShapeType::Line;
    else config.type = ShapeType::Rectangle;

    if (params.contains("fill_color")) config.fillColor = QColor(params["fill_color"].toString());
    if (params.contains("stroke_color")) config.strokeColor = QColor(params["stroke_color"].toString());
    if (params.contains("stroke_width")) config.strokeWidth = params["stroke_width"].toInt(2);
    if (params.contains("text")) config.text = params["text"].toString();
    if (params.contains("text_color")) config.textColor = QColor(params["text_color"].toString());
    if (params.contains("font_size")) config.fontSize = params["font_size"].toInt(12);
    if (params.contains("opacity")) config.opacity = static_cast<float>(params["opacity"].toDouble(1.0));

    auto* shape = new ShapeWidget(m_spreadsheetView->viewport());
    shape->setConfig(config);

    int w = params["width"].toInt(160);
    int h = params["height"].toInt(120);
    QRect viewRect = m_spreadsheetView->viewport()->rect();
    int x = (viewRect.width() - w) / 2;
    int y = (viewRect.height() - h) / 2;
    shape->setGeometry(qMax(10, x), qMax(10, y), w, h);

    connect(shape, &ShapeWidget::editRequested, this, &MainWindow::onEditShape);
    connect(shape, &ShapeWidget::deleteRequested, this, &MainWindow::onDeleteShape);
    connect(shape, &ShapeWidget::shapeSelected, this, [this](ShapeWidget* s) {
        int si = s->property("sheetIndex").toInt();
        bool multiSelect = QApplication::keyboardModifiers() & Qt::ControlModifier;
        if (!multiSelect) {
            for (auto* other : m_shapes) if (other != s && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            for (auto* c : m_charts) if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
            for (auto* img : m_images) if (img->property("sheetIndex").toInt() == si) img->setSelected(false);
        }
        OverlayGroup* grp = findGroupContaining(s);
        if (grp) {
            for (auto* w : grp->members) {
                if (auto* chart = qobject_cast<ChartWidget*>(w)) chart->setSelected(true);
                else if (auto* shape2 = qobject_cast<ShapeWidget*>(w)) shape2->setSelected(true);
                else if (auto* img = qobject_cast<ImageWidget*>(w)) img->setSelected(true);
            }
        }
        m_selectedChart = nullptr;
    });

    shape->setProperty("sheetIndex", m_activeSheetIndex);
    shape->show();
    m_shapes.append(shape);
    addOverlay(shape);
    applyZOrder();
}

void MainWindow::insertImageFromChat(const QJsonObject& params) {
    QString path = params["path"].toString();
    if (path.isEmpty()) return;

    QPixmap pixmap(path);
    if (pixmap.isNull()) return;

    auto* image = new ImageWidget(m_spreadsheetView->viewport());
    image->setImageFromFile(path);

    int w = params["width"].toInt(0);
    int h = params["height"].toInt(0);
    if (w <= 0 || h <= 0) {
        w = qMin(pixmap.width(), 400);
        h = qMin(pixmap.height(), 300);
        if (pixmap.width() > 400 || pixmap.height() > 300) {
            double scale = qMin(400.0 / pixmap.width(), 300.0 / pixmap.height());
            w = static_cast<int>(pixmap.width() * scale);
            h = static_cast<int>(pixmap.height() * scale);
        }
    }

    QRect viewRect = m_spreadsheetView->viewport()->rect();
    int x = (viewRect.width() - w) / 2;
    int y = (viewRect.height() - h) / 2;
    image->setGeometry(qMax(10, x), qMax(10, y), w, h);

    connect(image, &ImageWidget::editRequested, this, &MainWindow::onEditImage);
    connect(image, &ImageWidget::deleteRequested, this, &MainWindow::onDeleteImage);
    connect(image, &ImageWidget::imageMoved, this, [this]() { setDirty(); });
    connect(image, &ImageWidget::imageSelected, this, [this](ImageWidget* img) {
        int si = img->property("sheetIndex").toInt();
        bool multiSelect = QApplication::keyboardModifiers() & Qt::ControlModifier;
        if (!multiSelect) {
            for (auto* other : m_images) if (other != img && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            for (auto* c : m_charts) if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
            for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
        }
        OverlayGroup* grp = findGroupContaining(img);
        if (grp) {
            for (auto* w : grp->members) {
                if (auto* chart = qobject_cast<ChartWidget*>(w)) chart->setSelected(true);
                else if (auto* shape = qobject_cast<ShapeWidget*>(w)) shape->setSelected(true);
                else if (auto* img2 = qobject_cast<ImageWidget*>(w)) img2->setSelected(true);
            }
        }
        m_selectedChart = nullptr;
    });

    image->setProperty("sheetIndex", m_activeSheetIndex);
    image->show();
    m_images.append(image);
    addOverlay(image);
    applyZOrder();
}

// ============== Chart Properties Panel ==============

void MainWindow::onChartPropertiesRequested(ChartWidget* chart) {
    if (!chart || !m_chartPropsPanel || !m_chartPropsDock) return;

    m_chartPropsPanel->setChart(chart);
    m_chartPropsDock->show();
    m_chartPropsDock->raise();
}

// ============== Pivot Table ==============

void MainWindow::onCreatePivotTable() {
    if (m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];

    // Detect data range from selection or auto-detect
    QString rangeStr = getSelectionRange();
    CellRange sourceRange;
    if (!rangeStr.isEmpty()) {
        sourceRange = CellRange(rangeStr);
        // Single cell selected — auto-detect the contiguous data region
        if (sourceRange.isSingleCell()) {
            CellRange detected = m_spreadsheetView->detectDataRegion(
                sourceRange.getStart().row, sourceRange.getStart().col);
            if (!detected.isSingleCell()) {
                sourceRange = detected;
            }
        }
    } else {
        int maxRow = sheet->getMaxRow();
        int maxCol = sheet->getMaxColumn();
        if (maxRow < 0 || maxCol < 0) {
            QMessageBox::information(this, "Pivot Table",
                "Please select a data range or enter data first.");
            return;
        }
        sourceRange = CellRange(0, 0, maxRow, maxCol);
    }

    PivotTableDialog dialog(sheet, sourceRange, this);
    if (dialog.exec() == QDialog::Accepted) {
        PivotConfig config = dialog.getConfig();
        config.sourceSheetIndex = m_activeSheetIndex;

        if (config.valueFields.empty()) {
            QMessageBox::warning(this, "Pivot Table", "Please add at least one value field.");
            return;
        }

        PivotEngine engine;
        engine.setSource(sheet, config);
        PivotResult result = engine.compute();

        // Create a new sheet for the pivot output
        auto pivotSheet = std::make_shared<Spreadsheet>();
        pivotSheet->setSheetName("Pivot - " + sheet->getSheetName());
        engine.writeToSheet(pivotSheet, result, config);

        // Store pivot config for refresh
        pivotSheet->setPivotConfig(std::make_unique<PivotConfig>(config));

        // Add the pivot sheet
        m_sheets.push_back(pivotSheet);
        m_sheetTabBar->addTab(pivotSheet->getSheetName());
        int pivotSheetIdx = static_cast<int>(m_sheets.size()) - 1;
        m_sheetTabBar->setCurrentIndex(pivotSheetIdx);

        // Auto-generate chart if requested
        if (config.autoChart && !result.rowLabels.empty()) {
            ChartConfig chartCfg;
            chartCfg.type = static_cast<ChartType>(config.chartType);
            chartCfg.title = config.valueFields[0].displayName();
            chartCfg.showLegend = true;
            chartCfg.showGridLines = true;

            // Account for filter rows written at top of pivot sheet (all filter fields)
            int filterRowOffset = static_cast<int>(config.filterFields.size());
            if (filterRowOffset > 0) filterRowOffset++; // blank separator row

            // Build chart data range from pivot output (exclude Grand Total column)
            int headerRow = filterRowOffset + result.dataStartRow - 1;
            if (headerRow < 0) headerRow = 0;
            int endRow = headerRow + static_cast<int>(result.rowLabels.size());
            int numDataCols = static_cast<int>(result.columnLabels.size());
            if (config.showGrandTotalColumn)
                numDataCols -= static_cast<int>(config.valueFields.size());
            int endCol = result.numRowHeaderColumns + numDataCols - 1;
            if (endCol < result.numRowHeaderColumns) endCol = result.numRowHeaderColumns;
            chartCfg.dataRange = CellRange(headerRow, 0, endRow, endCol).toString();

            // Pre-populate category labels from pivot row labels
            for (const auto& rowLabel : result.rowLabels) {
                chartCfg.categoryLabels.append(rowLabel.empty() ? "" : rowLabel[0]);
            }

            auto* chart = createChartWidget(m_spreadsheetView->viewport());
            chart->setSpreadsheet(pivotSheet);
            chart->setConfig(chartCfg);
            chart->loadDataFromRange(chartCfg.dataRange);

            QRect viewRect = m_spreadsheetView->viewport()->rect();
            chart->setGeometry(qMax(10, viewRect.width() / 2 - 50), 20, 420, 320);

            connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
            connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
            connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
            connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
            connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
            connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

            chart->setProperty("sheetIndex", pivotSheetIdx);
            chart->show();
            chart->startEntryAnimation();
            m_charts.append(chart);
            addOverlay(chart);
            applyZOrder();
        }

        // Enable auto-filter on pivot output header row
        {
            // Account for filter rows (all filter fields are always written)
            int filterRowOffset = static_cast<int>(config.filterFields.size());
            if (filterRowOffset > 0) filterRowOffset++; // blank row after filter rows

            // Header row is at filterRowOffset + (dataStartRow - 1)
            // dataStartRow is the number of column header rows (typically 1)
            int headerRow = filterRowOffset + result.dataStartRow - 1;
            if (headerRow < 0) headerRow = 0;
            int endRow = headerRow + static_cast<int>(result.rowLabels.size());
            int endCol = result.numRowHeaderColumns + static_cast<int>(result.columnLabels.size()) - 1;
            if (endCol < 0) endCol = 0;
            m_spreadsheetView->enableAutoFilter(CellRange(headerRow, 0, endRow, endCol));
        }

        statusBar()->showMessage("Pivot table created on sheet: " + pivotSheet->getSheetName());
    }
}

void MainWindow::onRefreshPivotTable() {
    if (m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];
    const PivotConfig* config = sheet->getPivotConfig();
    if (!config) {
        QMessageBox::information(this, "Refresh Pivot Table",
            "The current sheet is not a pivot table.");
        return;
    }

    int srcIdx = config->sourceSheetIndex;
    if (srcIdx < 0 || srcIdx >= static_cast<int>(m_sheets.size())) {
        QMessageBox::warning(this, "Refresh Pivot Table",
            "Source sheet no longer exists.");
        return;
    }

    auto sourceSheet = m_sheets[srcIdx];
    PivotEngine engine;
    engine.setSource(sourceSheet, *config);
    PivotResult result = engine.compute();

    // Clear and rewrite the pivot sheet
    sheet->clearRange(CellRange(0, 0, sheet->getMaxRow() + 1, sheet->getMaxColumn() + 1));
    engine.writeToSheet(sheet, result, *config);

    m_spreadsheetView->setSpreadsheet(sheet); // refresh view
    statusBar()->showMessage("Pivot table refreshed");
}

void MainWindow::onPivotFilterChanged(int filterIndex, QStringList /*selectedValues*/) {
    if (m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];
    const PivotConfig* existing = sheet->getPivotConfig();
    if (!existing || filterIndex < 0
        || filterIndex >= static_cast<int>(existing->filterFields.size()))
        return;

    int srcIdx = existing->sourceSheetIndex;
    if (srcIdx < 0 || srcIdx >= static_cast<int>(m_sheets.size())) return;

    auto sourceSheet = m_sheets[srcIdx];
    const PivotFilterField& ff = existing->filterFields[filterIndex];

    // Get all unique values for this filter field from the source
    PivotEngine engine;
    QStringList uniqueVals = engine.getUniqueValues(
        sourceSheet, existing->sourceRange, ff.sourceColumnIndex);
    if (uniqueVals.isEmpty()) return;

    // Show value picker dialog
    QDialog dlg(this);
    dlg.setWindowTitle("Filter: " + ff.name);
    dlg.setMinimumSize(280, 360);
    dlg.setStyleSheet(ThemeManager::dialogStylesheet());

    QVBoxLayout* layout = new QVBoxLayout(&dlg);
    layout->addWidget(new QLabel("Select values to include:"));

    // Search box
    QLineEdit* searchBox = new QLineEdit();
    searchBox->setPlaceholderText("Search...");
    layout->addWidget(searchBox);

    QListWidget* valList = new QListWidget();
    valList->setUniformItemSizes(true);
    for (const auto& val : uniqueVals) {
        auto* item = new QListWidgetItem(val, valList);
        item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
        // Pre-check: if selectedValues is empty (= all) or contains this value
        bool checked = ff.selectedValues.isEmpty()
                       || ff.selectedValues.contains(val);
        item->setCheckState(checked ? Qt::Checked : Qt::Unchecked);
        item->setForeground(QColor("#1D2939"));
    }
    layout->addWidget(valList);

    // Search filtering
    connect(searchBox, &QLineEdit::textChanged, [valList](const QString& text) {
        for (int i = 0; i < valList->count(); ++i) {
            auto* item = valList->item(i);
            item->setHidden(!text.isEmpty()
                && !item->text().contains(text, Qt::CaseInsensitive));
        }
    });

    // Select All / Clear buttons
    QHBoxLayout* selBtnLayout = new QHBoxLayout();
    auto* selectAllBtn = new QPushButton("Select All");
    auto* selectNoneBtn = new QPushButton("Clear");
    selBtnLayout->addWidget(selectAllBtn);
    selBtnLayout->addWidget(selectNoneBtn);
    selBtnLayout->addStretch();
    layout->addLayout(selBtnLayout);

    connect(selectAllBtn, &QPushButton::clicked, [valList]() {
        for (int i = 0; i < valList->count(); ++i)
            if (!valList->item(i)->isHidden())
                valList->item(i)->setCheckState(Qt::Checked);
    });
    connect(selectNoneBtn, &QPushButton::clicked, [valList]() {
        for (int i = 0; i < valList->count(); ++i)
            if (!valList->item(i)->isHidden())
                valList->item(i)->setCheckState(Qt::Unchecked);
    });

    QDialogButtonBox* btns = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    connect(btns, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(btns, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);
    layout->addWidget(btns);

    if (dlg.exec() != QDialog::Accepted) return;

    // Collect selected values
    QStringList selected;
    for (int i = 0; i < valList->count(); ++i) {
        if (valList->item(i)->checkState() == Qt::Checked)
            selected.append(valList->item(i)->text());
    }

    // Build updated config
    PivotConfig newConfig = *existing;
    if (selected.size() == uniqueVals.size()) {
        newConfig.filterFields[filterIndex].selectedValues.clear(); // all = no filter
    } else {
        newConfig.filterFields[filterIndex].selectedValues = selected;
    }

    // Recompute pivot
    engine.setSource(sourceSheet, newConfig);
    PivotResult result = engine.compute();

    // Clear and rewrite pivot sheet
    sheet->clearRange(CellRange(0, 0, sheet->getMaxRow() + 1, sheet->getMaxColumn() + 1));
    engine.writeToSheet(sheet, result, newConfig);
    sheet->setPivotConfig(std::make_unique<PivotConfig>(newConfig));

    m_spreadsheetView->setSpreadsheet(sheet);

    // Update chart if auto-chart is on
    if (newConfig.autoChart && !result.rowLabels.empty()) {
        // Remove existing charts on this sheet
        for (int i = m_charts.size() - 1; i >= 0; --i) {
            if (m_charts[i]->property("sheetIndex").toInt() == m_activeSheetIndex) {
                removeOverlay(m_charts[i]);
                m_charts[i]->deleteLater();
                m_charts.removeAt(i);
            }
        }

        ChartConfig chartCfg;
        chartCfg.type = static_cast<ChartType>(newConfig.chartType);
        chartCfg.title = newConfig.valueFields[0].displayName();
        chartCfg.showLegend = true;
        chartCfg.showGridLines = true;

        int filterRowOffset = static_cast<int>(newConfig.filterFields.size());
        if (filterRowOffset > 0) filterRowOffset++; // blank separator

        int headerRow = filterRowOffset + result.dataStartRow - 1;
        if (headerRow < 0) headerRow = 0;
        int endRow = headerRow + static_cast<int>(result.rowLabels.size());
        int numDataCols = static_cast<int>(result.columnLabels.size());
        if (newConfig.showGrandTotalColumn)
            numDataCols -= static_cast<int>(newConfig.valueFields.size());
        int endCol = result.numRowHeaderColumns + numDataCols - 1;
        if (endCol < result.numRowHeaderColumns) endCol = result.numRowHeaderColumns;
        chartCfg.dataRange = CellRange(headerRow, 0, endRow, endCol).toString();

        for (const auto& rowLabel : result.rowLabels)
            chartCfg.categoryLabels.append(rowLabel.empty() ? "" : rowLabel[0]);

        auto* chart = createChartWidget(m_spreadsheetView->viewport());
        chart->setSpreadsheet(sheet);
        chart->setConfig(chartCfg);
        chart->loadDataFromRange(chartCfg.dataRange);

        QRect viewRect = m_spreadsheetView->viewport()->rect();
        chart->setGeometry(qMax(10, viewRect.width() / 2 - 50), 20, 420, 320);
        chart->setProperty("sheetIndex", m_activeSheetIndex);
        connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
        connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
        connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
        connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

        chart->show();
        chart->startEntryAnimation();
        m_charts.append(chart);
        addOverlay(chart);
        applyZOrder();
    }

    // Re-enable auto-filter on pivot header row
    {
        int filterRowOffset = static_cast<int>(newConfig.filterFields.size());
        if (filterRowOffset > 0) filterRowOffset++;
        int hdrRow = filterRowOffset + result.dataStartRow - 1;
        if (hdrRow < 0) hdrRow = 0;
        int endRow = hdrRow + static_cast<int>(result.rowLabels.size());
        int endCol = result.numRowHeaderColumns
                     + static_cast<int>(result.columnLabels.size()) - 1;
        if (endCol < 0) endCol = 0;
        m_spreadsheetView->enableAutoFilter(CellRange(hdrRow, 0, endRow, endCol));
    }

    setDirty();
    statusBar()->showMessage("Pivot filter updated");
}

void MainWindow::onEditPivotTable() {
    if (m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];
    const PivotConfig* existing = sheet->getPivotConfig();
    if (!existing) {
        QMessageBox::information(this, "Edit Pivot Table",
            "The current sheet is not a pivot table.");
        return;
    }

    int srcIdx = existing->sourceSheetIndex;
    if (srcIdx < 0 || srcIdx >= static_cast<int>(m_sheets.size())) {
        QMessageBox::warning(this, "Edit Pivot Table",
            "Source sheet no longer exists.");
        return;
    }

    auto sourceSheet = m_sheets[srcIdx];
    PivotTableDialog dialog(sourceSheet, existing->sourceRange, this);
    dialog.setWindowTitle("Edit Pivot Table");
    dialog.loadConfig(*existing);

    if (dialog.exec() == QDialog::Accepted) {
        PivotConfig config = dialog.getConfig();
        config.sourceSheetIndex = srcIdx;

        if (config.valueFields.empty()) {
            QMessageBox::warning(this, "Edit Pivot Table",
                "Please add at least one value field.");
            return;
        }

        PivotEngine engine;
        engine.setSource(sourceSheet, config);
        PivotResult result = engine.compute();

        // Clear and rewrite the pivot sheet
        sheet->clearRange(CellRange(0, 0, sheet->getMaxRow() + 1, sheet->getMaxColumn() + 1));
        engine.writeToSheet(sheet, result, config);

        // Update stored config
        sheet->setPivotConfig(std::make_unique<PivotConfig>(config));

        m_spreadsheetView->setSpreadsheet(sheet);

        // Update chart if auto-chart is on
        if (config.autoChart && !result.rowLabels.empty()) {
            // Remove existing charts on this sheet
            for (int i = m_charts.size() - 1; i >= 0; --i) {
                if (m_charts[i]->property("sheetIndex").toInt() == m_activeSheetIndex) {
                    removeOverlay(m_charts[i]);
                    m_charts[i]->deleteLater();
                    m_charts.removeAt(i);
                }
            }

            ChartConfig chartCfg;
            chartCfg.type = static_cast<ChartType>(config.chartType);
            chartCfg.title = config.valueFields[0].displayName();
            chartCfg.showLegend = true;
            chartCfg.showGridLines = true;

            // Account for filter rows (all filter fields are always written)
            int filterRowOffset = static_cast<int>(config.filterFields.size());
            if (filterRowOffset > 0) filterRowOffset++; // blank separator row

            int headerRow = filterRowOffset + result.dataStartRow - 1;
            if (headerRow < 0) headerRow = 0;
            int endRow = headerRow + static_cast<int>(result.rowLabels.size());
            int numDataCols = static_cast<int>(result.columnLabels.size());
            if (config.showGrandTotalColumn)
                numDataCols -= static_cast<int>(config.valueFields.size());
            int endCol = result.numRowHeaderColumns + numDataCols - 1;
            if (endCol < result.numRowHeaderColumns) endCol = result.numRowHeaderColumns;
            chartCfg.dataRange = CellRange(headerRow, 0, endRow, endCol).toString();

            // Pre-populate category labels from pivot row labels
            for (const auto& rowLabel : result.rowLabels) {
                chartCfg.categoryLabels.append(rowLabel.empty() ? "" : rowLabel[0]);
            }

            auto* chart = createChartWidget(m_spreadsheetView->viewport());
            chart->setSpreadsheet(sheet);
            chart->setConfig(chartCfg);
            chart->loadDataFromRange(chartCfg.dataRange);

            QRect viewRect = m_spreadsheetView->viewport()->rect();
            chart->setGeometry(qMax(10, viewRect.width() / 2 - 50), 20, 420, 320);
            chart->setProperty("sheetIndex", m_activeSheetIndex);
            connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
            connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
            connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
            connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

            chart->show();
            chart->startEntryAnimation();
            m_charts.append(chart);
            addOverlay(chart);
            applyZOrder();
        }

        statusBar()->showMessage("Pivot table updated");
    }
}

// ============== Template Gallery ==============

void MainWindow::onTemplateGallery() {
    TemplateGallery gallery(this);
    if (gallery.exec() == QDialog::Accepted) {
        applyTemplate(gallery.getResult());
    }
}

void MainWindow::applyTemplate(const TemplateResult& result) {
    if (result.sheets.empty()) return;

    // Templates hide gridlines for a cleaner look
    for (auto& sheet : result.sheets) {
        sheet->setShowGridlines(false);
    }

    setSheets(result.sheets);
    setWindowTitle("Nexel Pro - " + result.sheets[0]->getSheetName());

    // Create chart widgets from template charts
    for (int i = 0; i < static_cast<int>(result.charts.size()); ++i) {
        int sheetIdx = (i < static_cast<int>(result.chartSheetIndices.size()))
                       ? result.chartSheetIndices[i] : 0;
        if (sheetIdx >= static_cast<int>(m_sheets.size())) continue;

        auto* chart = createChartWidget(m_spreadsheetView->viewport());
        chart->setSpreadsheet(m_sheets[sheetIdx]);
        chart->setLazyLoad(m_lazyLoadCharts);
        chart->setConfig(result.charts[i]);

        if (!result.charts[i].dataRange.isEmpty()) {
            chart->loadDataFromRange(result.charts[i].dataRange);
        }

        int x = 450 + (i % 2) * 20;
        int y = 20 + (i / 2) * 340;
        chart->setGeometry(x, y, 420, 320);

        connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
        connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
        connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
        connect(chart, &ChartWidget::chartSelected, this, &MainWindow::onChartSelected);
        connect(chart, &ChartWidget::chartMoved, this, [this]() { setDirty(); });
        connect(chart, &ChartWidget::chartResized, this, [this]() { setDirty(); });

        chart->setProperty("sheetIndex", sheetIdx);
        chart->setVisible(sheetIdx == m_activeSheetIndex);
        if (sheetIdx == m_activeSheetIndex) {
            chart->show();
            chart->startEntryAnimation();
        }
        m_charts.append(chart);
        addOverlay(chart);
    }
    applyZOrder();

    statusBar()->showMessage("Template applied: " + result.sheets[0]->getSheetName());
}

ChartWidget* MainWindow::createChartWidget(QWidget* parent) {
#ifdef HAS_DATA2APP
    if (m_chartBackend == ChartBackend::Data2App)
        return new NativeChartWidget(parent);
#endif
    return new ChartWidget(parent);
}

void MainWindow::loadVisibleLazyCharts() {
    QRect viewport = m_spreadsheetView->viewport()->rect();
    for (auto* chart : m_charts) {
        if (!chart->hasLazyPending()) continue;
        if (chart->property("sheetIndex").toInt() != m_activeSheetIndex) continue;
        // Check if chart geometry intersects the visible viewport
        if (viewport.intersects(chart->geometry())) {
            chart->loadPendingData();
        }
    }
}
