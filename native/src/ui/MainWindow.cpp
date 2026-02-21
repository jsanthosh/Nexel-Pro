#include "MainWindow.h"
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
#include "ShapeWidget.h"
#include "ChartDialog.h"
#include "ShapePropertiesDialog.h"
#include "ChartPropertiesPanel.h"
#include "../core/Spreadsheet.h"
#include "../core/UndoManager.h"
#include "../core/CellRange.h"
#include "../services/DocumentService.h"
#include "../services/CsvService.h"
#include "../services/XlsxService.h"
#include "../core/PivotEngine.h"
#include "PivotTableDialog.h"
#include "TemplateGallery.h"
#include "ImageWidget.h"
#include "SparklineDialog.h"
#include "../core/MacroEngine.h"
#include "MacroEditorDialog.h"
#include "../core/SparklineConfig.h"
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
#include <QTabBar>
#include <QToolButton>
#include <QScrollBar>
#include <QDockWidget>
#include <QJsonObject>
#include <QJsonArray>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent) {
    setWindowTitle("Nexel");
    setGeometry(100, 100, 1280, 800);

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
    m_chatDock->setStyleSheet(
        "QDockWidget { border: none; }"
        "QDockWidget::title { background: #1B5E3B; color: white; padding: 6px; font-weight: bold; text-align: center; }"
        "QDockWidget::close-button { background: transparent; }"
    );
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

    createMenuBar();
    createStatusBar();
    connectSignals();

    // Deselect charts/shapes when clicking on the spreadsheet
    m_spreadsheetView->viewport()->installEventFilter(this);

    setAcceptDrops(true);

    setStyleSheet(
        "QMainWindow { background-color: #F0F2F5; }"
        "QMenuBar { background-color: #1B5E3B; color: white; border: none; padding: 2px; font-size: 12px; }"
        "QMenuBar::item { padding: 4px 12px; border-radius: 3px; }"
        "QMenuBar::item:selected { background-color: #155030; }"
        "QMenu { background-color: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 30px 6px 20px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 4px 8px; }"
    );
}

void MainWindow::createSheetTabBar() {
    m_bottomBar = new QWidget(this);
    m_bottomBar->setFixedHeight(28);
    m_bottomBar->setStyleSheet(
        "QWidget { background-color: #F3F3F3; border-top: 1px solid #D0D0D0; }");

    QHBoxLayout* bottomLayout = new QHBoxLayout(m_bottomBar);
    bottomLayout->setContentsMargins(4, 0, 0, 0);
    bottomLayout->setSpacing(2);

    // Add sheet button
    m_addSheetBtn = new QToolButton(m_bottomBar);
    m_addSheetBtn->setText("+");
    m_addSheetBtn->setFixedSize(24, 22);
    m_addSheetBtn->setToolTip("Add New Sheet");
    m_addSheetBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; "
        "border-radius: 3px; font-size: 16px; font-weight: bold; color: #555; }"
        "QToolButton:hover { background-color: #E0E0E0; border-color: #C0C0C0; }");
    bottomLayout->addWidget(m_addSheetBtn);
    connect(m_addSheetBtn, &QToolButton::clicked, this, &MainWindow::onAddSheet);

    // Tab bar
    m_sheetTabBar = new QTabBar(m_bottomBar);
    m_sheetTabBar->setExpanding(false);
    m_sheetTabBar->setMovable(true);
    m_sheetTabBar->setTabsClosable(false);
    m_sheetTabBar->setDocumentMode(true);
    m_sheetTabBar->setStyleSheet(
        "QTabBar { background: transparent; border: none; }"
        "QTabBar::tab {"
        "   background-color: #E8E8E8;"
        "   border: 1px solid #C8C8C8;"
        "   border-bottom: none;"
        "   padding: 3px 16px;"
        "   margin-right: 2px;"
        "   font-size: 11px;"
        "   min-width: 60px;"
        "   border-top-left-radius: 3px;"
        "   border-top-right-radius: 3px;"
        "}"
        "QTabBar::tab:selected {"
        "   background-color: white;"
        "   border-bottom: 2px solid #217346;"
        "   font-weight: bold;"
        "}"
        "QTabBar::tab:hover:!selected {"
        "   background-color: #D8D8D8;"
        "}"
    );

    // Populate tabs from m_sheets
    for (const auto& sheet : m_sheets) {
        m_sheetTabBar->addTab(sheet->getSheetName());
    }

    bottomLayout->addWidget(m_sheetTabBar);
    bottomLayout->addStretch();

    // Connect signals
    connect(m_sheetTabBar, &QTabBar::currentChanged, this, &MainWindow::onSheetTabChanged);
    connect(m_sheetTabBar, &QTabBar::tabBarDoubleClicked, this, &MainWindow::onSheetTabDoubleClicked);

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
    m_spreadsheetView->setSpreadsheet(m_sheets[index]);
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

    // Update macro engine's spreadsheet reference
    if (m_macroEngine) {
        m_macroEngine->setSpreadsheet(m_sheets[index]);
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
    }
}

void MainWindow::onAddSheet() {
    int num = nextSheetNumber();
    QString name = QString("Sheet%1").arg(num);
    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setSheetName(name);
    m_sheets.push_back(sheet);
    m_sheetTabBar->addTab(name);
    m_sheetTabBar->setCurrentIndex(m_sheetTabBar->count() - 1);
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
            m_charts[i]->hide();
            m_charts[i]->deleteLater();
            m_charts.removeAt(i);
        }
    }
    for (int i = m_shapes.size() - 1; i >= 0; --i) {
        if (m_shapes[i]->property("sheetIndex").toInt() == idx) {
            m_shapes[i]->hide();
            m_shapes[i]->deleteLater();
            m_shapes.removeAt(i);
        }
    }
    for (int i = m_images.size() - 1; i >= 0; --i) {
        if (m_images[i]->property("sheetIndex").toInt() == idx) {
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
    m_sheetTabBar->insertTab(idx + 1, copy->getSheetName());
    m_sheetTabBar->setCurrentIndex(idx + 1);
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

    m_sheets = sheets;
    m_activeSheetIndex = 0;

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
            setWindowTitle("Nexel - " + newName);
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
    insertMenu->addAction("&Row Above", m_spreadsheetView, &SpreadsheetView::insertEntireRow);
    insertMenu->addAction("&Column Left", m_spreadsheetView, &SpreadsheetView::insertEntireColumn);

    QMenu* dataMenu = menuBar->addMenu("&Data");
    dataMenu->addAction("Sort &Ascending", m_spreadsheetView, &SpreadsheetView::sortAscending);
    dataMenu->addAction("Sort &Descending", m_spreadsheetView, &SpreadsheetView::sortDescending);
    dataMenu->addSeparator();
    dataMenu->addAction("&Data Validation...", this, &MainWindow::onDataValidation);
    dataMenu->addSeparator();
    dataMenu->addAction("Create &Pivot Table...", this, &MainWindow::onCreatePivotTable);
    dataMenu->addAction("&Refresh Pivot Table", this, &MainWindow::onRefreshPivotTable,
                        QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_R));
    dataMenu->addSeparator();
    QAction* highlightAction = dataMenu->addAction("&Circle Invalid Data", this, &MainWindow::onHighlightInvalidCells);
    highlightAction->setCheckable(true);

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

    // ===== Tools Menu =====
    QMenu* toolsMenu = menuBar->addMenu("&Tools");
    toolsMenu->addAction("Macro &Editor...", this, &MainWindow::onMacroEditor,
                         QKeySequence(Qt::CTRL | Qt::ALT | Qt::Key_M));
    toolsMenu->addAction("Run &Last Macro", this, &MainWindow::onRunLastMacro,
                         QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_M));
}

void MainWindow::createToolBar() {}

void MainWindow::createStatusBar() {
    statusBar()->showMessage("Ready");
    statusBar()->setStyleSheet(
        "QStatusBar { background-color: #217346; color: white; border: none; font-size: 11px; padding: 2px 8px; }"
    );
}

void MainWindow::connectSignals() {
    connect(m_toolbar, &Toolbar::newDocument, this, &MainWindow::onNewDocument);
    connect(m_toolbar, &Toolbar::saveDocument, this, &MainWindow::onSaveDocument);
    connect(m_toolbar, &Toolbar::undo, this, &MainWindow::onUndo);
    connect(m_toolbar, &Toolbar::redo, this, &MainWindow::onRedo);

    connect(m_toolbar, &Toolbar::bold, m_spreadsheetView, &SpreadsheetView::applyBold);
    connect(m_toolbar, &Toolbar::italic, m_spreadsheetView, &SpreadsheetView::applyItalic);
    connect(m_toolbar, &Toolbar::underline, m_spreadsheetView, &SpreadsheetView::applyUnderline);
    connect(m_toolbar, &Toolbar::strikethrough, m_spreadsheetView, &SpreadsheetView::applyStrikethrough);
    connect(m_toolbar, &Toolbar::fontFamilyChanged, m_spreadsheetView, &SpreadsheetView::applyFontFamily);
    connect(m_toolbar, &Toolbar::fontSizeChanged, m_spreadsheetView, &SpreadsheetView::applyFontSize);

    connect(m_toolbar, &Toolbar::foregroundColorChanged, m_spreadsheetView, &SpreadsheetView::applyForegroundColor);
    connect(m_toolbar, &Toolbar::backgroundColorChanged, m_spreadsheetView, &SpreadsheetView::applyBackgroundColor);

    connect(m_toolbar, &Toolbar::hAlignChanged, m_spreadsheetView, &SpreadsheetView::applyHAlign);
    connect(m_toolbar, &Toolbar::vAlignChanged, m_spreadsheetView, &SpreadsheetView::applyVAlign);

    connect(m_toolbar, &Toolbar::thousandSeparatorToggled, m_spreadsheetView, &SpreadsheetView::applyThousandSeparator);
    connect(m_toolbar, &Toolbar::numberFormatChanged, m_spreadsheetView, &SpreadsheetView::applyNumberFormat);
    connect(m_toolbar, &Toolbar::formatCellsRequested, this, &MainWindow::onFormatCells);

    connect(m_toolbar, &Toolbar::formatPainterToggled, m_spreadsheetView, &SpreadsheetView::activateFormatPainter);

    connect(m_toolbar, &Toolbar::sortAscending, m_spreadsheetView, &SpreadsheetView::sortAscending);
    connect(m_toolbar, &Toolbar::sortDescending, m_spreadsheetView, &SpreadsheetView::sortDescending);
    connect(m_toolbar, &Toolbar::filterToggled, m_spreadsheetView, &SpreadsheetView::toggleAutoFilter);

    connect(m_toolbar, &Toolbar::tableStyleSelected, m_spreadsheetView, &SpreadsheetView::applyTableStyle);

    connect(m_toolbar, &Toolbar::borderStyleSelected, m_spreadsheetView, &SpreadsheetView::applyBorderStyle);
    connect(m_toolbar, &Toolbar::mergeCellsRequested, m_spreadsheetView, &SpreadsheetView::mergeCells);
    connect(m_toolbar, &Toolbar::unmergeCellsRequested, m_spreadsheetView, &SpreadsheetView::unmergeCells);
    connect(m_toolbar, &Toolbar::increaseIndent, m_spreadsheetView, &SpreadsheetView::increaseIndent);
    connect(m_toolbar, &Toolbar::decreaseIndent, m_spreadsheetView, &SpreadsheetView::decreaseIndent);

    connect(m_toolbar, &Toolbar::conditionalFormatRequested, this, &MainWindow::onConditionalFormat);
    connect(m_toolbar, &Toolbar::dataValidationRequested, this, &MainWindow::onDataValidation);

    // Chart and shape insertion from toolbar
    connect(m_toolbar, &Toolbar::insertChartRequested, this, &MainWindow::onInsertChart);
    connect(m_toolbar, &Toolbar::insertShapeRequested, this, &MainWindow::onInsertShape);

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

    connect(m_spreadsheetView, &SpreadsheetView::cellSelected,
            this, [this](int, int, const QString& content, const QString& address) {
        // Don't update formula bar if we're in formula editing mode (would overwrite the formula)
        if (m_spreadsheetView->isFormulaEditMode() || m_formulaBar->isFormulaEditing()) {
            return;
        }
        m_formulaBar->setCellAddress(address);
        m_formulaBar->setCellContent(content);

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
            this, &MainWindow::refreshActiveCharts);
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

    m_currentFilePath = fileName;
    QString ext = QFileInfo(fileName).suffix().toLower();

    if (ext == "xlsx" || ext == "xls") {
        auto result = XlsxService::importFromFile(fileName);
        if (!result.sheets.empty()) {
            setSheets(result.sheets);
            setWindowTitle("Nexel - " + QFileInfo(fileName).fileName());

            // Create chart widgets from imported charts
            static const QVector<QColor> excelColors = {
                QColor("#4472C4"), QColor("#ED7D31"), QColor("#A5A5A5"),
                QColor("#FFC000"), QColor("#5B9BD5"), QColor("#70AD47"),
                QColor("#264478"), QColor("#9E480E"), QColor("#636363")
            };

            for (const auto& imported : result.charts) {
                ChartConfig config;

                // Map chart type string to enum
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

                // Convert imported series to ChartSeries
                for (int i = 0; i < imported.series.size(); ++i) {
                    ChartSeries s;
                    s.name = imported.series[i].name;
                    s.yValues = imported.series[i].values;

                    // Use numeric x values if available (scatter), otherwise indices
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

                int si = imported.sheetIndex;
                if (si < 0 || si >= static_cast<int>(m_sheets.size())) continue;

                auto* chart = new ChartWidget(m_spreadsheetView->viewport());
                chart->setSpreadsheet(m_sheets[si]);
                chart->setConfig(config);
                chart->setGeometry(imported.x, imported.y, imported.width, imported.height);

                connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
                connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
                connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
                connect(chart, &ChartWidget::chartSelected, this, [this](ChartWidget* c) {
                    int idx = c->property("sheetIndex").toInt();
                    for (auto* other : m_charts) if (other != c && other->property("sheetIndex").toInt() == idx) other->setSelected(false);
                    for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == idx) s->setSelected(false);
                });

                chart->setProperty("sheetIndex", si);
                chart->setVisible(si == m_activeSheetIndex);
                if (si == m_activeSheetIndex) {
                    chart->show();
                    chart->raise();
                    chart->startEntryAnimation();
                }
                m_charts.append(chart);
            }

            int chartCount = static_cast<int>(result.charts.size());
            if (chartCount > 0) {
                statusBar()->showMessage(QString("Opened: %1 (%2 chart(s) imported)").arg(fileName).arg(chartCount));
            } else {
                statusBar()->showMessage("Opened: " + fileName);
            }
        } else {
            QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
        }
    } else {
        auto spreadsheet = CsvService::importFromFile(fileName);
        if (spreadsheet) {
            spreadsheet->setSheetName(QFileInfo(fileName).baseName());
            std::vector<std::shared_ptr<Spreadsheet>> sheets = { spreadsheet };
            setSheets(sheets);
            setWindowTitle("Nexel - " + QFileInfo(fileName).fileName());
            statusBar()->showMessage("Opened: " + fileName);
        } else {
            QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
        }
    }
}

void MainWindow::onNewDocument() {
    if (DocumentService::instance().getCurrentDocument()) {
        if (QMessageBox::question(this, "New Document", "Save current document before creating new one?")
            == QMessageBox::Yes) {
            onSaveDocument();
        }
    }

    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setSheetName("Sheet1");
    std::vector<std::shared_ptr<Spreadsheet>> sheets = { sheet };
    setSheets(sheets);

    DocumentService::instance().createNewDocument("Untitled");
    setWindowTitle("Nexel");
    statusBar()->showMessage("New document created");
}

void MainWindow::onOpenDocument() {
    QString fileName = QFileDialog::getOpenFileName(this, "Open Document", "",
        "All Spreadsheet Files (*.xlsx *.csv *.txt);;Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    openFile(fileName);
}

void MainWindow::onSaveDocument() {
    if (m_currentFilePath.isEmpty()) {
        onSaveAs();
        return;
    }

    QString ext = QFileInfo(m_currentFilePath).suffix().toLower();
    bool success = false;

    if (ext == "xlsx" || ext == "xls") {
        success = XlsxService::exportToFile(m_sheets, m_currentFilePath);
    } else {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) success = CsvService::exportToFile(*spreadsheet, m_currentFilePath);
    }

    if (success) {
        statusBar()->showMessage("Saved: " + m_currentFilePath);
    } else {
        QMessageBox::warning(this, "Save Failed", "Could not save file.");
    }
}

void MainWindow::onSaveAs() {
    QString fileName = QFileDialog::getSaveFileName(this, "Save Document As", "",
        "Excel Workbook (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    if (fileName.isEmpty()) return;

    QString ext = QFileInfo(fileName).suffix().toLower();
    bool success = false;

    if (ext == "xlsx") {
        success = XlsxService::exportToFile(m_sheets, fileName);
    } else {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) success = CsvService::exportToFile(*spreadsheet, fileName);
    }

    if (success) {
        m_currentFilePath = fileName;
        setWindowTitle("Nexel - " + QFileInfo(fileName).fileName());
        statusBar()->showMessage("Saved: " + fileName);
    } else {
        QMessageBox::warning(this, "Save Failed", "Could not save file.");
    }
}

void MainWindow::onUndo() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canUndo()) {
        sheet->getUndoManager().undo(sheet.get());
        CellAddress target = sheet->getUndoManager().lastUndoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            model->resetModel();
            QModelIndex idx = model->index(target.row, target.col);
            m_spreadsheetView->setCurrentIndex(idx);
            m_spreadsheetView->scrollTo(idx);
        }
        refreshActiveCharts();
        statusBar()->showMessage("Undo");
    }
}

void MainWindow::onRedo() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canRedo()) {
        sheet->getUndoManager().redo(sheet.get());
        CellAddress target = sheet->getUndoManager().lastRedoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            model->resetModel();
            QModelIndex idx = model->index(target.row, target.col);
            m_spreadsheetView->setCurrentIndex(idx);
            m_spreadsheetView->scrollTo(idx);
        }
        refreshActiveCharts();
        statusBar()->showMessage("Redo");
    }
}

void MainWindow::onCut() { if (m_spreadsheetView) m_spreadsheetView->cut(); }
void MainWindow::onCopy() { if (m_spreadsheetView) m_spreadsheetView->copy(); }
void MainWindow::onPaste() { if (m_spreadsheetView) m_spreadsheetView->paste(); }
void MainWindow::onDelete() { if (m_spreadsheetView) m_spreadsheetView->deleteSelection(); }
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
    if (saveCurrentDocument()) event->accept();
    else event->ignore();
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
    int maxCol = sheet->getMaxColumn();
    if (maxRow < 0 || maxCol < 0) {
        m_findReplaceDialog->setStatus("No data to search.");
        return;
    }

    QModelIndex current = m_spreadsheetView->currentIndex();
    int startRow = current.isValid() ? current.row() : 0;
    int startCol = current.isValid() ? current.column() + 1 : 0;

    // Search forward: row by row, column by column
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

    // Wrap around from top
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

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    // Get selection bounding box
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    ConditionalFormatDialog dialog(range, sheet->getConditionalFormatting(), this);
    if (dialog.exec() == QDialog::Accepted) {
        m_spreadsheetView->refreshView();
        statusBar()->showMessage("Conditional formatting applied");
    }
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

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.size() <= 1) {
        statusBar()->showMessage("Ready");
        return;
    }

    // Compute SUM, AVERAGE, COUNT for numeric values (like Excel status bar)
    double sum = 0;
    int numericCount = 0;
    int nonEmptyCount = 0;

    // Limit calculation to avoid lag on huge selections
    int limit = qMin(selected.size(), static_cast<qsizetype>(50000));
    for (int i = 0; i < limit; ++i) {
        const auto& idx = selected[i];
        auto val = sheet->getCellValue(CellAddress(idx.row(), idx.column()));
        QString text = val.toString();
        if (!text.isEmpty()) {
            nonEmptyCount++;
            bool ok;
            double num = text.toDouble(&ok);
            if (ok) {
                sum += num;
                numericCount++;
            }
        }
    }

    if (numericCount > 0) {
        double avg = sum / numericCount;
        statusBar()->showMessage(
            QString("Average: %1   Count: %2   Sum: %3")
                .arg(QString::number(avg, 'f', 2))
                .arg(nonEmptyCount)
                .arg(QString::number(sum, 'f', 2)));
    } else if (nonEmptyCount > 0) {
        statusBar()->showMessage(QString("Count: %1").arg(nonEmptyCount));
    } else {
        statusBar()->showMessage("Ready");
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

// ============== Chart and Shape Insertion ==============

QString MainWindow::getSelectionRange() const {
    if (!m_spreadsheetView) return "";

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return "";

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    return CellAddress(minRow, minCol).toString() + ":" + CellAddress(maxRow, maxCol).toString();
}

void MainWindow::onInsertChart() {
    ChartDialog dialog(this);
    dialog.setSpreadsheet(m_sheets[m_activeSheetIndex]);

    // Pre-fill with current selection
    QString range = getSelectionRange();
    if (!range.isEmpty()) {
        dialog.setDataRange(range);
    }

    if (dialog.exec() == QDialog::Accepted) {
        ChartConfig config = dialog.getConfig();

        // Auto-generate titles from data headers if not specified
        ChartWidget::autoGenerateTitles(config, m_sheets[m_activeSheetIndex]);

        auto* chart = new ChartWidget(m_spreadsheetView->viewport());
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
        connect(chart, &ChartWidget::chartSelected, this, [this](ChartWidget* c) {
            int si = c->property("sheetIndex").toInt();
            for (auto* other : m_charts) {
                if (other != c && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            }
            for (auto* s : m_shapes) {
                if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
            }
        });

        chart->setProperty("sheetIndex", m_activeSheetIndex);
        chart->show();
        chart->raise();
        m_charts.append(chart);

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
        connect(shape, &ShapeWidget::shapeSelected, this, [this](ShapeWidget* s) {
            int si = s->property("sheetIndex").toInt();
            for (auto* other : m_shapes) {
                if (other != s && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            }
            for (auto* c : m_charts) {
                if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
            }
        });

        shape->setProperty("sheetIndex", m_activeSheetIndex);
        shape->show();
        shape->raise();
        m_shapes.append(shape);

        statusBar()->showMessage("Shape inserted");
    }
}

void MainWindow::onEditChart(ChartWidget* chart) {
    if (!chart) return;
    // Use the side panel for chart editing instead of a dialog
    onChartPropertiesRequested(chart);
}

void MainWindow::onDeleteChart(ChartWidget* chart) {
    if (!chart) return;

    m_charts.removeOne(chart);
    chart->hide();
    chart->deleteLater();
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

    m_shapes.removeOne(shape);
    shape->hide();
    shape->deleteLater();
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
    connect(image, &ImageWidget::imageSelected, this, [this](ImageWidget* img) {
        int si = img->property("sheetIndex").toInt();
        for (auto* other : m_images) {
            if (other != img && other->property("sheetIndex").toInt() == si) other->setSelected(false);
        }
        for (auto* c : m_charts) {
            if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
        }
        for (auto* s : m_shapes) {
            if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
        }
    });

    image->setProperty("sheetIndex", m_activeSheetIndex);
    image->show();
    image->raise();
    m_images.append(image);

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
        m_images.removeOne(img);
        img->hide();
        img->deleteLater();
    }

    int total = chartsToDelete.size() + shapesToDelete.size() + imagesToDelete.size();
    if (total > 0) {
        statusBar()->showMessage(QString("Deleted %1 object(s)").arg(total));
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

    auto* chart = new ChartWidget(m_spreadsheetView->viewport());
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
    connect(chart, &ChartWidget::chartSelected, this, [this](ChartWidget* c) {
        int si = c->property("sheetIndex").toInt();
        for (auto* other : m_charts) if (other != c && other->property("sheetIndex").toInt() == si) other->setSelected(false);
        for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
    });

    chart->setProperty("sheetIndex", m_activeSheetIndex);
    chart->show();
    chart->raise();
    m_charts.append(chart);
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
        for (auto* other : m_shapes) if (other != s && other->property("sheetIndex").toInt() == si) other->setSelected(false);
        for (auto* c : m_charts) if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
    });

    shape->setProperty("sheetIndex", m_activeSheetIndex);
    shape->show();
    shape->raise();
    m_shapes.append(shape);
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
    connect(image, &ImageWidget::imageSelected, this, [this](ImageWidget* img) {
        int si = img->property("sheetIndex").toInt();
        for (auto* other : m_images) if (other != img && other->property("sheetIndex").toInt() == si) other->setSelected(false);
        for (auto* c : m_charts) if (c->property("sheetIndex").toInt() == si) c->setSelected(false);
        for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
    });

    image->setProperty("sheetIndex", m_activeSheetIndex);
    image->show();
    image->raise();
    m_images.append(image);
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

            // Build chart data range from pivot output
            int headerRow = result.dataStartRow - 1;
            if (headerRow < 0) headerRow = 0;
            int endRow = headerRow + static_cast<int>(result.rowLabels.size());
            int endCol = result.numRowHeaderColumns + static_cast<int>(result.columnLabels.size()) - 1;
            chartCfg.dataRange = CellRange(headerRow, 0, endRow, endCol).toString();

            auto* chart = new ChartWidget(m_spreadsheetView->viewport());
            chart->setSpreadsheet(pivotSheet);
            chart->setConfig(chartCfg);
            chart->loadDataFromRange(chartCfg.dataRange);

            QRect viewRect = m_spreadsheetView->viewport()->rect();
            chart->setGeometry(qMax(10, viewRect.width() / 2 - 50), 20, 420, 320);

            connect(chart, &ChartWidget::editRequested, this, &MainWindow::onEditChart);
            connect(chart, &ChartWidget::deleteRequested, this, &MainWindow::onDeleteChart);
            connect(chart, &ChartWidget::propertiesRequested, this, &MainWindow::onChartPropertiesRequested);
            connect(chart, &ChartWidget::chartSelected, this, [this](ChartWidget* c) {
                int si = c->property("sheetIndex").toInt();
                for (auto* other : m_charts) if (other != c && other->property("sheetIndex").toInt() == si) other->setSelected(false);
                for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
            });

            chart->setProperty("sheetIndex", pivotSheetIdx);
            chart->show();
            chart->raise();
            chart->startEntryAnimation();
            m_charts.append(chart);
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
    setWindowTitle("Nexel - " + result.sheets[0]->getSheetName());

    // Create chart widgets from template charts
    for (int i = 0; i < static_cast<int>(result.charts.size()); ++i) {
        int sheetIdx = (i < static_cast<int>(result.chartSheetIndices.size()))
                       ? result.chartSheetIndices[i] : 0;
        if (sheetIdx >= static_cast<int>(m_sheets.size())) continue;

        auto* chart = new ChartWidget(m_spreadsheetView->viewport());
        chart->setSpreadsheet(m_sheets[sheetIdx]);
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
        connect(chart, &ChartWidget::chartSelected, this, [this](ChartWidget* c) {
            int si = c->property("sheetIndex").toInt();
            for (auto* other : m_charts) if (other != c && other->property("sheetIndex").toInt() == si) other->setSelected(false);
            for (auto* s : m_shapes) if (s->property("sheetIndex").toInt() == si) s->setSelected(false);
        });

        chart->setProperty("sheetIndex", sheetIdx);
        chart->setVisible(sheetIdx == m_activeSheetIndex);
        if (sheetIdx == m_activeSheetIndex) {
            chart->show();
            chart->raise();
            chart->startEntryAnimation();
        }
        m_charts.append(chart);
    }

    statusBar()->showMessage("Template applied: " + result.sheets[0]->getSheetName());
}
