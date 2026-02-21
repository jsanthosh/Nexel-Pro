#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QCloseEvent>
#include <QKeyEvent>
#include <QTabBar>
#include <QToolButton>
#include <QJsonArray>
#include <memory>
#include <vector>

class Spreadsheet;
class SpreadsheetView;
class FormulaBar;
class Toolbar;
class FormatCellsDialog;
class FindReplaceDialog;
class ChatPanel;
class ChartWidget;
class ShapeWidget;
class ImageWidget;
class ChartPropertiesPanel;
class MacroEngine;
struct TemplateResult;

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow() = default;
    void openFile(const QString& fileName);

protected:
    void closeEvent(QCloseEvent* event) override;
    void dragEnterEvent(QDragEnterEvent* event) override;
    void dropEvent(QDropEvent* event) override;
    void keyPressEvent(QKeyEvent* event) override;
    bool eventFilter(QObject* obj, QEvent* event) override;

private slots:
    void onNewDocument();
    void onOpenDocument();
    void onSaveDocument();
    void onSaveAs();
    void onImportCsv();
    void onExportCsv();
    void onUndo();
    void onRedo();
    void onCut();
    void onCopy();
    void onPaste();
    void onDelete();
    void onSelectAll();
    void onFormatCells();
    void onFindReplace();
    void onGoTo();
    void onConditionalFormat();
    void onDataValidation();
    void onFreezePane();

    // Find/Replace operations
    void onFindNext();
    void onFindPrevious();
    void onReplaceOne();
    void onReplaceAll();

    // Sheet tab management
    void onSheetTabChanged(int index);
    void onSheetTabDoubleClicked(int index);
    void onAddSheet();
    void onDeleteSheet();
    void onDuplicateSheet();
    void showSheetContextMenu(const QPoint& pos);
    void onHighlightInvalidCells();
    void onChatActions(const QJsonArray& actions);

    // Chart and Shape insertion
    void onInsertChart();
    void onInsertShape();
    void onEditChart(ChartWidget* chart);
    void onDeleteChart(ChartWidget* chart);
    void onEditShape(ShapeWidget* shape);
    void onDeleteShape(ShapeWidget* shape);
    void onChartPropertiesRequested(ChartWidget* chart);

    // Image insertion
    void onInsertImage();
    void onEditImage(ImageWidget* image);
    void onDeleteImage(ImageWidget* image);

    // Sparkline insertion
    void onInsertSparkline();

    // Macros
    void onMacroEditor();
    void onRunLastMacro();

    // Pivot table
    void onCreatePivotTable();
    void onRefreshPivotTable();

    // Templates
    void onTemplateGallery();

private:
    void createMenuBar();
    void createToolBar();
    void createStatusBar();
    void createSheetTabBar();
    void connectSignals();
    bool saveCurrentDocument();
    void setSheets(const std::vector<std::shared_ptr<Spreadsheet>>& sheets);
    void switchToSheet(int index);
    int nextSheetNumber() const;
    bool cellMatchesSearch(int row, int col, const QString& searchText, bool matchCase, bool wholeCell) const;
    void updateStatusBarSummary();

    SpreadsheetView* m_spreadsheetView;
    FormulaBar* m_formulaBar;
    Toolbar* m_toolbar;
    QWidget* m_bottomBar;
    QTabBar* m_sheetTabBar;
    QToolButton* m_addSheetBtn;
    FindReplaceDialog* m_findReplaceDialog = nullptr;
    ChatPanel* m_chatPanel = nullptr;
    QDockWidget* m_chatDock = nullptr;
    ChartPropertiesPanel* m_chartPropsPanel = nullptr;
    QDockWidget* m_chartPropsDock = nullptr;
    QString m_currentFilePath;  // Track the file path for Ctrl+S

    // Multi-sheet storage
    std::vector<std::shared_ptr<Spreadsheet>> m_sheets;
    int m_activeSheetIndex = 0;
    bool m_frozenPanes = false;
    QAction* m_gridlinesAction = nullptr;

    // Charts, shapes, and images (flat lists; each widget has "sheetIndex" property)
    QVector<ChartWidget*> m_charts;
    QVector<ShapeWidget*> m_shapes;
    QVector<ImageWidget*> m_images;
    QMetaObject::Connection m_dataChangedConnection;
    QMetaObject::Connection m_modelResetConnection;
    QString getSelectionRange() const;
    void deleteSelectedOverlays();
    void deselectAllOverlays();
    void insertChartFromChat(const QJsonObject& params);
    void insertShapeFromChat(const QJsonObject& params);
    void insertImageFromChat(const QJsonObject& params);
    void reconnectDataChanged();
    void refreshActiveCharts();
    void applyTemplate(const TemplateResult& result);

    // Macro engine
    MacroEngine* m_macroEngine = nullptr;
};

#endif // MAINWINDOW_H
