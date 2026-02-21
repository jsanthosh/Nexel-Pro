#ifndef TEMPLATEGALLERY_H
#define TEMPLATEGALLERY_H

#include <QDialog>
#include <QIcon>
#include <QColor>
#include <QString>
#include <QStringList>
#include <vector>
#include <memory>
#include "ChartWidget.h"

class QListWidget;
class QListWidgetItem;
class QLabel;
class Spreadsheet;

enum class TemplateCategory {
    All,
    Finance,
    Business,
    Personal,
    Education,
    ProjectManagement
};

struct TemplateInfo {
    QString id;
    QString name;
    QString description;
    TemplateCategory category;
    QColor accentColor;
};

struct TemplateResult {
    std::vector<std::shared_ptr<Spreadsheet>> sheets;
    std::vector<ChartConfig> charts;
    std::vector<int> chartSheetIndices;
};

class TemplateGallery : public QDialog {
    Q_OBJECT

public:
    explicit TemplateGallery(QWidget* parent = nullptr);

    TemplateResult getResult() const { return m_result; }

private slots:
    void onCategoryChanged(int row);
    void onTemplateSelected(QListWidgetItem* item);
    void onTemplateDoubleClicked(QListWidgetItem* item);

private:
    void createLayout();
    void populateTemplates();
    void filterByCategory(TemplateCategory cat);
    QIcon generateThumbnail(const QString& templateId, const QColor& accent);

    // Template builders
    TemplateResult buildBudgetTracker();
    TemplateResult buildInvoice();
    TemplateResult buildExpenseReport();
    TemplateResult buildFinancialDashboard();
    TemplateResult buildSalesReport();
    TemplateResult buildProjectTimeline();
    TemplateResult buildEmployeeDirectory();
    TemplateResult buildMeetingAgenda();
    TemplateResult buildWorkoutLog();
    TemplateResult buildMealPlanner();
    TemplateResult buildTravelItinerary();
    TemplateResult buildHabitTracker();
    TemplateResult buildGradeTracker();
    TemplateResult buildClassSchedule();
    TemplateResult buildStudentRoster();
    TemplateResult buildProjectTaskBoard();
    TemplateResult buildGanttChart();
    TemplateResult buildFamilyBudget();
    TemplateResult buildWeddingPlanner();
    TemplateResult buildHomeInventory();
    TemplateResult buildClientTracker();
    TemplateResult buildEventPlanner();
    TemplateResult buildInventoryTracker();
    TemplateResult buildComparisonChart();
    TemplateResult buildKPIDashboard();

    // Helpers
    void applyHeaderStyle(Spreadsheet* sheet, int row, int startCol, int endCol,
                          const QString& bgColor, const QString& fgColor,
                          int fontSize = 12, bool bold = true);
    void applyBorders(Spreadsheet* sheet, int startRow, int startCol, int endRow, int endCol,
                      const QString& color = "#D0D5DD");
    void applyCurrencyFormat(Spreadsheet* sheet, int startRow, int startCol, int endRow, int endCol);
    void applyPercentFormat(Spreadsheet* sheet, int startRow, int startCol, int endRow, int endCol);
    void setColumnWidths(Spreadsheet* sheet, const std::vector<std::pair<int,int>>& colWidths);
    void setRowHeights(Spreadsheet* sheet, const std::vector<std::pair<int,int>>& rowHeights);
    void setCellStyleRange(Spreadsheet* sheet, int startRow, int startCol, int endRow, int endCol,
                           const QString& bgColor);
    void applyBandedRows(Spreadsheet* sheet, int startRow, int endRow, int startCol, int endCol,
                         const QString& evenColor = "#F8FAFC", const QString& oddColor = "#FFFFFF");
    void applyTitleRow(Spreadsheet* sheet, int row, int startCol, int endCol,
                       const QString& bgColor, const QString& fgColor, int fontSize, int rowHeight = 40);
    void applySectionTitle(Spreadsheet* sheet, int row, int startCol, int endCol,
                           const QString& text, const QString& color, int fontSize = 18);

    TemplateResult buildTemplate(const QString& id);

    QListWidget* m_categoryList = nullptr;
    QListWidget* m_templateGrid = nullptr;
    QLabel* m_descriptionLabel = nullptr;
    QLabel* m_previewLabel = nullptr;

    std::vector<TemplateInfo> m_allTemplates;
    TemplateResult m_result;
};

#endif // TEMPLATEGALLERY_H
