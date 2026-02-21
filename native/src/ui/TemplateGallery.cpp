#include "TemplateGallery.h"
#include "../core/Spreadsheet.h"
#include "../core/Cell.h"
#include "../core/CellRange.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QListWidget>
#include <QLabel>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QPainter>
#include <QPainterPath>

// ============== Thumbnail Generation ==============

QIcon TemplateGallery::generateThumbnail(const QString& templateId, const QColor& accent) {
    qreal dpr = 2.0;
    QPixmap pix(120 * dpr, 90 * dpr);
    pix.setDevicePixelRatio(dpr);
    pix.fill(Qt::white);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing);

    // Grid background
    p.setPen(QPen(QColor("#E8ECF0"), 0.5));
    for (int x = 0; x <= 120; x += 20) p.drawLine(x, 14, x, 90);
    for (int y = 14; y <= 90; y += 10) p.drawLine(0, y, 120, y);

    // Colored header bar
    p.fillRect(0, 0, 120, 14, accent);

    // Template-specific decoration
    if (templateId.contains("budget") || templateId.contains("expense")) {
        // Mini pie chart
        p.setPen(QPen(Qt::white, 1));
        p.setBrush(accent); p.drawPie(70, 25, 40, 40, 0, 200 * 16);
        p.setBrush(accent.lighter(130)); p.drawPie(70, 25, 40, 40, 200 * 16, 100 * 16);
        p.setBrush(accent.lighter(170)); p.drawPie(70, 25, 40, 40, 300 * 16, 60 * 16);
        // Mini data lines
        p.setPen(QPen(QColor("#666"), 0.8));
        for (int y = 24; y < 70; y += 10) p.drawLine(6, y, 60, y);
    } else if (templateId.contains("invoice")) {
        // Document with lines
        p.setPen(QPen(accent, 1.5));
        p.drawLine(6, 20, 30, 20);
        p.setPen(QPen(QColor("#AAA"), 0.6));
        for (int y = 32; y < 78; y += 8) p.drawLine(6, y, 114, y);
        p.fillRect(6, 72, 108, 8, accent.lighter(180));
    } else if (templateId.contains("dashboard") || templateId.contains("sales")) {
        // Mini bar chart
        int bars[] = {30, 45, 25, 50, 35};
        for (int i = 0; i < 5; ++i) {
            QColor c = (i % 2 == 0) ? accent : accent.lighter(140);
            p.fillRect(10 + i * 20, 75 - bars[i], 14, bars[i], c);
        }
    } else if (templateId.contains("timeline") || templateId.contains("gantt")) {
        // Horizontal bars
        int widths[] = {60, 45, 80, 35, 55};
        for (int i = 0; i < 5; ++i) {
            QColor c = accent;
            c.setAlpha(180 - i * 25);
            p.fillRect(30, 20 + i * 12, widths[i], 8, c);
            p.setPen(QPen(QColor("#666"), 0.6));
            p.drawLine(6, 20 + i * 12 + 4, 28, 20 + i * 12 + 4);
        }
    } else if (templateId.contains("directory") || templateId.contains("roster")) {
        // Table rows
        p.fillRect(4, 18, 112, 10, accent.lighter(170));
        for (int y = 28; y < 80; y += 10) {
            p.fillRect(4, y, 112, 10, (y / 10) % 2 == 0 ? QColor("#F8F9FA") : Qt::white);
        }
        p.setPen(QPen(QColor("#DDD"), 0.5));
        for (int y = 18; y < 80; y += 10) p.drawLine(4, y, 116, y);
    } else if (templateId.contains("habit")) {
        // Checkmark grid
        p.setPen(QPen(QColor("#DDD"), 0.5));
        for (int x = 30; x < 115; x += 7) for (int y = 22; y < 82; y += 8) p.drawRect(x, y, 6, 7);
        p.setPen(QPen(accent, 1.5));
        for (int x = 30; x < 110; x += 14) for (int y = 22; y < 75; y += 16)
            p.drawText(QRect(x, y, 6, 7), Qt::AlignCenter, QString::fromUtf8("\xe2\x9c\x93"));
    } else if (templateId.contains("grade")) {
        // Column chart for grades
        int h[] = {20, 40, 55, 30, 10};
        for (int i = 0; i < 5; ++i) {
            p.fillRect(15 + i * 20, 80 - h[i], 14, h[i], accent.lighter(100 + i * 20));
        }
    } else if (templateId.contains("schedule")) {
        // Time grid with color blocks
        QColor colors[] = {accent, accent.lighter(130), accent.lighter(160), QColor("#ED7D31")};
        for (int i = 0; i < 4; ++i) {
            int x = 25 + (i % 4) * 22, y = 22 + (i / 2) * 25;
            p.fillRect(x, y, 20, 20, colors[i]);
        }
        p.setPen(QPen(QColor("#AAA"), 0.5));
        for (int y = 18; y < 82; y += 12) p.drawLine(4, y, 116, y);
    } else if (templateId.contains("workout") || templateId.contains("meal")) {
        // Bars
        p.fillRect(8, 24, 104, 10, accent.lighter(170));
        for (int i = 0; i < 5; ++i) {
            int w = 20 + (i * 17) % 60;
            p.fillRect(8, 38 + i * 10, w, 7, accent.lighter(130 + i * 10));
        }
    } else if (templateId.contains("task")) {
        // Kanban-ish columns
        p.fillRect(4, 18, 35, 66, QColor("#E8F5E9"));
        p.fillRect(42, 18, 35, 66, QColor("#FFF3E0"));
        p.fillRect(80, 18, 35, 66, QColor("#E3F2FD"));
        for (int i = 0; i < 3; ++i) {
            p.fillRect(8 + i * 38, 26, 27, 8, accent.lighter(140));
            p.fillRect(8 + i * 38, 38, 27, 8, accent.lighter(160));
        }
    } else {
        // Generic: lines + accent block
        p.fillRect(6, 22, 50, 8, accent.lighter(170));
        p.setPen(QPen(QColor("#CCC"), 0.6));
        for (int y = 36; y < 80; y += 8) p.drawLine(6, y, 114, y);
    }

    // Border
    p.setPen(QPen(QColor("#D0D5DD"), 1));
    p.setBrush(Qt::NoBrush);
    p.drawRoundedRect(0, 0, 119, 89, 4, 4);

    p.end();
    return QIcon(pix);
}

// ============== Dialog Construction ==============

TemplateGallery::TemplateGallery(QWidget* parent) : QDialog(parent) {
    setWindowTitle("Template Gallery");
    setMinimumSize(780, 520);
    resize(840, 560);

    populateTemplates();
    createLayout();

    setStyleSheet(
        "QDialog { background: #FAFBFC; }"
        "QListWidget { border: 1px solid #D0D5DD; border-radius: 6px; background: white; outline: none; }"
        "QListWidget::item { padding: 6px 8px; border-radius: 4px; }"
        "QListWidget::item:selected { background-color: #E8F0FE; color: #1A1A1A; }"
        "QListWidget::item:hover:!selected { background-color: #F5F5F5; }"
    );
}

void TemplateGallery::populateTemplates() {
    m_allTemplates = {
        {"finance_budget",     "Budget Tracker",       "Track monthly income and expenses with formulas and a pie chart.", TemplateCategory::Finance, QColor("#1B5E3B")},
        {"finance_invoice",    "Invoice",              "Professional invoice template with line items and tax calculation.", TemplateCategory::Finance, QColor("#2B5797")},
        {"finance_expense",    "Expense Report",       "Quarterly expense report with categorized entries and a column chart.", TemplateCategory::Finance, QColor("#4F46E5")},
        {"finance_dashboard",  "Financial Dashboard",  "Multi-sheet financial dashboard with KPIs, trends, and charts.", TemplateCategory::Finance, QColor("#0E7C6B")},
        {"biz_sales",          "Sales Report",         "Multi-region sales data with subtotals, bar chart, and pie chart.", TemplateCategory::Business, QColor("#4472C4")},
        {"biz_timeline",       "Project Timeline",     "Project phases with tasks, owners, dates, and status tracking.", TemplateCategory::Business, QColor("#ED7D31")},
        {"biz_directory",      "Employee Directory",   "Company-wide employee directory with departments and contact info.", TemplateCategory::Business, QColor("#5B6B7D")},
        {"biz_agenda",         "Meeting Agenda",       "Structured meeting agenda with time slots and presenters.", TemplateCategory::Business, QColor("#217346")},
        {"personal_workout",   "Workout Log",          "Weekly workout tracker with exercises, sets, reps, and calorie chart.", TemplateCategory::Personal, QColor("#D94166")},
        {"personal_meal",      "Meal Planner",         "Weekly meal planner with nutrition tracking and calorie breakdown.", TemplateCategory::Personal, QColor("#2D8C4E")},
        {"personal_travel",    "Travel Itinerary",     "Day-by-day trip planner with activities, costs, and booking info.", TemplateCategory::Personal, QColor("#E67E22")},
        {"personal_habit",     "Habit Tracker",        "Monthly habit tracker with daily checkmarks and completion rates.", TemplateCategory::Personal, QColor("#0EA5E9")},
        {"edu_grades",         "Grade Tracker",        "Student grade tracker with formulas for averages and letter grades.", TemplateCategory::Education, QColor("#4F46E5")},
        {"edu_schedule",       "Class Schedule",       "Weekly class schedule grid with color-coded time blocks.", TemplateCategory::Education, QColor("#7C3AED")},
        {"edu_roster",         "Student Roster",       "Class roster with student IDs, contact info, and GPA.", TemplateCategory::Education, QColor("#9333EA")},
        {"pm_taskboard",       "Project Task Board",   "Sprint task board with priorities, status, and story points.", TemplateCategory::ProjectManagement, QColor("#D97706")},
        {"pm_gantt",           "Gantt Chart",          "Visual Gantt chart with task bars spanning weeks.", TemplateCategory::ProjectManagement, QColor("#059669")},
        {"finance_family",     "Family Budget",        "Family budget with cash flow, income, expenses — projected vs actual with variance.", TemplateCategory::Finance, QColor("#42A5A1")},
        {"personal_wedding",   "Wedding Planner",      "Complete wedding planning checklist with vendors, budget, and timeline.", TemplateCategory::Personal, QColor("#D4508B")},
        {"personal_home",      "Home Inventory",       "Room-by-room home inventory with item values and insurance info.", TemplateCategory::Personal, QColor("#6366F1")},
        {"biz_clients",        "Client Tracker",       "CRM-style client tracking with deal pipeline, status, and revenue.", TemplateCategory::Business, QColor("#0891B2")},
        {"biz_event",          "Event Planner",        "Event planning tracker with tasks, vendors, budget, and deadlines.", TemplateCategory::Business, QColor("#9333EA")},
        {"biz_inventory",      "Inventory Tracker",    "Product inventory with stock levels, reorder points, and valuation.", TemplateCategory::Business, QColor("#EA580C")},
        {"pm_comparison",      "Comparison Matrix",    "Side-by-side comparison of options with scoring and weighted ranking.", TemplateCategory::ProjectManagement, QColor("#2563EB")},
        {"finance_kpi",        "KPI Dashboard",        "Executive KPI dashboard with targets, actuals, and performance indicators.", TemplateCategory::Finance, QColor("#DC2626")},
    };
}

void TemplateGallery::createLayout() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(10);

    QLabel* titleLabel = new QLabel("Choose a Template");
    titleLabel->setStyleSheet("font-size: 16px; font-weight: bold; color: #1B5E3B; padding: 4px 0;");
    mainLayout->addWidget(titleLabel);

    QHBoxLayout* contentLayout = new QHBoxLayout();
    contentLayout->setSpacing(10);

    // Left: Category list
    m_categoryList = new QListWidget();
    m_categoryList->setFixedWidth(150);
    m_categoryList->addItem("All Templates");
    m_categoryList->addItem("Finance");
    m_categoryList->addItem("Business");
    m_categoryList->addItem("Personal");
    m_categoryList->addItem("Education");
    m_categoryList->addItem("Project Mgmt");
    m_categoryList->setCurrentRow(0);
    connect(m_categoryList, &QListWidget::currentRowChanged, this, &TemplateGallery::onCategoryChanged);
    contentLayout->addWidget(m_categoryList);

    // Center: Template grid
    m_templateGrid = new QListWidget();
    m_templateGrid->setViewMode(QListView::IconMode);
    m_templateGrid->setIconSize(QSize(120, 90));
    m_templateGrid->setGridSize(QSize(145, 125));
    m_templateGrid->setResizeMode(QListView::Adjust);
    m_templateGrid->setWordWrap(true);
    m_templateGrid->setSpacing(6);
    m_templateGrid->setMovement(QListView::Static);

    for (const auto& tmpl : m_allTemplates) {
        auto* item = new QListWidgetItem(generateThumbnail(tmpl.id, tmpl.accentColor), tmpl.name);
        item->setData(Qt::UserRole, tmpl.id);
        item->setData(Qt::UserRole + 1, static_cast<int>(tmpl.category));
        item->setTextAlignment(Qt::AlignHCenter);
        m_templateGrid->addItem(item);
    }

    connect(m_templateGrid, &QListWidget::currentItemChanged,
            this, [this](QListWidgetItem* item, QListWidgetItem*) {
        if (item) onTemplateSelected(item);
    });
    connect(m_templateGrid, &QListWidget::itemDoubleClicked, this, &TemplateGallery::onTemplateDoubleClicked);
    contentLayout->addWidget(m_templateGrid, 1);

    // Right: Preview + description
    QVBoxLayout* previewLayout = new QVBoxLayout();
    m_previewLabel = new QLabel();
    m_previewLabel->setFixedSize(240, 180);
    m_previewLabel->setAlignment(Qt::AlignCenter);
    m_previewLabel->setStyleSheet("border: 1px solid #D0D5DD; border-radius: 6px; background: white;");
    previewLayout->addWidget(m_previewLabel);

    m_descriptionLabel = new QLabel("Select a template to preview.");
    m_descriptionLabel->setWordWrap(true);
    m_descriptionLabel->setStyleSheet("color: #667085; font-size: 12px; padding: 6px;");
    m_descriptionLabel->setMaximumWidth(240);
    previewLayout->addWidget(m_descriptionLabel);
    previewLayout->addStretch();
    contentLayout->addLayout(previewLayout);

    mainLayout->addLayout(contentLayout, 1);

    // Buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Create from Template");
    buttons->button(QDialogButtonBox::Ok)->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 8px 24px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }");
    buttons->button(QDialogButtonBox::Cancel)->setStyleSheet(
        "QPushButton { background: #F0F2F5; border: 1px solid #D0D5DD; border-radius: 4px; "
        "padding: 8px 20px; }"
        "QPushButton:hover { background: #E8ECF0; }");
    connect(buttons, &QDialogButtonBox::accepted, this, [this]() {
        auto* item = m_templateGrid->currentItem();
        if (item) {
            m_result = buildTemplate(item->data(Qt::UserRole).toString());
            accept();
        }
    });
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);
}

void TemplateGallery::onCategoryChanged(int row) {
    TemplateCategory cat = static_cast<TemplateCategory>(row);
    filterByCategory(cat);
}

void TemplateGallery::filterByCategory(TemplateCategory cat) {
    for (int i = 0; i < m_templateGrid->count(); ++i) {
        auto* item = m_templateGrid->item(i);
        if (cat == TemplateCategory::All) {
            item->setHidden(false);
        } else {
            int itemCat = item->data(Qt::UserRole + 1).toInt();
            item->setHidden(itemCat != static_cast<int>(cat));
        }
    }
}

void TemplateGallery::onTemplateSelected(QListWidgetItem* item) {
    QString id = item->data(Qt::UserRole).toString();
    for (const auto& tmpl : m_allTemplates) {
        if (tmpl.id == id) {
            m_descriptionLabel->setText("<b>" + tmpl.name + "</b><br><br>" + tmpl.description);
            // Scale up thumbnail for preview
            QIcon icon = item->icon();
            QPixmap pix = icon.pixmap(QSize(240, 180));
            m_previewLabel->setPixmap(pix.scaled(236, 176, Qt::KeepAspectRatio, Qt::SmoothTransformation));
            break;
        }
    }
}

void TemplateGallery::onTemplateDoubleClicked(QListWidgetItem* item) {
    if (item) {
        m_result = buildTemplate(item->data(Qt::UserRole).toString());
        accept();
    }
}

TemplateResult TemplateGallery::buildTemplate(const QString& id) {
    if (id == "finance_budget") return buildBudgetTracker();
    if (id == "finance_invoice") return buildInvoice();
    if (id == "finance_expense") return buildExpenseReport();
    if (id == "finance_dashboard") return buildFinancialDashboard();
    if (id == "biz_sales") return buildSalesReport();
    if (id == "biz_timeline") return buildProjectTimeline();
    if (id == "biz_directory") return buildEmployeeDirectory();
    if (id == "biz_agenda") return buildMeetingAgenda();
    if (id == "personal_workout") return buildWorkoutLog();
    if (id == "personal_meal") return buildMealPlanner();
    if (id == "personal_travel") return buildTravelItinerary();
    if (id == "personal_habit") return buildHabitTracker();
    if (id == "edu_grades") return buildGradeTracker();
    if (id == "edu_schedule") return buildClassSchedule();
    if (id == "edu_roster") return buildStudentRoster();
    if (id == "pm_taskboard") return buildProjectTaskBoard();
    if (id == "pm_gantt") return buildGanttChart();
    if (id == "finance_family") return buildFamilyBudget();
    if (id == "personal_wedding") return buildWeddingPlanner();
    if (id == "personal_home") return buildHomeInventory();
    if (id == "biz_clients") return buildClientTracker();
    if (id == "biz_event") return buildEventPlanner();
    if (id == "biz_inventory") return buildInventoryTracker();
    if (id == "pm_comparison") return buildComparisonChart();
    if (id == "finance_kpi") return buildKPIDashboard();
    return {};
}

// ============== Helpers ==============

void TemplateGallery::applyHeaderStyle(Spreadsheet* s, int row, int startCol, int endCol,
                                        const QString& bgColor, const QString& fgColor,
                                        int fontSize, bool bold) {
    for (int c = startCol; c <= endCol; ++c) {
        auto cell = s->getCell(CellAddress(row, c));
        CellStyle style = cell->getStyle();
        style.backgroundColor = bgColor;
        style.foregroundColor = fgColor;
        style.fontSize = fontSize;
        style.bold = bold;
        style.hAlign = HorizontalAlignment::Center;
        cell->setStyle(style);
    }
}

void TemplateGallery::applyBorders(Spreadsheet* s, int r1, int c1, int r2, int c2,
                                    const QString& color) {
    BorderStyle bs; bs.enabled = true; bs.color = color; bs.width = 1;
    for (int r = r1; r <= r2; ++r) {
        for (int c = c1; c <= c2; ++c) {
            auto cell = s->getCell(CellAddress(r, c));
            CellStyle st = cell->getStyle();
            st.borderTop = bs; st.borderBottom = bs; st.borderLeft = bs; st.borderRight = bs;
            cell->setStyle(st);
        }
    }
}

void TemplateGallery::applyCurrencyFormat(Spreadsheet* s, int r1, int c1, int r2, int c2) {
    for (int r = r1; r <= r2; ++r) {
        for (int c = c1; c <= c2; ++c) {
            auto cell = s->getCell(CellAddress(r, c));
            CellStyle st = cell->getStyle();
            st.numberFormat = "Currency"; st.decimalPlaces = 0; st.hAlign = HorizontalAlignment::Right;
            cell->setStyle(st);
        }
    }
}

void TemplateGallery::applyPercentFormat(Spreadsheet* s, int r1, int c1, int r2, int c2) {
    for (int r = r1; r <= r2; ++r) {
        for (int c = c1; c <= c2; ++c) {
            auto cell = s->getCell(CellAddress(r, c));
            CellStyle st = cell->getStyle();
            st.numberFormat = "Percentage"; st.decimalPlaces = 1; st.hAlign = HorizontalAlignment::Right;
            cell->setStyle(st);
        }
    }
}

void TemplateGallery::setColumnWidths(Spreadsheet* s, const std::vector<std::pair<int,int>>& cw) {
    for (const auto& [col, width] : cw) {
        s->setColumnWidth(col, width);
    }
}

void TemplateGallery::setCellStyleRange(Spreadsheet* s, int r1, int c1, int r2, int c2,
                                         const QString& bgColor) {
    for (int r = r1; r <= r2; ++r) {
        for (int c = c1; c <= c2; ++c) {
            auto cell = s->getCell(CellAddress(r, c));
            CellStyle st = cell->getStyle();
            st.backgroundColor = bgColor;
            cell->setStyle(st);
        }
    }
}

void TemplateGallery::setRowHeights(Spreadsheet* s, const std::vector<std::pair<int,int>>& rh) {
    for (const auto& [row, height] : rh) {
        s->setRowHeight(row, height);
    }
}

void TemplateGallery::applyBandedRows(Spreadsheet* s, int startRow, int endRow,
                                       int startCol, int endCol,
                                       const QString& evenColor, const QString& oddColor) {
    for (int r = startRow; r <= endRow; ++r) {
        QString bg = ((r - startRow) % 2 == 0) ? oddColor : evenColor;
        for (int c = startCol; c <= endCol; ++c) {
            auto cell = s->getCell(CellAddress(r, c));
            CellStyle st = cell->getStyle();
            st.backgroundColor = bg;
            cell->setStyle(st);
        }
    }
}

void TemplateGallery::applyTitleRow(Spreadsheet* s, int row, int startCol, int endCol,
                                     const QString& bgColor, const QString& fgColor,
                                     int fontSize, int rowHeight) {
    s->setRowHeight(row, rowHeight);
    s->mergeCells(CellRange(row, startCol, row, endCol));
    for (int c = startCol; c <= endCol; ++c) {
        auto cell = s->getCell(CellAddress(row, c));
        CellStyle st = cell->getStyle();
        st.backgroundColor = bgColor;
        st.foregroundColor = fgColor;
        st.fontSize = fontSize;
        st.bold = true;
        st.vAlign = VerticalAlignment::Middle;
        cell->setStyle(st);
    }
}

// Helper: set large section title with merge
void TemplateGallery::applySectionTitle(Spreadsheet* s, int row, int startCol, int endCol,
                                         const QString& text, const QString& color, int fontSize) {
    s->setCellValue(CellAddress(row, startCol), text);
    s->mergeCells(CellRange(row, startCol, row, endCol));
    auto cell = s->getCell(CellAddress(row, startCol));
    CellStyle st = cell->getStyle();
    st.fontSize = fontSize;
    st.bold = true;
    st.foregroundColor = color;
    st.vAlign = VerticalAlignment::Middle;
    cell->setStyle(st);
}

// ============== Template Builders ==============

TemplateResult TemplateGallery::buildBudgetTracker() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Monthly Budget");

    setColumnWidths(s.get(), {{0,160},{1,110},{2,110},{3,110},{4,90},{5,90}});

    // Fill all visible cells with white background (hides gridlines)
    setCellStyleRange(s.get(), 0, 0, 28, 5, "#FFFFFF");

    // Title row - tall and bold
    s->setCellValue(CellAddress(0, 0), "Monthly Budget 2026");
    applyTitleRow(s.get(), 0, 0, 5, "#1B5E3B", "#FFFFFF", 18, 44);
    { auto c = s->getCell(CellAddress(0, 0)); CellStyle st = c->getStyle(); st.hAlign = HorizontalAlignment::Left; c->setStyle(st); }

    // Spacer row
    s->setRowHeight(1, 8);
    setCellStyleRange(s.get(), 1, 0, 1, 5, "#1B5E3B");

    // Column headers
    QStringList headers = {"Category", "Budget", "Actual", "Difference", "% Spent", "Status"};
    for (int c = 0; c < headers.size(); ++c) s->setCellValue(CellAddress(2, c), headers[c]);
    applyHeaderStyle(s.get(), 2, 0, 5, "#E8F5E9", "#1B5E3B", 11, true);
    s->setRowHeight(2, 30);

    // Income section header
    s->setCellValue(CellAddress(3, 0), "INCOME");
    s->mergeCells(CellRange(3, 0, 3, 5));
    setCellStyleRange(s.get(), 3, 0, 3, 5, "#F0F7F2");
    { auto cell = s->getCell(CellAddress(3, 0)); CellStyle st = cell->getStyle(); st.bold = true; st.foregroundColor = "#1B5E3B"; st.fontSize = 12; cell->setStyle(st); }
    s->setRowHeight(3, 28);

    struct Entry { QString name; double budget; double actual; };
    std::vector<Entry> income = {{"Salary", 5200, 5200}, {"Freelance", 800, 650}, {"Investments", 300, 320}, {"Other", 100, 75}};
    for (int i = 0; i < (int)income.size(); ++i) {
        int r = 4 + i;
        s->setCellValue(CellAddress(r, 0), income[i].name);
        s->setCellValue(CellAddress(r, 1), income[i].budget);
        s->setCellValue(CellAddress(r, 2), income[i].actual);
        s->setCellFormula(CellAddress(r, 3), QString("=C%1-B%1").arg(r+1));
        s->setCellFormula(CellAddress(r, 4), QString("=C%1/B%1").arg(r+1));
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 4, 7, 0, 5, "#F8FAFC", "#FFFFFF");
    applyCurrencyFormat(s.get(), 4, 1, 7, 3);
    applyPercentFormat(s.get(), 4, 4, 7, 4);

    // Total income
    int totalIncomeRow = 9;
    s->setCellValue(CellAddress(totalIncomeRow, 0), "Total Income");
    s->setCellFormula(CellAddress(totalIncomeRow, 1), "=SUM(B5:B8)");
    s->setCellFormula(CellAddress(totalIncomeRow, 2), "=SUM(C5:C8)");
    s->setCellFormula(CellAddress(totalIncomeRow, 3), "=C10-B10");
    applyHeaderStyle(s.get(), totalIncomeRow, 0, 5, "#D5E8D4", "#1A1A1A", 11, true);
    applyCurrencyFormat(s.get(), totalIncomeRow, 1, totalIncomeRow, 3);
    s->setRowHeight(totalIncomeRow, 28);
    applyBorders(s.get(), totalIncomeRow, 0, totalIncomeRow, 5, "#1B5E3B");

    // Spacer
    s->setRowHeight(10, 10);

    // Expenses section header
    s->setCellValue(CellAddress(11, 0), "EXPENSES");
    s->mergeCells(CellRange(11, 0, 11, 5));
    setCellStyleRange(s.get(), 11, 0, 11, 5, "#FFF5F5");
    { auto cell = s->getCell(CellAddress(11, 0)); CellStyle st = cell->getStyle(); st.bold = true; st.foregroundColor = "#CC3333"; st.fontSize = 12; cell->setStyle(st); }
    s->setRowHeight(11, 28);

    std::vector<Entry> expenses = {
        {"Housing", 1500, 1500}, {"Utilities", 200, 185}, {"Groceries", 400, 420},
        {"Transportation", 150, 140}, {"Insurance", 300, 300}, {"Entertainment", 200, 250},
        {"Dining Out", 250, 280}, {"Subscriptions", 80, 80}, {"Savings", 500, 500}, {"Misc", 150, 120}
    };
    for (int i = 0; i < (int)expenses.size(); ++i) {
        int r = 12 + i;
        s->setCellValue(CellAddress(r, 0), expenses[i].name);
        s->setCellValue(CellAddress(r, 1), expenses[i].budget);
        s->setCellValue(CellAddress(r, 2), expenses[i].actual);
        s->setCellFormula(CellAddress(r, 3), QString("=C%1-B%1").arg(r+1));
        s->setCellFormula(CellAddress(r, 4), QString("=C%1/B%1").arg(r+1));
        s->setCellFormula(CellAddress(r, 5), QString("=IF(C%1>B%1,\"Over\",\"Under\")").arg(r+1));
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 12, 21, 0, 5, "#FFF8F8", "#FFFFFF");
    applyCurrencyFormat(s.get(), 12, 1, 21, 3);
    applyPercentFormat(s.get(), 12, 4, 21, 4);
    applyBorders(s.get(), 2, 0, 21, 5, "#E0E5EA");

    // Status column - color code Over/Under
    for (int r = 12; r <= 21; ++r) {
        auto cell = s->getCell(CellAddress(r, 5));
        CellStyle st = cell->getStyle();
        st.hAlign = HorizontalAlignment::Center;
        st.bold = true;
        cell->setStyle(st);
    }

    int totalExpRow = 23;
    s->setCellValue(CellAddress(totalExpRow, 0), "Total Expenses");
    s->setCellFormula(CellAddress(totalExpRow, 1), "=SUM(B13:B22)");
    s->setCellFormula(CellAddress(totalExpRow, 2), "=SUM(C13:C22)");
    s->setCellFormula(CellAddress(totalExpRow, 3), "=C24-B24");
    applyHeaderStyle(s.get(), totalExpRow, 0, 5, "#FDEAEA", "#CC3333", 11, true);
    applyCurrencyFormat(s.get(), totalExpRow, 1, totalExpRow, 3);
    s->setRowHeight(totalExpRow, 28);
    applyBorders(s.get(), totalExpRow, 0, totalExpRow, 5, "#CC3333");

    // Net summary
    s->setRowHeight(24, 6);
    s->setCellValue(CellAddress(25, 0), "NET BALANCE");
    s->mergeCells(CellRange(25, 0, 25, 1));
    s->setCellFormula(CellAddress(25, 2), "=C10-C24");
    applyHeaderStyle(s.get(), 25, 0, 5, "#1B5E3B", "#FFFFFF", 13, true);
    applyCurrencyFormat(s.get(), 25, 2, 25, 2);
    s->setRowHeight(25, 34);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Pie;
    cfg.title = "Expense Breakdown";
    cfg.dataRange = "A12:C22";
    cfg.showLegend = true;
    cfg.themeIndex = 0;
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildInvoice() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Invoice");

    setColumnWidths(s.get(), {{0,180},{1,200},{2,70},{3,100},{4,120}});
    setCellStyleRange(s.get(), 0, 0, 28, 4, "#FFFFFF");

    // Accent bar at top
    s->setRowHeight(0, 6);
    setCellStyleRange(s.get(), 0, 0, 0, 4, "#2B5797");

    // Company info
    s->setRowHeight(1, 32);
    s->setCellValue(CellAddress(1, 0), "Acme Corporation");
    { auto c = s->getCell(CellAddress(1, 0)); CellStyle st = c->getStyle(); st.bold = true; st.fontSize = 16; st.foregroundColor = "#2B5797"; c->setStyle(st); }
    s->setCellValue(CellAddress(2, 0), "123 Business Ave, Suite 100");
    s->setCellValue(CellAddress(3, 0), "San Francisco, CA 94105");
    s->setCellValue(CellAddress(4, 0), "Phone: (555) 123-4567");
    for (int r = 2; r <= 4; ++r) {
        auto c = s->getCell(CellAddress(r, 0)); CellStyle st = c->getStyle(); st.foregroundColor = "#667085"; c->setStyle(st);
    }

    // Invoice title
    s->setCellValue(CellAddress(1, 4), "INVOICE");
    { auto c = s->getCell(CellAddress(1, 4)); CellStyle st = c->getStyle(); st.bold = true; st.fontSize = 28; st.foregroundColor = "#2B5797"; st.hAlign = HorizontalAlignment::Right; c->setStyle(st); }

    // Invoice details with light blue background
    setCellStyleRange(s.get(), 6, 0, 7, 4, "#F0F5FB");
    s->setCellValue(CellAddress(6, 0), "Invoice #:"); s->setCellValue(CellAddress(6, 1), "INV-2026-001");
    s->setCellValue(CellAddress(6, 3), "Date:"); s->setCellValue(CellAddress(6, 4), "02/21/2026");
    s->setCellValue(CellAddress(7, 0), "Terms:"); s->setCellValue(CellAddress(7, 1), "Net 30");
    s->setCellValue(CellAddress(7, 3), "Due Date:"); s->setCellValue(CellAddress(7, 4), "03/21/2026");
    for (int r = 6; r <= 7; ++r) for (int c : {0, 3}) {
        auto cell = s->getCell(CellAddress(r, c)); CellStyle st = cell->getStyle(); st.bold = true; st.foregroundColor = "#2B5797"; cell->setStyle(st);
    }
    s->setRowHeight(6, 26); s->setRowHeight(7, 26);

    // Bill To
    s->setRowHeight(9, 28);
    s->setCellValue(CellAddress(9, 0), "Bill To:");
    setCellStyleRange(s.get(), 9, 0, 9, 1, "#2B5797");
    { auto c = s->getCell(CellAddress(9, 0)); CellStyle st = c->getStyle(); st.bold = true; st.foregroundColor = "#FFFFFF"; st.fontSize = 11; c->setStyle(st); }
    s->setCellValue(CellAddress(10, 0), "Client Corp");
    { auto c = s->getCell(CellAddress(10, 0)); CellStyle st = c->getStyle(); st.bold = true; st.fontSize = 12; c->setStyle(st); }
    s->setCellValue(CellAddress(11, 0), "456 Client St, New York, NY 10001");

    // Line items header
    QStringList ih = {"Item", "Description", "Qty", "Unit Price", "Amount"};
    for (int c = 0; c < ih.size(); ++c) s->setCellValue(CellAddress(13, c), ih[c]);
    applyHeaderStyle(s.get(), 13, 0, 4, "#2B5797", "#FFFFFF", 11, true);
    s->setRowHeight(13, 30);

    struct LineItem { QString item; QString desc; int qty; double price; };
    std::vector<LineItem> items = {
        {"Web Development", "Frontend redesign & responsive layout", 40, 150},
        {"API Integration", "REST API setup & authentication", 20, 175},
        {"Database Design", "Schema optimization & migration", 15, 200},
        {"Testing & QA", "Automated tests & manual QA", 10, 125},
        {"Documentation", "Technical docs & API reference", 8, 100},
    };
    for (int i = 0; i < (int)items.size(); ++i) {
        int r = 14 + i;
        s->setCellValue(CellAddress(r, 0), items[i].item);
        s->setCellValue(CellAddress(r, 1), items[i].desc);
        s->setCellValue(CellAddress(r, 2), items[i].qty);
        s->setCellValue(CellAddress(r, 3), items[i].price);
        s->setCellFormula(CellAddress(r, 4), QString("=C%1*D%1").arg(r+1));
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 14, 18, 0, 4, "#F0F5FB", "#FFFFFF");
    applyCurrencyFormat(s.get(), 14, 3, 18, 4);
    applyBorders(s.get(), 13, 0, 18, 4, "#D0D8E8");

    // Totals section
    s->setRowHeight(19, 6);
    s->setCellValue(CellAddress(20, 3), "Subtotal:"); s->setCellFormula(CellAddress(20, 4), "=SUM(E15:E19)");
    s->setCellValue(CellAddress(21, 3), "Tax (8.5%):"); s->setCellFormula(CellAddress(21, 4), "=E21*0.085");
    for (int r = 20; r <= 21; ++r) {
        auto c = s->getCell(CellAddress(r, 3)); CellStyle st = c->getStyle(); st.bold = true; st.hAlign = HorizontalAlignment::Right; c->setStyle(st);
        s->setRowHeight(r, 26);
    }
    applyCurrencyFormat(s.get(), 20, 4, 21, 4);

    // Grand total row
    s->setCellValue(CellAddress(22, 3), "TOTAL:");
    s->setCellFormula(CellAddress(22, 4), "=E21+E22");
    applyHeaderStyle(s.get(), 22, 3, 4, "#2B5797", "#FFFFFF", 14, true);
    applyCurrencyFormat(s.get(), 22, 4, 22, 4);
    s->setRowHeight(22, 34);

    // Footer
    s->setRowHeight(24, 6);
    setCellStyleRange(s.get(), 24, 0, 24, 4, "#2B5797");
    s->setCellValue(CellAddress(25, 0), "Payment: Wire to Acme Corp, Account #1234567890");
    s->setCellValue(CellAddress(26, 0), "Thank you for your business!");
    { auto c = s->getCell(CellAddress(25, 0)); CellStyle st = c->getStyle(); st.foregroundColor = "#667085"; st.fontSize = 10; c->setStyle(st); }
    { auto c = s->getCell(CellAddress(26, 0)); CellStyle st = c->getStyle(); st.italic = true; st.foregroundColor = "#2B5797"; st.fontSize = 11; c->setStyle(st); }

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildExpenseReport() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Expense Report");

    setColumnWidths(s.get(), {{0,100},{1,120},{2,200},{3,100},{4,80},{5,80}});
    setCellStyleRange(s.get(), 0, 0, 24, 5, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Expense Report - Q1 2026");
    applyTitleRow(s.get(), 0, 0, 5, "#4F46E5", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6);
    setCellStyleRange(s.get(), 1, 0, 1, 5, "#4F46E5");

    setCellStyleRange(s.get(), 2, 0, 3, 5, "#F0F0FF");
    s->setCellValue(CellAddress(2, 0), "Employee:"); s->setCellValue(CellAddress(2, 1), "Jane Smith");
    s->setCellValue(CellAddress(3, 0), "Department:"); s->setCellValue(CellAddress(3, 1), "Engineering");
    for (int r = 2; r <= 3; ++r) {
        auto c = s->getCell(CellAddress(r, 0)); CellStyle st = c->getStyle(); st.bold = true; st.foregroundColor = "#4F46E5"; c->setStyle(st);
        s->setRowHeight(r, 26);
    }

    QStringList h = {"Date", "Category", "Description", "Amount", "Receipt", "Approved"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(5, c), h[c]);
    applyHeaderStyle(s.get(), 5, 0, 5, "#E8E0FF", "#4F46E5", 11, true);
    s->setRowHeight(5, 30);

    struct Exp { QString date; QString cat; QString desc; double amt; };
    std::vector<Exp> exps = {
        {"01/05", "Travel", "Flight to NYC - Business trip", 450}, {"01/05", "Hotel", "Marriott 2 nights", 380},
        {"01/06", "Meals", "Client dinner - NYC", 120}, {"01/15", "Software", "IDE license annual", 99},
        {"02/03", "Travel", "Uber rides - week total", 65}, {"02/04", "Conference", "Tech Summit 2026 pass", 799},
        {"02/04", "Meals", "Team lunch celebration", 85}, {"02/20", "Office", "Ergonomic monitor stand", 45},
        {"03/01", "Travel", "Train tickets round-trip", 120}, {"03/10", "Meals", "Working dinner w/ team", 95},
        {"03/15", "Software", "Cloud hosting monthly", 150}, {"03/22", "Office", "Mechanical keyboard", 129},
    };
    for (int i = 0; i < (int)exps.size(); ++i) {
        int r = 6 + i;
        s->setCellValue(CellAddress(r, 0), exps[i].date);
        s->setCellValue(CellAddress(r, 1), exps[i].cat);
        s->setCellValue(CellAddress(r, 2), exps[i].desc);
        s->setCellValue(CellAddress(r, 3), exps[i].amt);
        s->setCellValue(CellAddress(r, 4), "Yes");
        s->setCellValue(CellAddress(r, 5), "Yes");
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 6, 17, 0, 5, "#F8F6FF", "#FFFFFF");
    applyCurrencyFormat(s.get(), 6, 3, 17, 3);
    applyBorders(s.get(), 5, 0, 17, 5, "#D8D0F0");

    s->setRowHeight(18, 6);
    s->setCellValue(CellAddress(19, 2), "Total:"); s->setCellFormula(CellAddress(19, 3), "=SUM(D7:D18)");
    applyHeaderStyle(s.get(), 19, 2, 5, "#4F46E5", "#FFFFFF", 12, true);
    applyCurrencyFormat(s.get(), 19, 3, 19, 3);
    s->setRowHeight(19, 32);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Expenses by Category";
    cfg.dataRange = "B6:D18";
    cfg.showLegend = false;
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildFinancialDashboard() {
    TemplateResult res;

    // Data sheet
    auto data = std::make_shared<Spreadsheet>();
    data->setAutoRecalculate(false);
    data->setSheetName("Data");
    setColumnWidths(data.get(), {{0,80},{1,110},{2,110},{3,110}});
    setCellStyleRange(data.get(), 0, 0, 14, 3, "#FFFFFF");

    QStringList months = {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"};
    double revenues[] = {42000,45000,48000,44000,52000,55000,53000,58000,56000,62000,60000,65000};
    double expenses[] = {35000,36000,38000,37000,40000,42000,41000,44000,43000,46000,45000,48000};

    data->setCellValue(CellAddress(0, 0), "Month"); data->setCellValue(CellAddress(0, 1), "Revenue");
    data->setCellValue(CellAddress(0, 2), "Expenses"); data->setCellValue(CellAddress(0, 3), "Profit");
    applyHeaderStyle(data.get(), 0, 0, 3, "#0E7C6B", "#FFFFFF", 11, true);
    data->setRowHeight(0, 30);

    for (int i = 0; i < 12; ++i) {
        data->setCellValue(CellAddress(i+1, 0), months[i]);
        data->setCellValue(CellAddress(i+1, 1), revenues[i]);
        data->setCellValue(CellAddress(i+1, 2), expenses[i]);
        data->setCellFormula(CellAddress(i+1, 3), QString("=B%1-C%1").arg(i+2));
        data->setRowHeight(i+1, 26);
    }
    applyBandedRows(data.get(), 1, 12, 0, 3, "#F0FAF8", "#FFFFFF");
    applyCurrencyFormat(data.get(), 1, 1, 12, 3);
    applyBorders(data.get(), 0, 0, 12, 3, "#C0E0D8");
    data->setAutoRecalculate(true);

    // Dashboard sheet
    auto dash = std::make_shared<Spreadsheet>();
    dash->setAutoRecalculate(false);
    dash->setSheetName("Dashboard");
    setColumnWidths(dash.get(), {{0,140},{1,140},{2,140},{3,140}});
    setCellStyleRange(dash.get(), 0, 0, 14, 3, "#FFFFFF");

    dash->setCellValue(CellAddress(0, 0), "Financial Dashboard FY 2026");
    applyTitleRow(dash.get(), 0, 0, 3, "#0E7C6B", "#FFFFFF", 18, 48);
    dash->setRowHeight(1, 6);
    setCellStyleRange(dash.get(), 1, 0, 1, 3, "#0E7C6B");

    // KPI cards
    QStringList kpis = {"Revenue YTD", "Expenses YTD", "Net Profit", "Profit Margin"};
    QStringList values = {"$600,000", "$475,000", "$125,000", "20.8%"};
    dash->setRowHeight(2, 8);
    dash->setRowHeight(3, 24);
    dash->setRowHeight(4, 40);
    for (int i = 0; i < 4; ++i) {
        dash->setCellValue(CellAddress(3, i), kpis[i]);
        { auto cell = dash->getCell(CellAddress(3, i)); CellStyle st = cell->getStyle();
          st.backgroundColor = "#F0FAF8"; st.foregroundColor = "#667085"; st.fontSize = 10;
          st.hAlign = HorizontalAlignment::Center; st.bold = false; cell->setStyle(st); }
        dash->setCellValue(CellAddress(4, i), values[i]);
        { auto cell = dash->getCell(CellAddress(4, i)); CellStyle st = cell->getStyle();
          st.backgroundColor = "#F0FAF8"; st.bold = true; st.fontSize = 20; st.foregroundColor = "#0E7C6B";
          st.hAlign = HorizontalAlignment::Center; st.vAlign = VerticalAlignment::Middle; cell->setStyle(st); }
    }
    applyBorders(dash.get(), 3, 0, 4, 3, "#C0E0D8");

    // Quarterly summary
    dash->setRowHeight(5, 10);
    dash->setCellValue(CellAddress(6, 0), "Quarter"); dash->setCellValue(CellAddress(6, 1), "Revenue");
    dash->setCellValue(CellAddress(6, 2), "Expenses"); dash->setCellValue(CellAddress(6, 3), "Profit");
    applyHeaderStyle(dash.get(), 6, 0, 3, "#0E7C6B", "#FFFFFF", 11, true);
    dash->setRowHeight(6, 30);

    QStringList quarters = {"Q1", "Q2", "Q3", "Q4"};
    double qRev[] = {135000, 151000, 167000, 187000};
    double qExp[] = {109000, 119000, 128000, 139000};
    for (int i = 0; i < 4; ++i) {
        int r = 7 + i;
        dash->setCellValue(CellAddress(r, 0), quarters[i]);
        dash->setCellValue(CellAddress(r, 1), qRev[i]);
        dash->setCellValue(CellAddress(r, 2), qExp[i]);
        dash->setCellValue(CellAddress(r, 3), qRev[i] - qExp[i]);
        dash->setRowHeight(r, 28);
    }
    applyBandedRows(dash.get(), 7, 10, 0, 3, "#F0FAF8", "#FFFFFF");
    applyCurrencyFormat(dash.get(), 7, 1, 10, 3);
    applyBorders(dash.get(), 6, 0, 10, 3, "#C0E0D8");
    dash->setAutoRecalculate(true);

    res.sheets.push_back(dash);
    res.sheets.push_back(data);

    ChartConfig lineCfg;
    lineCfg.type = ChartType::Line;
    lineCfg.title = "Revenue vs Expenses";
    lineCfg.dataRange = "A1:C13";
    lineCfg.showLegend = true;
    res.charts.push_back(lineCfg);
    res.chartSheetIndices.push_back(1);
    return res;
}

TemplateResult TemplateGallery::buildSalesReport() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Sales Report");

    setColumnWidths(s.get(), {{0,90},{1,110},{2,110},{3,70},{4,100},{5,100},{6,100},{7,80}});
    setCellStyleRange(s.get(), 0, 0, 20, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Quarterly Sales Report - Q1 2026");
    applyTitleRow(s.get(), 0, 0, 7, "#4472C4", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#4472C4");

    QStringList h = {"Region","Rep","Product","Units","Revenue","Cost","Profit","Margin"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#D6E4F0", "#1A1A1A", 11, true);
    s->setRowHeight(2, 30);

    struct Sale { QString region, rep, product; int units; double rev, cost; };
    std::vector<Sale> sales = {
        {"North","Alice","Widget A",120,24000,16800},{"North","Alice","Widget B",85,21250,14875},
        {"North","Bob","Widget A",95,19000,13300},{"South","Carol","Widget A",110,22000,15400},
        {"South","Carol","Widget C",60,18000,12600},{"South","Dave","Widget B",75,18750,13125},
        {"East","Eve","Widget A",130,26000,18200},{"East","Eve","Widget C",45,13500,9450},
        {"East","Frank","Widget B",90,22500,15750},{"West","Grace","Widget A",105,21000,14700},
        {"West","Grace","Widget C",55,16500,11550},{"West","Hank","Widget B",80,20000,14000},
    };
    for (int i = 0; i < (int)sales.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), sales[i].region);
        s->setCellValue(CellAddress(r, 1), sales[i].rep);
        s->setCellValue(CellAddress(r, 2), sales[i].product);
        s->setCellValue(CellAddress(r, 3), sales[i].units);
        s->setCellValue(CellAddress(r, 4), sales[i].rev);
        s->setCellValue(CellAddress(r, 5), sales[i].cost);
        s->setCellFormula(CellAddress(r, 6), QString("=E%1-F%1").arg(r+1));
        s->setCellFormula(CellAddress(r, 7), QString("=G%1/E%1").arg(r+1));
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 3, 14, 0, 7, "#EDF2FA", "#FFFFFF");
    applyCurrencyFormat(s.get(), 3, 4, 14, 6);
    applyPercentFormat(s.get(), 3, 7, 14, 7);
    applyBorders(s.get(), 2, 0, 14, 7, "#C8D8EC");

    s->setRowHeight(15, 6);
    s->setCellValue(CellAddress(16, 0), "Grand Total");
    s->setCellFormula(CellAddress(16, 3), "=SUM(D4:D15)");
    s->setCellFormula(CellAddress(16, 4), "=SUM(E4:E15)");
    s->setCellFormula(CellAddress(16, 5), "=SUM(F4:F15)");
    s->setCellFormula(CellAddress(16, 6), "=SUM(G4:G15)");
    applyHeaderStyle(s.get(), 16, 0, 7, "#4472C4", "#FFFFFF", 12, true);
    applyCurrencyFormat(s.get(), 16, 4, 16, 6);
    s->setRowHeight(16, 32);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Bar;
    cfg.title = "Revenue by Region";
    cfg.dataRange = "A2:E15";
    cfg.showLegend = true;
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildProjectTimeline() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Project Timeline");

    setColumnWidths(s.get(), {{0,90},{1,180},{2,110},{3,100},{4,100},{5,80},{6,90},{7,70}});
    setCellStyleRange(s.get(), 0, 0, 16, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Project Alpha - Timeline");
    applyTitleRow(s.get(), 0, 0, 7, "#ED7D31", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#ED7D31");

    QStringList h = {"Phase","Task","Owner","Start","End","Duration","Status","Progress"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#FDE8D0", "#5D3A1A", 11, true);
    s->setRowHeight(2, 30);

    struct Task { QString phase, task, owner, start, end; int dur; QString status; int progress; };
    std::vector<Task> tasks = {
        {"Planning","Requirements","Alice","Jan 6","Jan 17",10,"Complete",100},
        {"Planning","Architecture","Bob","Jan 13","Jan 24",10,"Complete",100},
        {"Design","UI Mockups","Carol","Jan 27","Feb 7",10,"Complete",100},
        {"Design","API Design","Dave","Feb 3","Feb 14",10,"In Progress",80},
        {"Develop","Frontend","Eve","Feb 10","Mar 7",20,"In Progress",60},
        {"Develop","Backend","Frank","Feb 17","Mar 14",20,"In Progress",45},
        {"Develop","Database","Grace","Feb 24","Mar 7",10,"Not Started",0},
        {"Testing","Unit Tests","Hank","Mar 10","Mar 21",10,"Not Started",0},
        {"Testing","Integration","Alice","Mar 17","Mar 28",10,"Not Started",0},
        {"Launch","Deployment","Bob","Mar 31","Apr 4",5,"Not Started",0},
    };
    for (int i = 0; i < (int)tasks.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), tasks[i].phase);
        s->setCellValue(CellAddress(r, 1), tasks[i].task);
        s->setCellValue(CellAddress(r, 2), tasks[i].owner);
        s->setCellValue(CellAddress(r, 3), tasks[i].start);
        s->setCellValue(CellAddress(r, 4), tasks[i].end);
        s->setCellValue(CellAddress(r, 5), tasks[i].dur);
        s->setCellValue(CellAddress(r, 6), tasks[i].status);
        s->setCellValue(CellAddress(r, 7), QString("%1%").arg(tasks[i].progress));
        s->setRowHeight(r, 26);
        // Color code status
        QString bg = tasks[i].status == "Complete" ? "#D4EDDA" :
                     tasks[i].status == "In Progress" ? "#FFF3CD" : "#F5F5F5";
        setCellStyleRange(s.get(), r, 6, r, 6, bg);
        auto sc = s->getCell(CellAddress(r, 6)); CellStyle sst = sc->getStyle(); sst.hAlign = HorizontalAlignment::Center; sst.bold = true; sc->setStyle(sst);
    }
    applyBandedRows(s.get(), 3, 12, 0, 5, "#FFF8F0", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 12, 7, "#E8D0B0");
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildEmployeeDirectory() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Employee Directory");

    setColumnWidths(s.get(), {{0,50},{1,130},{2,110},{3,130},{4,190},{5,110},{6,100},{7,90}});
    setCellStyleRange(s.get(), 0, 0, 20, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Company Employee Directory");
    applyTitleRow(s.get(), 0, 0, 7, "#5B6B7D", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#5B6B7D");

    QStringList h = {"ID","Name","Department","Title","Email","Phone","Start Date","Location"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#E8ECF0", "#3A4A5C", 11, true);
    s->setRowHeight(2, 30);

    struct Emp { int id; QString name, dept, title, email, phone, date, loc; };
    std::vector<Emp> emps = {
        {1001,"Alice Johnson","Engineering","Sr. Developer","alice@acme.com","555-0101","2020-03-15","SF"},
        {1002,"Bob Smith","Engineering","Tech Lead","bob@acme.com","555-0102","2019-07-01","SF"},
        {1003,"Carol Williams","Marketing","Marketing Mgr","carol@acme.com","555-0103","2021-01-10","NY"},
        {1004,"Dave Brown","Sales","Account Exec","dave@acme.com","555-0104","2022-05-20","NY"},
        {1005,"Eve Davis","Engineering","Jr. Developer","eve@acme.com","555-0105","2023-09-01","SF"},
        {1006,"Frank Miller","HR","HR Manager","frank@acme.com","555-0106","2018-11-15","SF"},
        {1007,"Grace Wilson","Finance","Controller","grace@acme.com","555-0107","2020-06-01","NY"},
        {1008,"Hank Moore","Sales","Sales Dir","hank@acme.com","555-0108","2019-02-14","CHI"},
        {1009,"Ivy Taylor","Engineering","DevOps Eng","ivy@acme.com","555-0109","2021-08-01","SF"},
        {1010,"Jack Anderson","Marketing","Designer","jack@acme.com","555-0110","2022-03-15","NY"},
        {1011,"Kate Thomas","Finance","Accountant","kate@acme.com","555-0111","2023-01-10","SF"},
        {1012,"Leo Jackson","Engineering","Backend Dev","leo@acme.com","555-0112","2021-11-20","Remote"},
        {1013,"Mia White","HR","Recruiter","mia@acme.com","555-0113","2022-07-01","SF"},
        {1014,"Noah Harris","Sales","BDR","noah@acme.com","555-0114","2023-04-15","CHI"},
        {1015,"Olivia Martin","Engineering","QA Engineer","olivia@acme.com","555-0115","2020-09-01","SF"},
    };
    for (int i = 0; i < (int)emps.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), emps[i].id);
        s->setCellValue(CellAddress(r, 1), emps[i].name);
        s->setCellValue(CellAddress(r, 2), emps[i].dept);
        s->setCellValue(CellAddress(r, 3), emps[i].title);
        s->setCellValue(CellAddress(r, 4), emps[i].email);
        s->setCellValue(CellAddress(r, 5), emps[i].phone);
        s->setCellValue(CellAddress(r, 6), emps[i].date);
        s->setCellValue(CellAddress(r, 7), emps[i].loc);
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 3, 17, 0, 7, "#F0F2F5", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 17, 7, "#D0D5DD");
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildMeetingAgenda() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Meeting Agenda");

    setColumnWidths(s.get(), {{0,90},{1,220},{2,110},{3,80},{4,200}});
    setCellStyleRange(s.get(), 0, 0, 16, 4, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Weekly Team Meeting");
    applyTitleRow(s.get(), 0, 0, 4, "#217346", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 28);
    setCellStyleRange(s.get(), 1, 0, 1, 4, "#F0F7F2");
    s->setCellValue(CellAddress(1, 0), "Date: Feb 21, 2026  |  Time: 9:00 AM  |  Room: Conference A");
    { auto c = s->getCell(CellAddress(1, 0)); CellStyle st = c->getStyle(); st.foregroundColor = "#217346"; st.italic = true; c->setStyle(st); }

    QStringList h = {"Time","Topic","Presenter","Duration","Notes"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(3, c), h[c]);
    applyHeaderStyle(s.get(), 3, 0, 4, "#D4EDDA", "#1A5C2A", 11, true);
    s->setRowHeight(3, 30);

    struct Item { QString time, topic, presenter, dur; };
    std::vector<Item> items = {
        {"9:00","Opening & Updates","Alice","10 min"},
        {"9:10","Sprint Review","Bob","15 min"},
        {"9:25","Blockers Discussion","Team","15 min"},
        {"9:40","Feature Demo: Dashboard","Carol","10 min"},
        {"9:50","Customer Feedback","Dave","10 min"},
        {"10:00","Break","","5 min"},
        {"10:05","Architecture Review","Eve","20 min"},
        {"10:25","Action Items & Wrap-up","Alice","5 min"},
    };
    for (int i = 0; i < (int)items.size(); ++i) {
        int r = 4 + i;
        s->setCellValue(CellAddress(r, 0), items[i].time);
        s->setCellValue(CellAddress(r, 1), items[i].topic);
        s->setCellValue(CellAddress(r, 2), items[i].presenter);
        s->setCellValue(CellAddress(r, 3), items[i].dur);
        s->setRowHeight(r, 28);
    }
    applyBandedRows(s.get(), 4, 11, 0, 4, "#F0F7F2", "#FFFFFF");
    applyBorders(s.get(), 3, 0, 11, 4, "#C0D8C4");
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildWorkoutLog() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Workout Log");

    setColumnWidths(s.get(), {{0,90},{1,150},{2,60},{3,60},{4,90},{5,80},{6,80},{7,140}});
    setCellStyleRange(s.get(), 0, 0, 16, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Weekly Workout Log");
    applyTitleRow(s.get(), 0, 0, 7, "#D94166", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#D94166");

    QStringList h = {"Day","Exercise","Sets","Reps","Weight (lbs)","Duration","Calories","Notes"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#FDE8EE", "#8B1A3A", 11, true);
    s->setRowHeight(2, 30);

    struct W { QString day, exercise; int sets, reps, weight, dur, cal; QString notes; };
    std::vector<W> workouts = {
        {"Monday","Bench Press",4,10,135,8,80,""},{"Monday","Squats",4,8,185,10,120,"PR attempt"},
        {"Tuesday","Running",0,0,0,30,350,"5K pace"},
        {"Wednesday","Deadlift",4,6,225,10,110,""},{"Wednesday","Pull-ups",3,12,0,6,60,"Bodyweight"},
        {"Thursday","Yoga",0,0,0,45,200,"Flexibility focus"},
        {"Friday","Shoulder Press",4,10,95,8,75,""},{"Friday","Lunges",3,12,50,8,90,""},
        {"Saturday","HIIT",0,0,0,25,400,"Tabata"},
        {"Sunday","Rest",0,0,0,0,0,"Active recovery"},
    };
    for (int i = 0; i < (int)workouts.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), workouts[i].day);
        s->setCellValue(CellAddress(r, 1), workouts[i].exercise);
        if (workouts[i].sets > 0) s->setCellValue(CellAddress(r, 2), workouts[i].sets);
        if (workouts[i].reps > 0) s->setCellValue(CellAddress(r, 3), workouts[i].reps);
        if (workouts[i].weight > 0) s->setCellValue(CellAddress(r, 4), workouts[i].weight);
        s->setCellValue(CellAddress(r, 5), workouts[i].dur);
        s->setCellValue(CellAddress(r, 6), workouts[i].cal);
        s->setCellValue(CellAddress(r, 7), workouts[i].notes);
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 3, 12, 0, 7, "#FFF0F4", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 12, 7, "#E8C0CC");
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Calories per Day";
    cfg.dataRange = "A2:G13";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildMealPlanner() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Meal Planner");

    setColumnWidths(s.get(), {{0,100},{1,160},{2,160},{3,160},{4,120}});
    setCellStyleRange(s.get(), 0, 0, 12, 4, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Weekly Meal Planner");
    applyTitleRow(s.get(), 0, 0, 4, "#2D8C4E", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 4, "#2D8C4E");

    QStringList h = {"Day","Breakfast","Lunch","Dinner","Snack"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 4, "#D4EDDA", "#1A5C2A", 11, true);
    s->setRowHeight(2, 30);

    QStringList days = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"};
    QString meals[7][4] = {
        {"Oatmeal + Berries","Grilled Chicken Salad","Salmon + Quinoa","Greek Yogurt"},
        {"Eggs + Toast","Turkey Wrap","Pasta Primavera","Apple + PB"},
        {"Smoothie Bowl","Sushi Bowl","Stir Fry Tofu","Trail Mix"},
        {"Pancakes","Caesar Salad","Grilled Steak","Hummus + Veggies"},
        {"Avocado Toast","Soup + Sandwich","Fish Tacos","Protein Bar"},
        {"French Toast","Poke Bowl","Pizza (homemade)","Fruit Salad"},
        {"Brunch: Eggs Benedict","Leftover Pizza","Roast Chicken","Nuts + Dark Chocolate"},
    };
    for (int i = 0; i < 7; ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), days[i]);
        for (int m = 0; m < 4; ++m) s->setCellValue(CellAddress(r, 1+m), meals[i][m]);
        s->setRowHeight(r, 30);
    }
    applyBandedRows(s.get(), 3, 9, 0, 4, "#F0FFF4", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 9, 4, "#B0D8B8");
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildTravelItinerary() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Travel Itinerary");

    setColumnWidths(s.get(), {{0,50},{1,90},{2,70},{3,180},{4,130},{5,90},{6,110},{7,130}});
    setCellStyleRange(s.get(), 0, 0, 18, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Trip to Tokyo - March 2026");
    applyTitleRow(s.get(), 0, 0, 7, "#E67E22", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 28);
    setCellStyleRange(s.get(), 1, 0, 1, 7, "#FFF5EB");
    s->setCellValue(CellAddress(1, 0), "Dates: Mar 1-7  |  Budget: $3,500");
    { auto c = s->getCell(CellAddress(1, 0)); CellStyle st = c->getStyle(); st.italic = true; st.foregroundColor = "#E67E22"; c->setStyle(st); }

    QStringList h = {"Day","Date","Time","Activity","Location","Cost","Confirm #","Notes"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(3, c), h[c]);
    applyHeaderStyle(s.get(), 3, 0, 7, "#FDE8D0", "#7A4A1A", 11, true);
    s->setRowHeight(3, 30);

    struct Act { int day; QString date, time, activity, location; double cost; QString confirm, notes; };
    std::vector<Act> acts = {
        {1,"Mar 1","8:00 AM","Flight SFO-NRT","SFO Airport",850,"AA1234","Direct 11h"},
        {2,"Mar 2","10:00 AM","Hotel Check-in","Shinjuku Hotel",0,"HT5678",""},
        {2,"Mar 2","2:00 PM","Meiji Shrine","Harajuku",0,"","Walk from Shinjuku"},
        {3,"Mar 3","9:00 AM","Tsukiji Market","Tsukiji",50,"","Sushi breakfast"},
        {3,"Mar 3","2:00 PM","teamLab Borderless","Odaiba",35,"TL9012","Book ahead"},
        {4,"Mar 4","10:00 AM","Day trip: Hakone","Hakone",80,"","Round trip pass"},
        {5,"Mar 5","9:00 AM","Akihabara","Akihabara",100,"","Shopping"},
        {5,"Mar 5","6:00 PM","Shibuya Crossing","Shibuya",60,"","Dinner nearby"},
        {6,"Mar 6","10:00 AM","Asakusa Temple","Asakusa",0,"","Senso-ji"},
        {7,"Mar 7","8:00 AM","Flight NRT-SFO","Narita Airport",0,"AA5678","Check-out 6AM"},
    };
    for (int i = 0; i < (int)acts.size(); ++i) {
        int r = 4 + i;
        s->setCellValue(CellAddress(r, 0), acts[i].day);
        s->setCellValue(CellAddress(r, 1), acts[i].date);
        s->setCellValue(CellAddress(r, 2), acts[i].time);
        s->setCellValue(CellAddress(r, 3), acts[i].activity);
        s->setCellValue(CellAddress(r, 4), acts[i].location);
        s->setCellValue(CellAddress(r, 5), acts[i].cost);
        s->setCellValue(CellAddress(r, 6), acts[i].confirm);
        s->setCellValue(CellAddress(r, 7), acts[i].notes);
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 4, 13, 0, 7, "#FFF8F0", "#FFFFFF");
    applyCurrencyFormat(s.get(), 4, 5, 13, 5);
    applyBorders(s.get(), 3, 0, 13, 7, "#E8C8A0");

    s->setRowHeight(14, 6);
    s->setCellValue(CellAddress(15, 4), "Total Cost:");
    s->setCellFormula(CellAddress(15, 5), "=SUM(F5:F14)");
    applyHeaderStyle(s.get(), 15, 4, 7, "#E67E22", "#FFFFFF", 12, true);
    applyCurrencyFormat(s.get(), 15, 5, 15, 5);
    s->setRowHeight(15, 32);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildHabitTracker() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Habit Tracker");

    // Col 0 = habit name, cols 1-28 = days, col 29 = count, col 30 = %
    setColumnWidths(s.get(), {{0,130}});
    for (int c = 1; c <= 28; ++c) s->setColumnWidth(c, 32);
    s->setColumnWidth(29, 60); s->setColumnWidth(30, 50);
    setCellStyleRange(s.get(), 0, 0, 12, 30, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "February 2026 Habit Tracker");
    applyTitleRow(s.get(), 0, 0, 30, "#0EA5E9", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 30, "#0EA5E9");

    // Day numbers header
    s->setCellValue(CellAddress(2, 0), "Habit");
    for (int d = 1; d <= 28; ++d) s->setCellValue(CellAddress(2, d), d);
    s->setCellValue(CellAddress(2, 29), "Count"); s->setCellValue(CellAddress(2, 30), "%");
    applyHeaderStyle(s.get(), 2, 0, 30, "#E0F2FE", "#1A1A1A", 10, true);

    QStringList habits = {"Exercise", "Reading", "Meditation", "Water 8 cups", "Sleep 8hrs", "No Sugar", "Journaling"};
    QString check = QString::fromUtf8("\xe2\x9c\x93");

    for (int h = 0; h < habits.size(); ++h) {
        int r = 3 + h;
        s->setCellValue(CellAddress(r, 0), habits[h]);
        { auto cell = s->getCell(CellAddress(r, 0)); CellStyle st = cell->getStyle(); st.bold = true; cell->setStyle(st); }
        // Random pattern of checks
        int count = 0;
        for (int d = 1; d <= 28; ++d) {
            bool done = ((d + h * 3) % 3 != 0) && (d <= 21 || h < 4);
            if (done) {
                s->setCellValue(CellAddress(r, d), check);
                setCellStyleRange(s.get(), r, d, r, d, "#D4EDDA");
                count++;
            }
        }
        s->setCellValue(CellAddress(r, 29), count);
        s->setCellValue(CellAddress(r, 30), QString("%1%").arg(qRound(count * 100.0 / 28)));
    }
    applyBorders(s.get(), 2, 0, 9, 30);
    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Bar;
    cfg.title = "Habit Completion";
    cfg.dataRange = "A2:AD3";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildGradeTracker() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Grade Tracker");

    setColumnWidths(s.get(), {{0,140},{1,65},{2,65},{3,65},{4,75},{5,65},{6,65},{7,65},{8,80},{9,55}});
    setCellStyleRange(s.get(), 0, 0, 24, 9, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Student Grade Tracker - Spring 2026");
    applyTitleRow(s.get(), 0, 0, 9, "#4F46E5", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 9, "#4F46E5");

    QStringList h = {"Student","HW1","HW2","HW3","Midterm","HW4","HW5","Final","Average","Grade"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 9, "#E8E0FF", "#3A2E8A", 11, true);
    s->setRowHeight(2, 30);

    QStringList students = {"Emma Anderson","Liam Brown","Sophia Clark","Noah Davis","Olivia Evans",
        "William Foster","Ava Garcia","James Harris","Isabella Johnson","Benjamin Kim",
        "Mia Lee","Lucas Martin","Charlotte Nelson","Henry Ortiz","Amelia Patel"};
    int grades[15][7] = {
        {92,88,95,90,87,93,91},{78,82,75,80,85,79,77},{95,97,92,98,94,96,99},
        {70,72,68,75,73,71,74},{88,85,90,82,91,87,86},{93,90,95,88,92,94,91},
        {85,88,82,90,87,84,89},{76,79,74,81,78,75,80},{98,95,97,93,96,99,94},
        {82,80,85,78,83,81,84},{90,92,88,94,91,89,93},{73,75,70,77,74,72,76},
        {87,89,84,91,86,88,90},{95,93,96,92,94,97,95},{81,83,79,85,82,80,84},
    };
    for (int i = 0; i < 15; ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), students[i]);
        for (int g = 0; g < 7; ++g) s->setCellValue(CellAddress(r, 1+g), grades[i][g]);
        s->setCellFormula(CellAddress(r, 8), QString("=AVERAGE(B%1:H%1)").arg(r+1));
        s->setCellFormula(CellAddress(r, 9),
            QString("=IF(I%1>=90,\"A\",IF(I%1>=80,\"B\",IF(I%1>=70,\"C\",IF(I%1>=60,\"D\",\"F\"))))").arg(r+1));
        s->setRowHeight(r, 26);
    }
    applyBandedRows(s.get(), 3, 17, 0, 9, "#F5F3FF", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 17, 9, "#D0C8F0");

    // Spacer + Class stats
    s->setRowHeight(18, 8);
    s->setCellValue(CellAddress(19, 0), "Class Average");
    s->setCellFormula(CellAddress(19, 8), "=AVERAGE(I4:I18)");
    s->setCellValue(CellAddress(20, 0), "Highest");
    s->setCellFormula(CellAddress(20, 8), "=MAX(I4:I18)");
    s->setCellValue(CellAddress(21, 0), "Lowest");
    s->setCellFormula(CellAddress(21, 8), "=MIN(I4:I18)");
    for (int r = 19; r <= 21; ++r) {
        applyHeaderStyle(s.get(), r, 0, 9, "#EDE9FE", "#3A2E8A", 11, true);
        s->setRowHeight(r, 28);
    }

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Student Averages";
    cfg.dataRange = "A2:I18";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildClassSchedule() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Class Schedule");

    setColumnWidths(s.get(), {{0,90},{1,130},{2,130},{3,130},{4,130},{5,130}});
    setCellStyleRange(s.get(), 0, 0, 14, 5, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Spring 2026 Class Schedule");
    applyTitleRow(s.get(), 0, 0, 5, "#7C3AED", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 5, "#7C3AED");

    QStringList days = {"Time","Monday","Tuesday","Wednesday","Thursday","Friday"};
    for (int c = 0; c < days.size(); ++c) s->setCellValue(CellAddress(2, c), days[c]);
    applyHeaderStyle(s.get(), 2, 0, 5, "#EDE9FE", "#3A2E8A", 11, true);
    s->setRowHeight(2, 30);

    QStringList times = {"8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM"};
    for (int i = 0; i < times.size(); ++i) {
        s->setCellValue(CellAddress(3+i, 0), times[i]);
        auto cell = s->getCell(CellAddress(3+i, 0));
        CellStyle st = cell->getStyle(); st.bold = true; st.foregroundColor = "#4A4A4A"; cell->setStyle(st);
        s->setRowHeight(3+i, 38);
    }

    // Place classes with colors
    struct Class { int row; int col; QString name; QString color; };
    std::vector<Class> classes = {
        {3,1,"CS 301\nAlgorithms","#DBEAFE"},{3,3,"CS 301\nAlgorithms","#DBEAFE"},{3,5,"CS 301\nAlgorithms","#DBEAFE"},
        {5,1,"MATH 201\nLinear Algebra","#FEF3C7"},{5,3,"MATH 201\nLinear Algebra","#FEF3C7"},
        {7,2,"PHYS 101\nPhysics Lab","#D1FAE5"},{7,4,"PHYS 101\nPhysics Lab","#D1FAE5"},
        {9,1,"ENG 102\nTech Writing","#FCE7F3"},{9,3,"ENG 102\nTech Writing","#FCE7F3"},
        {4,2,"CS 350\nDatabases","#E0E7FF"},{4,4,"CS 350\nDatabases","#E0E7FF"},
    };
    for (const auto& c : classes) {
        s->setCellValue(CellAddress(c.row, c.col), c.name);
        setCellStyleRange(s.get(), c.row, c.col, c.row, c.col, c.color);
    }
    applyBorders(s.get(), 2, 0, 11, 5, "#D0C8F0");

    // Course legend
    s->setRowHeight(12, 8);
    s->setCellValue(CellAddress(13, 0), "Course Legend");
    applyHeaderStyle(s.get(), 13, 0, 5, "#7C3AED", "#FFFFFF", 11, true);
    s->setRowHeight(13, 28);
    struct Legend { QString name, room, prof, color; };
    std::vector<Legend> legend = {
        {"CS 301 - Algorithms","Room 204","Dr. Smith","#DBEAFE"},
        {"CS 350 - Databases","Room 310","Dr. Jones","#E0E7FF"},
        {"MATH 201 - Linear Algebra","Room 105","Prof. Lee","#FEF3C7"},
        {"PHYS 101 - Physics Lab","Lab 102","Dr. Chen","#D1FAE5"},
        {"ENG 102 - Tech Writing","Room 401","Prof. Davis","#FCE7F3"},
    };
    for (int i = 0; i < (int)legend.size(); ++i) {
        int r = 14 + i;
        setCellStyleRange(s.get(), r, 0, r, 0, legend[i].color);
        s->setCellValue(CellAddress(r, 0), legend[i].name);
        s->setCellValue(CellAddress(r, 1), legend[i].room);
        s->setCellValue(CellAddress(r, 2), legend[i].prof);
        s->setRowHeight(r, 24);
    }

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildStudentRoster() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Student Roster");

    setColumnWidths(s.get(), {{0,40},{1,85},{2,140},{3,210},{4,100},{5,50},{6,55},{7,75}});
    setCellStyleRange(s.get(), 0, 0, 30, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "CS 301 - Student Roster - Spring 2026");
    applyTitleRow(s.get(), 0, 0, 7, "#9333EA", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#9333EA");

    QStringList h = {"#","ID","Name","Email","Major","Year","GPA","Status"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#F3E8FF", "#5B21B6", 11, true);
    s->setRowHeight(2, 30);

    QStringList names = {"Alice Wang","Bob Chen","Carol Kim","Dave Patel","Eve Johnson",
        "Frank Liu","Grace Lee","Hank Martinez","Ivy Thompson","Jack Williams",
        "Kate Brown","Leo Garcia","Mia Davis","Noah Wilson","Olivia Moore",
        "Pete Taylor","Quinn Anderson","Rachel Thomas","Sam Jackson","Tina White",
        "Uma Harris","Victor Martin","Wendy Clark","Xander Lewis","Yuki Robinson"};
    QStringList majors = {"CS","CS","CE","CS","Math","CS","CE","CS","CS","Math",
        "CS","CE","CS","CS","Math","CS","CE","CS","CS","Math","CS","CE","CS","CS","Math"};
    double gpas[] = {3.8,3.5,3.9,3.2,3.7,3.4,3.6,3.1,3.8,3.3,3.5,3.7,3.9,3.0,3.6,3.4,3.8,3.2,3.5,3.7,3.3,3.6,3.8,3.1,3.9};

    for (int i = 0; i < 25; ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), i + 1);
        s->setCellValue(CellAddress(r, 1), QString("S%1").arg(20260001 + i));
        s->setCellValue(CellAddress(r, 2), names[i]);
        s->setCellValue(CellAddress(r, 3), names[i].toLower().replace(" ", ".") + "@university.edu");
        s->setCellValue(CellAddress(r, 4), majors[i]);
        s->setCellValue(CellAddress(r, 5), (i % 4) + 1);
        s->setCellValue(CellAddress(r, 6), gpas[i]);
        s->setCellValue(CellAddress(r, 7), "Active");
        s->setRowHeight(r, 25);
    }
    applyBandedRows(s.get(), 3, 27, 0, 7, "#FAF5FF", "#FFFFFF");
    applyBorders(s.get(), 2, 0, 27, 7, "#D8C8F0");

    // Summary row
    s->setRowHeight(28, 6);
    s->setCellValue(CellAddress(29, 0), "Total Students: 25");
    s->setCellValue(CellAddress(29, 4), "Avg GPA:");
    s->setCellFormula(CellAddress(29, 6), "=AVERAGE(G4:G28)");
    applyHeaderStyle(s.get(), 29, 0, 7, "#F3E8FF", "#5B21B6", 11, true);
    s->setRowHeight(29, 30);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

TemplateResult TemplateGallery::buildProjectTaskBoard() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Task Board");

    setColumnWidths(s.get(), {{0,70},{1,200},{2,100},{3,80},{4,95},{5,65},{6,85},{7,140}});
    setCellStyleRange(s.get(), 0, 0, 18, 7, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Sprint 14 Task Board");
    applyTitleRow(s.get(), 0, 0, 7, "#D97706", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 7, "#D97706");

    QStringList h = {"Task ID","Title","Assignee","Priority","Status","Points","Due Date","Notes"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    applyHeaderStyle(s.get(), 2, 0, 7, "#FEF3C7", "#78350F", 11, true);
    s->setRowHeight(2, 30);

    struct Task { QString id, title, assignee, priority, status; int pts; QString due, notes; };
    std::vector<Task> tasks = {
        {"SP14-01","User authentication","Alice","High","Done",8,"Feb 14",""},
        {"SP14-02","Dashboard redesign","Bob","High","Done",13,"Feb 14",""},
        {"SP14-03","API rate limiting","Carol","Medium","Done",5,"Feb 17",""},
        {"SP14-04","Search feature","Dave","High","Review",8,"Feb 19",""},
        {"SP14-05","Email notifications","Eve","Medium","Review",5,"Feb 19",""},
        {"SP14-06","File upload","Frank","High","In Progress",8,"Feb 21",""},
        {"SP14-07","Report export","Grace","Medium","In Progress",5,"Feb 21",""},
        {"SP14-08","Dark mode","Hank","Low","In Progress",3,"Feb 24",""},
        {"SP14-09","Performance audit","Alice","High","To Do",8,"Feb 26",""},
        {"SP14-10","Mobile responsive","Bob","Medium","To Do",5,"Feb 26",""},
        {"SP14-11","Error handling","Carol","Medium","To Do",5,"Feb 28",""},
        {"SP14-12","Documentation","Dave","Low","Backlog",3,"Mar 3",""},
    };
    for (int i = 0; i < (int)tasks.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), tasks[i].id);
        s->setCellValue(CellAddress(r, 1), tasks[i].title);
        s->setCellValue(CellAddress(r, 2), tasks[i].assignee);
        s->setCellValue(CellAddress(r, 3), tasks[i].priority);
        s->setCellValue(CellAddress(r, 4), tasks[i].status);
        s->setCellValue(CellAddress(r, 5), tasks[i].pts);
        s->setCellValue(CellAddress(r, 6), tasks[i].due);
        s->setCellValue(CellAddress(r, 7), tasks[i].notes);
        s->setRowHeight(r, 26);
        // Priority colors
        QString pColor = tasks[i].priority == "High" ? "#FEE2E2" :
                         tasks[i].priority == "Medium" ? "#FEF3C7" : "#D1FAE5";
        setCellStyleRange(s.get(), r, 3, r, 3, pColor);
        // Status colors
        QString sColor = tasks[i].status == "Done" ? "#D1FAE5" :
                         tasks[i].status == "In Progress" ? "#DBEAFE" :
                         tasks[i].status == "Review" ? "#FEF3C7" : "#F5F5F5";
        setCellStyleRange(s.get(), r, 4, r, 4, sColor);
    }
    applyBorders(s.get(), 2, 0, 14, 7, "#E8D0A0");

    // Summary section
    s->setRowHeight(15, 8);
    s->setCellValue(CellAddress(16, 0), "Sprint Summary");
    applyHeaderStyle(s.get(), 16, 0, 7, "#D97706", "#FFFFFF", 12, true);
    s->setRowHeight(16, 30);
    s->setCellValue(CellAddress(17, 0), "Total Points:");
    s->setCellFormula(CellAddress(17, 5), "=SUM(F4:F15)");
    s->setCellValue(CellAddress(17, 3), "Velocity:");
    s->setCellValue(CellAddress(17, 4), "76 pts");
    applyHeaderStyle(s.get(), 17, 0, 7, "#FFFBEB", "#78350F", 11, true);
    s->setRowHeight(17, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Pie;
    cfg.title = "Tasks by Status";
    cfg.dataRange = "B2:F15";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildGanttChart() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Gantt Chart");

    setColumnWidths(s.get(), {{0,160},{1,75},{2,75},{3,55},{4,90}});
    for (int c = 5; c <= 16; ++c) s->setColumnWidth(c, 50);
    setCellStyleRange(s.get(), 0, 0, 15, 16, "#FFFFFF");

    s->setCellValue(CellAddress(0, 0), "Project Gantt Chart - Q1 2026");
    applyTitleRow(s.get(), 0, 0, 16, "#059669", "#FFFFFF", 16, 44);
    s->setRowHeight(1, 6); setCellStyleRange(s.get(), 1, 0, 1, 16, "#059669");

    // Headers
    QStringList h = {"Task","Start","End","Weeks","Owner"};
    for (int c = 0; c < h.size(); ++c) s->setCellValue(CellAddress(2, c), h[c]);
    // Week headers
    for (int w = 1; w <= 12; ++w) s->setCellValue(CellAddress(2, 4 + w), QString("W%1").arg(w));
    applyHeaderStyle(s.get(), 2, 0, 16, "#D1FAE5", "#065F46", 10, true);
    s->setRowHeight(2, 30);

    struct GTask { QString task, owner; int startWeek, dur; QString color; };
    std::vector<GTask> tasks = {
        {"Requirements","Alice",1,2,"#BFDBFE"},
        {"Architecture","Bob",2,2,"#C7D2FE"},
        {"UI Design","Carol",3,3,"#DDD6FE"},
        {"Backend Setup","Dave",3,2,"#FBCFE8"},
        {"Database Design","Eve",4,2,"#FED7AA"},
        {"Frontend Dev","Frank",5,4,"#BBF7D0"},
        {"Backend Dev","Grace",5,4,"#A7F3D0"},
        {"API Integration","Hank",8,2,"#FDE68A"},
        {"Testing","Alice",9,3,"#FCA5A5"},
        {"Deployment","Bob",11,2,"#E9D5FF"},
    };
    for (int i = 0; i < (int)tasks.size(); ++i) {
        int r = 3 + i;
        s->setCellValue(CellAddress(r, 0), tasks[i].task);
        auto nameCell = s->getCell(CellAddress(r, 0));
        CellStyle ns = nameCell->getStyle(); ns.bold = true; nameCell->setStyle(ns);
        s->setCellValue(CellAddress(r, 1), QString("Week %1").arg(tasks[i].startWeek));
        s->setCellValue(CellAddress(r, 2), QString("Week %1").arg(tasks[i].startWeek + tasks[i].dur - 1));
        s->setCellValue(CellAddress(r, 3), tasks[i].dur);
        s->setCellValue(CellAddress(r, 4), tasks[i].owner);
        s->setRowHeight(r, 28);

        // Color the Gantt bars
        for (int w = 0; w < tasks[i].dur; ++w) {
            int col = 4 + tasks[i].startWeek + w;
            if (col <= 16) {
                setCellStyleRange(s.get(), r, col, r, col, tasks[i].color);
            }
        }
    }
    applyBorders(s.get(), 2, 0, 12, 16, "#A8D8B8");

    // Legend
    s->setRowHeight(13, 8);
    s->setCellValue(CellAddress(14, 0), "Timeline: Jan 5 - Mar 27, 2026");
    applyHeaderStyle(s.get(), 14, 0, 4, "#059669", "#FFFFFF", 11, true);
    s->setRowHeight(14, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);
    return res;
}

// ============== New Template Builders ==============

TemplateResult TemplateGallery::buildFamilyBudget() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Family Budget");

    setColumnWidths(s.get(), {{0,180},{1,70},{2,120},{3,70},{4,120},{5,70},{6,120}});
    setCellStyleRange(s.get(), 0, 0, 45, 6, "#FFFFFF");

    // Title section - large colored text like Excel template
    s->setRowHeight(0, 10);
    applySectionTitle(s.get(), 1, 0, 6, "Family Budget", "#42A5A1", 26);
    s->setRowHeight(1, 50);

    applySectionTitle(s.get(), 2, 0, 6, "[Month]", "#42A5A1", 16);
    s->setCellValue(CellAddress(3, 0), "[Year]");
    s->mergeCells(CellRange(3, 0, 3, 6));
    { auto c = s->getCell(CellAddress(3, 0)); CellStyle st = c->getStyle();
      st.fontSize = 12; st.foregroundColor = "#888888"; c->setStyle(st); }
    s->setRowHeight(2, 30); s->setRowHeight(3, 22);

    // Cash Flow Section
    s->setRowHeight(5, 6); setCellStyleRange(s.get(), 5, 0, 5, 6, "#42A5A1");

    applySectionTitle(s.get(), 6, 0, 6, "Cash Flow", "#42A5A1", 20);
    s->setRowHeight(6, 40);

    QStringList cfh = {"", "", "Projected", "", "Actual", "", "Variance"};
    for (int c = 0; c < cfh.size(); ++c) s->setCellValue(CellAddress(7, c), cfh[c]);
    applyHeaderStyle(s.get(), 7, 0, 6, "#42A5A1", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    s->setCellValue(CellAddress(8, 0), "Total Income");
    s->setCellFormula(CellAddress(8, 2), "=C17"); s->setCellFormula(CellAddress(8, 4), "=E17");
    s->setCellFormula(CellAddress(8, 6), "=E8-C8");
    s->setCellValue(CellAddress(9, 0), "Total Expenses");
    s->setCellFormula(CellAddress(9, 2), "=C31"); s->setCellFormula(CellAddress(9, 4), "=E31");
    s->setCellFormula(CellAddress(9, 6), "=E9-C9");
    s->setCellValue(CellAddress(10, 0), "Total Cash");
    s->setCellFormula(CellAddress(10, 2), "=C8-C9"); s->setCellFormula(CellAddress(10, 4), "=E8-E9");
    s->setCellFormula(CellAddress(10, 6), "=E10-C10");
    for (int r = 8; r <= 10; ++r) {
        s->setRowHeight(r, 26);
        applyCurrencyFormat(s.get(), r, 2, r, 2);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        applyCurrencyFormat(s.get(), r, 6, r, 6);
    }
    for (int c = 0; c <= 6; ++c) {
        auto cell = s->getCell(CellAddress(10, c));
        CellStyle st = cell->getStyle(); st.bold = true; cell->setStyle(st);
    }
    applyBandedRows(s.get(), 8, 10, 0, 6, "#EEF9F8", "#FFFFFF");
    applyBorders(s.get(), 7, 0, 10, 6, "#B0D8D4");

    // Monthly Income Section
    s->setRowHeight(12, 10);
    applySectionTitle(s.get(), 13, 0, 6, "Monthly Income", "#42A5A1", 20);
    s->setRowHeight(13, 40);

    QStringList inh = {"", "", "Projected", "", "Actual", "", "Variance"};
    for (int c = 0; c < inh.size(); ++c) s->setCellValue(CellAddress(14, c), inh[c]);
    applyHeaderStyle(s.get(), 14, 0, 6, "#42A5A1", "#FFFFFF", 11, true);
    s->setRowHeight(14, 28);

    struct IncomeItem { QString name; int projected, actual; };
    std::vector<IncomeItem> income = {
        {"Salary", 4500, 4500}, {"Partner Salary", 3200, 3200},
        {"Freelance", 800, 650}, {"Investments", 200, 280},
    };
    for (int i = 0; i < (int)income.size(); ++i) {
        int r = 15 + i;
        s->setCellValue(CellAddress(r, 0), income[i].name);
        s->setCellValue(CellAddress(r, 2), income[i].projected);
        s->setCellValue(CellAddress(r, 4), income[i].actual);
        s->setCellFormula(CellAddress(r, 6), QString("=E%1-C%1").arg(r+1));
        applyCurrencyFormat(s.get(), r, 2, r, 2);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        applyCurrencyFormat(s.get(), r, 6, r, 6);
        s->setRowHeight(r, 25);
    }
    int itot = 15 + (int)income.size();
    s->setCellValue(CellAddress(itot, 0), "Total Income");
    s->setCellFormula(CellAddress(itot, 2), QString("=SUM(C16:C%1)").arg(itot));
    s->setCellFormula(CellAddress(itot, 4), QString("=SUM(E16:E%1)").arg(itot));
    s->setCellFormula(CellAddress(itot, 6), QString("=E%1-C%1").arg(itot+1));
    for (int c = 0; c <= 6; ++c) {
        auto cell = s->getCell(CellAddress(itot, c));
        CellStyle st = cell->getStyle(); st.bold = true; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), itot, 2, itot, 2);
    applyCurrencyFormat(s.get(), itot, 4, itot, 4);
    applyCurrencyFormat(s.get(), itot, 6, itot, 6);
    s->setRowHeight(itot, 28);
    applyBandedRows(s.get(), 15, itot - 1, 0, 6, "#EEF9F8", "#FFFFFF");
    applyBorders(s.get(), 14, 0, itot, 6, "#B0D8D4");

    // Monthly Expenses Section
    int eStart = itot + 2;
    applySectionTitle(s.get(), eStart, 0, 6, "Monthly Expenses", "#42A5A1", 20);
    s->setRowHeight(eStart, 40);

    int ehRow = eStart + 1;
    QStringList exh = {"", "", "Projected", "", "Actual", "", "Variance"};
    for (int c = 0; c < exh.size(); ++c) s->setCellValue(CellAddress(ehRow, c), exh[c]);
    applyHeaderStyle(s.get(), ehRow, 0, 6, "#42A5A1", "#FFFFFF", 11, true);
    s->setRowHeight(ehRow, 28);

    struct ExpItem { QString name; int projected, actual; };
    std::vector<ExpItem> expenses = {
        {"Mortgage / Rent", 1800, 1800}, {"Utilities", 250, 275},
        {"Groceries", 600, 650}, {"Transportation", 350, 320},
        {"Insurance", 400, 400}, {"Healthcare", 150, 180},
        {"Childcare", 800, 800}, {"Entertainment", 200, 250},
        {"Dining Out", 250, 310}, {"Clothing", 100, 85},
        {"Subscriptions", 80, 80}, {"Savings", 500, 500},
    };
    for (int i = 0; i < (int)expenses.size(); ++i) {
        int r = ehRow + 1 + i;
        s->setCellValue(CellAddress(r, 0), expenses[i].name);
        s->setCellValue(CellAddress(r, 2), expenses[i].projected);
        s->setCellValue(CellAddress(r, 4), expenses[i].actual);
        s->setCellFormula(CellAddress(r, 6), QString("=E%1-C%1").arg(r+1));
        applyCurrencyFormat(s.get(), r, 2, r, 2);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        applyCurrencyFormat(s.get(), r, 6, r, 6);
        s->setRowHeight(r, 25);
    }
    int etot = ehRow + 1 + (int)expenses.size();
    s->setCellValue(CellAddress(etot, 0), "Total Expenses");
    s->setCellFormula(CellAddress(etot, 2), QString("=SUM(C%1:C%2)").arg(ehRow+2).arg(etot));
    s->setCellFormula(CellAddress(etot, 4), QString("=SUM(E%1:E%2)").arg(ehRow+2).arg(etot));
    s->setCellFormula(CellAddress(etot, 6), QString("=E%1-C%1").arg(etot+1));
    for (int c = 0; c <= 6; ++c) {
        auto cell = s->getCell(CellAddress(etot, c));
        CellStyle st = cell->getStyle(); st.bold = true; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), etot, 2, etot, 2);
    applyCurrencyFormat(s.get(), etot, 4, etot, 4);
    applyCurrencyFormat(s.get(), etot, 6, etot, 6);
    s->setRowHeight(etot, 28);
    applyBandedRows(s.get(), ehRow + 1, etot - 1, 0, 6, "#EEF9F8", "#FFFFFF");
    applyBorders(s.get(), ehRow, 0, etot, 6, "#B0D8D4");

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Budget Overview";
    cfg.dataRange = "A7:G11";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildWeddingPlanner() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Wedding Planner");

    setColumnWidths(s.get(), {{0,160},{1,100},{2,130},{3,90},{4,100},{5,90},{6,120}});
    setCellStyleRange(s.get(), 0, 0, 35, 6, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 6, "Wedding Planner", "#D4508B", 26);
    s->setCellValue(CellAddress(2, 0), "Sarah & James  |  June 15, 2026");
    s->mergeCells(CellRange(2, 0, 2, 6));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 13; st.italic = true; st.foregroundColor = "#D4508B"; c->setStyle(st); }
    s->setRowHeight(2, 26);

    // Budget Summary
    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 6, "#D4508B");

    applySectionTitle(s.get(), 4, 0, 6, "Budget Summary", "#D4508B", 18);
    s->setRowHeight(4, 36);

    s->setCellValue(CellAddress(5, 0), "Total Budget:"); s->setCellValue(CellAddress(5, 1), 35000);
    s->setCellValue(CellAddress(5, 3), "Spent:"); s->setCellFormula(CellAddress(5, 4), "=SUM(E11:E22)");
    s->setCellValue(CellAddress(5, 5), "Remaining:");
    s->setCellFormula(CellAddress(5, 6), "=B6-E6");
    for (int c = 0; c <= 6; ++c) {
        auto cell = s->getCell(CellAddress(5, c));
        CellStyle st = cell->getStyle(); st.bold = true; st.fontSize = 12; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), 5, 1, 5, 1);
    applyCurrencyFormat(s.get(), 5, 4, 5, 4);
    applyCurrencyFormat(s.get(), 5, 6, 5, 6);
    setCellStyleRange(s.get(), 5, 0, 5, 6, "#FDF2F8");
    s->setRowHeight(5, 32);

    // Vendors & Expenses
    s->setRowHeight(7, 10);
    applySectionTitle(s.get(), 8, 0, 6, "Vendors & Expenses", "#D4508B", 18);
    s->setRowHeight(8, 36);

    QStringList vh = {"Category","Vendor","Contact","Due Date","Cost","Paid","Notes"};
    for (int c = 0; c < vh.size(); ++c) s->setCellValue(CellAddress(9, c), vh[c]);
    applyHeaderStyle(s.get(), 9, 0, 6, "#D4508B", "#FFFFFF", 11, true);
    s->setRowHeight(9, 28);

    struct Vendor { QString cat, vendor, contact, due; int cost; QString paid, notes; };
    std::vector<Vendor> vendors = {
        {"Venue","Grand Hall","555-0101","Jan 15",12000,"Yes","Deposit paid"},
        {"Catering","Gourmet Co.","555-0102","Mar 1",8500,"Partial","Tasting done"},
        {"Photography","Studio A","555-0103","Feb 15",3500,"No","Engagement shoot incl."},
        {"Flowers","Bloom & Co.","555-0104","Apr 1",2200,"No","Centerpieces + bouquet"},
        {"Music/DJ","DJ Mike","555-0105","Mar 15",1800,"Yes",""},
        {"Cake","Sweet Treats","555-0106","May 1",800,"No","3-tier vanilla"},
        {"Dress","Bridal Shop","555-0107","Feb 1",2500,"Yes","Alterations incl."},
        {"Invitations","Print Co.","555-0108","Jan 30",600,"Yes","150 guests"},
        {"Decor","Event Style","555-0109","Apr 15",1500,"No","Outdoor theme"},
        {"Hair & Makeup","Glam Team","555-0110","Jun 14",500,"No","Bridal party"},
        {"Rings","Jeweler","555-0111","May 15",1200,"No","Engraving"},
        {"Transportation","Limo Co.","555-0112","Jun 1",400,"No","2 vehicles"},
    };
    for (int i = 0; i < (int)vendors.size(); ++i) {
        int r = 10 + i;
        s->setCellValue(CellAddress(r, 0), vendors[i].cat);
        s->setCellValue(CellAddress(r, 1), vendors[i].vendor);
        s->setCellValue(CellAddress(r, 2), vendors[i].contact);
        s->setCellValue(CellAddress(r, 3), vendors[i].due);
        s->setCellValue(CellAddress(r, 4), vendors[i].cost);
        s->setCellValue(CellAddress(r, 5), vendors[i].paid);
        s->setCellValue(CellAddress(r, 6), vendors[i].notes);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        s->setRowHeight(r, 25);
        QString pColor = vendors[i].paid == "Yes" ? "#D1FAE5" :
                         vendors[i].paid == "Partial" ? "#FEF3C7" : "#FEE2E2";
        setCellStyleRange(s.get(), r, 5, r, 5, pColor);
    }
    applyBandedRows(s.get(), 10, 21, 0, 6, "#FDF2F8", "#FFFFFF");
    applyBorders(s.get(), 9, 0, 21, 6, "#E8B0CC");

    int vtot = 22;
    s->setCellValue(CellAddress(vtot, 0), "Total");
    s->setCellFormula(CellAddress(vtot, 4), "=SUM(E11:E22)");
    applyHeaderStyle(s.get(), vtot, 0, 6, "#D4508B", "#FFFFFF", 11, true);
    applyCurrencyFormat(s.get(), vtot, 4, vtot, 4);
    s->setRowHeight(vtot, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Pie;
    cfg.title = "Wedding Budget Breakdown";
    cfg.dataRange = "A9:E22";
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildHomeInventory() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Home Inventory");

    setColumnWidths(s.get(), {{0,130},{1,170},{2,80},{3,100},{4,90},{5,100},{6,100},{7,120}});
    setCellStyleRange(s.get(), 0, 0, 30, 7, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 7, "Home Inventory", "#6366F1", 26);
    s->setCellValue(CellAddress(2, 0), "Insurance Policy: #HI-2026-4521  |  Updated: Feb 2026");
    s->mergeCells(CellRange(2, 0, 2, 7));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 11; st.italic = true; st.foregroundColor = "#6366F1"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    // Summary
    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 7, "#6366F1");
    s->setCellValue(CellAddress(4, 0), "Total Value:");
    s->setCellFormula(CellAddress(4, 1), "=SUM(E9:E26)");
    s->setCellValue(CellAddress(4, 3), "Items:");
    s->setCellValue(CellAddress(4, 4), 18);
    s->setCellValue(CellAddress(4, 6), "Rooms:");
    s->setCellValue(CellAddress(4, 7), 5);
    setCellStyleRange(s.get(), 4, 0, 4, 7, "#EEF2FF");
    for (int c = 0; c <= 7; ++c) {
        auto cell = s->getCell(CellAddress(4, c));
        CellStyle st = cell->getStyle(); st.bold = true; st.fontSize = 12; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), 4, 1, 4, 1);
    s->setRowHeight(4, 32);

    // Items table
    s->setRowHeight(5, 10);
    applySectionTitle(s.get(), 6, 0, 7, "Inventory Items", "#6366F1", 18);
    s->setRowHeight(6, 36);

    QStringList hdr = {"Room","Item","Qty","Brand/Model","Value","Purchase Date","Condition","Serial/Notes"};
    for (int c = 0; c < hdr.size(); ++c) s->setCellValue(CellAddress(7, c), hdr[c]);
    applyHeaderStyle(s.get(), 7, 0, 7, "#6366F1", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    struct Item { QString room, item; int qty; QString brand; int value; QString date, condition, notes; };
    std::vector<Item> items = {
        {"Living Room","Sofa",1,"West Elm",2800,"2024-03","Excellent",""},
        {"Living Room","TV 65\"",1,"Samsung QN65",1200,"2025-01","Excellent","SN: SM65Q1234"},
        {"Living Room","Coffee Table",1,"IKEA",350,"2023-06","Good",""},
        {"Living Room","Bookshelf",2,"Custom",600,"2022-11","Good",""},
        {"Kitchen","Refrigerator",1,"LG French Door",2200,"2024-08","Excellent","SN: LG8812"},
        {"Kitchen","Dishwasher",1,"Bosch 500",900,"2024-08","Excellent",""},
        {"Kitchen","Cookware Set",1,"All-Clad",500,"2023-12","Good","10-piece"},
        {"Kitchen","Stand Mixer",1,"KitchenAid",350,"2025-06","New",""},
        {"Bedroom","Bed + Mattress",1,"Casper King",2400,"2024-01","Excellent",""},
        {"Bedroom","Dresser",1,"Pottery Barn",1100,"2023-05","Good",""},
        {"Bedroom","Nightstands",2,"Target",200,"2023-05","Good",""},
        {"Office","Desk",1,"Uplift V2",800,"2024-06","Excellent","Standing desk"},
        {"Office","Chair",1,"Herman Miller",1400,"2024-06","Excellent","Aeron"},
        {"Office","MacBook Pro",1,"Apple M3 16\"",3500,"2025-02","Excellent","SN: C02X1234"},
        {"Office","Monitor",2,"Dell 27\"",700,"2024-09","Excellent",""},
        {"Garage","Power Tools",1,"DeWalt",800,"2022-04","Good","Drill, saw, etc."},
        {"Garage","Bicycles",2,"Trek",1600,"2023-07","Good",""},
        {"Garage","Lawn Mower",1,"Honda",450,"2021-05","Fair",""},
    };
    for (int i = 0; i < (int)items.size(); ++i) {
        int r = 8 + i;
        s->setCellValue(CellAddress(r, 0), items[i].room);
        s->setCellValue(CellAddress(r, 1), items[i].item);
        s->setCellValue(CellAddress(r, 2), items[i].qty);
        s->setCellValue(CellAddress(r, 3), items[i].brand);
        s->setCellValue(CellAddress(r, 4), items[i].value);
        s->setCellValue(CellAddress(r, 5), items[i].date);
        s->setCellValue(CellAddress(r, 6), items[i].condition);
        s->setCellValue(CellAddress(r, 7), items[i].notes);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        s->setRowHeight(r, 24);
        QString cColor = items[i].condition == "Excellent" ? "#D1FAE5" :
                         items[i].condition == "Good" ? "#FEF3C7" :
                         items[i].condition == "New" ? "#DBEAFE" : "#FEE2E2";
        setCellStyleRange(s.get(), r, 6, r, 6, cColor);
    }
    int iEnd = 8 + (int)items.size() - 1;
    applyBandedRows(s.get(), 8, iEnd, 0, 7, "#EEF2FF", "#FFFFFF");
    applyBorders(s.get(), 7, 0, iEnd, 7, "#C0C8F0");

    s->setCellValue(CellAddress(iEnd + 1, 0), "Total Value");
    s->setCellFormula(CellAddress(iEnd + 1, 4), QString("=SUM(E9:E%1)").arg(iEnd + 1));
    applyHeaderStyle(s.get(), iEnd + 1, 0, 7, "#6366F1", "#FFFFFF", 11, true);
    applyCurrencyFormat(s.get(), iEnd + 1, 4, iEnd + 1, 4);
    s->setRowHeight(iEnd + 1, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Pie;
    cfg.title = "Value by Room";
    cfg.dataRange = QString("A7:E%1").arg(iEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildClientTracker() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Client Tracker");

    setColumnWidths(s.get(), {{0,140},{1,130},{2,120},{3,90},{4,100},{5,90},{6,100},{7,130}});
    setCellStyleRange(s.get(), 0, 0, 25, 7, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 7, "Client Tracker", "#0891B2", 26);
    s->setCellValue(CellAddress(2, 0), "Sales Pipeline  |  Q1 2026");
    s->mergeCells(CellRange(2, 0, 2, 7));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 12; st.italic = true; st.foregroundColor = "#0891B2"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    // KPI bar
    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 7, "#0891B2");
    setCellStyleRange(s.get(), 4, 0, 4, 7, "#ECFEFF");
    s->setCellValue(CellAddress(4, 0), "Active Deals:"); s->setCellValue(CellAddress(4, 1), 15);
    s->setCellValue(CellAddress(4, 2), "Pipeline Value:"); s->setCellFormula(CellAddress(4, 3), "=SUM(F9:F23)");
    s->setCellValue(CellAddress(4, 4), "Won:"); s->setCellValue(CellAddress(4, 5), 5);
    s->setCellValue(CellAddress(4, 6), "Win Rate:"); s->setCellValue(CellAddress(4, 7), "33%");
    for (int c = 0; c <= 7; ++c) {
        auto cell = s->getCell(CellAddress(4, c));
        CellStyle st = cell->getStyle(); st.bold = true; st.fontSize = 12; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), 4, 3, 4, 3);
    s->setRowHeight(4, 32);

    // Pipeline table
    s->setRowHeight(5, 10);
    applySectionTitle(s.get(), 6, 0, 7, "Deal Pipeline", "#0891B2", 18);
    s->setRowHeight(6, 36);

    QStringList dh = {"Company","Contact","Email","Stage","Close Date","Deal Value","Owner","Notes"};
    for (int c = 0; c < dh.size(); ++c) s->setCellValue(CellAddress(7, c), dh[c]);
    applyHeaderStyle(s.get(), 7, 0, 7, "#0891B2", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    struct Deal { QString company, contact, email, stage, closeDate; int value; QString owner, notes; };
    std::vector<Deal> deals = {
        {"Acme Corp","John Smith","john@acme.com","Won","Jan 15",45000,"Alice","Signed"},
        {"TechStart","Sara Lee","sara@techstart.com","Won","Jan 20",28000,"Bob","Annual"},
        {"GlobalInc","Mike Chen","mike@global.com","Won","Feb 1",62000,"Alice","Enterprise"},
        {"DataFlow","Lisa Park","lisa@dataflow.com","Won","Feb 10",35000,"Carol","3-year"},
        {"CloudNet","Tom Davis","tom@cloudnet.com","Won","Feb 18",18000,"Bob","Starter"},
        {"MedTech","Amy Wu","amy@medtech.com","Negotiation","Mar 5",55000,"Alice","Pending legal"},
        {"FinServe","Dan Brown","dan@finserve.com","Negotiation","Mar 12",42000,"Carol","Demo done"},
        {"RetailCo","Eve Jones","eve@retailco.com","Proposal","Mar 20",30000,"Bob","Sent Feb 25"},
        {"LogiTech","Ray Kim","ray@logitech.com","Proposal","Mar 25",25000,"Alice","Follow up"},
        {"EduPlatform","Mia Lin","mia@edu.com","Discovery","Apr 1",38000,"Carol","First call"},
        {"HealthApp","Sam Patel","sam@health.com","Discovery","Apr 10",22000,"Bob","Qualified"},
        {"AutoDrive","Jess Tang","jess@auto.com","Lead","Apr 15",50000,"Alice","Inbound"},
        {"FoodChain","Alex Rios","alex@food.com","Lead","Apr 20",15000,"Carol","Referral"},
        {"GameStudio","Pat Cho","pat@games.com","Lost","Feb 28",35000,"Bob","Budget cut"},
        {"MediaGroup","Kim West","kim@media.com","Lost","Mar 1",20000,"Alice","Chose competitor"},
    };
    for (int i = 0; i < (int)deals.size(); ++i) {
        int r = 8 + i;
        s->setCellValue(CellAddress(r, 0), deals[i].company);
        s->setCellValue(CellAddress(r, 1), deals[i].contact);
        s->setCellValue(CellAddress(r, 2), deals[i].email);
        s->setCellValue(CellAddress(r, 3), deals[i].stage);
        s->setCellValue(CellAddress(r, 4), deals[i].closeDate);
        s->setCellValue(CellAddress(r, 5), deals[i].value);
        s->setCellValue(CellAddress(r, 6), deals[i].owner);
        s->setCellValue(CellAddress(r, 7), deals[i].notes);
        applyCurrencyFormat(s.get(), r, 5, r, 5);
        s->setRowHeight(r, 25);
        QString sColor = deals[i].stage == "Won" ? "#D1FAE5" :
                         deals[i].stage == "Negotiation" ? "#FEF3C7" :
                         deals[i].stage == "Proposal" ? "#DBEAFE" :
                         deals[i].stage == "Discovery" ? "#E0E7FF" :
                         deals[i].stage == "Lead" ? "#F3E8FF" : "#FEE2E2";
        setCellStyleRange(s.get(), r, 3, r, 3, sColor);
    }
    int dEnd = 8 + (int)deals.size() - 1;
    applyBandedRows(s.get(), 8, dEnd, 0, 7, "#ECFEFF", "#FFFFFF");
    applyBorders(s.get(), 7, 0, dEnd, 7, "#A0D8E0");

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Bar;
    cfg.title = "Pipeline by Stage";
    cfg.dataRange = QString("A7:F%1").arg(dEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildEventPlanner() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Event Planner");

    setColumnWidths(s.get(), {{0,160},{1,120},{2,90},{3,90},{4,90},{5,80},{6,130}});
    setCellStyleRange(s.get(), 0, 0, 30, 6, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 6, "Event Planner", "#9333EA", 26);
    s->setCellValue(CellAddress(2, 0), "Annual Company Conference  |  March 28, 2026");
    s->mergeCells(CellRange(2, 0, 2, 6));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 12; st.italic = true; st.foregroundColor = "#9333EA"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    // Info bar
    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 6, "#9333EA");
    setCellStyleRange(s.get(), 4, 0, 4, 6, "#F3E8FF");
    s->setCellValue(CellAddress(4, 0), "Venue: Grand Convention Center");
    s->setCellValue(CellAddress(4, 2), "Attendees: 250");
    s->setCellValue(CellAddress(4, 4), "Budget:");
    s->setCellFormula(CellAddress(4, 5), "=SUM(E9:E21)");
    for (int c = 0; c <= 6; ++c) {
        auto cell = s->getCell(CellAddress(4, c));
        CellStyle st = cell->getStyle(); st.bold = true; st.fontSize = 11; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), 4, 5, 4, 5);
    s->setRowHeight(4, 30);

    // Tasks section
    s->setRowHeight(5, 10);
    applySectionTitle(s.get(), 6, 0, 6, "Planning Tasks", "#9333EA", 18);
    s->setRowHeight(6, 36);

    QStringList th = {"Task","Assigned To","Deadline","Status","Budget","Spent","Notes"};
    for (int c = 0; c < th.size(); ++c) s->setCellValue(CellAddress(7, c), th[c]);
    applyHeaderStyle(s.get(), 7, 0, 6, "#9333EA", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    struct ETask { QString task, assigned, deadline, status; int budget, spent; QString notes; };
    std::vector<ETask> tasks = {
        {"Book venue","Sarah","Jan 15","Complete",8000,8000,"Confirmed"},
        {"Hire caterer","Mike","Feb 1","Complete",6000,5800,"Menu finalized"},
        {"AV equipment","Tom","Feb 15","Complete",3000,2900,"Projectors + mics"},
        {"Print materials","Lisa","Feb 20","In Progress",1500,800,"Brochures + badges"},
        {"Book speakers","Sarah","Feb 25","In Progress",5000,2000,"3 of 5 confirmed"},
        {"Photography","Amy","Mar 1","In Progress",2000,0,"Getting quotes"},
        {"Decorations","Lisa","Mar 10","Pending",2500,0,"Theme: Innovation"},
        {"Gift bags","Mike","Mar 15","Pending",1000,0,"Branded merch"},
        {"Transport","Tom","Mar 20","Pending",1500,0,"Shuttle service"},
        {"Marketing","Amy","Feb 10","Complete",2000,1900,"Email + social"},
        {"Registration","Sarah","Jan 20","Complete",500,450,"Online portal"},
        {"Insurance","Mike","Feb 5","Complete",800,800,"Event liability"},
        {"Entertainment","Tom","Mar 5","Pending",1200,0,"Live band"},
    };
    for (int i = 0; i < (int)tasks.size(); ++i) {
        int r = 8 + i;
        s->setCellValue(CellAddress(r, 0), tasks[i].task);
        s->setCellValue(CellAddress(r, 1), tasks[i].assigned);
        s->setCellValue(CellAddress(r, 2), tasks[i].deadline);
        s->setCellValue(CellAddress(r, 3), tasks[i].status);
        s->setCellValue(CellAddress(r, 4), tasks[i].budget);
        s->setCellValue(CellAddress(r, 5), tasks[i].spent);
        s->setCellValue(CellAddress(r, 6), tasks[i].notes);
        applyCurrencyFormat(s.get(), r, 4, r, 4);
        applyCurrencyFormat(s.get(), r, 5, r, 5);
        s->setRowHeight(r, 25);
        QString sColor = tasks[i].status == "Complete" ? "#D1FAE5" :
                         tasks[i].status == "In Progress" ? "#DBEAFE" : "#F5F5F5";
        setCellStyleRange(s.get(), r, 3, r, 3, sColor);
    }
    int eEnd = 8 + (int)tasks.size() - 1;
    applyBandedRows(s.get(), 8, eEnd, 0, 6, "#F3E8FF", "#FFFFFF");
    applyBorders(s.get(), 7, 0, eEnd, 6, "#D0B8F0");

    s->setCellValue(CellAddress(eEnd + 1, 0), "Total");
    s->setCellFormula(CellAddress(eEnd + 1, 4), QString("=SUM(E9:E%1)").arg(eEnd + 1));
    s->setCellFormula(CellAddress(eEnd + 1, 5), QString("=SUM(F9:F%1)").arg(eEnd + 1));
    applyHeaderStyle(s.get(), eEnd + 1, 0, 6, "#9333EA", "#FFFFFF", 11, true);
    applyCurrencyFormat(s.get(), eEnd + 1, 4, eEnd + 1, 4);
    applyCurrencyFormat(s.get(), eEnd + 1, 5, eEnd + 1, 5);
    s->setRowHeight(eEnd + 1, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Column;
    cfg.title = "Budget vs Spent";
    cfg.dataRange = QString("A7:F%1").arg(eEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildInventoryTracker() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Inventory");

    setColumnWidths(s.get(), {{0,70},{1,170},{2,100},{3,70},{4,80},{5,80},{6,90},{7,120}});
    setCellStyleRange(s.get(), 0, 0, 28, 7, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 7, "Inventory Tracker", "#EA580C", 26);
    s->setCellValue(CellAddress(2, 0), "Warehouse A  |  Last Updated: Feb 21, 2026");
    s->mergeCells(CellRange(2, 0, 2, 7));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 11; st.italic = true; st.foregroundColor = "#EA580C"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    // Summary
    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 7, "#EA580C");
    setCellStyleRange(s.get(), 4, 0, 4, 7, "#FFF7ED");
    s->setCellValue(CellAddress(4, 0), "Total SKUs:"); s->setCellValue(CellAddress(4, 1), 16);
    s->setCellValue(CellAddress(4, 2), "Total Units:"); s->setCellFormula(CellAddress(4, 3), "=SUM(D9:D24)");
    s->setCellValue(CellAddress(4, 4), "Total Value:"); s->setCellFormula(CellAddress(4, 5), "=SUM(G9:G24)");
    s->setCellValue(CellAddress(4, 6), "Low Stock:"); s->setCellValue(CellAddress(4, 7), "3 items");
    for (int c = 0; c <= 7; ++c) {
        auto cell = s->getCell(CellAddress(4, c));
        CellStyle st = cell->getStyle(); st.bold = true; st.fontSize = 11; cell->setStyle(st);
    }
    applyCurrencyFormat(s.get(), 4, 5, 4, 5);
    s->setRowHeight(4, 30);

    // Products table
    s->setRowHeight(5, 10);
    applySectionTitle(s.get(), 6, 0, 7, "Product Inventory", "#EA580C", 18);
    s->setRowHeight(6, 36);

    QStringList ph = {"SKU","Product Name","Category","In Stock","Reorder At","Unit Price","Total Value","Status"};
    for (int c = 0; c < ph.size(); ++c) s->setCellValue(CellAddress(7, c), ph[c]);
    applyHeaderStyle(s.get(), 7, 0, 7, "#EA580C", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    struct Prod { QString sku, name, category; int stock, reorder; double price; QString status; };
    std::vector<Prod> products = {
        {"WH-001","Wireless Headphones","Electronics",145,50,79.99,"In Stock"},
        {"KB-002","Mechanical Keyboard","Electronics",82,30,129.99,"In Stock"},
        {"MS-003","Ergonomic Mouse","Electronics",210,40,49.99,"In Stock"},
        {"MN-004","27\" Monitor","Electronics",35,20,349.99,"In Stock"},
        {"LP-005","Laptop Stand","Accessories",95,25,39.99,"In Stock"},
        {"CB-006","USB-C Cable 6ft","Accessories",420,100,12.99,"In Stock"},
        {"WC-007","Webcam HD","Electronics",28,30,69.99,"Low Stock"},
        {"DK-008","Standing Desk","Furniture",12,10,599.99,"In Stock"},
        {"CH-009","Office Chair","Furniture",8,15,449.99,"Low Stock"},
        {"MP-010","Mouse Pad XL","Accessories",310,50,19.99,"In Stock"},
        {"SP-011","Speakers","Electronics",55,20,89.99,"In Stock"},
        {"HB-012","USB Hub 7-port","Accessories",180,40,29.99,"In Stock"},
        {"BG-013","Laptop Backpack","Accessories",65,25,59.99,"In Stock"},
        {"WB-014","Whiteboard 4x6","Office",22,10,149.99,"In Stock"},
        {"PH-015","Phone Stand","Accessories",200,50,14.99,"In Stock"},
        {"TB-016","Tablet Case","Accessories",5,20,34.99,"Low Stock"},
    };
    for (int i = 0; i < (int)products.size(); ++i) {
        int r = 8 + i;
        s->setCellValue(CellAddress(r, 0), products[i].sku);
        s->setCellValue(CellAddress(r, 1), products[i].name);
        s->setCellValue(CellAddress(r, 2), products[i].category);
        s->setCellValue(CellAddress(r, 3), products[i].stock);
        s->setCellValue(CellAddress(r, 4), products[i].reorder);
        s->setCellValue(CellAddress(r, 5), products[i].price);
        s->setCellFormula(CellAddress(r, 6), QString("=D%1*F%1").arg(r+1));
        s->setCellValue(CellAddress(r, 7), products[i].status);
        applyCurrencyFormat(s.get(), r, 5, r, 5);
        applyCurrencyFormat(s.get(), r, 6, r, 6);
        s->setRowHeight(r, 25);
        QString stColor = products[i].status == "Low Stock" ? "#FEE2E2" : "#D1FAE5";
        setCellStyleRange(s.get(), r, 7, r, 7, stColor);
    }
    int pEnd = 8 + (int)products.size() - 1;
    applyBandedRows(s.get(), 8, pEnd, 0, 7, "#FFF7ED", "#FFFFFF");
    applyBorders(s.get(), 7, 0, pEnd, 7, "#E8C0A0");

    s->setCellValue(CellAddress(pEnd + 1, 0), "Total");
    s->setCellFormula(CellAddress(pEnd + 1, 3), QString("=SUM(D9:D%1)").arg(pEnd + 1));
    s->setCellFormula(CellAddress(pEnd + 1, 6), QString("=SUM(G9:G%1)").arg(pEnd + 1));
    applyHeaderStyle(s.get(), pEnd + 1, 0, 7, "#EA580C", "#FFFFFF", 11, true);
    applyCurrencyFormat(s.get(), pEnd + 1, 6, pEnd + 1, 6);
    s->setRowHeight(pEnd + 1, 28);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Bar;
    cfg.title = "Stock Levels";
    cfg.dataRange = QString("B7:D%1").arg(pEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildComparisonChart() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("Comparison");

    setColumnWidths(s.get(), {{0,160},{1,90},{2,90},{3,90},{4,90},{5,100}});
    setCellStyleRange(s.get(), 0, 0, 25, 5, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 5, "Comparison Matrix", "#2563EB", 26);
    s->setCellValue(CellAddress(2, 0), "Project Management Tool Selection  |  Feb 2026");
    s->mergeCells(CellRange(2, 0, 2, 5));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 11; st.italic = true; st.foregroundColor = "#2563EB"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 5, "#2563EB");

    // Scoring guide
    s->setCellValue(CellAddress(4, 0), "Scoring: 1 = Poor  |  2 = Fair  |  3 = Good  |  4 = Very Good  |  5 = Excellent");
    s->mergeCells(CellRange(4, 0, 4, 5));
    setCellStyleRange(s.get(), 4, 0, 4, 5, "#EFF6FF");
    { auto c = s->getCell(CellAddress(4, 0)); CellStyle st = c->getStyle();
      st.italic = true; st.foregroundColor = "#2563EB"; c->setStyle(st); }
    s->setRowHeight(4, 28);

    // Feature comparison
    s->setRowHeight(5, 10);
    applySectionTitle(s.get(), 6, 0, 5, "Feature Comparison", "#2563EB", 18);
    s->setRowHeight(6, 36);

    QStringList ch = {"Criteria","Weight","Option A","Option B","Option C","Winner"};
    for (int c = 0; c < ch.size(); ++c) s->setCellValue(CellAddress(7, c), ch[c]);
    applyHeaderStyle(s.get(), 7, 0, 5, "#2563EB", "#FFFFFF", 11, true);
    s->setRowHeight(7, 28);

    // Option names sub-header
    s->setCellValue(CellAddress(8, 2), "Jira"); s->setCellValue(CellAddress(8, 3), "Asana");
    s->setCellValue(CellAddress(8, 4), "Monday.com");
    applyHeaderStyle(s.get(), 8, 0, 5, "#DBEAFE", "#1E40AF", 11, true);
    s->setRowHeight(8, 26);

    struct Crit { QString name; int weight, a, b, c; };
    std::vector<Crit> criteria = {
        {"Ease of Use",20,3,5,4},
        {"Customization",15,5,3,4},
        {"Integrations",15,5,4,3},
        {"Reporting",10,4,3,4},
        {"Price",20,3,4,3},
        {"Scalability",10,5,3,3},
        {"Mobile App",5,3,4,5},
        {"Support",5,4,4,3},
    };
    for (int i = 0; i < (int)criteria.size(); ++i) {
        int r = 9 + i;
        s->setCellValue(CellAddress(r, 0), criteria[i].name);
        s->setCellValue(CellAddress(r, 1), QString("%1%").arg(criteria[i].weight));
        s->setCellValue(CellAddress(r, 2), criteria[i].a);
        s->setCellValue(CellAddress(r, 3), criteria[i].b);
        s->setCellValue(CellAddress(r, 4), criteria[i].c);
        int best = qMax(criteria[i].a, qMax(criteria[i].b, criteria[i].c));
        QString winner;
        if (criteria[i].a == best) winner += "A ";
        if (criteria[i].b == best) winner += "B ";
        if (criteria[i].c == best) winner += "C";
        s->setCellValue(CellAddress(r, 5), winner.trimmed());
        s->setRowHeight(r, 26);

        for (int sc = 2; sc <= 4; ++sc) {
            int val = (sc == 2) ? criteria[i].a : (sc == 3) ? criteria[i].b : criteria[i].c;
            QString color = val >= 5 ? "#D1FAE5" : val >= 4 ? "#ECFCCB" : val >= 3 ? "#FEF3C7" : val >= 2 ? "#FED7AA" : "#FEE2E2";
            setCellStyleRange(s.get(), r, sc, r, sc, color);
        }
    }
    int cEnd = 9 + (int)criteria.size() - 1;
    applyBandedRows(s.get(), 9, cEnd, 0, 1, "#EFF6FF", "#FFFFFF");
    applyBorders(s.get(), 7, 0, cEnd, 5, "#B0C8F0");

    // Weighted scores
    int wRow = cEnd + 1;
    s->setRowHeight(wRow, 6);
    int sRow = wRow + 1;
    s->setCellValue(CellAddress(sRow, 0), "Weighted Score");
    s->setCellValue(CellAddress(sRow, 1), "100%");
    double wa = 0, wb = 0, wc = 0;
    for (auto& cr : criteria) {
        wa += cr.a * cr.weight / 100.0;
        wb += cr.b * cr.weight / 100.0;
        wc += cr.c * cr.weight / 100.0;
    }
    s->setCellValue(CellAddress(sRow, 2), QString::number(wa, 'f', 1));
    s->setCellValue(CellAddress(sRow, 3), QString::number(wb, 'f', 1));
    s->setCellValue(CellAddress(sRow, 4), QString::number(wc, 'f', 1));
    double bestScore = qMax(wa, qMax(wb, wc));
    QString overallWinner;
    if (wa == bestScore) overallWinner = "A - Jira";
    else if (wb == bestScore) overallWinner = "B - Asana";
    else overallWinner = "C - Monday.com";
    s->setCellValue(CellAddress(sRow, 5), overallWinner);
    applyHeaderStyle(s.get(), sRow, 0, 5, "#2563EB", "#FFFFFF", 12, true);
    s->setRowHeight(sRow, 32);

    // Recommendation
    int rRow = sRow + 2;
    s->setCellValue(CellAddress(rRow, 0), QString("Recommendation: %1 with weighted score %2/5.0").arg(overallWinner).arg(QString::number(bestScore, 'f', 1)));
    s->mergeCells(CellRange(rRow, 0, rRow, 5));
    setCellStyleRange(s.get(), rRow, 0, rRow, 5, "#EFF6FF");
    { auto c = s->getCell(CellAddress(rRow, 0)); CellStyle st = c->getStyle();
      st.bold = true; st.fontSize = 12; st.foregroundColor = "#1E40AF"; c->setStyle(st); }
    s->setRowHeight(rRow, 30);

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Bar;
    cfg.title = "Feature Scores Comparison";
    cfg.dataRange = QString("A8:E%1").arg(cEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}

TemplateResult TemplateGallery::buildKPIDashboard() {
    TemplateResult res;
    auto s = std::make_shared<Spreadsheet>();
    s->setAutoRecalculate(false);
    s->setSheetName("KPI Dashboard");

    setColumnWidths(s.get(), {{0,160},{1,100},{2,100},{3,100},{4,90},{5,100},{6,120}});
    setCellStyleRange(s.get(), 0, 0, 30, 6, "#FFFFFF");

    s->setRowHeight(0, 8); s->setRowHeight(1, 50);
    applySectionTitle(s.get(), 1, 0, 6, "KPI Dashboard", "#DC2626", 26);
    s->setCellValue(CellAddress(2, 0), "Executive Summary  |  Q1 2026");
    s->mergeCells(CellRange(2, 0, 2, 6));
    { auto c = s->getCell(CellAddress(2, 0)); CellStyle st = c->getStyle();
      st.fontSize = 12; st.italic = true; st.foregroundColor = "#DC2626"; c->setStyle(st); }
    s->setRowHeight(2, 24);

    s->setRowHeight(3, 6); setCellStyleRange(s.get(), 3, 0, 3, 6, "#DC2626");

    // KPI table
    applySectionTitle(s.get(), 4, 0, 6, "Key Performance Indicators", "#DC2626", 18);
    s->setRowHeight(4, 36);

    QStringList kh = {"KPI","Target","Actual","Variance","Status","Trend","Notes"};
    for (int c = 0; c < kh.size(); ++c) s->setCellValue(CellAddress(5, c), kh[c]);
    applyHeaderStyle(s.get(), 5, 0, 6, "#DC2626", "#FFFFFF", 11, true);
    s->setRowHeight(5, 28);

    struct KPI { QString name; double target, actual; QString trend, notes; };
    std::vector<KPI> kpis = {
        {"Revenue ($M)",12.5,13.2,"Up","Beat by $700K"},
        {"Gross Margin %",65.0,62.3,"Down","Material costs up"},
        {"New Customers",500,485,"Flat","Marketing ramp"},
        {"Customer Retention %",92.0,94.5,"Up","Loyalty program"},
        {"NPS Score",70,73,"Up","Improved support"},
        {"Avg Deal Size ($K)",25.0,27.8,"Up","Enterprise growth"},
        {"Sales Cycle (days)",45,42,"Up","Faster close"},
        {"Employee Satisfaction",4.2,4.0,"Down","Survey pending"},
        {"Ticket Resolution (hrs)",4.0,3.5,"Up","Automation helped"},
        {"Uptime %",99.9,99.95,"Up","Zero incidents"},
    };
    QString checkMark = QString::fromUtf8("\xe2\x9c\x93");
    QString crossMark = QString::fromUtf8("\xe2\x9c\x97");
    for (int i = 0; i < (int)kpis.size(); ++i) {
        int r = 6 + i;
        s->setCellValue(CellAddress(r, 0), kpis[i].name);
        s->setCellValue(CellAddress(r, 1), kpis[i].target);
        s->setCellValue(CellAddress(r, 2), kpis[i].actual);
        double variance = kpis[i].actual - kpis[i].target;
        s->setCellValue(CellAddress(r, 3), QString::number(variance, 'f', 1));
        bool lowerBetter = (kpis[i].name.contains("days") || kpis[i].name.contains("hrs"));
        bool onTarget = lowerBetter ? (kpis[i].actual <= kpis[i].target) : (kpis[i].actual >= kpis[i].target);
        s->setCellValue(CellAddress(r, 4), onTarget ? checkMark + " On Track" : crossMark + " Below");
        s->setCellValue(CellAddress(r, 5), kpis[i].trend);
        s->setCellValue(CellAddress(r, 6), kpis[i].notes);
        s->setRowHeight(r, 26);
        setCellStyleRange(s.get(), r, 4, r, 4, onTarget ? "#D1FAE5" : "#FEE2E2");
        QString tColor = kpis[i].trend == "Up" ? "#D1FAE5" : kpis[i].trend == "Down" ? "#FEE2E2" : "#FEF3C7";
        setCellStyleRange(s.get(), r, 5, r, 5, tColor);
    }
    int kEnd = 6 + (int)kpis.size() - 1;
    applyBandedRows(s.get(), 6, kEnd, 0, 6, "#FEF2F2", "#FFFFFF");
    applyBorders(s.get(), 5, 0, kEnd, 6, "#E8B0B0");

    // Quarterly Trends
    int qStart = kEnd + 2;
    applySectionTitle(s.get(), qStart, 0, 6, "Quarterly Revenue Trend", "#DC2626", 18);
    s->setRowHeight(qStart, 36);

    QStringList qh = {"Quarter","Revenue ($M)","Expenses ($M)","Profit ($M)","Margin %","Headcount","Revenue/Head ($K)"};
    int qhRow = qStart + 1;
    for (int c = 0; c < qh.size(); ++c) s->setCellValue(CellAddress(qhRow, c), qh[c]);
    applyHeaderStyle(s.get(), qhRow, 0, 6, "#DC2626", "#FFFFFF", 11, true);
    s->setRowHeight(qhRow, 28);

    struct Quarter { QString name; double rev, exp; int heads; };
    std::vector<Quarter> quarters = {
        {"Q1 2025",10.2,7.1,85}, {"Q2 2025",11.0,7.5,90},
        {"Q3 2025",11.8,7.9,95}, {"Q4 2025",12.1,8.2,98},
        {"Q1 2026",13.2,8.8,102},
    };
    for (int i = 0; i < (int)quarters.size(); ++i) {
        int r = qhRow + 1 + i;
        s->setCellValue(CellAddress(r, 0), quarters[i].name);
        s->setCellValue(CellAddress(r, 1), quarters[i].rev);
        s->setCellValue(CellAddress(r, 2), quarters[i].exp);
        double profit = quarters[i].rev - quarters[i].exp;
        s->setCellValue(CellAddress(r, 3), QString::number(profit, 'f', 1));
        s->setCellValue(CellAddress(r, 4), QString("%1%").arg(qRound(profit / quarters[i].rev * 100)));
        s->setCellValue(CellAddress(r, 5), quarters[i].heads);
        s->setCellValue(CellAddress(r, 6), QString::number(quarters[i].rev / quarters[i].heads * 1000, 'f', 0));
        s->setRowHeight(r, 26);
    }
    int qEnd = qhRow + (int)quarters.size();
    applyBandedRows(s.get(), qhRow + 1, qEnd, 0, 6, "#FEF2F2", "#FFFFFF");
    applyBorders(s.get(), qhRow, 0, qEnd, 6, "#E8B0B0");

    s->setAutoRecalculate(true);
    res.sheets.push_back(s);

    ChartConfig cfg;
    cfg.type = ChartType::Line;
    cfg.title = "Revenue vs Expenses";
    cfg.dataRange = QString("A%1:C%2").arg(qhRow + 1).arg(qEnd + 1);
    res.charts.push_back(cfg);
    res.chartSheetIndices.push_back(0);
    return res;
}
