#include "SpreadsheetModel.h"
#include "../core/Spreadsheet.h"
#include "../core/NumberFormat.h"
#include "../core/UndoManager.h"
#include "../core/TableStyle.h"
#include "../core/ConditionalFormatting.h"
#include "../core/SparklineConfig.h"
#include "../core/MacroEngine.h"
#include "../core/DocumentTheme.h"
#include <QFont>
#include <QMessageBox>
#include <QApplication>
#include <QTimer>
#include <QColor>
#include <QDate>
#include <QRegularExpression>
#include <limits>

// Convert a QDate to Excel serial number (days since 1899-12-30)
static double dateToSerial(const QDate& date) {
    static const QDate epoch(1899, 12, 30);
    return epoch.daysTo(date);
}

// Convert Excel serial number to QDate
static QDate serialToDate(double serial) {
    static const QDate epoch(1899, 12, 30);
    return epoch.addDays(static_cast<int>(serial));
}

// Try to parse user input as a date. Returns true + QDate + suggested format ID.
static bool tryParseDate(const QString& input, QDate& outDate, QString& outFormatId) {
    QString trimmed = input.trimmed();
    if (trimmed.isEmpty()) return false;

    int currentYear = QDate::currentDate().year();

    // Try ISO format: 2026-01-15
    QDate d = QDate::fromString(trimmed, Qt::ISODate);
    if (d.isValid()) { outDate = d; outFormatId = "dd/MM/yyyy"; return true; }

    // Try MM/dd/yyyy
    d = QDate::fromString(trimmed, "MM/dd/yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "MM/dd/yyyy"; return true; }

    // Try dd/MM/yyyy
    d = QDate::fromString(trimmed, "dd/MM/yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "dd/MM/yyyy"; return true; }

    // Try M/d/yyyy
    d = QDate::fromString(trimmed, "M/d/yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "MM/dd/yyyy"; return true; }

    // Try M/d/yy
    d = QDate::fromString(trimmed, "M/d/yy");
    if (d.isValid()) {
        // Qt parses 2-digit year as 1900s — adjust to current century
        if (d.year() < 100) d = d.addYears(2000);
        outDate = d; outFormatId = "MM/dd/yyyy"; return true;
    }

    // Try M-d-yyyy
    d = QDate::fromString(trimmed, "M-d-yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "MM/dd/yyyy"; return true; }

    // Try d-MMM-yy (e.g., 2-Dec-26)
    d = QDate::fromString(trimmed, "d-MMM-yy");
    if (d.isValid()) {
        if (d.year() < 100) d = d.addYears(2000);
        outDate = d; outFormatId = "d MMM, yyyy"; return true;
    }

    // Try d-MMM-yyyy (e.g., 2-Dec-2026)
    d = QDate::fromString(trimmed, "d-MMM-yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "d MMM, yyyy"; return true; }

    // Try "MMM d" (e.g., "Dec 2", "Jan 15") — use current year
    d = QDate::fromString(trimmed, "MMM d");
    if (d.isValid()) {
        outDate = QDate(currentYear, d.month(), d.day());
        outFormatId = "d MMM, yyyy";
        return true;
    }

    // Try "MMM d, yyyy" (e.g., "Dec 2, 2026")
    d = QDate::fromString(trimmed, "MMM d, yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "d MMM, yyyy"; return true; }

    // Try "MMMM d" (e.g., "December 2") — use current year
    d = QDate::fromString(trimmed, "MMMM d");
    if (d.isValid()) {
        outDate = QDate(currentYear, d.month(), d.day());
        outFormatId = "d MMMM, yyyy";
        return true;
    }

    // Try "MMMM d, yyyy" (e.g., "December 2, 2026")
    d = QDate::fromString(trimmed, "MMMM d, yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "d MMMM, yyyy"; return true; }

    // Try "d MMM" (e.g., "2 Dec") — use current year
    d = QDate::fromString(trimmed, "d MMM");
    if (d.isValid()) {
        outDate = QDate(currentYear, d.month(), d.day());
        outFormatId = "d MMM, yyyy";
        return true;
    }

    // Try "d MMM yyyy" (e.g., "2 Dec 2026")
    d = QDate::fromString(trimmed, "d MMM yyyy");
    if (d.isValid()) { outDate = d; outFormatId = "d MMM, yyyy"; return true; }

    // Try M/d (e.g., "12/2") — use current year; only if both parts look like month/day
    static QRegularExpression mdRe("^(\\d{1,2})/(\\d{1,2})$");
    auto match = mdRe.match(trimmed);
    if (match.hasMatch()) {
        int m = match.captured(1).toInt();
        int day = match.captured(2).toInt();
        if (m >= 1 && m <= 12 && day >= 1 && day <= 31) {
            d = QDate(currentYear, m, day);
            if (d.isValid()) { outDate = d; outFormatId = "MM/dd/yyyy"; return true; }
        }
    }

    return false;
}

SpreadsheetModel::SpreadsheetModel(std::shared_ptr<Spreadsheet> spreadsheet, QObject* parent)
    : QAbstractTableModel(parent), m_spreadsheet(spreadsheet) {
}

int SpreadsheetModel::rowCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return m_spreadsheet ? m_spreadsheet->getRowCount() : 100;
}

int SpreadsheetModel::columnCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return m_spreadsheet ? m_spreadsheet->getColumnCount() : 26;
}

QVariant SpreadsheetModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid() || !m_spreadsheet) {
        return QVariant();
    }

    // Use getCellIfExists to avoid creating Cell objects for empty cells
    auto cell = m_spreadsheet->getCellIfExists(index.row(), index.column());

    // Fast path for empty cells — only check table styling, skip everything else
    if (!cell) {
        switch (role) {
            case Qt::BackgroundRole: {
                auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
                if (table) {
                    int startRow = table->range.getStart().row;
                    if (table->hasHeaderRow && index.row() == startRow)
                        return table->theme.headerBg;
                    int dataRow = index.row() - startRow - (table->hasHeaderRow ? 1 : 0);
                    if (table->bandedRows)
                        return (dataRow % 2 == 0) ? table->theme.bandedRow1 : table->theme.bandedRow2;
                    return table->theme.bandedRow1;
                }
                return QVariant();
            }
            case Qt::FontRole: {
                auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
                if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                    QFont font("Arial", 11);
                    font.setBold(true);
                    return font;
                }
                return QVariant();
            }
            case Qt::ForegroundRole: {
                auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
                if (table && table->hasHeaderRow && index.row() == table->range.getStart().row)
                    return table->theme.headerFg;
                return QVariant();
            }
            case Qt::TextAlignmentRole: {
                if (m_spreadsheet->hasDefaultCellStyle()) {
                    const auto& ds = m_spreadsheet->getDefaultCellStyle();
                    int alignment = 0;
                    if (ds.vAlign == VerticalAlignment::Top) alignment |= Qt::AlignTop;
                    else if (ds.vAlign == VerticalAlignment::Bottom) alignment |= Qt::AlignBottom;
                    else alignment |= Qt::AlignVCenter;
                    if (ds.hAlign == HorizontalAlignment::Left) alignment |= Qt::AlignLeft;
                    else if (ds.hAlign == HorizontalAlignment::Right) alignment |= Qt::AlignRight;
                    else if (ds.hAlign == HorizontalAlignment::Center) alignment |= Qt::AlignHCenter;
                    else alignment |= Qt::AlignLeft;
                    return alignment;
                }
                return QVariant();
            }
            default:
                return QVariant();
        }
    }

    // Shared state: get cell style once, get cell value once (avoid repeated hash lookups)
    const auto& baseStyle = cell->getStyle();
    const bool hasCustomStyle = cell->hasCustomStyle();

    switch (role) {
        case Qt::DisplayRole: {
            auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
            if (baseStyle.numberFormat != "General" && !value.toString().isEmpty()) {
                NumberFormatOptions opts;
                opts.type = NumberFormat::typeFromString(baseStyle.numberFormat);
                opts.decimalPlaces = baseStyle.decimalPlaces;
                opts.useThousandsSeparator = baseStyle.useThousandsSeparator;
                opts.currencyCode = baseStyle.currencyCode;
                opts.dateFormatId = baseStyle.dateFormatId;
                return NumberFormat::format(value.toString(), opts);
            }
            return value;
        }
        case Qt::EditRole: {
            if (cell->getType() == CellType::Formula) {
                return cell->getFormula();
            }
            // For date-formatted cells, show the date string in the editor (not serial number)
            if (baseStyle.numberFormat == "Date") {
                auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
                bool ok;
                double serial = value.toDouble(&ok);
                if (ok && serial > 0 && serial < 200000) {
                    QDate date = serialToDate(serial);
                    if (date.isValid()) {
                        return date.toString("MM/dd/yyyy");
                    }
                }
            }
            return m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
        }
        case Qt::FontRole: {
            // Fast path: cells with default style (vast majority) — return cached default font
            if (!hasCustomStyle && m_spreadsheet->getConditionalFormatting().getAllRules().empty()) {
                static const QFont s_defaultFont("Arial", 11);
                return s_defaultFont;
            }
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);

            QFont font(style.fontName);
            font.setPointSize(style.fontSize);
            font.setBold(style.bold);
            font.setItalic(style.italic);
            font.setUnderline(style.underline);
            font.setStrikeOut(style.strikethrough);
            // Table header row: force bold
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                font.setBold(true);
            }
            return font;
        }
        case Qt::ForegroundRole: {
            // Fast path: default style = black text
            if (!hasCustomStyle && m_spreadsheet->getConditionalFormatting().getAllRules().empty()) {
                return QVariant(); // Default foreground (black)
            }
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                return table->theme.headerFg;
            }
            QColor fg(style.foregroundColor);
            if (!fg.isValid()) {
                fg = m_spreadsheet->getDocumentTheme().resolveColor(style.foregroundColor);
            }
            return fg.isValid() ? QVariant(fg) : QVariant();
        }
        case Qt::BackgroundRole: {
            // Check table first (common case for styled regions)
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table) {
                int tableStartRow = table->range.getStart().row;
                if (table->hasHeaderRow && index.row() == tableStartRow) {
                    return table->theme.headerBg;
                }
                int dataRow = index.row() - tableStartRow - (table->hasHeaderRow ? 1 : 0);
                if (table->bandedRows) {
                    return (dataRow % 2 == 0) ? table->theme.bandedRow1 : table->theme.bandedRow2;
                }
                return table->theme.bandedRow1;
            }
            // Highlight invalid cells with red tint
            if (m_highlightInvalid) {
                CellAddress addr(index.row(), index.column());
                QString cellText = m_spreadsheet->getCellValue(addr).toString();
                if (!cellText.isEmpty() && !m_spreadsheet->validateCell(index.row(), index.column(), cellText)) {
                    return QColor(255, 200, 200);
                }
            }
            // Fast path: default style = white background
            if (!hasCustomStyle && m_spreadsheet->getConditionalFormatting().getAllRules().empty()) {
                return QVariant(); // Default background (white)
            }
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            QColor bg(style.backgroundColor);
            if (!bg.isValid()) {
                bg = m_spreadsheet->getDocumentTheme().resolveColor(style.backgroundColor);
            }
            if (bg.isValid() && bg != QColor("#FFFFFF") && bg != Qt::white) {
                return bg;
            }
            return QVariant();
        }
        case Qt::TextAlignmentRole: {
            // Vertical alignment
            int alignment = 0;
            if (baseStyle.vAlign == VerticalAlignment::Top) {
                alignment |= Qt::AlignTop;
            } else if (baseStyle.vAlign == VerticalAlignment::Bottom) {
                alignment |= Qt::AlignBottom;
            } else {
                alignment |= Qt::AlignVCenter;
            }

            // Horizontal alignment
            if (baseStyle.hAlign == HorizontalAlignment::Left) {
                alignment |= Qt::AlignLeft;
            } else if (baseStyle.hAlign == HorizontalAlignment::Right) {
                alignment |= Qt::AlignRight;
            } else if (baseStyle.hAlign == HorizontalAlignment::Center) {
                alignment |= Qt::AlignHCenter;
            } else {
                // General: right-align numbers and dates, left-align text
                bool isRightAligned = false;
                // Date-formatted cells are always right-aligned
                if (baseStyle.numberFormat == "Date" || baseStyle.numberFormat == "Time" ||
                    baseStyle.numberFormat == "Currency" || baseStyle.numberFormat == "Accounting" ||
                    baseStyle.numberFormat == "Percentage" || baseStyle.numberFormat == "Number" ||
                    baseStyle.numberFormat == "Scientific" || baseStyle.numberFormat == "Fraction") {
                    isRightAligned = true;
                } else {
                    auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
                    if (value.typeId() == QMetaType::Double || value.typeId() == QMetaType::Int ||
                        value.typeId() == QMetaType::LongLong || value.typeId() == QMetaType::Float) {
                        isRightAligned = true;
                    } else {
                        value.toString().toDouble(&isRightAligned);
                    }
                }
                alignment |= isRightAligned ? Qt::AlignRight : Qt::AlignLeft;
            }
            return alignment;
        }
        // Custom roles for indent and borders
        case Qt::UserRole + 10: { // Indent level
            return cell->getStyle().indentLevel;
        }
        case Qt::UserRole + 11: { // Border top
            const auto& b = cell->getStyle().borderTop;
            if (b.enabled) {
                QString resolvedColor = DocumentTheme::isThemeColor(b.color)
                    ? m_spreadsheet->getDocumentTheme().resolveColor(b.color).name() : b.color;
                return QString("%1,%2,%3").arg(b.width).arg(resolvedColor).arg(b.penStyle);
            }
            return QVariant();
        }
        case Qt::UserRole + 12: { // Border bottom
            const auto& b = cell->getStyle().borderBottom;
            if (b.enabled) {
                QString resolvedColor = DocumentTheme::isThemeColor(b.color)
                    ? m_spreadsheet->getDocumentTheme().resolveColor(b.color).name() : b.color;
                return QString("%1,%2,%3").arg(b.width).arg(resolvedColor).arg(b.penStyle);
            }
            return QVariant();
        }
        case Qt::UserRole + 13: { // Border left
            const auto& b = cell->getStyle().borderLeft;
            if (b.enabled) {
                QString resolvedColor = DocumentTheme::isThemeColor(b.color)
                    ? m_spreadsheet->getDocumentTheme().resolveColor(b.color).name() : b.color;
                return QString("%1,%2,%3").arg(b.width).arg(resolvedColor).arg(b.penStyle);
            }
            return QVariant();
        }
        case Qt::UserRole + 14: { // Border right
            const auto& b = cell->getStyle().borderRight;
            if (b.enabled) {
                QString resolvedColor = DocumentTheme::isThemeColor(b.color)
                    ? m_spreadsheet->getDocumentTheme().resolveColor(b.color).name() : b.color;
                return QString("%1,%2,%3").arg(b.width).arg(resolvedColor).arg(b.penStyle);
            }
            return QVariant();
        }
        case Qt::UserRole + 16: { // Text rotation
            return cell->getStyle().textRotation;
        }
        case SparklineRole: { // Sparkline render data
            auto* sparkline = m_spreadsheet->getSparkline(CellAddress(index.row(), index.column()));
            if (!sparkline) return QVariant();

            SparklineRenderData rd;
            rd.type = sparkline->type;
            rd.lineColor = sparkline->lineColor;
            rd.highPointColor = sparkline->highPointColor;
            rd.lowPointColor = sparkline->lowPointColor;
            rd.negativeColor = sparkline->negativeColor;
            rd.showHighPoint = sparkline->showHighPoint;
            rd.showLowPoint = sparkline->showLowPoint;
            rd.lineWidth = sparkline->lineWidth;
            rd.minVal = std::numeric_limits<double>::max();
            rd.maxVal = std::numeric_limits<double>::lowest();

            CellRange range(sparkline->dataRange);
            for (const auto& addr : range.getCells()) {
                auto val = m_spreadsheet->getCellValue(addr);
                bool ok;
                double num = val.toString().toDouble(&ok);
                if (ok) {
                    rd.values.append(num);
                    if (num < rd.minVal) { rd.minVal = num; rd.lowIndex = rd.values.size() - 1; }
                    if (num > rd.maxVal) { rd.maxVal = num; rd.highIndex = rd.values.size() - 1; }
                } else {
                    rd.values.append(0.0);
                }
            }
            if (rd.values.isEmpty()) return QVariant();
            return QVariant::fromValue(rd);
        }
        default:
            return QVariant();
    }
}

QVariant SpreadsheetModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole) {
        if (orientation == Qt::Horizontal) {
            return columnIndexToLetter(section);
        } else {
            return section + 1;
        }
    }

    if (role == Qt::TextAlignmentRole) {
        return Qt::AlignCenter;
    }

    return QVariant();
}

Qt::ItemFlags SpreadsheetModel::flags(const QModelIndex& index) const {
    if (!index.isValid()) {
        return Qt::NoItemFlags;
    }
    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

bool SpreadsheetModel::setData(const QModelIndex& index, const QVariant& value, int role) {
    if (!index.isValid() || role != Qt::EditRole || !m_spreadsheet) {
        return false;
    }

    CellAddress addr(index.row(), index.column());
    QString strValue = value.toString();

    // Auto-complete unmatched parentheses in formulas (like Excel)
    if (strValue.startsWith("=")) {
        int open = 0;
        for (QChar ch : strValue) {
            if (ch == '(') ++open;
            else if (ch == ')') --open;
        }
        while (open > 0) {
            strValue.append(')');
            --open;
        }
    }

    // Data validation check (skip for formulas)
    if (!strValue.startsWith("=") && !m_suppressUndo) {
        if (!m_spreadsheet->validateCell(index.row(), index.column(), strValue)) {
            const auto* rule = m_spreadsheet->getValidationAt(index.row(), index.column());
            if (rule && rule->showErrorAlert) {
                QString errorMsg = rule->errorMessage.isEmpty()
                    ? "The value you entered is not valid.\nA user has restricted values that can be entered into this cell."
                    : rule->errorMessage;
                QString errorTitle = rule->errorTitle.isEmpty()
                    ? "Invalid Input" : rule->errorTitle;
                auto errorStyle = rule->errorStyle;

                // Defer the dialog to avoid re-entrant event loop crash
                QTimer::singleShot(0, qApp, [errorTitle, errorMsg, errorStyle]() {
                    QWidget* parent = QApplication::activeWindow();
                    if (errorStyle == Spreadsheet::DataValidationRule::Stop) {
                        QMessageBox::critical(parent, errorTitle, errorMsg);
                    } else if (errorStyle == Spreadsheet::DataValidationRule::Warning) {
                        QMessageBox::warning(parent, errorTitle, errorMsg);
                    } else {
                        QMessageBox::information(parent, errorTitle, errorMsg);
                    }
                });

                // For Stop style, reject the value; for Warning/Info, allow it
                if (rule->errorStyle == Spreadsheet::DataValidationRule::Stop) {
                    return false;
                }
            }
        }
    }

    // If the cell didn't exist before and there's a sheet-level default style, apply it
    bool wasNew = (m_spreadsheet->getCellIfExists(addr) == nullptr);

    // Auto-detect date input (before setting value, so we can convert to serial)
    QDate parsedDate;
    QString dateFormatId;
    bool isDateInput = false;
    if (!strValue.startsWith("=")) {
        // Only try date detection if not already a plain number
        bool isNum = false;
        strValue.toDouble(&isNum);
        if (!isNum) {
            isDateInput = tryParseDate(strValue, parsedDate, dateFormatId);
        }
    }

    if (!m_suppressUndo) {
        // Single-cell edit: capture before/after for undo
        CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);

        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else if (isDateInput) {
            // Store as Excel serial number
            double serial = dateToSerial(parsedDate);
            m_spreadsheet->setCellValue(addr, QVariant(serial));
            // Set date format on the cell
            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            style.numberFormat = "Date";
            style.dateFormatId = dateFormatId;
            cell->setStyle(style);
        } else {
            m_spreadsheet->setCellValue(addr, value);
        }

        // Apply default style to newly created cells
        if (wasNew && m_spreadsheet->hasDefaultCellStyle()) {
            auto cell = m_spreadsheet->getCell(addr);
            if (cell && !cell->hasCustomStyle()) {
                CellStyle defaultStyle = m_spreadsheet->getDefaultCellStyle();
                // Preserve date format if we just set it
                if (isDateInput) {
                    defaultStyle.numberFormat = "Date";
                    defaultStyle.dateFormatId = dateFormatId;
                }
                cell->setStyle(defaultStyle);
            }
        }

        CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<CellEditCommand>(before, after));
    } else {
        // Bulk operation: caller handles undo tracking
        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else if (isDateInput) {
            double serial = dateToSerial(parsedDate);
            m_spreadsheet->setCellValue(addr, QVariant(serial));
            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            style.numberFormat = "Date";
            style.dateFormatId = dateFormatId;
            cell->setStyle(style);
        } else {
            m_spreadsheet->setCellValue(addr, value);
        }

        if (wasNew && m_spreadsheet->hasDefaultCellStyle()) {
            auto cell = m_spreadsheet->getCell(addr);
            if (cell && !cell->hasCustomStyle()) {
                CellStyle defaultStyle = m_spreadsheet->getDefaultCellStyle();
                if (isDateInput) {
                    defaultStyle.numberFormat = "Date";
                    defaultStyle.dateFormatId = dateFormatId;
                }
                cell->setStyle(defaultStyle);
            }
        }
    }

    // Record action for macro recording
    if (m_macroEngine && m_macroEngine->isRecording()) {
        QString cellRef = addr.toString();
        if (strValue.startsWith("=")) {
            m_macroEngine->recordAction(
                QString("sheet.setCellFormula(\"%1\", \"%2\");")
                    .arg(cellRef, strValue));
        } else {
            // Try to preserve numeric types
            bool isNum = false;
            double numVal = strValue.toDouble(&isNum);
            if (isNum) {
                m_macroEngine->recordAction(
                    QString("sheet.setCellValue(\"%1\", %2);")
                        .arg(cellRef).arg(numVal));
            } else {
                QString escaped = strValue;
                escaped.replace("\\", "\\\\").replace("\"", "\\\"");
                m_macroEngine->recordAction(
                    QString("sheet.setCellValue(\"%1\", \"%2\");")
                        .arg(cellRef, escaped));
            }
        }
    }

    emit dataChanged(index, index, {Qt::DisplayRole, Qt::EditRole});
    return true;
}

QString SpreadsheetModel::columnIndexToLetter(int column) const {
    QString result;
    while (column >= 0) {
        result = QChar('A' + (column % 26)) + result;
        column = column / 26 - 1;
    }
    return result;
}
