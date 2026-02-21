#include "SpreadsheetModel.h"
#include "../core/Spreadsheet.h"
#include "../core/NumberFormat.h"
#include "../core/UndoManager.h"
#include "../core/TableStyle.h"
#include "../core/ConditionalFormatting.h"
#include "../core/SparklineConfig.h"
#include <QFont>
#include <QMessageBox>
#include <QApplication>
#include <QTimer>
#include <QColor>
#include <limits>

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
            default:
                return QVariant();
        }
    }

    switch (role) {
        case Qt::DisplayRole: {
            auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
            const auto& style = cell->getStyle();
            if (style.numberFormat != "General" && !value.toString().isEmpty()) {
                NumberFormatOptions opts;
                opts.type = NumberFormat::typeFromString(style.numberFormat);
                opts.decimalPlaces = style.decimalPlaces;
                opts.useThousandsSeparator = style.useThousandsSeparator;
                opts.currencyCode = style.currencyCode;
                opts.dateFormatId = style.dateFormatId;
                return NumberFormat::format(value.toString(), opts);
            }
            return value;
        }
        case Qt::EditRole: {
            return m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
        }
        case Qt::FontRole: {
            const auto& baseStyle = cell->getStyle();
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
            const auto& baseStyle = cell->getStyle();
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            // Table header row: use header foreground
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                return table->theme.headerFg;
            }
            return QColor(style.foregroundColor);
        }
        case Qt::BackgroundRole: {
            const auto& baseStyle = cell->getStyle();
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            // Check if cell is in a table
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
                QString cellText = m_spreadsheet->getCellValue(addr).toString();
                if (!cellText.isEmpty() && !m_spreadsheet->validateCell(index.row(), index.column(), cellText)) {
                    return QColor(255, 200, 200); // Light red
                }
            }
            return QColor(style.backgroundColor);
        }
        case Qt::TextAlignmentRole: {
            const auto& style = cell->getStyle();

            // Vertical alignment
            int alignment = 0;
            if (style.vAlign == VerticalAlignment::Top) {
                alignment |= Qt::AlignTop;
            } else if (style.vAlign == VerticalAlignment::Bottom) {
                alignment |= Qt::AlignBottom;
            } else {
                alignment |= Qt::AlignVCenter;
            }

            // Horizontal alignment
            if (style.hAlign == HorizontalAlignment::Left) {
                alignment |= Qt::AlignLeft;
            } else if (style.hAlign == HorizontalAlignment::Right) {
                alignment |= Qt::AlignRight;
            } else if (style.hAlign == HorizontalAlignment::Center) {
                alignment |= Qt::AlignHCenter;
            } else {
                // General: right-align numbers, left-align text
                auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
                bool isNumber = false;
                if (value.typeId() == QMetaType::Double || value.typeId() == QMetaType::Int ||
                    value.typeId() == QMetaType::LongLong || value.typeId() == QMetaType::Float) {
                    isNumber = true;
                } else {
                    value.toString().toDouble(&isNumber);
                }
                alignment |= isNumber ? Qt::AlignRight : Qt::AlignLeft;
            }
            return alignment;
        }
        // Custom roles for indent and borders
        case Qt::UserRole + 10: { // Indent level
            return cell->getStyle().indentLevel;
        }
        case Qt::UserRole + 11: { // Border top
            const auto& b = cell->getStyle().borderTop;
            if (b.enabled) return QString("%1,%2").arg(b.width).arg(b.color);
            return QVariant();
        }
        case Qt::UserRole + 12: { // Border bottom
            const auto& b = cell->getStyle().borderBottom;
            if (b.enabled) return QString("%1,%2").arg(b.width).arg(b.color);
            return QVariant();
        }
        case Qt::UserRole + 13: { // Border left
            const auto& b = cell->getStyle().borderLeft;
            if (b.enabled) return QString("%1,%2").arg(b.width).arg(b.color);
            return QVariant();
        }
        case Qt::UserRole + 14: { // Border right
            const auto& b = cell->getStyle().borderRight;
            if (b.enabled) return QString("%1,%2").arg(b.width).arg(b.color);
            return QVariant();
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

    if (!m_suppressUndo) {
        // Single-cell edit: capture before/after for undo
        CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);

        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else {
            m_spreadsheet->setCellValue(addr, value);
        }

        CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<CellEditCommand>(before, after));
    } else {
        // Bulk operation: caller handles undo tracking
        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else {
            m_spreadsheet->setCellValue(addr, value);
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
