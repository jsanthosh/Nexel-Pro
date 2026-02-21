#include "PivotEngine.h"
#include "Spreadsheet.h"
#include "Cell.h"
#include <algorithm>
#include <numeric>

// ============== AggregateAccumulator ==============

void AggregateAccumulator::addValue(double val, const QString& rawVal) {
    sum += val;
    if (val < min) min = val;
    if (val > max) max = val;
    count++;
    if (!rawVal.isEmpty()) distinctValues.insert(rawVal);
}

double AggregateAccumulator::result(AggregationFunction func) const {
    switch (func) {
        case AggregationFunction::Sum: return sum;
        case AggregationFunction::Count: return count;
        case AggregationFunction::Average: return count > 0 ? sum / count : 0.0;
        case AggregationFunction::Min: return count > 0 ? min : 0.0;
        case AggregationFunction::Max: return count > 0 ? max : 0.0;
        case AggregationFunction::CountDistinct: return static_cast<double>(distinctValues.size());
    }
    return 0.0;
}

// ============== PivotValueField ==============

QString PivotValueField::displayName() const {
    QString prefix;
    switch (aggregation) {
        case AggregationFunction::Sum: prefix = "Sum of "; break;
        case AggregationFunction::Count: prefix = "Count of "; break;
        case AggregationFunction::Average: prefix = "Avg of "; break;
        case AggregationFunction::Min: prefix = "Min of "; break;
        case AggregationFunction::Max: prefix = "Max of "; break;
        case AggregationFunction::CountDistinct: prefix = "Distinct "; break;
    }
    return prefix + name;
}

// ============== PivotEngine ==============

PivotEngine::PivotEngine() = default;

void PivotEngine::setSource(std::shared_ptr<Spreadsheet> sheet, const PivotConfig& config) {
    m_sourceSheet = sheet;
    m_config = config;
}

QStringList PivotEngine::detectColumnHeaders(std::shared_ptr<Spreadsheet> sheet,
                                              const CellRange& range) {
    QStringList headers;
    int headerRow = range.getStart().row;
    for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
        QVariant val = sheet->getCellValue(CellAddress(headerRow, c));
        QString text = val.toString().trimmed();
        if (text.isEmpty()) text = CellAddress(headerRow, c).toString();
        headers.append(text);
    }
    return headers;
}

QStringList PivotEngine::getUniqueValues(std::shared_ptr<Spreadsheet> sheet,
                                          const CellRange& range, int columnIndex) {
    std::set<QString> unique;
    int absCol = range.getStart().col + columnIndex;
    for (int r = range.getStart().row + 1; r <= range.getEnd().row; ++r) {
        QVariant val = sheet->getCellValue(CellAddress(r, absCol));
        QString text = val.toString().trimmed();
        if (!text.isEmpty()) unique.insert(text);
    }
    QStringList result;
    for (const auto& v : unique) result.append(v);
    return result;
}

std::vector<PivotEngine::DataRow> PivotEngine::extractSourceData(
    std::shared_ptr<Spreadsheet> sheet,
    const CellRange& range,
    const std::vector<PivotFilterField>& filters) {

    std::vector<DataRow> rows;
    int numCols = range.getColumnCount();
    int startCol = range.getStart().col;

    // Skip header row (first row of range)
    for (int r = range.getStart().row + 1; r <= range.getEnd().row; ++r) {
        DataRow row;
        row.values.resize(numCols);

        bool isEmpty = true;
        for (int c = 0; c < numCols; ++c) {
            QVariant val = sheet->getCellValue(CellAddress(r, startCol + c));
            row.values[c] = val;
            if (!val.toString().trimmed().isEmpty()) isEmpty = false;
        }
        if (isEmpty) continue; // skip completely empty rows

        // Apply filters
        bool passesFilter = true;
        for (const auto& filter : filters) {
            if (filter.selectedValues.isEmpty()) continue; // no filter = all values
            int colIdx = filter.sourceColumnIndex;
            if (colIdx < 0 || colIdx >= numCols) continue;
            QString cellVal = row.values[colIdx].toString().trimmed();
            if (!filter.selectedValues.contains(cellVal)) {
                passesFilter = false;
                break;
            }
        }
        if (!passesFilter) continue;

        rows.push_back(std::move(row));
    }
    return rows;
}

QString PivotEngine::buildKey(const DataRow& row, const std::vector<PivotField>& fields) {
    QStringList parts;
    for (const auto& f : fields) {
        if (f.sourceColumnIndex >= 0 && f.sourceColumnIndex < static_cast<int>(row.values.size())) {
            parts.append(row.values[f.sourceColumnIndex].toString().trimmed());
        } else {
            parts.append("");
        }
    }
    return parts.join("\x1F"); // unit separator as delimiter
}

std::vector<QString> PivotEngine::splitKey(const QString& key, int fieldCount) {
    QStringList parts = key.split("\x1F");
    std::vector<QString> result;
    for (int i = 0; i < fieldCount; ++i) {
        result.push_back(i < parts.size() ? parts[i] : "");
    }
    return result;
}

PivotResult PivotEngine::compute() {
    PivotResult result;
    if (!m_sourceSheet || m_config.valueFields.empty()) return result;

    auto dataRows = extractSourceData(m_sourceSheet, m_config.sourceRange, m_config.filterFields);
    if (dataRows.empty()) return result;

    int numValueFields = static_cast<int>(m_config.valueFields.size());
    bool hasColFields = !m_config.columnFields.empty();

    // Collect unique row and column keys
    std::set<QString> uniqueRowKeys;
    std::set<QString> uniqueColKeys;

    // rowKey -> colKey -> accumulators (one per value field)
    using AccumVec = std::vector<AggregateAccumulator>;
    std::map<QString, std::map<QString, AccumVec>> accumMap;

    // Grand total accumulators
    std::map<QString, AccumVec> rowTotals;   // per row key
    std::map<QString, AccumVec> colTotals;   // per col key
    AccumVec grandTotals(numValueFields);

    for (const auto& row : dataRows) {
        QString rowKey = buildKey(row, m_config.rowFields);
        QString colKey = hasColFields ? buildKey(row, m_config.columnFields) : "";

        uniqueRowKeys.insert(rowKey);
        if (hasColFields) uniqueColKeys.insert(colKey);

        auto& accums = accumMap[rowKey][colKey];
        if (accums.empty()) accums.resize(numValueFields);

        auto& rowTotal = rowTotals[rowKey];
        if (rowTotal.empty()) rowTotal.resize(numValueFields);

        auto& colTotal = colTotals[colKey];
        if (colTotal.empty()) colTotal.resize(numValueFields);

        for (int v = 0; v < numValueFields; ++v) {
            int srcCol = m_config.valueFields[v].sourceColumnIndex;
            if (srcCol < 0 || srcCol >= static_cast<int>(row.values.size())) continue;

            bool ok;
            double val = row.values[srcCol].toDouble(&ok);
            if (!ok) val = 0.0;
            QString rawVal = row.values[srcCol].toString();

            accums[v].addValue(val, rawVal);
            rowTotal[v].addValue(val, rawVal);
            colTotal[v].addValue(val, rawVal);
            grandTotals[v].addValue(val, rawVal);
        }
    }

    // Sort keys
    std::vector<QString> sortedRowKeys(uniqueRowKeys.begin(), uniqueRowKeys.end());
    std::sort(sortedRowKeys.begin(), sortedRowKeys.end());

    std::vector<QString> sortedColKeys;
    if (hasColFields) {
        sortedColKeys.assign(uniqueColKeys.begin(), uniqueColKeys.end());
        std::sort(sortedColKeys.begin(), sortedColKeys.end());
    }

    // Build column labels
    int numRowFields = static_cast<int>(m_config.rowFields.size());
    if (hasColFields) {
        // Each col key × each value field = one result column
        for (const auto& ck : sortedColKeys) {
            auto keyParts = splitKey(ck, static_cast<int>(m_config.columnFields.size()));
            for (int v = 0; v < numValueFields; ++v) {
                std::vector<QString> label = keyParts;
                if (numValueFields > 1) {
                    label.push_back(m_config.valueFields[v].displayName());
                }
                result.columnLabels.push_back(label);
            }
        }
        // Grand total column labels
        if (m_config.showGrandTotalColumn) {
            for (int v = 0; v < numValueFields; ++v) {
                std::vector<QString> label;
                label.push_back("Grand Total");
                if (numValueFields > 1) label.push_back(m_config.valueFields[v].displayName());
                result.columnLabels.push_back(label);
            }
        }
    } else {
        // No column fields: one column per value field
        for (int v = 0; v < numValueFields; ++v) {
            result.columnLabels.push_back({m_config.valueFields[v].displayName()});
        }
    }

    // Build row labels and data
    int numDataCols = static_cast<int>(result.columnLabels.size());

    for (const auto& rk : sortedRowKeys) {
        auto keyParts = splitKey(rk, numRowFields);
        result.rowLabels.push_back(keyParts);

        std::vector<QVariant> rowData(numDataCols);
        int colIdx = 0;

        if (hasColFields) {
            for (const auto& ck : sortedColKeys) {
                auto it = accumMap[rk].find(ck);
                for (int v = 0; v < numValueFields; ++v) {
                    if (it != accumMap[rk].end() && !it->second.empty()) {
                        double val = it->second[v].result(m_config.valueFields[v].aggregation);
                        rowData[colIdx] = val;
                    } else {
                        rowData[colIdx] = 0.0;
                    }
                    colIdx++;
                }
            }
            // Grand total column for this row
            if (m_config.showGrandTotalColumn) {
                for (int v = 0; v < numValueFields; ++v) {
                    double val = rowTotals[rk][v].result(m_config.valueFields[v].aggregation);
                    rowData[colIdx++] = val;
                }
            }
        } else {
            for (int v = 0; v < numValueFields; ++v) {
                // Sum across all col keys (there's only "")
                auto it = accumMap[rk].find("");
                if (it != accumMap[rk].end() && !it->second.empty()) {
                    rowData[colIdx] = it->second[v].result(m_config.valueFields[v].aggregation);
                }
                colIdx++;
            }
        }

        result.data.push_back(rowData);
    }

    // Grand total row
    if (m_config.showGrandTotalRow) {
        result.grandTotalRow.resize(numDataCols);
        int colIdx = 0;
        if (hasColFields) {
            for (const auto& ck : sortedColKeys) {
                for (int v = 0; v < numValueFields; ++v) {
                    result.grandTotalRow[colIdx++] = colTotals[ck][v].result(m_config.valueFields[v].aggregation);
                }
            }
            if (m_config.showGrandTotalColumn) {
                for (int v = 0; v < numValueFields; ++v) {
                    result.grandTotalRow[colIdx++] = grandTotals[v].result(m_config.valueFields[v].aggregation);
                }
            }
        } else {
            for (int v = 0; v < numValueFields; ++v) {
                result.grandTotalRow[colIdx++] = grandTotals[v].result(m_config.valueFields[v].aggregation);
            }
        }
        result.grandTotal = grandTotals.empty() ? QVariant() :
                            grandTotals[0].result(m_config.valueFields[0].aggregation);
    }

    result.numRowHeaderColumns = numRowFields;
    result.numColHeaderRows = hasColFields ? (numValueFields > 1 ? 2 : 1) : 1;
    result.dataStartRow = result.numColHeaderRows;
    result.dataStartCol = numRowFields;

    return result;
}

void PivotEngine::writeToSheet(std::shared_ptr<Spreadsheet> targetSheet,
                                const PivotResult& result,
                                const PivotConfig& config) {
    targetSheet->setAutoRecalculate(false);

    int currentRow = 0;

    // Write filter summary
    for (const auto& filter : config.filterFields) {
        if (filter.selectedValues.isEmpty()) continue;
        targetSheet->setCellValue(CellAddress(currentRow, 0), filter.name + ":");
        auto cell = targetSheet->getCell(CellAddress(currentRow, 0));
        CellStyle style = cell->getStyle();
        style.bold = true;
        cell->setStyle(style);

        targetSheet->setCellValue(CellAddress(currentRow, 1), filter.selectedValues.join(", "));
        currentRow++;
    }
    bool hasFilterRows = currentRow > 0;
    if (hasFilterRows) currentRow++;

    int headerRow = currentRow;
    int dataColStart = result.numRowHeaderColumns;

    // Write row field headers
    for (size_t rf = 0; rf < config.rowFields.size(); ++rf) {
        targetSheet->setCellValue(CellAddress(headerRow, static_cast<int>(rf)), config.rowFields[rf].name);
    }

    // Write column headers (use the last element of each label vector for simplicity)
    for (size_t c = 0; c < result.columnLabels.size(); ++c) {
        int col = dataColStart + static_cast<int>(c);
        const auto& label = result.columnLabels[c];
        QString text = label.empty() ? "" : label.back();
        targetSheet->setCellValue(CellAddress(headerRow, col), text);
    }

    // If multi-level column headers, write the first level above
    bool hasColFields = !config.columnFields.empty();
    if (hasColFields && result.numColHeaderRows > 1 && headerRow > 0) {
        // Write top-level column headers (merged conceptually)
        int colIdx = dataColStart;
        int numValueFields = static_cast<int>(config.valueFields.size());
        for (size_t c = 0; c < result.columnLabels.size(); ++c) {
            const auto& label = result.columnLabels[c];
            if (label.size() > 1) {
                targetSheet->setCellValue(CellAddress(headerRow - 1, colIdx), label[0]);
                auto cell = targetSheet->getCell(CellAddress(headerRow - 1, colIdx));
                CellStyle style = cell->getStyle();
                style.bold = true;
                style.hAlign = HorizontalAlignment::Center;
                style.backgroundColor = "#D6E4F0";
                cell->setStyle(style);
            }
            colIdx++;
        }
    }

    // Style header row
    int totalCols = dataColStart + static_cast<int>(result.columnLabels.size());
    for (int c = 0; c < totalCols; ++c) {
        auto cell = targetSheet->getCell(CellAddress(headerRow, c));
        CellStyle style = cell->getStyle();
        style.backgroundColor = "#4472C4";
        style.foregroundColor = "#FFFFFF";
        style.bold = true;
        style.hAlign = HorizontalAlignment::Center;
        BorderStyle bs;
        bs.enabled = true;
        bs.color = "#2B5797";
        bs.width = 1;
        style.borderBottom = bs;
        cell->setStyle(style);
    }

    currentRow = headerRow + 1;

    // Write data rows
    for (size_t r = 0; r < result.rowLabels.size(); ++r) {
        // Row labels
        for (size_t rf = 0; rf < result.rowLabels[r].size(); ++rf) {
            targetSheet->setCellValue(CellAddress(currentRow, static_cast<int>(rf)),
                                      result.rowLabels[r][rf]);
            auto cell = targetSheet->getCell(CellAddress(currentRow, static_cast<int>(rf)));
            CellStyle style = cell->getStyle();
            style.bold = true;
            cell->setStyle(style);
        }

        // Data values
        for (size_t c = 0; c < result.data[r].size(); ++c) {
            int col = dataColStart + static_cast<int>(c);
            targetSheet->setCellValue(CellAddress(currentRow, col), result.data[r][c]);

            auto cell = targetSheet->getCell(CellAddress(currentRow, col));
            CellStyle style = cell->getStyle();
            style.numberFormat = "Number";
            style.useThousandsSeparator = true;
            style.decimalPlaces = 0;
            style.hAlign = HorizontalAlignment::Right;
            cell->setStyle(style);
        }

        // Banded row coloring
        if (r % 2 == 1) {
            for (int c = 0; c < totalCols; ++c) {
                auto cell = targetSheet->getCell(CellAddress(currentRow, c));
                CellStyle style = cell->getStyle();
                style.backgroundColor = "#D9E2F3";
                cell->setStyle(style);
            }
        }

        currentRow++;
    }

    // Grand total row
    if (config.showGrandTotalRow && !result.grandTotalRow.empty()) {
        targetSheet->setCellValue(CellAddress(currentRow, 0), "Grand Total");
        auto cell = targetSheet->getCell(CellAddress(currentRow, 0));
        CellStyle style = cell->getStyle();
        style.bold = true;
        style.fontSize = 12;
        BorderStyle bs;
        bs.enabled = true;
        bs.color = "#2B5797";
        bs.width = 2;
        style.borderTop = bs;
        cell->setStyle(style);

        for (size_t c = 0; c < result.grandTotalRow.size(); ++c) {
            int col = dataColStart + static_cast<int>(c);
            targetSheet->setCellValue(CellAddress(currentRow, col), result.grandTotalRow[c]);
            auto totalCell = targetSheet->getCell(CellAddress(currentRow, col));
            CellStyle s = totalCell->getStyle();
            s.bold = true;
            s.borderTop = bs;
            s.numberFormat = "Number";
            s.useThousandsSeparator = true;
            s.decimalPlaces = 0;
            s.hAlign = HorizontalAlignment::Right;
            totalCell->setStyle(s);
        }
    }

    // Set reasonable column widths
    for (int c = 0; c < dataColStart; ++c) {
        auto cell = targetSheet->getCell(CellAddress(0, c));
        CellStyle style = cell->getStyle();
        style.columnWidth = 120;
        cell->setStyle(style);
    }
    for (int c = dataColStart; c < totalCols; ++c) {
        auto cell = targetSheet->getCell(CellAddress(0, c));
        CellStyle style = cell->getStyle();
        style.columnWidth = 100;
        cell->setStyle(style);
    }

    targetSheet->setAutoRecalculate(true);
}
