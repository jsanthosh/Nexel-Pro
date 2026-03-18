#include "CsvService.h"
#include <QFile>
#include <cstring>
#include <cstdlib>
#include <thread>
#include <vector>
#include <algorithm>

namespace {

// Auto-detect delimiter by sampling first few KB (respects quoting)
char detectDelimiter(const char* data, qint64 size) {
    qint64 sampleSize = std::min(size, qint64(8192));
    int counts[4] = {0, 0, 0, 0}; // comma, tab, semicolon, pipe
    bool inQuotes = false;

    for (qint64 i = 0; i < sampleSize; i++) {
        char ch = data[i];
        if (ch == '"') { inQuotes = !inQuotes; continue; }
        if (inQuotes) continue;
        switch (ch) {
            case ',':  counts[0]++; break;
            case '\t': counts[1]++; break;
            case ';':  counts[2]++; break;
            case '|':  counts[3]++; break;
        }
    }

    const char delimiters[] = {',', '\t', ';', '|'};
    int maxIdx = 0;
    for (int i = 1; i < 4; i++) {
        if (counts[i] > counts[maxIdx]) maxIdx = i;
    }
    return (counts[maxIdx] > 0) ? delimiters[maxIdx] : ',';
}

// Parse rows from memory-mapped data into spreadsheet.
// Returns the number of rows parsed. Updates `pos` to resume position.
int parseRows(const char* data, qint64 dataSize, qint64& pos, char delim,
              Spreadsheet* spreadsheet, int startRow, int maxRows,
              int& maxCol) {
    QByteArray fieldBuf;
    fieldBuf.reserve(256);
    int row = startRow;
    int rowsParsed = 0;

    while (pos < dataSize && rowsParsed < maxRows) {
        int col = 0;

        // Parse one row
        while (pos < dataSize) {
            fieldBuf.clear();

            if (pos < dataSize && data[pos] == '"') {
                pos++;
                while (pos < dataSize) {
                    if (data[pos] == '"') {
                        if (pos + 1 < dataSize && data[pos + 1] == '"') {
                            fieldBuf.append('"');
                            pos += 2;
                        } else {
                            pos++;
                            break;
                        }
                    } else {
                        const char* qSearch = static_cast<const char*>(
                            memchr(data + pos, '"', dataSize - pos));
                        if (qSearch) {
                            fieldBuf.append(data + pos, static_cast<qsizetype>(qSearch - data - pos));
                            pos = qSearch - data;
                        } else {
                            fieldBuf.append(data + pos, static_cast<qsizetype>(dataSize - pos));
                            pos = dataSize;
                        }
                    }
                }
                while (pos < dataSize && data[pos] != delim && data[pos] != '\n' && data[pos] != '\r') {
                    pos++;
                }
            } else {
                qint64 start = pos;
                while (pos < dataSize && data[pos] != delim && data[pos] != '\n' && data[pos] != '\r') {
                    pos++;
                }
                if (pos > start) {
                    fieldBuf.append(data + start, static_cast<qsizetype>(pos - start));
                }
            }

            if (!fieldBuf.isEmpty()) {
                const char* fStart = fieldBuf.constData();
                const char* fEnd = fStart + fieldBuf.size();
                while (fStart < fEnd && (*fStart == ' ' || *fStart == '\t')) fStart++;
                while (fEnd > fStart && (*(fEnd - 1) == ' ' || *(fEnd - 1) == '\t')) fEnd--;

                int fLen = static_cast<int>(fEnd - fStart);
                if (fLen > 0) {
                    auto cell = spreadsheet->getOrCreateCellFast(row, col);

                    char firstCh = *fStart;
                    bool isNum = false;
                    if ((firstCh >= '0' && firstCh <= '9') || firstCh == '-' || firstCh == '+' || firstCh == '.') {
                        if (fLen < 63) {
                            char numBuf[64];
                            memcpy(numBuf, fStart, fLen);
                            numBuf[fLen] = '\0';
                            char* endPtr = nullptr;
                            double numValue = strtod(numBuf, &endPtr);
                            if (endPtr == numBuf + fLen) {
                                cell->setValueDirect(numValue);
                                isNum = true;
                            }
                        }
                    }

                    if (!isNum) {
                        QString strValue = QString::fromUtf8(fStart, fLen);
                        if (firstCh == '=') {
                            cell->setFormula(strValue);
                        } else {
                            cell->setValueDirect(strValue);
                        }
                    }
                }
            }

            col++;
            if (pos < dataSize && data[pos] == delim) {
                pos++;
            } else {
                break;
            }
        }

        if (col > maxCol) maxCol = col;
        row++;
        rowsParsed++;

        // Skip line endings
        if (pos < dataSize && data[pos] == '\r') {
            pos++;
            if (pos < dataSize && data[pos] == '\n') pos++;
        } else if (pos < dataSize && data[pos] == '\n') {
            pos++;
        }
    }

    return rowsParsed;
}

} // anonymous namespace

// CsvFileHandle implementation
bool CsvFileHandle::open(const QString& filePath) {
    file.setFileName(filePath);
    if (!file.open(QIODevice::ReadOnly)) return false;

    qint64 fileSize = file.size();
    if (fileSize == 0) return false;

    uchar* mapped = file.map(0, fileSize);
    if (mapped) {
        data = reinterpret_cast<const char*>(mapped);
        dataSize = fileSize;
    } else {
        rawData = file.readAll();
        data = rawData.constData();
        dataSize = rawData.size();
    }

    // Handle BOM
    startOffset = 0;
    if (dataSize >= 2) {
        auto u = reinterpret_cast<const unsigned char*>(data);
        if (dataSize >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) {
            startOffset = 3;
        } else if (u[0] == 0xFF && u[1] == 0xFE) {
            transcoded = QString::fromUtf16(
                reinterpret_cast<const char16_t*>(data + 2),
                static_cast<qsizetype>((dataSize - 2) / 2)).toUtf8();
            data = transcoded.constData();
            dataSize = transcoded.size();
            startOffset = 0;
        } else if (u[0] == 0xFE && u[1] == 0xFF) {
            QByteArray swapped(static_cast<qsizetype>(dataSize - 2), Qt::Uninitialized);
            const char* src = data + 2;
            char* dst = swapped.data();
            for (qint64 i = 0; i + 1 < dataSize - 2; i += 2) {
                dst[i] = src[i + 1];
                dst[i + 1] = src[i];
            }
            transcoded = QString::fromUtf16(
                reinterpret_cast<const char16_t*>(swapped.constData()),
                static_cast<qsizetype>(swapped.size() / 2)).toUtf8();
            data = transcoded.constData();
            dataSize = transcoded.size();
            startOffset = 0;
        }
    }

    delimiter = detectDelimiter(data + startOffset, dataSize - startOffset);

    // Count columns from first line
    estimatedCols = 1;
    bool inQ = false;
    for (qint64 p = startOffset; p < dataSize; p++) {
        if (data[p] == '"') { inQ = !inQ; continue; }
        if (inQ) continue;
        if (data[p] == delimiter) estimatedCols++;
        else if (data[p] == '\n' || data[p] == '\r') break;
    }

    return true;
}

std::shared_ptr<Spreadsheet> CsvService::importFromFile(const QString& filePath) {
    auto result = importProgressive(filePath, INT_MAX);
    return result.spreadsheet;
}

CsvProgressiveResult CsvService::importProgressive(const QString& filePath, int initialRows) {
    CsvProgressiveResult result;
    result.filePath = filePath;

    auto fh = std::make_shared<CsvFileHandle>();
    if (!fh->open(filePath)) {
        return result;
    }

    auto spreadsheet = std::make_shared<Spreadsheet>();
    spreadsheet->setAutoRecalculate(false);

    int estimatedRows = static_cast<int>(fh->dataSize / 50) + 100;
    spreadsheet->setRowCount(std::max(1000, estimatedRows));
    size_t estimatedCells = std::min(
        static_cast<size_t>(estimatedRows) * fh->estimatedCols, size_t(50'000'000));
    spreadsheet->reserveCells(estimatedCells);

    result.delimiter = fh->delimiter;
    result.maxCol = 0;

    qint64 pos = fh->startOffset;
    int rowsParsed = parseRows(fh->data, fh->dataSize, pos, fh->delimiter,
                               spreadsheet.get(), 0, initialRows, result.maxCol);

    result.currentRow = rowsParsed;
    result.spreadsheet = spreadsheet;

    if (pos >= fh->dataSize || initialRows == INT_MAX) {
        // All data loaded — use fast finish, skip O(n) scan
        result.resumeOffset = -1;
        spreadsheet->finishBulkImportWithMaxRowCol(
            rowsParsed - 1, result.maxCol - 1);
        spreadsheet->setRowCount(std::max(1000, rowsParsed + 100));
        spreadsheet->setColumnCount(std::max(26, result.maxCol + 10));
        spreadsheet->setAutoRecalculate(true);
        // fileHandle not needed — let it drop
    } else {
        // More data to load — keep file handle alive for continuation
        result.resumeOffset = pos;
        result.fileHandle = fh;
        spreadsheet->setRowCount(std::max(1000, estimatedRows));
        spreadsheet->setColumnCount(std::max(26, result.maxCol + 10));
    }

    return result;
}

int CsvService::continueImport(CsvProgressiveResult& result, int chunkRows) {
    if (result.resumeOffset < 0 || !result.spreadsheet || !result.fileHandle) return 0;

    auto& fh = *result.fileHandle;

    qint64 pos = result.resumeOffset;
    int rowsParsed = parseRows(fh.data, fh.dataSize, pos, result.delimiter,
                               result.spreadsheet.get(), result.currentRow,
                               chunkRows, result.maxCol);

    result.currentRow += rowsParsed;

    if (pos >= fh.dataSize) {
        // Done — use fast finish that skips O(n) cell scan
        result.resumeOffset = -1;
        result.fileHandle.reset();
        result.spreadsheet->finishBulkImportWithMaxRowCol(
            result.currentRow - 1, result.maxCol - 1);
        result.spreadsheet->setRowCount(std::max(1000, result.currentRow + 100));
        result.spreadsheet->setColumnCount(std::max(26, result.maxCol + 10));
        result.spreadsheet->setAutoRecalculate(true);
    } else {
        result.resumeOffset = pos;
    }

    return rowsParsed;
}

bool CsvService::exportToFile(const Spreadsheet& spreadsheet, const QString& filePath) {
    QFile file(filePath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        return false;
    }

    int maxRow = spreadsheet.getMaxRow();
    int maxCol = spreadsheet.getMaxColumn();

    // Write in chunks to avoid huge memory allocation and allow streaming
    constexpr int FLUSH_INTERVAL = 10000; // flush every 10K rows
    QByteArray buffer;
    buffer.reserve(FLUSH_INTERVAL * (maxCol + 1) * 15);

    for (int r = 0; r <= maxRow; ++r) {
        int lastNonEmpty = -1;
        for (int c = maxCol; c >= 0; --c) {
            auto cell = spreadsheet.getCellIfExists(r, c);
            if (cell && cell->getType() != CellType::Empty) {
                lastNonEmpty = c;
                break;
            }
        }

        for (int c = 0; c <= lastNonEmpty; ++c) {
            if (c > 0) buffer.append(',');

            auto cell = spreadsheet.getCellIfExists(r, c);
            if (cell && cell->getType() != CellType::Empty) {
                QVariant value = cell->getValue();
                if (cell->getType() == CellType::Formula) {
                    value = cell->getComputedValue();
                }
                const auto& style = cell->getStyle();
                if (style.numberFormat == "Picklist") {
                    value = value.toString().replace('|', ", ");
                }
                if (style.numberFormat == "Checkbox") {
                    bool checked = value.toBool() || value.toString().toLower() == "true" || value.toString() == "1";
                    value = checked ? "TRUE" : "FALSE";
                }
                QString str = value.toString();
                if (str.contains(',') || str.contains('"') || str.contains('\n')) {
                    str.replace('"', "\"\"");
                    buffer.append('"');
                    buffer.append(str.toUtf8());
                    buffer.append('"');
                } else {
                    buffer.append(str.toUtf8());
                }
            }
        }
        buffer.append('\n');

        // Flush buffer periodically to avoid huge memory usage
        if ((r + 1) % FLUSH_INTERVAL == 0) {
            file.write(buffer);
            buffer.clear();
        }
    }

    // Write remaining
    if (!buffer.isEmpty()) {
        file.write(buffer);
    }

    file.close();
    return true;
}

QStringList CsvService::parseCsvLine(const QString& line) {
    QStringList fields;
    QString field;
    bool inQuotes = false;

    for (int i = 0; i < line.length(); ++i) {
        QChar ch = line[i];

        if (inQuotes) {
            if (ch == '"') {
                if (i + 1 < line.length() && line[i + 1] == '"') {
                    field += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                field += ch;
            }
        } else {
            if (ch == '"') {
                inQuotes = true;
            } else if (ch == ',') {
                fields.append(field);
                field.clear();
            } else {
                field += ch;
            }
        }
    }
    fields.append(field);

    return fields;
}
