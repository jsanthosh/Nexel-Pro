#include "CsvService.h"
#include <QFile>
#include <QApplication>
#include <cstring>
#include <cstdlib>

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
    // If no delimiters found at all, default to comma
    return (counts[maxIdx] > 0) ? delimiters[maxIdx] : ',';
}

} // anonymous namespace

std::shared_ptr<Spreadsheet> CsvService::importFromFile(const QString& filePath) {
    QFile file(filePath);
    if (!file.open(QIODevice::ReadOnly)) {
        return nullptr;
    }

    qint64 fileSize = file.size();
    if (fileSize == 0) {
        return std::make_shared<Spreadsheet>();
    }

    // Memory-map for zero-copy access; fallback to readAll
    QByteArray rawData;
    const char* data;
    qint64 dataSize;

    uchar* mapped = file.map(0, fileSize);
    if (mapped) {
        data = reinterpret_cast<const char*>(mapped);
        dataSize = fileSize;
    } else {
        rawData = file.readAll();
        data = rawData.constData();
        dataSize = rawData.size();
    }

    // Handle BOM and encoding detection
    qint64 offset = 0;
    QByteArray transcoded;

    if (dataSize >= 2) {
        auto u = reinterpret_cast<const unsigned char*>(data);
        if (dataSize >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) {
            // UTF-8 BOM — skip it
            offset = 3;
        } else if (u[0] == 0xFF && u[1] == 0xFE) {
            // UTF-16 LE — transcode to UTF-8
            transcoded = QString::fromUtf16(
                reinterpret_cast<const char16_t*>(data + 2),
                static_cast<qsizetype>((dataSize - 2) / 2)).toUtf8();
            data = transcoded.constData();
            dataSize = transcoded.size();
            offset = 0;
        } else if (u[0] == 0xFE && u[1] == 0xFF) {
            // UTF-16 BE — byte-swap then decode
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
            offset = 0;
        }
    }

    // Auto-detect delimiter from first 8KB
    char delim = detectDelimiter(data + offset, dataSize - offset);


    auto spreadsheet = std::make_shared<Spreadsheet>();
    spreadsheet->setAutoRecalculate(false);

    // Pre-estimate row count from file size (avg ~50 bytes/row)
    int estimatedRows = static_cast<int>(fileSize / 50) + 100;
    spreadsheet->setRowCount(std::max(1000, estimatedRows));
    // Pre-reserve hash map to avoid rehashing during import (~10 cols avg)
    spreadsheet->reserveCells(static_cast<size_t>(estimatedRows) * 10);

    int row = 0;
    int maxCol = 0;
    qint64 pos = offset;
    QByteArray fieldBuf;
    fieldBuf.reserve(256);

    while (pos < dataSize) {
        int col = 0;

        // Parse one row (fields separated by delim, terminated by EOL/EOF)
        while (pos < dataSize) {
            fieldBuf.clear();

            if (pos < dataSize && data[pos] == '"') {
                // === Quoted field ===
                pos++; // skip opening quote
                while (pos < dataSize) {
                    if (data[pos] == '"') {
                        if (pos + 1 < dataSize && data[pos + 1] == '"') {
                            fieldBuf.append('"');
                            pos += 2;
                        } else {
                            pos++; // skip closing quote
                            break;
                        }
                    } else {
                        fieldBuf.append(data[pos]);
                        pos++;
                    }
                }
                // Skip any trailing chars after close-quote before delimiter/EOL (malformed CSV)
                while (pos < dataSize && data[pos] != delim && data[pos] != '\n' && data[pos] != '\r') {
                    pos++;
                }
            } else {
                // === Unquoted field — fast memcpy scan ===
                qint64 start = pos;
                while (pos < dataSize && data[pos] != delim && data[pos] != '\n' && data[pos] != '\r') {
                    pos++;
                }
                if (pos > start) {
                    fieldBuf.append(data + start, static_cast<qsizetype>(pos - start));
                }
            }

            // Set cell value if field is non-empty (inline trim)
            if (!fieldBuf.isEmpty()) {
                const char* fStart = fieldBuf.constData();
                const char* fEnd = fStart + fieldBuf.size();
                while (fStart < fEnd && (*fStart == ' ' || *fStart == '\t')) fStart++;
                while (fEnd > fStart && (*(fEnd - 1) == ' ' || *(fEnd - 1) == '\t')) fEnd--;

                int fLen = static_cast<int>(fEnd - fStart);
                if (fLen > 0) {
                    CellAddress addr(row, col);

                    // Fast numeric detection using strtod with stack buffer
                    bool isNum = false;
                    char firstCh = *fStart;
                    if ((firstCh >= '0' && firstCh <= '9') || firstCh == '-' || firstCh == '+' || firstCh == '.') {
                        if (fLen < 63) {
                            char numBuf[64];
                            memcpy(numBuf, fStart, fLen);
                            numBuf[fLen] = '\0';
                            char* endPtr = nullptr;
                            double numValue = strtod(numBuf, &endPtr);
                            if (endPtr == numBuf + fLen) {
                                spreadsheet->setCellValue(addr, numValue);
                                isNum = true;
                            }
                        }
                    }

                    if (!isNum) {
                        QString strValue = QString::fromUtf8(fStart, fLen);
                        if (firstCh == '=') {
                            spreadsheet->setCellFormula(addr, strValue);
                        } else {
                            spreadsheet->setCellValue(addr, strValue);
                        }
                    }
                }
            }

            col++;

            // Advance past delimiter, or stop at EOL/EOF
            if (pos < dataSize && data[pos] == delim) {
                pos++;
            } else {
                break;
            }
        }

        if (col > maxCol) maxCol = col;
        row++;

        // Skip line endings: \r\n, \r, or \n
        if (pos < dataSize && data[pos] == '\r') {
            pos++;
            if (pos < dataSize && data[pos] == '\n') pos++;
        } else if (pos < dataSize && data[pos] == '\n') {
            pos++;
        }

        // Keep UI responsive every 10K rows
        if (row % 10000 == 0) {
            QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
        }
    }

    // Set final dimensions
    spreadsheet->setRowCount(std::max(1000, row + 100));
    spreadsheet->setColumnCount(std::max(26, maxCol + 10));

    spreadsheet->setAutoRecalculate(true);
    file.close();
    return spreadsheet;
}

bool CsvService::exportToFile(const Spreadsheet& spreadsheet, const QString& filePath) {
    QFile file(filePath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        return false;
    }

    int maxRow = spreadsheet.getMaxRow();
    int maxCol = spreadsheet.getMaxColumn();

    // Pre-allocate a buffer for the entire output
    QByteArray output;
    output.reserve(static_cast<qsizetype>((maxRow + 1) * (maxCol + 1) * 10));

    for (int r = 0; r <= maxRow; ++r) {
        int lastNonEmpty = -1;
        // Find last non-empty column to avoid trailing commas
        for (int c = maxCol; c >= 0; --c) {
            auto cell = spreadsheet.getCellIfExists(r, c);
            if (cell && cell->getType() != CellType::Empty) {
                lastNonEmpty = c;
                break;
            }
        }

        for (int c = 0; c <= lastNonEmpty; ++c) {
            if (c > 0) output.append(',');

            auto cell = spreadsheet.getCellIfExists(r, c);
            if (cell && cell->getType() != CellType::Empty) {
                QVariant value = cell->getValue();
                if (cell->getType() == CellType::Formula) {
                    value = cell->getComputedValue();
                }
                QString str = value.toString();
                if (str.contains(',') || str.contains('"') || str.contains('\n')) {
                    str.replace('"', "\"\"");
                    output.append('"');
                    output.append(str.toUtf8());
                    output.append('"');
                } else {
                    output.append(str.toUtf8());
                }
            }
        }
        output.append('\n');
    }

    file.write(output);
    file.close();
    return true;
}

// RFC 4180 compliant CSV line parser (kept for compatibility)
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
                    i++; // Skip escaped quote
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
