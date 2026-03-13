#ifndef CSVSERVICE_H
#define CSVSERVICE_H

#include <QString>
#include <QStringList>
#include <QFile>
#include <QByteArray>
#include <memory>
#include <functional>
#include <atomic>
#include "../core/Spreadsheet.h"

// Persistent file handle for progressive CSV import (mmap kept alive across chunks)
struct CsvFileHandle {
    QFile file;
    QByteArray rawData;
    QByteArray transcoded;
    const char* data = nullptr;
    qint64 dataSize = 0;
    qint64 startOffset = 0;
    char delimiter = ',';
    int estimatedCols = 1;

    bool open(const QString& filePath);
};

// Result of progressive CSV import
struct CsvProgressiveResult {
    std::shared_ptr<Spreadsheet> spreadsheet;
    std::shared_ptr<CsvFileHandle> fileHandle; // kept alive across chunks
    qint64 resumeOffset = -1;   // byte offset to resume from (-1 = complete)
    char delimiter = ',';
    int maxCol = 0;
    int currentRow = 0;
    QString filePath;           // needed for continuation
};

class CsvService {
public:
    static std::shared_ptr<Spreadsheet> importFromFile(const QString& filePath);
    static bool exportToFile(const Spreadsheet& spreadsheet, const QString& filePath);

    // Progressive import: loads first `initialRows` rows, returns partial result.
    // Call continueImport() to load more rows in chunks.
    static CsvProgressiveResult importProgressive(const QString& filePath, int initialRows = 100000);

    // Continue loading from where importProgressive left off.
    // Returns number of rows loaded in this chunk. Sets result.resumeOffset = -1 when done.
    static int continueImport(CsvProgressiveResult& result, int chunkRows = 200000);

private:
    static QStringList parseCsvLine(const QString& line);
};

#endif // CSVSERVICE_H
