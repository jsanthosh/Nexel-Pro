// End-to-end streaming XLSX benchmark.
// Builds a spreadsheet, exports through ZipStreamWriter, re-imports through
// ZipStreamReader, and reports timing, file size, and peak RSS.
//
// Build:  cmake --build build --target bench_xlsx_streaming
// Run:    ./bench_xlsx_streaming [rows]   (default: 500000)

#include <QCoreApplication>
#include <QFileInfo>
#include <QDir>
#include <QFile>
#include <QElapsedTimer>
#include <QDebug>
#include <chrono>
#include <cstdio>
#include <sys/resource.h>

#include "../src/core/Spreadsheet.h"
#include "../src/services/XlsxService.h"

namespace {

double peakRssMB() {
    struct rusage ru{};
    getrusage(RUSAGE_SELF, &ru);
#if defined(__APPLE__)
    // On macOS ru_maxrss is in bytes.
    return ru.ru_maxrss / (1024.0 * 1024.0);
#else
    // On Linux ru_maxrss is in kilobytes.
    return ru.ru_maxrss / 1024.0;
#endif
}

double fileMB(const QString& path) {
    return QFileInfo(path).size() / (1024.0 * 1024.0);
}

void printSection(const char* title) {
    std::printf("\n== %s ==\n", title);
}

} // namespace

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);

    int rows = 500'000;
    if (argc > 1) {
        bool ok = false;
        const int n = QString::fromLocal8Bit(argv[1]).toInt(&ok);
        if (ok && n > 0) rows = n;
    }
    const int cols = 5;

    std::printf("Streaming XLSX benchmark: %d rows x %d cols\n", rows, cols);
    std::printf("Initial RSS: %.1f MB\n", peakRssMB());

    QString tempPath = QDir::tempPath() + "/nexel_bench_streaming.xlsx";
    QFile::remove(tempPath);

    // ---------- Build ----------
    printSection("Build");
    QElapsedTimer t;
    t.start();
    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setAutoRecalculate(false);
    sheet->setSheetName("Bench");

    // Headers
    sheet->setCellValue({0, 0}, QVariant("id"));
    sheet->setCellValue({0, 1}, QVariant("name"));
    sheet->setCellValue({0, 2}, QVariant("value"));
    sheet->setCellValue({0, 3}, QVariant("category"));
    sheet->setCellValue({0, 4}, QVariant("note"));

    static const QStringList kCategories = {"alpha", "beta", "gamma", "delta", "omega"};
    for (int r = 1; r <= rows; ++r) {
        sheet->setCellValue({r, 0}, QVariant(r));
        sheet->setCellValue({r, 1}, QVariant(QString("row_%1").arg(r)));
        sheet->setCellValue({r, 2}, QVariant(double(r) * 1.5 + 0.25));
        sheet->setCellValue({r, 3}, QVariant(kCategories[r % kCategories.size()]));
        sheet->setCellValue({r, 4}, QVariant(QString("note for row %1").arg(r)));
    }
    std::printf("  built in %lld ms\n", static_cast<long long>(t.elapsed()));
    std::printf("  RSS after build: %.1f MB\n", peakRssMB());

    // ---------- Export ----------
    printSection("Export (streaming writer)");
    t.restart();
    std::vector<std::shared_ptr<Spreadsheet>> sheets = { sheet };
    const bool exportedOk = XlsxService::exportToFile(sheets, tempPath);
    const qint64 exportMs = t.elapsed();
    std::printf("  ok=%s, time=%lld ms\n",
                exportedOk ? "true" : "false",
                static_cast<long long>(exportMs));
    std::printf("  file size: %.1f MB\n", fileMB(tempPath));
    std::printf("  RSS after export: %.1f MB\n", peakRssMB());
    if (!exportedOk) {
        std::printf("EXPORT FAILED — aborting\n");
        return 1;
    }

    // ---------- Free the source — proves the import path is independent ----------
    sheet.reset();
    sheets.clear();

    // ---------- Streaming import ----------
    printSection("Import (streaming reader)");
    int lastReportedRows = 0;
    int progressCallbacks = 0;
    t.restart();
    XlsxImportResult result = XlsxService::importFromFileStreaming(
        tempPath,
        [&](int rowsParsed, int sheetIdx) {
            (void)sheetIdx;
            lastReportedRows = rowsParsed;
            ++progressCallbacks;
        });
    const qint64 importMs = t.elapsed();
    std::printf("  time=%lld ms\n", static_cast<long long>(importMs));
    std::printf("  sheets imported: %zu\n", result.sheets.size());
    std::printf("  progress callbacks fired: %d (last reported %d rows)\n",
                progressCallbacks, lastReportedRows);
    std::printf("  RSS after import: %.1f MB\n", peakRssMB());

    if (result.sheets.empty()) {
        std::printf("IMPORT FAILED — no sheets\n");
        QFile::remove(tempPath);
        return 1;
    }

    // ---------- Spot-check fidelity ----------
    printSection("Spot-check");
    auto& s = result.sheets[0];
    int failures = 0;
    auto checkCell = [&](int r, int c, const QVariant& expected, const char* label) {
        QVariant got = s->getCellValue({r, c});
        if (got != expected) {
            std::printf("  FAIL %s: expected %s got %s\n",
                        label,
                        expected.toString().toUtf8().constData(),
                        got.toString().toUtf8().constData());
            ++failures;
        }
    };
    checkCell(0, 0, QVariant("id"),                                   "headers[0]");
    checkCell(0, 4, QVariant("note"),                                 "headers[4]");
    checkCell(1, 0, QVariant(1),                                       "row 1 id");
    checkCell(1, 2, QVariant(1.0 * 1.5 + 0.25),                       "row 1 value");
    checkCell(rows / 2, 1, QVariant(QString("row_%1").arg(rows / 2)), "middle row name");
    checkCell(rows, 0, QVariant(rows),                                "last row id");
    checkCell(rows, 2, QVariant(double(rows) * 1.5 + 0.25),           "last row value");

    std::printf("  failures: %d\n", failures);

    // ---------- Summary ----------
    printSection("Summary");
    std::printf("  rows: %d\n", rows);
    std::printf("  file: %.1f MB\n", fileMB(tempPath));
    std::printf("  export: %lld ms\n", static_cast<long long>(exportMs));
    std::printf("  import: %lld ms\n", static_cast<long long>(importMs));
    std::printf("  peak RSS: %.1f MB\n", peakRssMB());
    std::printf("  spot-check failures: %d\n", failures);

    QFile::remove(tempPath);
    return failures == 0 ? 0 : 1;
}
