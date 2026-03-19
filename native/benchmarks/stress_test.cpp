// Nexel Pro — Sort/Filter/Formula Stress Test
// Loads a real CSV and tests sort, filter, search, formula for crashes.
// Build: cmake --build build --target stress_test
// Run:   ./stress_test /path/to/large.csv

#include <iostream>
#include <chrono>
#include <iomanip>

#include "../src/core/Spreadsheet.h"
#include "../src/core/ColumnStore.h"
#include "../src/core/FilterEngine.h"
#include "../src/core/FormulaEngine.h"
#include "../src/core/FormulaAST.h"
#include "../src/services/CsvService.h"

#include <QCoreApplication>

using Clock = std::chrono::high_resolution_clock;

static double ms_since(Clock::time_point t) {
    return std::chrono::duration<double, std::milli>(Clock::now() - t).count();
}

static int g_pass = 0, g_fail = 0;

static void check(bool cond, const char* msg) {
    if (cond) { g_pass++; std::cout << "  PASS: " << msg << std::endl; }
    else { g_fail++; std::cout << "  FAIL: " << msg << std::endl; }
}

int main(int argc, char* argv[]) {
    QCoreApplication app(argc, argv);

    if (argc < 2) {
        std::cerr << "Usage: stress_test <csv_file>" << std::endl;
        return 1;
    }

    QString csvPath = QString::fromLocal8Bit(argv[1]);
    std::cout << "==============================================" << std::endl;
    std::cout << "  Nexel Pro — Sort/Filter/Formula Stress Test" << std::endl;
    std::cout << "  File: " << csvPath.toStdString() << std::endl;
    std::cout << "==============================================" << std::endl;

    // --- LOAD CSV ---
    std::cout << "\n--- Load CSV ---" << std::endl;
    auto t0 = Clock::now();
    auto result = CsvService::importProgressive(csvPath, INT_MAX);
    auto sheet = result.spreadsheet;
    check(sheet != nullptr, "Spreadsheet created");
    if (!sheet) return 1;

    int rowCount = sheet->getMaxRow() + 1;
    int colCount = sheet->getMaxColumn() + 1;
    std::cout << "  Rows: " << rowCount << "  Cols: " << colCount
              << "  Time: " << std::fixed << std::setprecision(1) << ms_since(t0) << " ms" << std::endl;
    check(rowCount > 0, "Has rows");

    // --- SUM FORMULA (all rows) ---
    std::cout << "\n--- SUM formula ---" << std::endl;
    t0 = Clock::now();
    auto& engine = sheet->getFormulaEngine();
    QString sumFormula = QString("=SUM(A1:A%1)").arg(rowCount);
    QVariant sumResult = engine.evaluate(sumFormula);
    double sumMs = ms_since(t0);
    std::cout << "  =SUM(A1:A" << rowCount << ") = " << sumResult.toString().toStdString()
              << "  Time: " << sumMs << " ms" << std::endl;
    check(sumResult.isValid(), "SUM valid result");
    check(sumMs < 5000, "SUM < 5 seconds");

    // --- AVERAGE FORMULA ---
    std::cout << "\n--- AVERAGE formula ---" << std::endl;
    t0 = Clock::now();
    QVariant avgResult = engine.evaluate(QString("=AVERAGE(B1:B%1)").arg(rowCount));
    double avgMs = ms_since(t0);
    std::cout << "  Result: " << avgResult.toString().toStdString()
              << "  Time: " << avgMs << " ms" << std::endl;
    check(avgResult.isValid(), "AVERAGE valid result");

    // --- COUNT FORMULA ---
    std::cout << "\n--- COUNT formula ---" << std::endl;
    t0 = Clock::now();
    QVariant countResult = engine.evaluate(QString("=COUNT(A1:A%1)").arg(rowCount));
    double countMs = ms_since(t0);
    std::cout << "  Result: " << countResult.toString().toStdString()
              << "  Time: " << countMs << " ms" << std::endl;
    check(countResult.isValid(), "COUNT valid result");

    // --- DIRECT COLUMNSTORE SUM ---
    std::cout << "\n--- Direct ColumnStore::sumColumn ---" << std::endl;
    t0 = Clock::now();
    auto& cs = sheet->getColumnStore();
    double directSum = cs.sumColumn(0, 0, rowCount - 1);
    double directMs = ms_since(t0);
    std::cout << "  Sum col 0: " << directSum << "  Time: " << directMs << " ms" << std::endl;
    check(directMs < 1000, "Direct sum < 1 second");

    // --- SORT ASCENDING ---
    std::cout << "\n--- Sort ascending (all rows, col 0) ---" << std::endl;
    t0 = Clock::now();
    CellRange sortRange(CellAddress(0, 0), CellAddress(rowCount - 1, colCount - 1));
    sheet->sortRange(sortRange, 0, true);
    double sortAscMs = ms_since(t0);
    QVariant v0 = sheet->getCellValue(CellAddress(0, 0));
    QVariant v1 = sheet->getCellValue(CellAddress(1, 0));
    std::cout << "  First: " << v0.toString().toStdString() << ", " << v1.toString().toStdString()
              << "  Time: " << sortAscMs << " ms" << std::endl;
    check(sortAscMs < 30000, "Sort ascending < 30 seconds");

    // --- SORT DESCENDING ---
    std::cout << "\n--- Sort descending (all rows, col 0) ---" << std::endl;
    t0 = Clock::now();
    sheet->sortRange(sortRange, 0, false);
    double sortDescMs = ms_since(t0);
    v0 = sheet->getCellValue(CellAddress(0, 0));
    v1 = sheet->getCellValue(CellAddress(1, 0));
    std::cout << "  First: " << v0.toString().toStdString() << ", " << v1.toString().toStdString()
              << "  Time: " << sortDescMs << " ms" << std::endl;
    check(sortDescMs < 30000, "Sort descending < 30 seconds");

    // --- FILTER (bitmap) ---
    std::cout << "\n--- FilterEngine bitmap filter ---" << std::endl;
    t0 = Clock::now();
    FilterEngine filter;
    filter.setColumnStore(&cs);
    filter.setRange(0, rowCount - 1);
    filter.applyConditionFilter(0, FilterEngine::Condition::Gt, 50.0);
    double filterMs = ms_since(t0);
    int filtered = filter.filteredRowCount();
    std::cout << "  Rows passing (col0 > 50): " << filtered << " / " << rowCount
              << "  Time: " << filterMs << " ms" << std::endl;
    check(filtered >= 0 && filtered <= rowCount, "Filter count valid");
    check(filterMs < 5000, "Filter < 5 seconds");
    filter.clearAllFilters();

    // --- SEARCH ---
    std::cout << "\n--- Search all cells ---" << std::endl;
    t0 = Clock::now();
    auto searchResults = sheet->searchAllCells("0.5", false, false);
    double searchMs = ms_since(t0);
    std::cout << "  Matches for '0.5': " << searchResults.size()
              << "  Time: " << searchMs << " ms" << std::endl;
    check(searchMs < 10000, "Search < 10 seconds");

    // --- FORMULA AST CACHE ---
    std::cout << "\n--- FormulaAST cache ---" << std::endl;
    auto& pool = FormulaASTPool::instance();
    t0 = Clock::now();
    for (int i = 0; i < 1000; i++) {
        pool.parse(QString("=SUM(A%1:A%2)+B%3").arg(i+1).arg(i+100).arg(i+1));
    }
    double parseMs = ms_since(t0);
    t0 = Clock::now();
    for (int i = 0; i < 100000; i++) {
        pool.parse(QString("=SUM(A%1:A%2)+B%3").arg(i%1000+1).arg(i%1000+100).arg(i%1000+1));
    }
    double cacheMs = ms_since(t0);
    std::cout << "  Parse 1K formulas: " << parseMs << " ms"
              << "  100K cache hits: " << cacheMs << " ms" << std::endl;
    check(cacheMs < 5000, "Cache lookups < 5 seconds");

    // --- RAPID EDITS ---
    std::cout << "\n--- Rapid edits (1000 cells) ---" << std::endl;
    t0 = Clock::now();
    for (int i = 0; i < 1000; i++) {
        sheet->setCellValue(CellAddress(i, 0), QVariant(static_cast<double>(i) * 1.1));
    }
    double editMs = ms_since(t0);
    std::cout << "  1000 edits: " << editMs << " ms" << std::endl;
    check(editMs < 10000, "1000 edits < 10 seconds");

    // --- SUMMARY ---
    std::cout << "\n==============================================" << std::endl;
    std::cout << "  RESULTS: " << g_pass << " passed, " << g_fail << " failed" << std::endl;
    if (g_fail == 0) {
        std::cout << "  ALL TESTS PASSED — no crashes detected" << std::endl;
    } else {
        std::cout << "  SOME TESTS FAILED" << std::endl;
    }
    std::cout << "==============================================" << std::endl;

    return g_fail > 0 ? 1 : 0;
}
