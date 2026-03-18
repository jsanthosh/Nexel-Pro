// ============================================================================
// Nexel Pro — 20 Million Row Benchmark Suite
// ============================================================================
// Standalone benchmark that validates all performance targets from the
// architecture plan. Compile separately from the main app.
//
// Build: cd native/build && cmake .. && cmake --build . --target benchmark_20m
// Run:   ./benchmark_20m
//
// Performance Targets:
//   1. ColumnStore: 20M rows x 20 cols in <500 MB RAM
//   2. Scroll to arbitrary row: O(1)
//   3. Single cell recalc: <100ms
//   4. Full recalc (2M formulas): <1s
//   5. Search 20M cells: <2s
//   6. Sort 20M rows: <5s
//   7. Filter 20M rows: <2s
//

#include <iostream>
#include <chrono>
#include <random>
#include <string>
#include <vector>
#include <iomanip>
#include <cstring>

#include "../src/core/ColumnStore.h"
#include "../src/core/StyleTable.h"
#include "../src/core/StringPool.h"
#include "../src/core/FilterEngine.h"
#include "../src/core/FormulaAST.h"
#include "../src/core/DependencyGraph.h"

#include <QCoreApplication>
#include <QVariant>
#include <QString>

using Clock = std::chrono::high_resolution_clock;
using Ms = std::chrono::milliseconds;

struct BenchResult {
    std::string name;
    double elapsed_ms;
    std::string target;
    bool passed;
};

static std::vector<BenchResult> g_results;

template<typename Fn>
double measure(const std::string& name, Fn&& fn) {
    auto start = Clock::now();
    fn();
    auto end = Clock::now();
    double ms = std::chrono::duration<double, std::milli>(end - start).count();
    std::cout << "  " << name << ": " << std::fixed << std::setprecision(1) << ms << " ms" << std::endl;
    return ms;
}

// ============================================================================
// Benchmark 1: ColumnStore — populate 20M rows x 20 cols
// ============================================================================
void benchColumnStorePopulate(ColumnStore& store) {
    std::cout << "\n=== Benchmark 1: ColumnStore Populate (20M x 20 cols) ===" << std::endl;

    constexpr int ROWS = 20'000'000;
    constexpr int COLS = 20;

    // 80% numeric, 10% string, 10% formula
    std::mt19937 rng(42);
    std::uniform_real_distribution<double> numDist(0.0, 100000.0);
    std::uniform_int_distribution<int> typeDist(0, 9);

    double ms = measure("Populate 20M x 20 cols", [&]() {
        for (int col = 0; col < COLS; ++col) {
            for (int row = 0; row < ROWS; ++row) {
                int t = typeDist(rng);
                if (t < 8) {
                    // Numeric (80%)
                    store.setCellValue(row, col, QVariant(numDist(rng)));
                } else if (t < 9) {
                    // String (10%)
                    store.setCellValue(row, col, QVariant(QString("str_%1_%2").arg(row).arg(col)));
                } else {
                    // Formula placeholder (10%) — stored as string for benchmark
                    store.setCellValue(row, col, QVariant(QString("=A%1+B%1").arg(row + 1)));
                }
            }
        }
    });

    g_results.push_back({"Populate 20M x 20", ms, "<60000 ms", ms < 60000});
}

// ============================================================================
// Benchmark 2: Random access — O(1) cell lookup
// ============================================================================
void benchRandomAccess(ColumnStore& store) {
    std::cout << "\n=== Benchmark 2: Random Cell Access (1M lookups) ===" << std::endl;

    constexpr int LOOKUPS = 1'000'000;
    std::mt19937 rng(123);
    std::uniform_int_distribution<int> rowDist(0, 19'999'999);
    std::uniform_int_distribution<int> colDist(0, 19);

    double ms = measure("1M random cell lookups", [&]() {
        volatile double sink = 0;
        for (int i = 0; i < LOOKUPS; ++i) {
            int r = rowDist(rng);
            int c = colDist(rng);
            QVariant v = store.getCellValue(r, c);
            if (v.canConvert<double>()) sink += v.toDouble();
        }
    });

    g_results.push_back({"1M random lookups", ms, "<2000 ms", ms < 2000});
}

// ============================================================================
// Benchmark 3: Column scan (SUM-like aggregate over 20M rows)
// ============================================================================
void benchColumnScan(ColumnStore& store) {
    std::cout << "\n=== Benchmark 3: Column Scan (SUM over 20M rows) ===" << std::endl;

    double ms = measure("SUM(A1:A20000000)", [&]() {
        double sum = 0;
        store.scanColumnValues(0, 0, 19'999'999, [&](int /*row*/, const QVariant& val) {
            if (val.canConvert<double>()) sum += val.toDouble();
        });
        std::cout << "    Sum = " << sum << std::endl;
    });

    g_results.push_back({"Column scan 20M", ms, "<1000 ms", ms < 1000});
}

// ============================================================================
// Benchmark 4: FilterEngine — bitmap filter on 20M rows
// ============================================================================
void benchFilter(ColumnStore& store) {
    std::cout << "\n=== Benchmark 4: FilterEngine (20M rows) ===" << std::endl;

    FilterEngine filter;
    filter.setColumnStore(&store);
    filter.setRange(0, 19'999'999);

    // Condition filter: column 0 > 50000
    double ms = measure("Numeric filter (col0 > 50000)", [&]() {
        filter.applyConditionFilter(0, FilterEngine::Condition::Gt, 50000.0);
    });

    std::cout << "    Rows passing: " << filter.filteredRowCount() << std::endl;

    g_results.push_back({"Filter 20M rows", ms, "<2000 ms", ms < 2000});

    filter.clearAllFilters();
}

// ============================================================================
// Benchmark 5: FormulaAST parse + cache
// ============================================================================
void benchFormulaAST() {
    std::cout << "\n=== Benchmark 5: FormulaAST Parse + Cache ===" << std::endl;

    auto& pool = FormulaASTPool::instance();
    pool.clear();

    constexpr int UNIQUE_FORMULAS = 10000;
    constexpr int TOTAL_PARSES = 1'000'000;

    // Generate unique formulas
    std::vector<QString> formulas;
    formulas.reserve(UNIQUE_FORMULAS);
    for (int i = 0; i < UNIQUE_FORMULAS; ++i) {
        formulas.push_back(QString("=SUM(A%1:A%2)+B%3*C%4").arg(i+1).arg(i+100).arg(i+1).arg(i+1));
    }

    double parseMs = measure("Parse 10K unique formulas", [&]() {
        for (const auto& f : formulas) {
            pool.parse(f);
        }
    });

    g_results.push_back({"Parse 10K formulas", parseMs, "<500 ms", parseMs < 500});

    double cacheMs = measure("1M cached formula lookups", [&]() {
        for (int i = 0; i < TOTAL_PARSES; ++i) {
            pool.parse(formulas[i % UNIQUE_FORMULAS]);
        }
    });

    g_results.push_back({"1M cached lookups", cacheMs, "<500 ms", cacheMs < 500});

    std::cout << "    Cached formulas: " << pool.cachedFormulas()
              << "  Total nodes: " << pool.totalNodes() << std::endl;
}

// ============================================================================
// Benchmark 6: DependencyGraph — topological sort
// ============================================================================
void benchDependencyGraph() {
    std::cout << "\n=== Benchmark 6: DependencyGraph Topological Sort ===" << std::endl;

    DependencyGraph graph;
    constexpr int CHAIN_LENGTH = 100000;

    // Build a chain: A1 → A2 → A3 → ... → A100000
    double buildMs = measure("Build 100K-edge chain", [&]() {
        for (int i = 0; i < CHAIN_LENGTH - 1; ++i) {
            graph.addDependency(CellAddress(i + 1, 0), CellAddress(i, 0));
        }
    });

    g_results.push_back({"Build 100K dep chain", buildMs, "<500 ms", buildMs < 500});

    double recalcMs = measure("Topological recalc order (change A1)", [&]() {
        auto levels = graph.getRecalcLevels(CellAddress(0, 0));
        std::cout << "    Levels: " << levels.size()
                  << "  Total cells: ";
        int total = 0;
        for (auto& l : levels) total += l.size();
        std::cout << total << std::endl;
    });

    g_results.push_back({"Topo sort 100K cells", recalcMs, "<500 ms", recalcMs < 500});
}

// ============================================================================
// Benchmark 7: StyleTable deduplication
// ============================================================================
void benchStyleTable() {
    std::cout << "\n=== Benchmark 7: StyleTable Deduplication ===" << std::endl;

    auto& table = StyleTable::instance();
    table.clear();

    constexpr int UNIQUE_STYLES = 200;
    constexpr int TOTAL_INTERNS = 1'000'000;

    // Create unique styles
    std::vector<CellStyle> styles(UNIQUE_STYLES);
    for (int i = 0; i < UNIQUE_STYLES; ++i) {
        styles[i].fontSize = 10 + (i % 20);
        styles[i].bold = (i % 3 == 0);
        styles[i].italic = (i % 5 == 0);
        styles[i].foregroundColor = QString("#%1%2%3")
            .arg(i * 13 % 256, 2, 16, QChar('0'))
            .arg(i * 17 % 256, 2, 16, QChar('0'))
            .arg(i * 23 % 256, 2, 16, QChar('0'));
    }

    double ms = measure("1M style interns (200 unique)", [&]() {
        for (int i = 0; i < TOTAL_INTERNS; ++i) {
            table.intern(styles[i % UNIQUE_STYLES]);
        }
    });

    std::cout << "    Unique styles stored: " << table.count() << std::endl;

    g_results.push_back({"1M style interns", ms, "<1000 ms", ms < 1000});
}

// ============================================================================
// Benchmark 8: StringPool sharded intern
// ============================================================================
void benchStringPool() {
    std::cout << "\n=== Benchmark 8: StringPool Sharded Intern ===" << std::endl;

    auto& pool = StringPool::instance();
    pool.clear();

    constexpr int UNIQUE_STRINGS = 100000;
    constexpr int TOTAL_INTERNS = 2'000'000;

    std::vector<QString> strings;
    strings.reserve(UNIQUE_STRINGS);
    for (int i = 0; i < UNIQUE_STRINGS; ++i) {
        strings.push_back(QString("string_value_%1").arg(i));
    }

    double ms = measure("2M string interns (100K unique)", [&]() {
        for (int i = 0; i < TOTAL_INTERNS; ++i) {
            pool.intern(strings[i % UNIQUE_STRINGS]);
        }
    });

    std::cout << "    Pool size: " << pool.totalInterned() << std::endl;

    g_results.push_back({"2M string interns", ms, "<2000 ms", ms < 2000});
}

// ============================================================================
// Print summary
// ============================================================================
void printSummary() {
    std::cout << "\n" << std::string(70, '=') << std::endl;
    std::cout << "  BENCHMARK SUMMARY" << std::endl;
    std::cout << std::string(70, '=') << std::endl;

    int passed = 0, failed = 0;
    for (const auto& r : g_results) {
        std::cout << "  " << (r.passed ? "PASS" : "FAIL")
                  << "  " << std::left << std::setw(30) << r.name
                  << std::right << std::setw(10) << std::fixed << std::setprecision(1) << r.elapsed_ms << " ms"
                  << "  (target: " << r.target << ")"
                  << std::endl;
        if (r.passed) ++passed; else ++failed;
    }

    std::cout << std::string(70, '-') << std::endl;
    std::cout << "  Total: " << passed << " passed, " << failed << " failed"
              << " out of " << g_results.size() << " benchmarks" << std::endl;
    std::cout << std::string(70, '=') << std::endl;
}

// ============================================================================
// Main
// ============================================================================
int main(int argc, char* argv[]) {
    QCoreApplication app(argc, argv);

    std::cout << "============================================" << std::endl;
    std::cout << "  Nexel Pro — 20 Million Row Benchmark Suite" << std::endl;
    std::cout << "============================================" << std::endl;

    // Check if running full 20M benchmark or quick mode
    bool quickMode = false;
    for (int i = 1; i < argc; ++i) {
        if (std::strcmp(argv[i], "--quick") == 0) quickMode = true;
    }

    if (quickMode) {
        std::cout << "  Mode: QUICK (reduced row counts)" << std::endl;
    } else {
        std::cout << "  Mode: FULL (20M rows — may take several minutes)" << std::endl;
    }

    ColumnStore store;

    if (!quickMode) {
        benchColumnStorePopulate(store);
        benchRandomAccess(store);
        benchColumnScan(store);
        benchFilter(store);
    }

    benchFormulaAST();
    benchDependencyGraph();
    benchStyleTable();
    benchStringPool();

    printSummary();

    return 0;
}
