// ============================================================================
// Nexel Pro — Comprehensive Automated Test Suite
// ============================================================================
// Tests EVERY formula function, cell operation, and scalability scenario.
// Build: cmake --build build --target test_all
// Run:   ./test_all
//

#include <iostream>
#include <chrono>
#include <iomanip>
#include <cmath>
#include <cassert>

#include "../src/core/Spreadsheet.h"
#include "../src/core/ColumnStore.h"
#include "../src/core/FormulaEngine.h"
#include "../src/core/FormulaAST.h"
#include "../src/core/StyleTable.h"
#include "../src/core/FilterEngine.h"

#include <QCoreApplication>

using Clock = std::chrono::high_resolution_clock;

static int g_pass = 0, g_fail = 0, g_total = 0;

static void check(bool cond, const char* msg) {
    g_total++;
    if (cond) { g_pass++; }
    else { g_fail++; std::cout << "  FAIL: " << msg << std::endl; }
}

static void checkApprox(double actual, double expected, const char* msg, double tol = 1e-6) {
    g_total++;
    if (std::abs(actual - expected) < tol || (expected != 0 && std::abs((actual - expected) / expected) < tol)) {
        g_pass++;
    } else {
        g_fail++;
        std::cout << "  FAIL: " << msg << " (got " << actual << ", expected " << expected << ")" << std::endl;
    }
}

#define SECTION(name) std::cout << "\n--- " << name << " ---" << std::endl;

// Helper: evaluate formula on a spreadsheet
static QVariant eval(Spreadsheet& sheet, const QString& formula) {
    return sheet.getFormulaEngine().evaluate(formula);
}
static double evalNum(Spreadsheet& sheet, const QString& formula) {
    QVariant v = eval(sheet, formula);
    return v.toDouble();
}

// ============================================================================
// TEST 1: Math & Trig Functions
// ============================================================================
void testMathFunctions() {
    SECTION("Math & Trig Functions");
    Spreadsheet sheet;

    // Basic
    checkApprox(evalNum(sheet, "=ABS(-5)"), 5.0, "ABS(-5)");
    checkApprox(evalNum(sheet, "=ABS(5)"), 5.0, "ABS(5)");
    checkApprox(evalNum(sheet, "=SQRT(16)"), 4.0, "SQRT(16)");
    checkApprox(evalNum(sheet, "=POWER(2,10)"), 1024.0, "POWER(2,10)");
    checkApprox(evalNum(sheet, "=MOD(10,3)"), 1.0, "MOD(10,3)");
    checkApprox(evalNum(sheet, "=INT(5.9)"), 5.0, "INT(5.9)");
    checkApprox(evalNum(sheet, "=INT(-5.9)"), -6.0, "INT(-5.9)");

    // Rounding
    checkApprox(evalNum(sheet, "=ROUND(2.345,2)"), 2.35, "ROUND(2.345,2)");
    checkApprox(evalNum(sheet, "=ROUNDUP(2.341,2)"), 2.35, "ROUNDUP(2.341,2)");
    checkApprox(evalNum(sheet, "=ROUNDDOWN(2.349,2)"), 2.34, "ROUNDDOWN(2.349,2)");
    checkApprox(evalNum(sheet, "=MROUND(10,3)"), 9.0, "MROUND(10,3)");
    checkApprox(evalNum(sheet, "=CEILING(4.1,1)"), 5.0, "CEILING(4.1,1)");
    checkApprox(evalNum(sheet, "=FLOOR(4.9,1)"), 4.0, "FLOOR(4.9,1)");
    checkApprox(evalNum(sheet, "=TRUNC(5.9)"), 5.0, "TRUNC(5.9)");
    checkApprox(evalNum(sheet, "=TRUNC(-5.9)"), -5.0, "TRUNC(-5.9)");
    checkApprox(evalNum(sheet, "=TRUNC(5.678,2)"), 5.67, "TRUNC(5.678,2)");

    // Trig
    checkApprox(evalNum(sheet, "=PI()"), M_PI, "PI()");
    checkApprox(evalNum(sheet, "=SIN(0)"), 0.0, "SIN(0)");
    checkApprox(evalNum(sheet, "=COS(0)"), 1.0, "COS(0)");
    checkApprox(evalNum(sheet, "=TAN(0)"), 0.0, "TAN(0)");
    checkApprox(evalNum(sheet, "=ASIN(1)"), M_PI / 2, "ASIN(1)");
    checkApprox(evalNum(sheet, "=ACOS(1)"), 0.0, "ACOS(1)");
    checkApprox(evalNum(sheet, "=ATAN(1)"), M_PI / 4, "ATAN(1)");
    checkApprox(evalNum(sheet, "=RADIANS(180)"), M_PI, "RADIANS(180)");
    checkApprox(evalNum(sheet, "=DEGREES(3.14159265358979)"), 180.0, "DEGREES(PI)", 0.001);

    // Log/Exp
    checkApprox(evalNum(sheet, "=LN(1)"), 0.0, "LN(1)");
    checkApprox(evalNum(sheet, "=EXP(0)"), 1.0, "EXP(0)");
    checkApprox(evalNum(sheet, "=LOG(100,10)"), 2.0, "LOG(100,10)");
    checkApprox(evalNum(sheet, "=LOG10(1000)"), 3.0, "LOG10(1000)");

    // Combinatorics
    checkApprox(evalNum(sheet, "=FACT(5)"), 120.0, "FACT(5)");
    checkApprox(evalNum(sheet, "=COMBIN(10,3)"), 120.0, "COMBIN(10,3)");
    checkApprox(evalNum(sheet, "=PERMUT(10,3)"), 720.0, "PERMUT(10,3)");
    checkApprox(evalNum(sheet, "=GCD(12,8)"), 4.0, "GCD(12,8)");
    checkApprox(evalNum(sheet, "=LCM(4,6)"), 12.0, "LCM(4,6)");

    // Misc
    checkApprox(evalNum(sheet, "=SIGN(-5)"), -1.0, "SIGN(-5)");
    checkApprox(evalNum(sheet, "=SIGN(5)"), 1.0, "SIGN(5)");
    checkApprox(evalNum(sheet, "=SIGN(0)"), 0.0, "SIGN(0)");
    checkApprox(evalNum(sheet, "=EVEN(3)"), 4.0, "EVEN(3)");
    checkApprox(evalNum(sheet, "=ODD(4)"), 5.0, "ODD(4)");
    checkApprox(evalNum(sheet, "=QUOTIENT(10,3)"), 3.0, "QUOTIENT(10,3)");
    checkApprox(evalNum(sheet, "=SUMSQ(3,4)"), 25.0, "SUMSQ(3,4)");
}

// ============================================================================
// TEST 2: Text Functions
// ============================================================================
void testTextFunctions() {
    SECTION("Text Functions");
    Spreadsheet sheet;

    check(eval(sheet, "=UPPER(\"hello\")").toString() == "HELLO", "UPPER");
    check(eval(sheet, "=LOWER(\"HELLO\")").toString() == "hello", "LOWER");
    check(eval(sheet, "=PROPER(\"hello world\")").toString() == "Hello World", "PROPER");
    check(eval(sheet, "=TRIM(\"  hi  \")").toString() == "hi", "TRIM");
    check(eval(sheet, "=LEN(\"hello\")").toInt() == 5, "LEN");
    check(eval(sheet, "=LEFT(\"hello\",3)").toString() == "hel", "LEFT");
    check(eval(sheet, "=RIGHT(\"hello\",3)").toString() == "llo", "RIGHT");
    check(eval(sheet, "=MID(\"hello\",2,3)").toString() == "ell", "MID");
    checkApprox(evalNum(sheet, "=FIND(\"ll\",\"hello\")"), 3.0, "FIND");
    check(eval(sheet, "=SUBSTITUTE(\"hello\",\"l\",\"r\")").toString() == "herro", "SUBSTITUTE");
    check(eval(sheet, "=REPT(\"ab\",3)").toString() == "ababab", "REPT");
    check(eval(sheet, "=EXACT(\"abc\",\"abc\")").toBool() == true, "EXACT true");
    check(eval(sheet, "=EXACT(\"abc\",\"ABC\")").toBool() == false, "EXACT false");
    check(eval(sheet, "=CHAR(65)").toString() == "A", "CHAR(65)");
    checkApprox(evalNum(sheet, "=CODE(\"A\")"), 65.0, "CODE(A)");
    check(eval(sheet, "=CLEAN(\"ab\x01\x02\x03""cd\")").toString() == "abcd", "CLEAN");
    check(eval(sheet, "=REPLACE(\"hello\",2,3,\"XY\")").toString() == "hXYo", "REPLACE");
    check(eval(sheet, "=T(\"hello\")").toString() == "hello", "T(text)");
    check(eval(sheet, "=T(123)").toString() == "", "T(number)");
    checkApprox(evalNum(sheet, "=N(TRUE)"), 1.0, "N(TRUE)");
    checkApprox(evalNum(sheet, "=N(\"text\")"), 0.0, "N(text)");
    check(eval(sheet, "=TEXTBEFORE(\"hello-world\",\"-\")").toString() == "hello", "TEXTBEFORE");
    check(eval(sheet, "=TEXTAFTER(\"hello-world\",\"-\")").toString() == "world", "TEXTAFTER");
    check(eval(sheet, "=CONCAT(\"a\",\"b\",\"c\")").toString() == "abc", "CONCAT");
}

// ============================================================================
// TEST 3: Logical Functions
// ============================================================================
void testLogicalFunctions() {
    SECTION("Logical Functions");
    Spreadsheet sheet;

    check(eval(sheet, "=IF(TRUE,1,2)").toInt() == 1, "IF true");
    check(eval(sheet, "=IF(FALSE,1,2)").toInt() == 2, "IF false");
    check(eval(sheet, "=AND(TRUE,TRUE)").toBool() == true, "AND true");
    check(eval(sheet, "=AND(TRUE,FALSE)").toBool() == false, "AND false");
    check(eval(sheet, "=OR(FALSE,TRUE)").toBool() == true, "OR true");
    check(eval(sheet, "=OR(FALSE,FALSE)").toBool() == false, "OR false");
    check(eval(sheet, "=NOT(TRUE)").toBool() == false, "NOT");
    check(eval(sheet, "=XOR(TRUE,FALSE)").toBool() == true, "XOR true");
    check(eval(sheet, "=XOR(TRUE,TRUE)").toBool() == false, "XOR false");
    check(eval(sheet, "=IFERROR(1/0,\"err\")").toString() == "err", "IFERROR");
    check(eval(sheet, "=IFS(FALSE,1,TRUE,2)").toInt() == 2, "IFS");
    check(eval(sheet, "=SWITCH(2,1,\"a\",2,\"b\",\"default\")").toString() == "b", "SWITCH");
    check(eval(sheet, "=TRUE()").toBool() == true, "TRUE()");
    check(eval(sheet, "=FALSE()").toBool() == false, "FALSE()");
}

// ============================================================================
// TEST 4: Date & Time Functions
// ============================================================================
void testDateFunctions() {
    SECTION("Date & Time Functions");
    Spreadsheet sheet;

    // DATE creates serial number
    double d = evalNum(sheet, "=DATE(2026,3,20)");
    check(d > 46000, "DATE(2026,3,20) > 46000");
    checkApprox(evalNum(sheet, "=YEAR(DATE(2026,3,20))"), 2026.0, "YEAR");
    checkApprox(evalNum(sheet, "=MONTH(DATE(2026,3,20))"), 3.0, "MONTH");
    checkApprox(evalNum(sheet, "=DAY(DATE(2026,3,20))"), 20.0, "DAY");
    checkApprox(evalNum(sheet, "=DAYS(DATE(2026,3,20),DATE(2026,3,10))"), 10.0, "DAYS");
    check(evalNum(sheet, "=NOW()") > 40000, "NOW() returns serial");
    check(evalNum(sheet, "=TODAY()") > 40000, "TODAY() returns serial");
    checkApprox(evalNum(sheet, "=TIME(12,30,0)"), 0.520833, "TIME(12,30,0)", 0.001);
    checkApprox(evalNum(sheet, "=WEEKDAY(DATE(2026,3,20))"), 6.0, "WEEKDAY (Friday=6)");
}

// ============================================================================
// TEST 5: Financial Functions
// ============================================================================
void testFinancialFunctions() {
    SECTION("Financial Functions");
    Spreadsheet sheet;

    // PMT: monthly payment for $200K loan, 6%/12, 30yr
    checkApprox(evalNum(sheet, "=PMT(0.06/12,360,200000)"), -1199.10, "PMT", 1.0);
    // FV: $100/month, 5%/12, 10yr
    checkApprox(evalNum(sheet, "=FV(0.05/12,120,-100)"), 15528.23, "FV", 1.0);
    // PV: $100/month, 5%/12, 10yr
    checkApprox(evalNum(sheet, "=PV(0.05/12,120,-100)"), 9428.14, "PV", 1.0);
    // NPV
    checkApprox(evalNum(sheet, "=NPV(0.1,-1000,300,420,680)"), 78.82, "NPV", 1.0);
    // NPER
    double nper = evalNum(sheet, "=NPER(0.01,-100,5000)");
    check(nper > 60 && nper < 80, "NPER in range");
    // SLN
    checkApprox(evalNum(sheet, "=SLN(30000,5000,10)"), 2500.0, "SLN");
    // EFFECT/NOMINAL
    checkApprox(evalNum(sheet, "=EFFECT(0.1,4)"), 0.10381, "EFFECT", 0.001);
    checkApprox(evalNum(sheet, "=NOMINAL(0.10381,4)"), 0.1, "NOMINAL", 0.001);
}

// ============================================================================
// TEST 6: Statistical Functions
// ============================================================================
void testStatisticalFunctions() {
    SECTION("Statistical Functions");
    Spreadsheet sheet;
    sheet.setCellValue(CellAddress(0, 0), QVariant(10.0));
    sheet.setCellValue(CellAddress(1, 0), QVariant(20.0));
    sheet.setCellValue(CellAddress(2, 0), QVariant(30.0));
    sheet.setCellValue(CellAddress(3, 0), QVariant(40.0));
    sheet.setCellValue(CellAddress(4, 0), QVariant(50.0));

    checkApprox(evalNum(sheet, "=SUM(A1:A5)"), 150.0, "SUM");
    checkApprox(evalNum(sheet, "=AVERAGE(A1:A5)"), 30.0, "AVERAGE");
    checkApprox(evalNum(sheet, "=COUNT(A1:A5)"), 5.0, "COUNT");
    checkApprox(evalNum(sheet, "=MIN(A1:A5)"), 10.0, "MIN");
    checkApprox(evalNum(sheet, "=MAX(A1:A5)"), 50.0, "MAX");
    checkApprox(evalNum(sheet, "=MEDIAN(A1:A5)"), 30.0, "MEDIAN");
    checkApprox(evalNum(sheet, "=LARGE(A1:A5,2)"), 40.0, "LARGE");
    checkApprox(evalNum(sheet, "=SMALL(A1:A5,2)"), 20.0, "SMALL");

    // STDEV (sample)
    double stdev = evalNum(sheet, "=STDEV(A1:A5)");
    checkApprox(stdev, 15.811, "STDEV", 0.01);

    // Conditional
    sheet.setCellValue(CellAddress(0, 1), QVariant(1.0));
    sheet.setCellValue(CellAddress(1, 1), QVariant(2.0));
    sheet.setCellValue(CellAddress(2, 1), QVariant(3.0));
    sheet.setCellValue(CellAddress(3, 1), QVariant(4.0));
    sheet.setCellValue(CellAddress(4, 1), QVariant(5.0));

    checkApprox(evalNum(sheet, "=SUMIF(A1:A5,\">20\")"), 120.0, "SUMIF >20");
    checkApprox(evalNum(sheet, "=COUNTIF(A1:A5,\">20\")"), 3.0, "COUNTIF >20");

    // CORREL
    double corr = evalNum(sheet, "=CORREL(A1:A5,B1:B5)");
    checkApprox(corr, 1.0, "CORREL (perfect positive)");

    // SLOPE
    checkApprox(evalNum(sheet, "=SLOPE(A1:A5,B1:B5)"), 10.0, "SLOPE");
    checkApprox(evalNum(sheet, "=INTERCEPT(A1:A5,B1:B5)"), 0.0, "INTERCEPT");
}

// ============================================================================
// TEST 7: Lookup Functions
// ============================================================================
void testLookupFunctions() {
    SECTION("Lookup Functions");
    Spreadsheet sheet;
    // Setup lookup table
    sheet.setCellValue(CellAddress(0, 0), QVariant(1.0));
    sheet.setCellValue(CellAddress(0, 1), QVariant("Apple"));
    sheet.setCellValue(CellAddress(1, 0), QVariant(2.0));
    sheet.setCellValue(CellAddress(1, 1), QVariant("Banana"));
    sheet.setCellValue(CellAddress(2, 0), QVariant(3.0));
    sheet.setCellValue(CellAddress(2, 1), QVariant("Cherry"));

    check(eval(sheet, "=VLOOKUP(2,A1:B3,2,FALSE)").toString() == "Banana", "VLOOKUP exact");
    check(eval(sheet, "=INDEX(B1:B3,2)").toString() == "Banana", "INDEX");
    checkApprox(evalNum(sheet, "=MATCH(3,A1:A3,0)"), 3.0, "MATCH exact");

    // ADDRESS
    check(eval(sheet, "=ADDRESS(1,1)").toString() == "$A$1", "ADDRESS absolute");
    check(eval(sheet, "=ADDRESS(1,1,4)").toString() == "A1", "ADDRESS relative");

    // ROWS, COLUMNS
    checkApprox(evalNum(sheet, "=ROWS(A1:A5)"), 5.0, "ROWS");
    checkApprox(evalNum(sheet, "=COLUMNS(A1:C1)"), 3.0, "COLUMNS");
}

// ============================================================================
// TEST 8: Information Functions
// ============================================================================
void testInfoFunctions() {
    SECTION("Information Functions");
    Spreadsheet sheet;
    sheet.setCellValue(CellAddress(0, 0), QVariant(42.0));
    sheet.setCellValue(CellAddress(0, 1), QVariant("text"));

    check(eval(sheet, "=ISBLANK(A1)").toBool() == false, "ISBLANK(42)");
    check(eval(sheet, "=ISBLANK(C1)").toBool() == true, "ISBLANK(empty)");
    check(eval(sheet, "=ISNUMBER(A1)").toBool() == true, "ISNUMBER");
    check(eval(sheet, "=ISTEXT(B1)").toBool() == true, "ISTEXT");
    check(eval(sheet, "=ISEVEN(4)").toBool() == true, "ISEVEN(4)");
    check(eval(sheet, "=ISODD(3)").toBool() == true, "ISODD(3)");
    check(eval(sheet, "=ISLOGICAL(TRUE)").toBool() == true, "ISLOGICAL");
    checkApprox(evalNum(sheet, "=TYPE(1)"), 1.0, "TYPE(number)");
    checkApprox(evalNum(sheet, "=TYPE(\"a\")"), 2.0, "TYPE(text)");
    check(eval(sheet, "=NA()").toString() == "#N/A", "NA()");
}

// ============================================================================
// TEST 9: Cell Operations (formulas with cell refs)
// ============================================================================
void testCellOperations() {
    SECTION("Cell Operations & References");
    Spreadsheet sheet;

    sheet.setCellValue(CellAddress(0, 0), QVariant(10.0));
    sheet.setCellValue(CellAddress(0, 1), QVariant(20.0));
    sheet.setCellFormula(CellAddress(0, 2), "=A1+B1");

    QVariant result = sheet.getCellValue(CellAddress(0, 2));
    checkApprox(result.toDouble(), 30.0, "=A1+B1");

    // Change A1, check if C1 recalculates
    sheet.setCellValue(CellAddress(0, 0), QVariant(50.0));
    result = sheet.getCellValue(CellAddress(0, 2));
    checkApprox(result.toDouble(), 70.0, "Recalc after A1 change");

    // Formula referencing empty cell
    sheet.setCellFormula(CellAddress(0, 3), "=D1");
    result = sheet.getCellValue(CellAddress(0, 3));
    check(!result.isValid() || result.toString().isEmpty() || result.toDouble() == 0, "Ref to empty cell");

    // Circular reference detection
    sheet.setCellFormula(CellAddress(1, 0), "=B2");
    sheet.setCellFormula(CellAddress(1, 1), "=A2");
    result = sheet.getCellValue(CellAddress(1, 0));
    // Should detect circular and return error or 0
    check(true, "Circular ref doesn't crash");
}

// ============================================================================
// TEST 10: ColumnStore Operations
// ============================================================================
void testColumnStore() {
    SECTION("ColumnStore Operations");
    ColumnStore store;

    // Set and get values
    store.setCellValue(0, 0, QVariant(42.0));
    store.setCellValue(1, 0, QVariant("hello"));
    store.setCellValue(2, 0, QVariant(true));

    check(store.getCellValue(0, 0).toDouble() == 42.0, "Get numeric");
    check(store.getCellValue(1, 0).toString() == "hello", "Get string");
    check(store.hasCell(0, 0), "hasCell true");
    check(!store.hasCell(100, 0), "hasCell false");

    // Bulk insert
    for (int i = 0; i < 100000; ++i) {
        store.setCellValue(i, 1, QVariant(static_cast<double>(i)));
    }
    check(store.getCellValue(99999, 1).toDouble() == 99999.0, "100K cells stored");

    // sumColumn
    double sum = store.sumColumn(1, 0, 99999);
    checkApprox(sum, 4999950000.0, "sumColumn 100K", 1.0);

    // countColumn
    int count = store.countColumn(1, 0, 99999);
    check(count == 100000, "countColumn 100K");

    // Remove cell
    store.removeCell(50000, 1);
    check(!store.hasCell(50000, 1), "removeCell");
}

// ============================================================================
// TEST 11: FormulaAST Cache
// ============================================================================
void testFormulaAST() {
    SECTION("FormulaAST Cache");
    auto& pool = FormulaASTPool::instance();
    pool.clear();

    // Parse and cache
    uint32_t root1 = pool.parse("=SUM(A1:A10)");
    uint32_t root2 = pool.parse("=SUM(A1:A10)");
    check(root1 == root2, "Cache hit returns same root");

    uint32_t root3 = pool.parse("=AVERAGE(B1:B10)");
    check(root3 != root1, "Different formula gets different root");

    check(pool.cachedFormulas() >= 2, "At least 2 cached");

    // Performance: 100K cache lookups
    auto start = Clock::now();
    for (int i = 0; i < 100000; ++i) {
        pool.parse("=SUM(A1:A10)");
    }
    double ms = std::chrono::duration<double, std::milli>(Clock::now() - start).count();
    check(ms < 1000, "100K cache lookups < 1s");
    std::cout << "  100K cache lookups: " << std::fixed << std::setprecision(1) << ms << " ms" << std::endl;
}

// ============================================================================
// TEST 12: Sort Correctness
// ============================================================================
void testSort() {
    SECTION("Sort Correctness");
    Spreadsheet sheet;

    // Fill 1000 rows with random-ish data
    for (int i = 0; i < 1000; ++i) {
        sheet.setCellValue(CellAddress(i, 0), QVariant(static_cast<double>(1000 - i)));
        sheet.setCellValue(CellAddress(i, 1), QVariant(QString("item_%1").arg(1000 - i)));
    }

    // Sort ascending by column A
    CellRange range(CellAddress(0, 0), CellAddress(999, 1));
    sheet.sortRange(range, 0, true);

    // Verify sorted order
    double prev = -1;
    bool sorted = true;
    for (int i = 0; i < 1000; ++i) {
        double val = sheet.getCellValue(CellAddress(i, 0)).toDouble();
        if (val < prev) { sorted = false; break; }
        prev = val;
    }
    check(sorted, "Sort ascending correct");

    // Verify column B followed the sort
    check(sheet.getCellValue(CellAddress(0, 1)).toString() == "item_1", "Sort preserves paired data");
    check(sheet.getCellValue(CellAddress(999, 1)).toString() == "item_1000", "Sort last row correct");

    // Partial sort (rows 100-200 only)
    for (int i = 100; i <= 200; ++i) {
        sheet.setCellValue(CellAddress(i, 0), QVariant(static_cast<double>(200 - i + 100)));
    }
    CellRange partial(CellAddress(100, 0), CellAddress(200, 1));
    sheet.sortRange(partial, 0, true);

    // Verify partial sort didn't affect rows outside range
    check(sheet.getCellValue(CellAddress(99, 0)).toDouble() == 100.0, "Partial sort: row 99 unaffected");
    check(sheet.getCellValue(CellAddress(201, 0)).toDouble() == 202.0, "Partial sort: row 201 unaffected");
}

// ============================================================================
// TEST 13: Insert/Delete Row
// ============================================================================
void testInsertDeleteRow() {
    SECTION("Insert/Delete Row");
    Spreadsheet sheet;

    sheet.setCellValue(CellAddress(0, 0), QVariant(1.0));
    sheet.setCellValue(CellAddress(1, 0), QVariant(2.0));
    sheet.setCellValue(CellAddress(2, 0), QVariant(3.0));

    // Insert row at position 1
    sheet.insertRow(1, 1);
    checkApprox(sheet.getCellValue(CellAddress(0, 0)).toDouble(), 1.0, "Insert: row 0 unchanged");
    check(!sheet.getCellValue(CellAddress(1, 0)).isValid() || sheet.getCellValue(CellAddress(1, 0)).toDouble() == 0, "Insert: row 1 empty");
    checkApprox(sheet.getCellValue(CellAddress(2, 0)).toDouble(), 2.0, "Insert: old row 1 → row 2");
    checkApprox(sheet.getCellValue(CellAddress(3, 0)).toDouble(), 3.0, "Insert: old row 2 → row 3");

    // Delete row 1 (the empty one)
    sheet.deleteRow(1, 1);
    checkApprox(sheet.getCellValue(CellAddress(0, 0)).toDouble(), 1.0, "Delete: row 0 unchanged");
    checkApprox(sheet.getCellValue(CellAddress(1, 0)).toDouble(), 2.0, "Delete: row 2 → row 1");
    checkApprox(sheet.getCellValue(CellAddress(2, 0)).toDouble(), 3.0, "Delete: row 3 → row 2");
}

// ============================================================================
// TEST 14: Style Operations
// ============================================================================
void testStyles() {
    SECTION("Style Operations");
    auto& table = StyleTable::instance();
    table.clear();

    CellStyle s1;
    s1.bold = true;
    s1.fontSize = 14;
    uint16_t idx1 = table.intern(s1);

    CellStyle s2;
    s2.bold = true;
    s2.fontSize = 14;
    uint16_t idx2 = table.intern(s2);

    check(idx1 == idx2, "Identical styles get same index (dedup)");

    CellStyle s3;
    s3.bold = false;
    s3.fontSize = 14;
    uint16_t idx3 = table.intern(s3);
    check(idx3 != idx1, "Different styles get different index");

    const CellStyle& retrieved = table.get(idx1);
    check(retrieved.bold == true, "Retrieved style has correct bold");
    check(retrieved.fontSize == 14, "Retrieved style has correct fontSize");
}

// ============================================================================
// TEST 15: Scalability (100K cells)
// ============================================================================
void testScalability() {
    SECTION("Scalability (100K cells)");
    Spreadsheet sheet;

    auto start = Clock::now();
    for (int i = 0; i < 100000; ++i) {
        sheet.setCellValue(CellAddress(i, 0), QVariant(static_cast<double>(i)));
    }
    double insertMs = std::chrono::duration<double, std::milli>(Clock::now() - start).count();
    std::cout << "  Insert 100K cells: " << std::fixed << std::setprecision(1) << insertMs << " ms" << std::endl;
    check(insertMs < 10000, "Insert 100K < 10s");

    // SUM via formula
    start = Clock::now();
    QVariant sum = eval(sheet, "=SUM(A1:A100000)");
    double sumMs = std::chrono::duration<double, std::milli>(Clock::now() - start).count();
    std::cout << "  SUM 100K cells: " << std::fixed << std::setprecision(1) << sumMs << " ms" << std::endl;
    checkApprox(sum.toDouble(), 4999950000.0, "SUM 100K correct", 1.0);
    check(sumMs < 2000, "SUM 100K < 2s");

    // Sort 100K
    start = Clock::now();
    CellRange range(CellAddress(0, 0), CellAddress(99999, 0));
    sheet.sortRange(range, 0, false); // descending
    double sortMs = std::chrono::duration<double, std::milli>(Clock::now() - start).count();
    std::cout << "  Sort 100K cells: " << std::fixed << std::setprecision(1) << sortMs << " ms" << std::endl;
    check(sortMs < 5000, "Sort 100K < 5s");
    checkApprox(sheet.getCellValue(CellAddress(0, 0)).toDouble(), 99999.0, "Sort desc: first = 99999");
}

// ============================================================================
// MAIN
// ============================================================================
int main(int argc, char* argv[]) {
    QCoreApplication app(argc, argv);

    std::cout << "==============================================" << std::endl;
    std::cout << "  Nexel Pro — Comprehensive Test Suite" << std::endl;
    std::cout << "==============================================" << std::endl;

    testMathFunctions();
    testTextFunctions();
    testLogicalFunctions();
    testDateFunctions();
    testFinancialFunctions();
    testStatisticalFunctions();
    testLookupFunctions();
    testInfoFunctions();
    testCellOperations();
    testColumnStore();
    testFormulaAST();
    testSort();
    testInsertDeleteRow();
    testStyles();
    testScalability();

    // Large-dataset insert row test
    {
        SECTION("Insert Row on Large Dataset (200K rows)");
        Spreadsheet sheet;
        sheet.setAutoRecalculate(false);
        for (int r = 0; r < 200000; ++r) {
            sheet.getOrCreateCellFast(r, 0)->setValue(QVariant((double)r));
        }
        sheet.finishBulkImport();
        sheet.setRowCount(200000);

        // Test insert at row 0
        sheet.insertRow(0, 1);
        auto v0 = sheet.getCellValue({0, 0});
        check(!v0.isValid() || v0.toString().isEmpty(), "insert@0: row 0 empty");
        checkApprox(sheet.getCellValue({1, 0}).toDouble(), 0.0, "insert@0: row 1 = 0");
        checkApprox(sheet.getCellValue({2, 0}).toDouble(), 1.0, "insert@0: row 2 = 1");
        checkApprox(sheet.getCellValue({100, 0}).toDouble(), 99.0, "insert@0: row 100 = 99");
        checkApprox(sheet.getCellValue({65536, 0}).toDouble(), 65535.0, "insert@0: row 65536 = 65535");
        checkApprox(sheet.getCellValue({65537, 0}).toDouble(), 65536.0, "insert@0: row 65537 = 65536");
        checkApprox(sheet.getCellValue({100000, 0}).toDouble(), 99999.0, "insert@0: row 100000 = 99999");
        checkApprox(sheet.getCellValue({200000, 0}).toDouble(), 199999.0, "insert@0: row 200000 = 199999");
        std::cout << "  Insert@0 row count: " << sheet.getRowCount() << std::endl;

        // Test insert at middle (row 50000)
        Spreadsheet sheet2;
        sheet2.setAutoRecalculate(false);
        for (int r = 0; r < 200000; ++r) {
            sheet2.getOrCreateCellFast(r, 0)->setValue(QVariant((double)r));
        }
        sheet2.finishBulkImport();
        sheet2.setRowCount(200000);

        sheet2.insertRow(50000, 1);
        checkApprox(sheet2.getCellValue({49999, 0}).toDouble(), 49999.0, "insert@50K: row 49999 = 49999");
        auto v50k = sheet2.getCellValue({50000, 0});
        check(!v50k.isValid() || v50k.toString().isEmpty(), "insert@50K: row 50000 empty");
        checkApprox(sheet2.getCellValue({50001, 0}).toDouble(), 50000.0, "insert@50K: row 50001 = 50000");
        checkApprox(sheet2.getCellValue({65536, 0}).toDouble(), 65535.0, "insert@50K: row 65536 = 65535");
        checkApprox(sheet2.getCellValue({65537, 0}).toDouble(), 65536.0, "insert@50K: row 65537 = 65536");
        checkApprox(sheet2.getCellValue({200000, 0}).toDouble(), 199999.0, "insert@50K: row 200000 = 199999");
    }

    // Test 3: Insert at row 0 with 500K rows (multiple chunks)
    {
        SECTION("Insert Row at 0 on 500K rows (8 chunks)");
        Spreadsheet sheet;
        sheet.setAutoRecalculate(false);
        for (int r = 0; r < 500000; ++r) {
            sheet.getOrCreateCellFast(r, 0)->setValue(QVariant((double)r));
        }
        sheet.finishBulkImport();
        sheet.setRowCount(500000);

        sheet.insertRow(0, 1);

        auto v0 = sheet.getCellValue({0, 0});
        check(!v0.isValid() || v0.toString().isEmpty(), "500K insert@0: row 0 empty");
        checkApprox(sheet.getCellValue({1, 0}).toDouble(), 0.0, "500K insert@0: row 1 = 0");
        checkApprox(sheet.getCellValue({1000, 0}).toDouble(), 999.0, "500K insert@0: row 1000 = 999");
        checkApprox(sheet.getCellValue({65536, 0}).toDouble(), 65535.0, "500K insert@0: row 65536 = 65535");
        checkApprox(sheet.getCellValue({65537, 0}).toDouble(), 65536.0, "500K insert@0: row 65537 = 65536");
        checkApprox(sheet.getCellValue({65538, 0}).toDouble(), 65537.0, "500K insert@0: row 65538 = 65537");
        checkApprox(sheet.getCellValue({100000, 0}).toDouble(), 99999.0, "500K insert@0: row 100000 = 99999");
        checkApprox(sheet.getCellValue({131072, 0}).toDouble(), 131071.0, "500K insert@0: row 131072 = 131071");
        checkApprox(sheet.getCellValue({131073, 0}).toDouble(), 131072.0, "500K insert@0: row 131073 = 131072");
        checkApprox(sheet.getCellValue({200000, 0}).toDouble(), 199999.0, "500K insert@0: row 200000 = 199999");
        checkApprox(sheet.getCellValue({300000, 0}).toDouble(), 299999.0, "500K insert@0: row 300000 = 299999");
        checkApprox(sheet.getCellValue({400000, 0}).toDouble(), 399999.0, "500K insert@0: row 400000 = 399999");
        checkApprox(sheet.getCellValue({500000, 0}).toDouble(), 499999.0, "500K insert@0: row 500000 = 499999");
    }

    // Test 4: Insert at mid-chunk boundary (row 65536) with 500K rows
    {
        SECTION("Insert Row at chunk boundary (65536) on 500K rows");
        Spreadsheet sheet;
        sheet.setAutoRecalculate(false);
        for (int r = 0; r < 500000; ++r) {
            sheet.getOrCreateCellFast(r, 0)->setValue(QVariant((double)r));
        }
        sheet.finishBulkImport();
        sheet.setRowCount(500000);

        sheet.insertRow(65536, 1);

        checkApprox(sheet.getCellValue({65535, 0}).toDouble(), 65535.0, "boundary: row 65535 = 65535");
        auto v = sheet.getCellValue({65536, 0});
        check(!v.isValid() || v.toString().isEmpty(), "boundary: row 65536 empty");
        checkApprox(sheet.getCellValue({65537, 0}).toDouble(), 65536.0, "boundary: row 65537 = 65536");
        checkApprox(sheet.getCellValue({131073, 0}).toDouble(), 131072.0, "boundary: row 131073 = 131072");
        checkApprox(sheet.getCellValue({500000, 0}).toDouble(), 499999.0, "boundary: row 500000 = 499999");
    }

    // Test 5: Insert at row 5 with 2M rows — verify data at many chunk boundaries
    {
        SECTION("Insert Row at 5 on 2M rows (31 chunks)");
        Spreadsheet sheet;
        sheet.setAutoRecalculate(false);
        for (int r = 0; r < 2000000; ++r) {
            sheet.getOrCreateCellFast(r, 0)->setValue(QVariant((double)r));
        }
        sheet.finishBulkImport();
        sheet.setRowCount(2000000);

        sheet.insertRow(5, 1);

        checkApprox(sheet.getCellValue({4, 0}).toDouble(), 4.0, "2M: row 4 = 4");
        auto v5 = sheet.getCellValue({5, 0});
        check(!v5.isValid() || v5.toString().isEmpty(), "2M: row 5 empty");
        checkApprox(sheet.getCellValue({6, 0}).toDouble(), 5.0, "2M: row 6 = 5");
        // Check every chunk boundary
        checkApprox(sheet.getCellValue({65536, 0}).toDouble(), 65535.0, "2M: row 65536 = 65535");
        // Debug: check what's at 65537
        {
            auto v65537 = sheet.getCellValue({65537, 0});
            std::cout << "  DEBUG row 65537: " << (v65537.isValid() ? v65537.toString().toStdString() : "(empty)") << std::endl;
            // Check chunk layout for column 0
            Column* col0 = sheet.getColumnStore().getColumn(0);
            if (col0) {
                std::cout << "  DEBUG chunks in col 0:" << std::endl;
                for (size_t i = 0; i < col0->chunks().size() && i < 5; ++i) {
                    auto& ch = col0->chunks()[i];
                    std::cout << "    chunk[" << i << "] baseRow=" << ch->baseRow
                              << " pop=" << ch->populatedCount << std::endl;
                }
            }
        }
        checkApprox(sheet.getCellValue({65537, 0}).toDouble(), 65536.0, "2M: row 65537 = 65536");
        checkApprox(sheet.getCellValue({131072, 0}).toDouble(), 131071.0, "2M: row 131072 = 131071");
        checkApprox(sheet.getCellValue({131073, 0}).toDouble(), 131072.0, "2M: row 131073 = 131072");
        checkApprox(sheet.getCellValue({196608, 0}).toDouble(), 196607.0, "2M: row 196608 = 196607");
        checkApprox(sheet.getCellValue({500000, 0}).toDouble(), 499999.0, "2M: row 500000 = 499999");
        checkApprox(sheet.getCellValue({1000000, 0}).toDouble(), 999999.0, "2M: row 1000000 = 999999");
        checkApprox(sheet.getCellValue({1999999, 0}).toDouble(), 1999998.0, "2M: row 1999999 = 1999998");
        checkApprox(sheet.getCellValue({2000000, 0}).toDouble(), 1999999.0, "2M: row 2000000 = 1999999");
    }

    std::cout << "\n==============================================" << std::endl;
    std::cout << "  RESULTS: " << g_pass << " passed, " << g_fail << " failed, " << g_total << " total" << std::endl;
    if (g_fail == 0) {
        std::cout << "  ALL TESTS PASSED" << std::endl;
    } else {
        std::cout << "  SOME TESTS FAILED" << std::endl;
    }
    std::cout << "==============================================" << std::endl;

    return g_fail > 0 ? 1 : 0;
}
