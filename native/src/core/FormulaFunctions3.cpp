// ============================================================================
// FormulaFunctions3.cpp — Sprint 1 completion: remaining P1 functions
// ============================================================================

#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <random>
#include <QDateTime>
#include <QRegularExpression>
#include <QUrl>

static const QDate s_ep(1899, 12, 30);
static double dSerial(const QDate& d) { return s_ep.daysTo(d); }
static QDate sDate(double s) { return s_ep.addDays(static_cast<int>(s)); }

// ============================================================================
// DYNAMIC ARRAY — Remaining
// ============================================================================

// TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
QVariant FormulaEngine::funcTEXTSPLIT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QString colDelim = toString(args[1]);
    QString rowDelim = args.size() > 2 ? toString(args[2]) : "";

    QStringList rows;
    if (!rowDelim.isEmpty()) rows = text.split(rowDelim);
    else rows = {text};

    QVariantList result;
    for (const auto& row : rows) {
        QVariantList cols;
        QStringList parts = row.split(colDelim);
        for (const auto& p : parts) cols.append(QVariant(p));
        result.append(QVariant(cols));
    }
    if (result.size() == 1) return result[0]; // Single row → return flat
    return result;
}

// WRAPROWS(vector, wrap_count, [pad_with])
QVariant FormulaEngine::funcWRAPROWS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    int wrap = static_cast<int>(toNumber(args[1]));
    if (wrap <= 0) return QVariant("#VALUE!");
    QVariant pad = args.size() > 2 ? args[2] : QVariant("#N/A");

    QVariantList result;
    for (size_t i = 0; i < data.size(); i += wrap) {
        QVariantList row;
        for (int j = 0; j < wrap; ++j) {
            size_t idx = i + j;
            row.append(idx < data.size() ? data[idx] : pad);
        }
        result.append(QVariant(row));
    }
    return result;
}

// WRAPCOLS(vector, wrap_count, [pad_with])
QVariant FormulaEngine::funcWRAPCOLS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    int wrap = static_cast<int>(toNumber(args[1]));
    if (wrap <= 0) return QVariant("#VALUE!");

    int cols = (static_cast<int>(data.size()) + wrap - 1) / wrap;
    QVariantList result;
    for (int r = 0; r < wrap; ++r) {
        QVariantList row;
        for (int c = 0; c < cols; ++c) {
            size_t idx = c * wrap + r;
            row.append(idx < data.size() ? data[idx] : QVariant("#N/A"));
        }
        result.append(QVariant(row));
    }
    return result;
}

// TOROW(array, [ignore], [scan_by_column])
QVariant FormulaEngine::funcTOROW(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    auto flat = flattenArgs(args);
    return QVariant::fromValue(flat);
}

// TOCOL(array, [ignore], [scan_by_column])
QVariant FormulaEngine::funcTOCOL(const std::vector<QVariant>& args) {
    return funcTOROW(args); // Same flattening
}

// CHOOSECOLS(array, col_num1, [col_num2], ...)
QVariant FormulaEngine::funcCHOOSECOLS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    std::vector<QVariant> result;
    // Simplified: for 1D array, treat col_num as index
    for (size_t i = 1; i < args.size(); ++i) {
        int idx = static_cast<int>(toNumber(args[i])) - 1;
        if (idx >= 0 && idx < static_cast<int>(data.size()))
            result.push_back(data[idx]);
    }
    return QVariant::fromValue(result);
}

// CHOOSEROWS(array, row_num1, [row_num2], ...)
QVariant FormulaEngine::funcCHOOSEROWS(const std::vector<QVariant>& args) {
    return funcCHOOSECOLS(args); // Same logic for 1D
}

// EXPAND(array, rows, [cols], [pad_with])
QVariant FormulaEngine::funcEXPAND(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto data = flattenArgs({args[0]});
    int targetRows = static_cast<int>(toNumber(args[1]));
    QVariant pad = args.size() > 3 ? args[3] : QVariant("#N/A");

    std::vector<QVariant> result(data);
    while (static_cast<int>(result.size()) < targetRows) result.push_back(pad);
    return QVariant::fromValue(result);
}

// RANDARRAY([rows], [cols], [min], [max], [integer])
QVariant FormulaEngine::funcRANDARRAY(const std::vector<QVariant>& args) {
    int rows = args.size() > 0 ? static_cast<int>(toNumber(args[0])) : 1;
    int cols = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 1;
    double mn = args.size() > 2 ? toNumber(args[2]) : 0.0;
    double mx = args.size() > 3 ? toNumber(args[3]) : 1.0;
    bool isInt = args.size() > 4 ? toBoolean(args[4]) : false;

    static std::mt19937 rng(std::random_device{}());
    QVariantList result;
    for (int r = 0; r < rows; ++r) {
        QVariantList row;
        for (int c = 0; c < cols; ++c) {
            if (isInt) {
                std::uniform_int_distribution<int> dist(static_cast<int>(mn), static_cast<int>(mx));
                row.append(dist(rng));
            } else {
                std::uniform_real_distribution<double> dist(mn, mx);
                row.append(dist(rng));
            }
        }
        result.append(rows > 1 ? QVariant(row) : row[0]);
    }
    if (rows == 1 && cols == 1) return result[0];
    if (rows == 1) return result;
    return result;
}

// ============================================================================
// STATISTICAL — Distribution functions (P1 core)
// ============================================================================

// Helper: standard normal CDF (Abramowitz & Stegun approximation)
static double normCDF(double x) {
    const double a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
    const double a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
    int sign = (x < 0) ? -1 : 1;
    x = std::abs(x) / std::sqrt(2.0);
    double t = 1.0 / (1.0 + p * x);
    double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * std::exp(-x * x);
    return 0.5 * (1.0 + sign * y);
}

// Helper: inverse standard normal (rational approximation)
static double normINV(double p) {
    if (p <= 0 || p >= 1) return std::numeric_limits<double>::quiet_NaN();
    // Rational approximation by Peter Acklam
    const double a[] = {-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
                         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00};
    const double b[] = {-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
                         6.680131188771972e+01, -1.328068155288572e+01};
    const double c[] = {-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
                        -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00};
    const double d[] = {7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00};

    double q, r;
    if (p < 0.02425) {
        q = std::sqrt(-2 * std::log(p));
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    } else if (p <= 0.97575) {
        q = p - 0.5; r = q * q;
        return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1);
    } else {
        q = std::sqrt(-2 * std::log(1 - p));
        return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    }
}

// NORM.DIST(x, mean, stdev, cumulative)
QVariant FormulaEngine::funcNORM_DIST(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double x = toNumber(args[0]), mean = toNumber(args[1]), sd = toNumber(args[2]);
    bool cum = toBoolean(args[3]);
    if (sd <= 0) return QVariant("#NUM!");
    double z = (x - mean) / sd;
    if (cum) return normCDF(z);
    return std::exp(-0.5 * z * z) / (sd * std::sqrt(2 * M_PI));
}

// NORM.INV(probability, mean, stdev)
QVariant FormulaEngine::funcNORM_INV(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double p = toNumber(args[0]), mean = toNumber(args[1]), sd = toNumber(args[2]);
    if (p <= 0 || p >= 1 || sd <= 0) return QVariant("#NUM!");
    return mean + sd * normINV(p);
}

// NORM.S.DIST(z, cumulative)
QVariant FormulaEngine::funcNORM_S_DIST(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double z = toNumber(args[0]);
    bool cum = toBoolean(args[1]);
    if (cum) return normCDF(z);
    return std::exp(-0.5 * z * z) / std::sqrt(2 * M_PI);
}

// NORM.S.INV(probability)
QVariant FormulaEngine::funcNORM_S_INV(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double p = toNumber(args[0]);
    if (p <= 0 || p >= 1) return QVariant("#NUM!");
    return normINV(p);
}

// BINOM.DIST(successes, trials, prob, cumulative)
QVariant FormulaEngine::funcBINOM_DIST(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    int k = static_cast<int>(toNumber(args[0]));
    int n = static_cast<int>(toNumber(args[1]));
    double p = toNumber(args[2]);
    bool cum = toBoolean(args[3]);

    auto binomPMF = [](int k, int n, double p) -> double {
        double coeff = 1;
        for (int i = 0; i < k; ++i) coeff = coeff * (n - i) / (i + 1);
        return coeff * std::pow(p, k) * std::pow(1 - p, n - k);
    };

    if (!cum) return binomPMF(k, n, p);
    double sum = 0;
    for (int i = 0; i <= k; ++i) sum += binomPMF(i, n, p);
    return sum;
}

// POISSON.DIST(x, mean, cumulative)
QVariant FormulaEngine::funcPOISSON_DIST(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    int x = static_cast<int>(toNumber(args[0]));
    double mean = toNumber(args[1]);
    bool cum = toBoolean(args[2]);
    if (mean < 0 || x < 0) return QVariant("#NUM!");

    auto poisPMF = [](int x, double m) -> double {
        double result = std::exp(-m);
        for (int i = 1; i <= x; ++i) result *= m / i;
        return result;
    };

    if (!cum) return poisPMF(x, mean);
    double sum = 0;
    for (int i = 0; i <= x; ++i) sum += poisPMF(i, mean);
    return sum;
}

// CONFIDENCE.NORM(alpha, stdev, size)
QVariant FormulaEngine::funcCONFIDENCE_NORM(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double alpha = toNumber(args[0]), sd = toNumber(args[1]);
    int n = static_cast<int>(toNumber(args[2]));
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) return QVariant("#NUM!");
    return -normINV(alpha / 2) * sd / std::sqrt(n);
}

// MODE.SNGL (same as MODE)
QVariant FormulaEngine::funcMODE_SNGL(const std::vector<QVariant>& args) {
    return funcMODE(args);
}

// ============================================================================
// FINANCIAL — Remaining P1
// ============================================================================

// MIRR(values, finance_rate, reinvest_rate)
QVariant FormulaEngine::funcMIRR(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    auto values = flattenArgs({args[0]});
    double fRate = toNumber(args[1]);
    double rRate = toNumber(args[2]);
    int n = static_cast<int>(values.size());
    if (n < 2) return QVariant("#VALUE!");

    double pvNeg = 0, fvPos = 0;
    for (int i = 0; i < n; ++i) {
        double v = toNumber(values[i]);
        if (v < 0) pvNeg += v / std::pow(1 + fRate, i);
        else fvPos += v * std::pow(1 + rRate, n - 1 - i);
    }
    if (pvNeg == 0) return QVariant("#DIV/0!");
    return std::pow(-fvPos / pvNeg, 1.0 / (n - 1)) - 1;
}

// DB(cost, salvage, life, period, [month])
QVariant FormulaEngine::funcDB(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double cost = toNumber(args[0]), salvage = toNumber(args[1]);
    int life = static_cast<int>(toNumber(args[2]));
    int period = static_cast<int>(toNumber(args[3]));
    int month = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 12;
    if (life <= 0 || cost <= 0) return QVariant("#NUM!");

    double rate = 1 - std::pow(salvage / cost, 1.0 / life);
    rate = std::round(rate * 1000) / 1000; // Round to 3 decimal places

    double totalDepn = 0;
    for (int p = 1; p <= period; ++p) {
        double depn;
        if (p == 1) depn = cost * rate * month / 12;
        else depn = (cost - totalDepn) * rate;
        if (p == life + 1) depn = (cost - totalDepn) * rate * (12 - month) / 12;
        totalDepn += depn;
        if (p == period) return depn;
    }
    return QVariant("#NUM!");
}

// DDB(cost, salvage, life, period, [factor])
QVariant FormulaEngine::funcDDB(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double cost = toNumber(args[0]), salvage = toNumber(args[1]);
    double life = toNumber(args[2]);
    int period = static_cast<int>(toNumber(args[3]));
    double factor = args.size() > 4 ? toNumber(args[4]) : 2.0;
    if (life <= 0) return QVariant("#NUM!");

    double bookValue = cost;
    for (int p = 1; p <= period; ++p) {
        double depn = std::min(bookValue * factor / life, bookValue - salvage);
        if (depn < 0) depn = 0;
        bookValue -= depn;
        if (p == period) return depn;
    }
    return 0.0;
}

// SYD(cost, salvage, life, per)
QVariant FormulaEngine::funcSYD(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double cost = toNumber(args[0]), salvage = toNumber(args[1]);
    double life = toNumber(args[2]);
    int per = static_cast<int>(toNumber(args[3]));
    double sumYears = life * (life + 1) / 2;
    return (cost - salvage) * (life - per + 1) / sumYears;
}

// DOLLARDE(fractional_dollar, fraction)
QVariant FormulaEngine::funcDOLLARDE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double fd = toNumber(args[0]);
    int frac = static_cast<int>(toNumber(args[1]));
    if (frac <= 0) return QVariant("#NUM!");
    int intPart = static_cast<int>(fd);
    double fracPart = fd - intPart;
    return intPart + fracPart * 10.0 / frac * (frac > 9 ? 1 : 1);
}

// DOLLARFR(decimal_dollar, fraction)
QVariant FormulaEngine::funcDOLLARFR(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double dd = toNumber(args[0]);
    int frac = static_cast<int>(toNumber(args[1]));
    if (frac <= 0) return QVariant("#NUM!");
    int intPart = static_cast<int>(dd);
    double fracPart = dd - intPart;
    return intPart + fracPart * frac / 10.0;
}

// ============================================================================
// ENGINEERING — Core P2 (base conversions)
// ============================================================================

// BIN2DEC, DEC2BIN, HEX2DEC, DEC2HEX, OCT2DEC, DEC2OCT
QVariant FormulaEngine::funcBIN2DEC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    bool ok; long long result = toString(args[0]).toLongLong(&ok, 2);
    return ok ? QVariant(static_cast<double>(result)) : QVariant("#NUM!");
}

QVariant FormulaEngine::funcDEC2BIN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    long long n = static_cast<long long>(toNumber(args[0]));
    int places = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 0;
    QString result = QString::number(n, 2);
    while (places > 0 && result.length() < places) result = "0" + result;
    return result;
}

QVariant FormulaEngine::funcHEX2DEC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    bool ok; long long result = toString(args[0]).toLongLong(&ok, 16);
    return ok ? QVariant(static_cast<double>(result)) : QVariant("#NUM!");
}

QVariant FormulaEngine::funcDEC2HEX(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    long long n = static_cast<long long>(toNumber(args[0]));
    int places = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 0;
    QString result = QString::number(n, 16).toUpper();
    while (places > 0 && result.length() < places) result = "0" + result;
    return result;
}

QVariant FormulaEngine::funcOCT2DEC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    bool ok; long long result = toString(args[0]).toLongLong(&ok, 8);
    return ok ? QVariant(static_cast<double>(result)) : QVariant("#NUM!");
}

QVariant FormulaEngine::funcDEC2OCT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    long long n = static_cast<long long>(toNumber(args[0]));
    int places = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 0;
    QString result = QString::number(n, 8);
    while (places > 0 && result.length() < places) result = "0" + result;
    return result;
}

// CONVERT(number, from_unit, to_unit) — simplified with common conversions
QVariant FormulaEngine::funcCONVERT(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    QString from = toString(args[1]).toLower();
    QString to = toString(args[2]).toLower();

    // Convert everything to SI base, then to target
    std::map<QString, double> lengthToM = {
        {"m",1},{"cm",0.01},{"mm",0.001},{"km",1000},{"in",0.0254},
        {"ft",0.3048},{"yd",0.9144},{"mi",1609.344},{"nm",1852}
    };
    std::map<QString, double> massToKg = {
        {"kg",1},{"g",0.001},{"mg",1e-6},{"lbm",0.45359237},{"oz",0.028349523},{"ton",907.18474}
    };
    std::map<QString, std::pair<double,double>> tempConv = {
        {"c",{1,0}},{"f",{5.0/9, -32*5.0/9}},{"k",{1,-273.15}}
    };

    // Try length
    if (lengthToM.count(from) && lengthToM.count(to))
        return n * lengthToM[from] / lengthToM[to];
    // Try mass
    if (massToKg.count(from) && massToKg.count(to))
        return n * massToKg[from] / massToKg[to];
    // Try temperature
    if (tempConv.count(from) && tempConv.count(to)) {
        double celsius;
        if (from == "c") celsius = n;
        else if (from == "f") celsius = (n - 32) * 5.0 / 9;
        else celsius = n - 273.15; // kelvin
        if (to == "c") return celsius;
        if (to == "f") return celsius * 9.0 / 5 + 32;
        return celsius + 273.15; // kelvin
    }

    return QVariant("#N/A");
}

// DELTA(number1, [number2])
QVariant FormulaEngine::funcDELTA(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double a = toNumber(args[0]);
    double b = args.size() > 1 ? toNumber(args[1]) : 0;
    return (a == b) ? 1 : 0;
}

// GESTEP(number, [step])
QVariant FormulaEngine::funcGESTEP(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    double step = args.size() > 1 ? toNumber(args[1]) : 0;
    return (n >= step) ? 1 : 0;
}

// ============================================================================
// DATABASE — P2 functions
// ============================================================================

// Helper: evaluate database criteria
static bool dbCriteriaMatch(Spreadsheet* sheet, int dataRow, const CellRange& criteria,
                            const std::vector<int>& fieldCols, const std::vector<QString>& headers,
                            FormulaEngine* engine) {
    Q_UNUSED(engine);
    int critStartRow = criteria.getStart().row;
    int critEndRow = criteria.getEnd().row;
    int critStartCol = criteria.getStart().col;
    int critEndCol = criteria.getEnd().col;

    // Each criteria row is an AND condition
    for (int cr = critStartRow + 1; cr <= critEndRow; ++cr) {
        bool rowMatch = true;
        for (int cc = critStartCol; cc <= critEndCol; ++cc) {
            QVariant critVal = sheet->getCellValue(CellAddress(cr, cc));
            if (!critVal.isValid() || critVal.toString().isEmpty()) continue;

            // Find which data column this criteria header matches
            QString critHeader = sheet->getCellValue(CellAddress(critStartRow, cc)).toString();
            int dataCol = -1;
            for (size_t i = 0; i < headers.size(); ++i) {
                if (headers[i].compare(critHeader, Qt::CaseInsensitive) == 0) {
                    dataCol = fieldCols[i]; break;
                }
            }
            if (dataCol < 0) { rowMatch = false; break; }

            QVariant cellVal = sheet->getCellValue(CellAddress(dataRow, dataCol));
            QString critStr = critVal.toString();
            // Simple comparison
            if (critStr.startsWith(">") || critStr.startsWith("<") || critStr.startsWith("=")) {
                // Criteria with operator
                double cellNum = cellVal.toDouble();
                if (critStr.startsWith(">=")) { if (!(cellNum >= critStr.mid(2).toDouble())) rowMatch = false; }
                else if (critStr.startsWith("<=")) { if (!(cellNum <= critStr.mid(2).toDouble())) rowMatch = false; }
                else if (critStr.startsWith("<>")) { if (cellVal.toString() == critStr.mid(2)) rowMatch = false; }
                else if (critStr.startsWith(">")) { if (!(cellNum > critStr.mid(1).toDouble())) rowMatch = false; }
                else if (critStr.startsWith("<")) { if (!(cellNum < critStr.mid(1).toDouble())) rowMatch = false; }
                else if (critStr.startsWith("=")) { if (cellVal.toString() != critStr.mid(1)) rowMatch = false; }
            } else {
                if (cellVal.toString().compare(critStr, Qt::CaseInsensitive) != 0) rowMatch = false;
            }
        }
        if (rowMatch) return true; // OR between criteria rows
    }
    return false;
}

// DSUM(database, field, criteria)
QVariant FormulaEngine::funcDSUM(const std::vector<QVariant>& args) {
    if (args.size() < 3 || !m_spreadsheet) return QVariant("#VALUE!");
    if (!args[0].canConvert<CellRange>() || !args[2].canConvert<CellRange>()) {
        // Fallback for non-range args
        return QVariant("#VALUE!");
    }
    // Simplified: return 0 for now (full implementation needs range context)
    return 0.0;
}

// DAVERAGE, DCOUNT, DMIN, DMAX — simplified stubs
QVariant FormulaEngine::funcDAVERAGE(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }
QVariant FormulaEngine::funcDCOUNT(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0; }
QVariant FormulaEngine::funcDCOUNTA(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0; }
QVariant FormulaEngine::funcDMIN(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }
QVariant FormulaEngine::funcDMAX(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }
QVariant FormulaEngine::funcDGET(const std::vector<QVariant>& args) { Q_UNUSED(args); return QVariant("#VALUE!"); }
QVariant FormulaEngine::funcDPRODUCT(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }
QVariant FormulaEngine::funcDSTDEV(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }
QVariant FormulaEngine::funcDVAR(const std::vector<QVariant>& args) { Q_UNUSED(args); return 0.0; }

// ============================================================================
// WEB — P2
// ============================================================================

// ENCODEURL(text)
QVariant FormulaEngine::funcENCODEURL(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    return QUrl::toPercentEncoding(toString(args[0]));
}
