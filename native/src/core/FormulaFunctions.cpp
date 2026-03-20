// ============================================================================
// FormulaFunctions.cpp — Extended function library (272+ functions)
// ============================================================================
// Implements all P1 Excel functions not already in FormulaEngine.cpp.
// Organized by category matching EXCEL_FEATURES.md Section 13.
//

#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <numeric>
#include <random>
#include <set>
#include <map>
#include <QDateTime>
#include <QRegularExpression>

// ============================================================================
// Helper: date serial conversion (Excel 1900 system)
// ============================================================================
static const QDate EXCEL_EPOCH(1899, 12, 30);

static double dateToSerial(const QDate& d) { return EXCEL_EPOCH.daysTo(d); }
static QDate serialToDate(double s) { return EXCEL_EPOCH.addDays(static_cast<int>(s)); }
static double timeToSerial(const QTime& t) {
    return (t.hour() * 3600 + t.minute() * 60 + t.second()) / 86400.0;
}

// ============================================================================
// 13.1 MATH & TRIG — Missing functions
// ============================================================================

// PI()
QVariant FormulaEngine::funcPI(const std::vector<QVariant>&) {
    return M_PI;
}

// SIGN(number)
QVariant FormulaEngine::funcSIGN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    return (n > 0) ? 1 : (n < 0) ? -1 : 0;
}

// TRUNC(number, [num_digits])
QVariant FormulaEngine::funcTRUNC(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    int digits = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 0;
    double factor = std::pow(10.0, digits);
    return std::trunc(n * factor) / factor;
}

// PRODUCT(number1, ...)
QVariant FormulaEngine::funcPRODUCT(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double product = 1.0;
    bool hasNum = false;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            product *= toNumber(a);
            hasNum = true;
        }
    }
    return hasNum ? QVariant(product) : QVariant(0.0);
}

// QUOTIENT(numerator, denominator)
QVariant FormulaEngine::funcQUOTIENT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double num = toNumber(args[0]), den = toNumber(args[1]);
    if (den == 0) return QVariant("#DIV/0!");
    return static_cast<int>(num / den);
}

// MROUND(number, multiple)
QVariant FormulaEngine::funcMROUND(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double n = toNumber(args[0]), m = toNumber(args[1]);
    if (m == 0) return 0.0;
    return std::round(n / m) * m;
}

// CEILING.MATH(number, [significance], [mode])
QVariant FormulaEngine::funcCEILING_MATH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    double sig = args.size() > 1 ? toNumber(args[1]) : 1.0;
    int mode = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 0;
    if (sig == 0) return 0.0;
    if (n < 0 && mode != 0) return -std::floor(-n / std::abs(sig)) * std::abs(sig);
    return std::ceil(n / sig) * sig;
}

// FLOOR.MATH(number, [significance], [mode])
QVariant FormulaEngine::funcFLOOR_MATH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    double sig = args.size() > 1 ? toNumber(args[1]) : 1.0;
    int mode = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 0;
    if (sig == 0) return 0.0;
    if (n < 0 && mode != 0) return -std::ceil(-n / std::abs(sig)) * std::abs(sig);
    return std::floor(n / sig) * sig;
}

// LOG10(number)
QVariant FormulaEngine::funcLOG10(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    if (n <= 0) return QVariant("#NUM!");
    return std::log10(n);
}

// SIN, COS, TAN, ASIN, ACOS, ATAN, ATAN2
QVariant FormulaEngine::funcSIN(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::sin(toNumber(args[0])));
}
QVariant FormulaEngine::funcCOS(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::cos(toNumber(args[0])));
}
QVariant FormulaEngine::funcTAN(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::tan(toNumber(args[0])));
}
QVariant FormulaEngine::funcASIN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    if (n < -1 || n > 1) return QVariant("#NUM!");
    return std::asin(n);
}
QVariant FormulaEngine::funcACOS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    if (n < -1 || n > 1) return QVariant("#NUM!");
    return std::acos(n);
}
QVariant FormulaEngine::funcATAN(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::atan(toNumber(args[0])));
}
QVariant FormulaEngine::funcATAN2(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    return std::atan2(toNumber(args[1]), toNumber(args[0]));  // Excel: ATAN2(x, y)
}

// RADIANS, DEGREES
QVariant FormulaEngine::funcRADIANS(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(toNumber(args[0]) * M_PI / 180.0);
}
QVariant FormulaEngine::funcDEGREES(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(toNumber(args[0]) * 180.0 / M_PI);
}

// FACT(number)
QVariant FormulaEngine::funcFACT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    if (n < 0) return QVariant("#NUM!");
    if (n > 170) return QVariant("#NUM!"); // overflow
    double result = 1;
    for (int i = 2; i <= n; ++i) result *= i;
    return result;
}

// COMBIN(n, k)
QVariant FormulaEngine::funcCOMBIN(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    int k = static_cast<int>(toNumber(args[1]));
    if (n < 0 || k < 0 || k > n) return QVariant("#NUM!");
    double result = 1;
    for (int i = 0; i < k; ++i) result = result * (n - i) / (i + 1);
    return result;
}

// PERMUT(n, k)
QVariant FormulaEngine::funcPERMUT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    int k = static_cast<int>(toNumber(args[1]));
    if (n < 0 || k < 0 || k > n) return QVariant("#NUM!");
    double result = 1;
    for (int i = 0; i < k; ++i) result *= (n - i);
    return result;
}

// GCD(number1, number2, ...)
QVariant FormulaEngine::funcGCD(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#VALUE!");
    long long result = static_cast<long long>(std::abs(toNumber(flat[0])));
    for (size_t i = 1; i < flat.size(); ++i) {
        long long b = static_cast<long long>(std::abs(toNumber(flat[i])));
        while (b) { long long t = b; b = result % b; result = t; }
    }
    return static_cast<double>(result);
}

// LCM(number1, number2, ...)
QVariant FormulaEngine::funcLCM(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#VALUE!");
    long long result = static_cast<long long>(std::abs(toNumber(flat[0])));
    for (size_t i = 1; i < flat.size(); ++i) {
        long long b = static_cast<long long>(std::abs(toNumber(flat[i])));
        if (b == 0) return 0.0;
        // Manual GCD for LCM calculation
        long long a2 = result, b2 = b;
        while (b2) { long long t = b2; b2 = a2 % b2; a2 = t; }
        result = result / a2 * b;
    }
    return static_cast<double>(result);
}

// EVEN(number), ODD(number)
QVariant FormulaEngine::funcEVEN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    double c = (n >= 0) ? std::ceil(n) : std::floor(n);
    int i = static_cast<int>(c);
    if (i % 2 != 0) i += (n >= 0) ? 1 : -1;
    return static_cast<double>(i);
}

QVariant FormulaEngine::funcODD(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    double c = (n >= 0) ? std::ceil(n) : std::floor(n);
    int i = static_cast<int>(c);
    if (i == 0) i = 1;
    else if (i % 2 == 0) i += (n >= 0) ? 1 : -1;
    return static_cast<double>(i);
}

// SUMSQ(number1, ...)
QVariant FormulaEngine::funcSUMSQ(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double sum = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) { double v = toNumber(a); sum += v * v; }
    }
    return sum;
}

// ============================================================================
// 13.2 LOOKUP & REFERENCE — Missing functions
// ============================================================================

// ROW([reference]), COLUMN([reference])
QVariant FormulaEngine::funcROW(const std::vector<QVariant>& args) {
    Q_UNUSED(args);
    // Without reference, returns current row — needs context; return 1 as default
    return 1;
}
QVariant FormulaEngine::funcCOLUMN(const std::vector<QVariant>& args) {
    Q_UNUSED(args);
    return 1;
}

// ROWS(array), COLUMNS(array)
QVariant FormulaEngine::funcROWS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    if (args[0].canConvert<CellRange>()) {
        CellRange r = args[0].value<CellRange>();
        return r.getRowCount();
    }
    if (args[0].canConvert<std::vector<QVariant>>()) {
        return static_cast<int>(args[0].value<std::vector<QVariant>>().size());
    }
    return 1;
}
QVariant FormulaEngine::funcCOLUMNS(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    if (args[0].canConvert<CellRange>()) {
        CellRange r = args[0].value<CellRange>();
        return r.getColumnCount();
    }
    return 1;
}

// ADDRESS(row, col, [abs_num], [a1], [sheet])
QVariant FormulaEngine::funcADDRESS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int row = static_cast<int>(toNumber(args[0]));
    int col = static_cast<int>(toNumber(args[1]));
    int absNum = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 1;
    // Convert col to letter
    QString colStr;
    int c = col - 1;
    while (c >= 0) { colStr = QChar('A' + c % 26) + colStr; c = c / 26 - 1; }
    QString result;
    switch (absNum) {
        case 1: result = "$" + colStr + "$" + QString::number(row); break;
        case 2: result = colStr + "$" + QString::number(row); break;
        case 3: result = "$" + colStr + QString::number(row); break;
        default: result = colStr + QString::number(row); break;
    }
    if (args.size() > 4 && !args[4].toString().isEmpty()) {
        result = args[4].toString() + "!" + result;
    }
    return result;
}

// TRANSPOSE(array) — returns transposed 2D array
QVariant FormulaEngine::funcTRANSPOSE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    // For CellRange argument, get 2D values
    if (args[0].canConvert<CellRange>()) {
        CellRange range = args[0].value<CellRange>();
        auto range2D = getRangeValues2D(range);
        if (range2D.empty()) return QVariant("#VALUE!");
        int rows = static_cast<int>(range2D.size());
        int cols = static_cast<int>(range2D[0].size());
        QVariantList result;
        for (int c = 0; c < cols; ++c) {
            QVariantList row;
            for (int r = 0; r < rows; ++r) {
                row.append(c < static_cast<int>(range2D[r].size()) ? range2D[r][c] : QVariant());
            }
            result.append(QVariant(row));
        }
        return result;
    }
    return args[0];
}

// ============================================================================
// 13.3 TEXT — Missing functions
// ============================================================================

// CHAR(number)
QVariant FormulaEngine::funcCHAR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int code = static_cast<int>(toNumber(args[0]));
    if (code < 1 || code > 65535) return QVariant("#VALUE!");
    return QString(QChar(code));
}

// CODE(text)
QVariant FormulaEngine::funcCODE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString s = toString(args[0]);
    if (s.isEmpty()) return QVariant("#VALUE!");
    return static_cast<int>(s[0].unicode());
}

// CLEAN(text) — removes non-printable characters
QVariant FormulaEngine::funcCLEAN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString s = toString(args[0]);
    QString result;
    for (const QChar& c : s) {
        if (c.unicode() >= 32) result += c;
    }
    return result;
}

// REPLACE(old_text, start_num, num_chars, new_text)
QVariant FormulaEngine::funcREPLACE(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    int start = static_cast<int>(toNumber(args[1])) - 1; // 1-based
    int count = static_cast<int>(toNumber(args[2]));
    QString newText = toString(args[3]);
    return text.left(start) + newText + text.mid(start + count);
}

// FIXED(number, [decimals], [no_commas])
QVariant FormulaEngine::funcFIXED(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double n = toNumber(args[0]);
    int decimals = args.size() > 1 ? static_cast<int>(toNumber(args[1])) : 2;
    bool noCommas = args.size() > 2 ? toBoolean(args[2]) : false;
    QString result = QString::number(n, 'f', decimals);
    if (!noCommas && decimals >= 0) {
        // Add thousands separator
        int dotPos = result.indexOf('.');
        if (dotPos < 0) dotPos = result.length();
        int startPos = (n < 0) ? 1 : 0;
        for (int i = dotPos - 3; i > startPos; i -= 3) {
            result.insert(i, ',');
        }
    }
    return result;
}

// T(value) — returns text if text, "" if not
QVariant FormulaEngine::funcT(const std::vector<QVariant>& args) {
    if (args.empty()) return QString();
    if (args[0].typeId() == QMetaType::QString) return args[0];
    return QString();
}

// N(value) — returns number if number, 0 if not
QVariant FormulaEngine::funcN(const std::vector<QVariant>& args) {
    if (args.empty()) return 0.0;
    if (args[0].typeId() == QMetaType::Double || args[0].typeId() == QMetaType::Int) return toNumber(args[0]);
    if (args[0].typeId() == QMetaType::Bool) return args[0].toBool() ? 1.0 : 0.0;
    return 0.0;
}

// NUMBERVALUE(text, [decimal_separator], [group_separator])
QVariant FormulaEngine::funcNUMBERVALUE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString text = toString(args[0]).trimmed();
    QString decSep = args.size() > 1 ? toString(args[1]) : ".";
    QString grpSep = args.size() > 2 ? toString(args[2]) : ",";
    text.remove(grpSep);
    text.replace(decSep, ".");
    bool ok;
    double result = text.toDouble(&ok);
    return ok ? QVariant(result) : QVariant("#VALUE!");
}

// UNICODE(text), UNICHAR(number)
QVariant FormulaEngine::funcUNICODE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString s = toString(args[0]);
    if (s.isEmpty()) return QVariant("#VALUE!");
    return static_cast<int>(s[0].unicode());
}
QVariant FormulaEngine::funcUNICHAR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int code = static_cast<int>(toNumber(args[0]));
    return QString(QChar(code));
}

// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
QVariant FormulaEngine::funcTEXTBEFORE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QString delim = toString(args[1]);
    int instance = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 1;
    if (instance <= 0) return QVariant("#VALUE!");
    int pos = -1;
    int searchFrom = 0;
    for (int i = 0; i < instance; ++i) {
        pos = text.indexOf(delim, searchFrom);
        if (pos < 0) {
            return args.size() > 5 ? args[5] : QVariant("#N/A");
        }
        searchFrom = pos + delim.length();
    }
    return text.left(pos);
}

// TEXTAFTER(text, delimiter, [instance_num], ...)
QVariant FormulaEngine::funcTEXTAFTER(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QString delim = toString(args[1]);
    int instance = args.size() > 2 ? static_cast<int>(toNumber(args[2])) : 1;
    if (instance <= 0) return QVariant("#VALUE!");
    int pos = -1;
    int searchFrom = 0;
    for (int i = 0; i < instance; ++i) {
        pos = text.indexOf(delim, searchFrom);
        if (pos < 0) {
            return args.size() > 5 ? args[5] : QVariant("#N/A");
        }
        searchFrom = pos + delim.length();
    }
    return text.mid(pos + delim.length());
}

// ============================================================================
// 13.4 LOGICAL — Missing functions
// ============================================================================

// XOR(logical1, ...)
QVariant FormulaEngine::funcXOR(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    int trueCount = 0;
    for (const auto& a : flat) {
        if (toBoolean(a)) trueCount++;
    }
    return (trueCount % 2 == 1);
}

// IFNA(value, value_if_na)
QVariant FormulaEngine::funcIFNA(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString s = args[0].toString();
    if (s == "#N/A") return args[1];
    return args[0];
}

// TRUE(), FALSE()
QVariant FormulaEngine::funcTRUE(const std::vector<QVariant>&) { return true; }
QVariant FormulaEngine::funcFALSE(const std::vector<QVariant>&) { return false; }

// ============================================================================
// 13.5 DATE & TIME — Missing functions
// ============================================================================

// TIME(hour, minute, second)
QVariant FormulaEngine::funcTIME(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    int h = static_cast<int>(toNumber(args[0]));
    int m = static_cast<int>(toNumber(args[1]));
    int s = static_cast<int>(toNumber(args[2]));
    return (h * 3600 + m * 60 + s) / 86400.0;
}

// TIMEVALUE(time_text)
QVariant FormulaEngine::funcTIMEVALUE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QTime t = QTime::fromString(toString(args[0]), "h:mm:ss");
    if (!t.isValid()) t = QTime::fromString(toString(args[0]), "h:mm");
    if (!t.isValid()) return QVariant("#VALUE!");
    return timeToSerial(t);
}

// DAYS(end_date, start_date)
QVariant FormulaEngine::funcDAYS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double end = toNumber(args[0]), start = toNumber(args[1]);
    return end - start;
}

// ISOWEEKNUM(date)
QVariant FormulaEngine::funcISOWEEKNUM(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate d = serialToDate(toNumber(args[0]));
    if (!d.isValid()) return QVariant("#VALUE!");
    return d.weekNumber();
}

// WEEKNUM(serial, [return_type])
QVariant FormulaEngine::funcWEEKNUM(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate d = serialToDate(toNumber(args[0]));
    if (!d.isValid()) return QVariant("#VALUE!");
    // Simplified: week starts on Sunday (return_type=1)
    int jan1Day = QDate(d.year(), 1, 1).dayOfWeek(); // 1=Mon, 7=Sun
    int dayOfYear = d.dayOfYear();
    return (dayOfYear + jan1Day - 1) / 7 + 1;
}

// WORKDAY(start_date, days, [holidays])
QVariant FormulaEngine::funcWORKDAY(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QDate d = serialToDate(toNumber(args[0]));
    int days = static_cast<int>(toNumber(args[1]));
    // Collect holidays
    QSet<QDate> holidays;
    if (args.size() > 2) {
        auto flat = flattenArgs({args[2]});
        for (const auto& h : flat) {
            holidays.insert(serialToDate(toNumber(h)));
        }
    }
    int step = (days > 0) ? 1 : -1;
    int remaining = std::abs(days);
    while (remaining > 0) {
        d = d.addDays(step);
        int dow = d.dayOfWeek(); // 1=Mon, 7=Sun
        if (dow <= 5 && !holidays.contains(d)) remaining--;
    }
    return dateToSerial(d);
}

// ============================================================================
// 13.7 FINANCIAL — Core functions (P1)
// ============================================================================

// PMT(rate, nper, pv, [fv], [type])
QVariant FormulaEngine::funcPMT(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double nper = toNumber(args[1]);
    double pv = toNumber(args[2]);
    double fv = args.size() > 3 ? toNumber(args[3]) : 0;
    int type = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 0;
    if (rate == 0) return -(pv + fv) / nper;
    double pvif = std::pow(1 + rate, nper);
    double pmt = rate * (pv * pvif + fv) / ((pvif - 1) * (1 + rate * type));
    return -pmt;
}

// FV(rate, nper, pmt, [pv], [type])
QVariant FormulaEngine::funcFV(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double nper = toNumber(args[1]);
    double pmt = toNumber(args[2]);
    double pv = args.size() > 3 ? toNumber(args[3]) : 0;
    int type = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 0;
    if (rate == 0) return -(pv + pmt * nper);
    double pvif = std::pow(1 + rate, nper);
    return -(pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate);
}

// PV(rate, nper, pmt, [fv], [type])
QVariant FormulaEngine::funcPV(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double nper = toNumber(args[1]);
    double pmt = toNumber(args[2]);
    double fv = args.size() > 3 ? toNumber(args[3]) : 0;
    int type = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 0;
    if (rate == 0) return -(fv + pmt * nper);
    double pvif = std::pow(1 + rate, nper);
    return -(fv / pvif + pmt * (1 + rate * type) * (pvif - 1) / (rate * pvif));
}

// NPV(rate, value1, value2, ...)
QVariant FormulaEngine::funcNPV(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    auto values = flattenArgs({args.begin() + 1, args.end()});
    double npv = 0;
    for (size_t i = 0; i < values.size(); ++i) {
        npv += toNumber(values[i]) / std::pow(1 + rate, i + 1);
    }
    return npv;
}

// NPER(rate, pmt, pv, [fv], [type])
QVariant FormulaEngine::funcNPER(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    double pmt = toNumber(args[1]);
    double pv = toNumber(args[2]);
    double fv = args.size() > 3 ? toNumber(args[3]) : 0;
    int type = args.size() > 4 ? static_cast<int>(toNumber(args[4])) : 0;
    if (rate == 0) return -(pv + fv) / pmt;
    double z = pmt * (1 + rate * type) / rate;
    return std::log((-fv + z) / (pv + z)) / std::log(1 + rate);
}

// IRR(values, [guess])
QVariant FormulaEngine::funcIRR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    auto values = flattenArgs({args[0]});
    double guess = args.size() > 1 ? toNumber(args[1]) : 0.1;
    // Newton-Raphson method
    double rate = guess;
    for (int iter = 0; iter < 100; ++iter) {
        double npv = 0, dnpv = 0;
        for (size_t i = 0; i < values.size(); ++i) {
            double cf = toNumber(values[i]);
            npv += cf / std::pow(1 + rate, i);
            dnpv -= i * cf / std::pow(1 + rate, i + 1);
        }
        if (std::abs(dnpv) < 1e-15) break;
        double newRate = rate - npv / dnpv;
        if (std::abs(newRate - rate) < 1e-10) return newRate;
        rate = newRate;
    }
    return rate;
}

// EFFECT(nominal_rate, npery)
QVariant FormulaEngine::funcEFFECT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double nom = toNumber(args[0]);
    int npery = static_cast<int>(toNumber(args[1]));
    if (npery < 1) return QVariant("#NUM!");
    return std::pow(1 + nom / npery, npery) - 1;
}

// NOMINAL(effect_rate, npery)
QVariant FormulaEngine::funcNOMINAL(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double eff = toNumber(args[0]);
    int npery = static_cast<int>(toNumber(args[1]));
    if (npery < 1) return QVariant("#NUM!");
    return npery * (std::pow(1 + eff, 1.0 / npery) - 1);
}

// IPMT(rate, per, nper, pv, [fv], [type])
QVariant FormulaEngine::funcIPMT(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    double rate = toNumber(args[0]);
    int per = static_cast<int>(toNumber(args[1]));
    double nper = toNumber(args[2]);
    double pv = toNumber(args[3]);
    double fv = args.size() > 4 ? toNumber(args[4]) : 0;
    int type = args.size() > 5 ? static_cast<int>(toNumber(args[5])) : 0;
    // PMT
    double pmt;
    if (rate == 0) pmt = -(pv + fv) / nper;
    else {
        double pvif = std::pow(1 + rate, nper);
        pmt = -(rate * (pv * pvif + fv) / ((pvif - 1) * (1 + rate * type)));
    }
    // Interest portion
    double fvBefore;
    if (rate == 0) fvBefore = pv + pmt * (per - 1);
    else fvBefore = pv * std::pow(1 + rate, per - 1) + pmt * (std::pow(1 + rate, per - 1) - 1) / rate;
    return -(fvBefore * rate);
}

// PPMT(rate, per, nper, pv, [fv], [type])
QVariant FormulaEngine::funcPPMT(const std::vector<QVariant>& args) {
    if (args.size() < 4) return QVariant("#VALUE!");
    // PPMT = PMT - IPMT
    auto pmtArgs = std::vector<QVariant>{args[0], args[2], args[3]};
    if (args.size() > 4) pmtArgs.push_back(args[4]);
    if (args.size() > 5) pmtArgs.push_back(args[5]);
    double pmt = toNumber(funcPMT(pmtArgs));
    double ipmt = toNumber(funcIPMT(args));
    return pmt - ipmt;
}

// SLN(cost, salvage, life)
QVariant FormulaEngine::funcSLN(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double cost = toNumber(args[0]), salvage = toNumber(args[1]), life = toNumber(args[2]);
    if (life == 0) return QVariant("#DIV/0!");
    return (cost - salvage) / life;
}

// ============================================================================
// 13.8 INFORMATION — Missing functions
// ============================================================================

// ISERR(value) — TRUE if error but NOT #N/A
QVariant FormulaEngine::funcISERR(const std::vector<QVariant>& args) {
    if (args.empty()) return false;
    QString s = args[0].toString();
    return (s.startsWith("#") && s != "#N/A");
}

// ISNA(value)
QVariant FormulaEngine::funcISNA(const std::vector<QVariant>& args) {
    if (args.empty()) return false;
    return args[0].toString() == "#N/A";
}

// ISLOGICAL(value)
QVariant FormulaEngine::funcISLOGICAL(const std::vector<QVariant>& args) {
    if (args.empty()) return false;
    return args[0].typeId() == QMetaType::Bool;
}

// ISNONTEXT(value)
QVariant FormulaEngine::funcISNONTEXT(const std::vector<QVariant>& args) {
    if (args.empty()) return true;
    return args[0].typeId() != QMetaType::QString;
}

// ISEVEN(value), ISODD(value)
QVariant FormulaEngine::funcISEVEN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    return (n % 2 == 0);
}
QVariant FormulaEngine::funcISODD(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int n = static_cast<int>(toNumber(args[0]));
    return (n % 2 != 0);
}

// TYPE(value) — returns type number
QVariant FormulaEngine::funcTYPE(const std::vector<QVariant>& args) {
    if (args.empty()) return 1;
    if (args[0].typeId() == QMetaType::Double || args[0].typeId() == QMetaType::Int) return 1;
    if (args[0].typeId() == QMetaType::QString) {
        if (args[0].toString().startsWith("#")) return 16; // error
        return 2; // text
    }
    if (args[0].typeId() == QMetaType::Bool) return 4;
    return 1;
}

// NA() — returns #N/A
QVariant FormulaEngine::funcNA(const std::vector<QVariant>&) {
    return QVariant("#N/A");
}

// ERROR.TYPE(error_val)
QVariant FormulaEngine::funcERROR_TYPE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#N/A");
    QString s = args[0].toString();
    if (s == "#NULL!") return 1;
    if (s == "#DIV/0!") return 2;
    if (s == "#VALUE!") return 3;
    if (s == "#REF!") return 4;
    if (s == "#NAME?") return 5;
    if (s == "#NUM!") return 6;
    if (s == "#N/A") return 7;
    return QVariant("#N/A");
}

// ============================================================================
// 13.6 STATISTICAL — Missing core functions
// ============================================================================

// AVERAGEA(value1, ...)
QVariant FormulaEngine::funcAVERAGEA(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0; int count = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            if (a.typeId() == QMetaType::Bool) sum += a.toBool() ? 1.0 : 0.0;
            else if (a.typeId() == QMetaType::QString) sum += 0;
            else sum += toNumber(a);
            count++;
        }
    }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

// MAXA, MINA
QVariant FormulaEngine::funcMAXA(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double mx = std::numeric_limits<double>::lowest();
    bool found = false;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = (a.typeId() == QMetaType::Bool) ? (a.toBool() ? 1.0 : 0.0) : toNumber(a);
            if (!found || v > mx) { mx = v; found = true; }
        }
    }
    return found ? QVariant(mx) : QVariant(0.0);
}
QVariant FormulaEngine::funcMINA(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    double mn = std::numeric_limits<double>::max();
    bool found = false;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            double v = (a.typeId() == QMetaType::Bool) ? (a.toBool() ? 1.0 : 0.0) : toNumber(a);
            if (!found || v < mn) { mn = v; found = true; }
        }
    }
    return found ? QVariant(mn) : QVariant(0.0);
}

// CORREL(array1, array2)
QVariant FormulaEngine::funcCORREL(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto x = flattenArgs({args[0]}), y = flattenArgs({args[1]});
    int n = std::min(x.size(), y.size());
    if (n < 2) return QVariant("#DIV/0!");
    double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;
    for (int i = 0; i < n; ++i) {
        double xi = toNumber(x[i]), yi = toNumber(y[i]);
        sumX += xi; sumY += yi; sumXY += xi * yi;
        sumX2 += xi * xi; sumY2 += yi * yi;
    }
    double denom = std::sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
    if (denom == 0) return QVariant("#DIV/0!");
    return (n * sumXY - sumX * sumY) / denom;
}

// SLOPE(known_y, known_x)
QVariant FormulaEngine::funcSLOPE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto y = flattenArgs({args[0]}), x = flattenArgs({args[1]});
    int n = std::min(x.size(), y.size());
    if (n < 2) return QVariant("#DIV/0!");
    double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
    for (int i = 0; i < n; ++i) {
        double xi = toNumber(x[i]), yi = toNumber(y[i]);
        sumX += xi; sumY += yi; sumXY += xi * yi; sumX2 += xi * xi;
    }
    double denom = n * sumX2 - sumX * sumX;
    if (denom == 0) return QVariant("#DIV/0!");
    return (n * sumXY - sumX * sumY) / denom;
}

// INTERCEPT(known_y, known_x)
QVariant FormulaEngine::funcINTERCEPT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto y = flattenArgs({args[0]}), x = flattenArgs({args[1]});
    int n = std::min(x.size(), y.size());
    if (n < 2) return QVariant("#DIV/0!");
    double sumX = 0, sumY = 0;
    for (int i = 0; i < n; ++i) { sumX += toNumber(x[i]); sumY += toNumber(y[i]); }
    double slope = toNumber(funcSLOPE(args));
    return sumY / n - slope * sumX / n;
}

// FORECAST.LINEAR(x, known_y, known_x) or FORECAST
QVariant FormulaEngine::funcFORECAST(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    double x = toNumber(args[0]);
    double slope = toNumber(funcSLOPE({args[1], args[2]}));
    double intercept = toNumber(funcINTERCEPT({args[1], args[2]}));
    return intercept + slope * x;
}
