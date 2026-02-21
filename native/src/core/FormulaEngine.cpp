#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <cctype>
#include <limits>
#include <numeric>
#include <random>
#include <QDateTime>
#include <QRegularExpression>

FormulaEngine::FormulaEngine(Spreadsheet* spreadsheet)
    : m_spreadsheet(spreadsheet) {
}

void FormulaEngine::setSpreadsheet(Spreadsheet* spreadsheet) {
    m_spreadsheet = spreadsheet;
}

QVariant FormulaEngine::evaluate(const QString& formula) {
    m_lastError.clear();
    m_lastDependencies.clear();
    m_lastRangeArgs.clear();

    if (formula.isEmpty()) return QVariant();

    QString expr = formula.startsWith('=') ? formula.mid(1) : formula;

    try {
        return parseExpression(expr);
    } catch (const std::exception& e) {
        m_lastError = QString::fromStdString(e.what());
        return QVariant("#ERROR!");
    }
}

QString FormulaEngine::getLastError() const { return m_lastError; }
bool FormulaEngine::hasError() const { return !m_lastError.isEmpty(); }
void FormulaEngine::clearCache() { m_cache.clear(); }
void FormulaEngine::invalidateCell(const CellAddress& addr) { m_cache.erase(addr.toString().toStdString()); }

void FormulaEngine::skipWhitespace(const QString& expr, int& pos) {
    while (pos < expr.length() && expr[pos].isSpace()) pos++;
}

std::vector<QVariant> FormulaEngine::flattenArgs(const std::vector<QVariant>& args) {
    std::vector<QVariant> flat;
    for (const auto& arg : args) {
        if (arg.canConvert<std::vector<QVariant>>()) {
            auto nested = arg.value<std::vector<QVariant>>();
            for (const auto& v : nested) flat.push_back(v);
        } else {
            flat.push_back(arg);
        }
    }
    return flat;
}

QVariant FormulaEngine::parseExpression(const QString& expr) {
    int pos = 0;
    QVariant result = evaluateComparison(expr, pos);
    skipWhitespace(expr, pos);
    return result;
}

QVariant FormulaEngine::evaluateComparison(const QString& expr, int& pos) {
    QVariant left = evaluateTerm(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (pos + 1 < expr.length() && expr[pos] == '<' && expr[pos + 1] == '>') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) != toNumber(right));
        } else if (pos + 1 < expr.length() && expr[pos] == '<' && expr[pos + 1] == '=') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) <= toNumber(right));
        } else if (pos + 1 < expr.length() && expr[pos] == '>' && expr[pos + 1] == '=') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) >= toNumber(right));
        } else if (expr[pos] == '<') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) < toNumber(right));
        } else if (expr[pos] == '>') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) > toNumber(right));
        } else if (expr[pos] == '=') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) == toNumber(right));
        } else break;
        skipWhitespace(expr, pos);
    }
    return left;
}

QVariant FormulaEngine::evaluateTerm(const QString& expr, int& pos) {
    QVariant result = evaluateMultiplicative(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (expr[pos] == '+') { pos++; result = toNumber(result) + toNumber(evaluateMultiplicative(expr, pos)); }
        else if (expr[pos] == '-') { pos++; result = toNumber(result) - toNumber(evaluateMultiplicative(expr, pos)); }
        else break;
        skipWhitespace(expr, pos);
    }
    return result;
}

QVariant FormulaEngine::evaluateMultiplicative(const QString& expr, int& pos) {
    QVariant result = evaluateUnary(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (expr[pos] == '*') { pos++; result = toNumber(result) * toNumber(evaluateUnary(expr, pos)); }
        else if (expr[pos] == '/') {
            pos++; double d = toNumber(evaluateUnary(expr, pos));
            if (d == 0.0) { m_lastError = "Division by zero"; return QVariant("#DIV/0!"); }
            result = toNumber(result) / d;
        } else break;
        skipWhitespace(expr, pos);
    }
    return result;
}

QVariant FormulaEngine::evaluateUnary(const QString& expr, int& pos) {
    skipWhitespace(expr, pos);
    if (pos < expr.length() && expr[pos] == '-') { pos++; return -toNumber(evaluateUnary(expr, pos)); }
    return evaluatePower(expr, pos);
}

QVariant FormulaEngine::evaluatePower(const QString& expr, int& pos) {
    QVariant base = evaluateFactor(expr, pos);
    skipWhitespace(expr, pos);
    if (pos < expr.length() && expr[pos] == '^') {
        pos++; return std::pow(toNumber(base), toNumber(evaluateUnary(expr, pos)));
    }
    return base;
}

QVariant FormulaEngine::evaluateFactor(const QString& expr, int& pos) {
    skipWhitespace(expr, pos);
    if (pos >= expr.length()) return QVariant();

    // Numbers
    if (expr[pos].isDigit() || (expr[pos] == '.' && pos + 1 < expr.length() && expr[pos + 1].isDigit())) {
        int start = pos;
        while (pos < expr.length() && (expr[pos].isDigit() || expr[pos] == '.')) pos++;
        return expr.mid(start, pos - start).toDouble();
    }

    // Strings
    if (expr[pos] == '"') {
        pos++; int start = pos;
        while (pos < expr.length() && expr[pos] != '"') pos++;
        QString result = expr.mid(start, pos - start);
        if (pos < expr.length()) pos++;
        return result;
    }

    // Letter tokens: functions, cell refs, ranges
    if (pos < expr.length() && (expr[pos].isLetter() || expr[pos] == '$')) {
        int start = pos;
        while (pos < expr.length() && (expr[pos].isLetterOrNumber() || expr[pos] == ':' || expr[pos] == '$' || expr[pos] == '_')) pos++;
        QString token = expr.mid(start, pos - start);
        skipWhitespace(expr, pos);

        // Function call
        if (pos < expr.length() && expr[pos] == '(') {
            pos++;
            std::vector<QVariant> args;
            skipWhitespace(expr, pos);
            while (pos < expr.length() && expr[pos] != ')') {
                args.push_back(evaluateComparison(expr, pos));
                skipWhitespace(expr, pos);
                if (pos < expr.length() && expr[pos] == ',') pos++;
                skipWhitespace(expr, pos);
            }
            if (pos < expr.length() && expr[pos] == ')') pos++;
            return evaluateFunction(token.toUpper(), args);
        }

        // Range
        if (token.contains(':')) {
            CellRange range(token);
            m_lastRangeArgs.push_back(range);
            auto cells = range.getCells();
            for (const auto& c : cells) m_lastDependencies.push_back(c);
            std::vector<QVariant> values = getRangeValues(range);
            return QVariant::fromValue(values);
        }

        // Cell ref
        CellAddress addr = CellAddress::fromString(token);
        m_lastDependencies.push_back(addr);
        return getCellValue(addr);
    }

    // Parentheses
    if (pos < expr.length() && expr[pos] == '(') {
        pos++;
        QVariant result = evaluateComparison(expr, pos);
        skipWhitespace(expr, pos);
        if (pos < expr.length() && expr[pos] == ')') pos++;
        return result;
    }

    return QVariant();
}

QVariant FormulaEngine::evaluateFunction(const QString& fn, const std::vector<QVariant>& args) {
    // Aggregate
    if (fn == "SUM") return funcSUM(args);
    if (fn == "AVERAGE") return funcAVERAGE(args);
    if (fn == "COUNT") return funcCOUNT(args);
    if (fn == "COUNTA") return funcCOUNTA(args);
    if (fn == "MIN") return funcMIN(args);
    if (fn == "MAX") return funcMAX(args);
    if (fn == "IF") return funcIF(args);
    if (fn == "CONCAT" || fn == "CONCATENATE") return funcCONCAT(args);
    if (fn == "LEN") return funcLEN(args);
    if (fn == "UPPER") return funcUPPER(args);
    if (fn == "LOWER") return funcLOWER(args);
    if (fn == "TRIM") return funcTRIM(args);
    // Math
    if (fn == "ROUND") return funcROUND(args);
    if (fn == "ABS") return funcABS(args);
    if (fn == "SQRT") return funcSQRT(args);
    if (fn == "POWER") return funcPOWER(args);
    if (fn == "MOD") return funcMOD(args);
    if (fn == "INT") return funcINT(args);
    if (fn == "CEILING") return funcCEILING(args);
    if (fn == "FLOOR") return funcFLOOR(args);
    // Logical
    if (fn == "AND") return funcAND(args);
    if (fn == "OR") return funcOR(args);
    if (fn == "NOT") return funcNOT(args);
    if (fn == "IFERROR") return funcIFERROR(args);
    // Text
    if (fn == "LEFT") return funcLEFT(args);
    if (fn == "RIGHT") return funcRIGHT(args);
    if (fn == "MID") return funcMID(args);
    if (fn == "FIND") return funcFIND(args);
    if (fn == "SUBSTITUTE") return funcSUBSTITUTE(args);
    if (fn == "TEXT") return funcTEXT(args);
    // Statistical
    if (fn == "COUNTIF") return funcCOUNTIF(args);
    if (fn == "SUMIF") return funcSUMIF(args);
    // Date
    if (fn == "NOW") return funcNOW(args);
    if (fn == "TODAY") return funcTODAY(args);
    if (fn == "YEAR") return funcYEAR(args);
    if (fn == "MONTH") return funcMONTH(args);
    if (fn == "DAY") return funcDAY(args);
    if (fn == "DATE") return funcDATE(args);
    if (fn == "HOUR") return funcHOUR(args);
    if (fn == "MINUTE") return funcMINUTE(args);
    if (fn == "SECOND") return funcSECOND(args);
    if (fn == "DATEDIF") return funcDATEDIF(args);
    if (fn == "NETWORKDAYS") return funcNETWORKDAYS(args);
    if (fn == "WEEKDAY") return funcWEEKDAY(args);
    if (fn == "EDATE") return funcEDATE(args);
    if (fn == "EOMONTH") return funcEOMONTH(args);
    if (fn == "DATEVALUE") return funcDATEVALUE(args);
    // Lookup
    if (fn == "VLOOKUP") return funcVLOOKUP(args);
    if (fn == "HLOOKUP") return funcHLOOKUP(args);
    if (fn == "XLOOKUP") return funcXLOOKUP(args);
    if (fn == "INDEX") return funcINDEX(args);
    if (fn == "MATCH") return funcMATCH(args);
    // Additional statistical
    if (fn == "AVERAGEIF") return funcAVERAGEIF(args);
    if (fn == "COUNTBLANK") return funcCOUNTBLANK(args);
    if (fn == "SUMPRODUCT") return funcSUMPRODUCT(args);
    if (fn == "MEDIAN") return funcMEDIAN(args);
    if (fn == "MODE") return funcMODE(args);
    if (fn == "STDEV") return funcSTDEV(args);
    if (fn == "VAR") return funcVAR(args);
    if (fn == "LARGE") return funcLARGE(args);
    if (fn == "SMALL") return funcSMALL(args);
    if (fn == "RANK") return funcRANK(args);
    if (fn == "PERCENTILE") return funcPERCENTILE(args);
    // Additional math
    if (fn == "ROUNDUP") return funcROUNDUP(args);
    if (fn == "ROUNDDOWN") return funcROUNDDOWN(args);
    if (fn == "LOG") return funcLOG(args);
    if (fn == "LN") return funcLN(args);
    if (fn == "EXP") return funcEXP(args);
    if (fn == "RAND") return funcRAND(args);
    if (fn == "RANDBETWEEN") return funcRANDBETWEEN(args);
    // Additional text
    if (fn == "PROPER") return funcPROPER(args);
    if (fn == "SEARCH") return funcSEARCH(args);
    if (fn == "REPT") return funcREPT(args);
    if (fn == "EXACT") return funcEXACT(args);
    if (fn == "VALUE") return funcVALUE(args);
    // Additional logical/info
    if (fn == "ISBLANK") return funcISBLANK(args);
    if (fn == "ISERROR") return funcISERROR(args);
    if (fn == "ISNUMBER") return funcISNUMBER(args);
    if (fn == "ISTEXT") return funcISTEXT(args);
    if (fn == "CHOOSE") return funcCHOOSE(args);
    if (fn == "SWITCH") return funcSWITCH(args);

    m_lastError = "Unknown function: " + fn;
    return QVariant("#NAME?");
}

// ---- Aggregate functions ----

QVariant FormulaEngine::funcSUM(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double sum = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) sum += toNumber(a);
    return sum;
}

QVariant FormulaEngine::funcAVERAGE(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0; int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { sum += toNumber(a); count++; }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

QVariant FormulaEngine::funcCOUNT(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); int count = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            bool ok = false; a.toDouble(&ok);
            if (ok || a.typeId() == QMetaType::Int || a.typeId() == QMetaType::Double) count++;
        }
    }
    return count;
}

QVariant FormulaEngine::funcCOUNTA(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid() && !a.toString().isEmpty()) count++;
    return count;
}

QVariant FormulaEngine::funcMIN(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double min = std::numeric_limits<double>::max(); bool found = false;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { double v = toNumber(a); if (!found || v < min) { min = v; found = true; } }
    return found ? QVariant(min) : QVariant();
}

QVariant FormulaEngine::funcMAX(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double max = std::numeric_limits<double>::lowest(); bool found = false;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { double v = toNumber(a); if (!found || v > max) { max = v; found = true; } }
    return found ? QVariant(max) : QVariant();
}

QVariant FormulaEngine::funcIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    bool cond = toBoolean(args[0]);
    if (cond) return args[1];
    return args.size() >= 3 ? args[2] : QVariant(false);
}

QVariant FormulaEngine::funcCONCAT(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); QString result;
    for (const auto& a : flat) result += toString(a);
    return result;
}

QVariant FormulaEngine::funcLEN(const std::vector<QVariant>& args) {
    if (args.empty()) return 0;
    return static_cast<int>(toString(args[0]).length());
}

QVariant FormulaEngine::funcUPPER(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).toUpper()); }
QVariant FormulaEngine::funcLOWER(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).toLower()); }
QVariant FormulaEngine::funcTRIM(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).trimmed()); }

// ---- Math functions ----

QVariant FormulaEngine::funcROUND(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    int decimals = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 0;
    double factor = std::pow(10.0, decimals);
    return std::round(val * factor) / factor;
}

QVariant FormulaEngine::funcABS(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::abs(toNumber(args[0])));
}

QVariant FormulaEngine::funcSQRT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    return val < 0 ? QVariant("#NUM!") : QVariant(std::sqrt(val));
}

QVariant FormulaEngine::funcPOWER(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    return std::pow(toNumber(args[0]), toNumber(args[1]));
}

QVariant FormulaEngine::funcMOD(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double d = toNumber(args[1]);
    return d == 0.0 ? QVariant("#DIV/0!") : QVariant(std::fmod(toNumber(args[0]), d));
}

QVariant FormulaEngine::funcINT(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::floor(toNumber(args[0])));
}

QVariant FormulaEngine::funcCEILING(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    double sig = args.size() >= 2 ? toNumber(args[1]) : 1.0;
    if (sig == 0.0) return QVariant(0.0);
    return std::ceil(val / sig) * sig;
}

QVariant FormulaEngine::funcFLOOR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    double sig = args.size() >= 2 ? toNumber(args[1]) : 1.0;
    if (sig == 0.0) return QVariant(0.0);
    return std::floor(val / sig) * sig;
}

// ---- Logical functions ----

QVariant FormulaEngine::funcAND(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    for (const auto& a : flat) if (!toBoolean(a)) return QVariant(false);
    return QVariant(true);
}

QVariant FormulaEngine::funcOR(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    for (const auto& a : flat) if (toBoolean(a)) return QVariant(true);
    return QVariant(false);
}

QVariant FormulaEngine::funcNOT(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(!toBoolean(args[0]));
}

QVariant FormulaEngine::funcIFERROR(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString str = toString(args[0]);
    if (str.startsWith('#')) return args[1];
    return args[0];
}

// ---- Text functions ----

QVariant FormulaEngine::funcLEFT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int count = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    return toString(args[0]).left(count);
}

QVariant FormulaEngine::funcRIGHT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int count = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    return toString(args[0]).right(count);
}

QVariant FormulaEngine::funcMID(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QString str = toString(args[0]);
    int start = static_cast<int>(toNumber(args[1])) - 1; // 1-based to 0-based
    int count = static_cast<int>(toNumber(args[2]));
    return str.mid(start, count);
}

QVariant FormulaEngine::funcFIND(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString search = toString(args[0]);
    QString text = toString(args[1]);
    int startPos = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) - 1 : 0;
    int idx = text.indexOf(search, startPos);
    return idx >= 0 ? QVariant(idx + 1) : QVariant("#VALUE!"); // 1-based
}

QVariant FormulaEngine::funcSUBSTITUTE(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QString oldText = toString(args[1]);
    QString newText = toString(args[2]);
    return text.replace(oldText, newText);
}

QVariant FormulaEngine::funcTEXT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    QString fmt = toString(args[1]);
    if (fmt.contains('#') || fmt.contains('0')) {
        // Simple number formatting
        int decimals = 0;
        int dotPos = fmt.indexOf('.');
        if (dotPos >= 0) decimals = fmt.length() - dotPos - 1;
        return QString::number(val, 'f', decimals);
    }
    return QString::number(val);
}

// ---- Statistical functions ----

bool FormulaEngine::matchesCriteria(const QVariant& value, const QString& criteria) {
    if (criteria.startsWith(">=")) return toNumber(value) >= criteria.mid(2).toDouble();
    if (criteria.startsWith("<=")) return toNumber(value) <= criteria.mid(2).toDouble();
    if (criteria.startsWith("<>")) return value.toString() != criteria.mid(2);
    if (criteria.startsWith(">")) return toNumber(value) > criteria.mid(1).toDouble();
    if (criteria.startsWith("<")) return toNumber(value) < criteria.mid(1).toDouble();
    if (criteria.startsWith("=")) return value.toString() == criteria.mid(1);

    // Direct comparison
    bool ok = false;
    double critNum = criteria.toDouble(&ok);
    if (ok) return toNumber(value) == critNum;
    return value.toString() == criteria;
}

QVariant FormulaEngine::funcCOUNTIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    QString criteria = toString(args[1]);
    int count = 0;
    for (const auto& v : flat) if (matchesCriteria(v, criteria)) count++;
    return count;
}

QVariant FormulaEngine::funcSUMIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto range = flattenArgs({args[0]});
    QString criteria = toString(args[1]);
    auto sumRange = args.size() >= 3 ? flattenArgs({args[2]}) : range;
    double sum = 0;
    for (size_t i = 0; i < range.size() && i < sumRange.size(); ++i) {
        if (matchesCriteria(range[i], criteria)) sum += toNumber(sumRange[i]);
    }
    return sum;
}

// ---- Date functions ----

QVariant FormulaEngine::funcNOW(const std::vector<QVariant>&) {
    return QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
}

QVariant FormulaEngine::funcTODAY(const std::vector<QVariant>&) {
    return QDate::currentDate().toString("yyyy-MM-dd");
}

QVariant FormulaEngine::funcYEAR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.year()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcMONTH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.month()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcDAY(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.day()) : QVariant("#VALUE!");
}

// ---- Helper functions ----

double FormulaEngine::toNumber(const QVariant& value) {
    if (value.typeId() == QMetaType::Double || value.typeId() == QMetaType::Int) return value.toDouble();
    if (value.typeId() == QMetaType::Bool) return value.toBool() ? 1.0 : 0.0;
    bool ok = false; double r = value.toString().toDouble(&ok);
    return ok ? r : 0.0;
}

QString FormulaEngine::toString(const QVariant& value) { return value.toString(); }

bool FormulaEngine::toBoolean(const QVariant& value) {
    if (value.typeId() == QMetaType::Bool) return value.toBool();
    return toNumber(value) != 0;
}

QVariant FormulaEngine::getCellValue(const CellAddress& addr) {
    if (!m_spreadsheet) return QVariant();
    return m_spreadsheet->getCellValue(addr);
}

std::vector<QVariant> FormulaEngine::getRangeValues(const CellRange& range) {
    std::vector<QVariant> values;
    if (!m_spreadsheet) return values;
    for (const auto& addr : range.getCells()) values.push_back(m_spreadsheet->getCellValue(addr));
    return values;
}

std::vector<std::vector<QVariant>> FormulaEngine::getRangeValues2D(const CellRange& range) {
    std::vector<std::vector<QVariant>> result;
    if (!m_spreadsheet) return result;
    CellAddress start = range.getStart();
    CellAddress end = range.getEnd();
    for (int r = start.row; r <= end.row; ++r) {
        std::vector<QVariant> row;
        for (int c = start.col; c <= end.col; ++c) {
            row.push_back(m_spreadsheet->getCellValue(CellAddress(r, c)));
        }
        result.push_back(std::move(row));
    }
    return result;
}

QDate FormulaEngine::parseDate(const QVariant& value) {
    QString str = toString(value);
    QDate d = QDate::fromString(str, "yyyy-MM-dd");
    if (d.isValid()) return d;
    d = QDate::fromString(str, "MM/dd/yyyy");
    if (d.isValid()) return d;
    d = QDate::fromString(str, "dd/MM/yyyy");
    if (d.isValid()) return d;
    d = QDate::fromString(str, "yyyy/MM/dd");
    return d;
}

// ---- Lookup functions ----

QVariant FormulaEngine::funcVLOOKUP(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QVariant lookupVal = args[0];
    int colIdx = static_cast<int>(toNumber(args[2]));
    bool rangeLookup = args.size() >= 4 ? toBoolean(args[3]) : true;

    // Find the range from m_lastRangeArgs (second arg should be the table range)
    if (m_lastRangeArgs.size() < 1) return QVariant("#REF!");
    // The table range is typically the first range arg encountered
    CellRange tableRange = m_lastRangeArgs[0];
    auto table = getRangeValues2D(tableRange);

    if (colIdx < 1 || colIdx > static_cast<int>(table.empty() ? 0 : table[0].size()))
        return QVariant("#REF!");

    for (size_t r = 0; r < table.size(); ++r) {
        if (table[r].empty()) continue;
        QVariant cellVal = table[r][0];
        bool match = false;
        if (!rangeLookup) {
            // Exact match
            match = (cellVal.toString().compare(lookupVal.toString(), Qt::CaseInsensitive) == 0);
            if (!match) {
                bool ok1, ok2;
                double d1 = cellVal.toDouble(&ok1);
                double d2 = lookupVal.toDouble(&ok2);
                if (ok1 && ok2) match = (d1 == d2);
            }
        } else {
            // Approximate match (sorted ascending) - find largest value <= lookup
            double cv = toNumber(cellVal);
            double lv = toNumber(lookupVal);
            if (cv <= lv) {
                // Check if next row exceeds
                if (r + 1 >= table.size() || toNumber(table[r + 1][0]) > lv) {
                    match = true;
                }
            }
        }
        if (match) return table[r][colIdx - 1];
    }
    return QVariant("#N/A");
}

QVariant FormulaEngine::funcHLOOKUP(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QVariant lookupVal = args[0];
    int rowIdx = static_cast<int>(toNumber(args[2]));
    bool rangeLookup = args.size() >= 4 ? toBoolean(args[3]) : true;

    if (m_lastRangeArgs.size() < 1) return QVariant("#REF!");
    CellRange tableRange = m_lastRangeArgs[0];
    auto table = getRangeValues2D(tableRange);

    if (table.empty() || rowIdx < 1 || rowIdx > static_cast<int>(table.size()))
        return QVariant("#REF!");

    // Search first row
    for (size_t c = 0; c < table[0].size(); ++c) {
        QVariant cellVal = table[0][c];
        bool match = false;
        if (!rangeLookup) {
            match = (cellVal.toString().compare(lookupVal.toString(), Qt::CaseInsensitive) == 0);
        } else {
            double cv = toNumber(cellVal);
            double lv = toNumber(lookupVal);
            if (cv <= lv) {
                if (c + 1 >= table[0].size() || toNumber(table[0][c + 1]) > lv)
                    match = true;
            }
        }
        if (match) return table[rowIdx - 1][c];
    }
    return QVariant("#N/A");
}

QVariant FormulaEngine::funcXLOOKUP(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QVariant lookupVal = args[0];
    QVariant ifNotFound = args.size() >= 4 ? args[3] : QVariant("#N/A");

    if (m_lastRangeArgs.size() < 2) return QVariant("#REF!");
    CellRange lookupRange = m_lastRangeArgs[0];
    CellRange returnRange = m_lastRangeArgs[1];

    auto lookupVals = getRangeValues(lookupRange);
    auto returnVals = getRangeValues(returnRange);

    for (size_t i = 0; i < lookupVals.size() && i < returnVals.size(); ++i) {
        bool match = (lookupVals[i].toString().compare(lookupVal.toString(), Qt::CaseInsensitive) == 0);
        if (!match) {
            bool ok1, ok2;
            double d1 = lookupVals[i].toDouble(&ok1);
            double d2 = lookupVal.toDouble(&ok2);
            if (ok1 && ok2) match = (d1 == d2);
        }
        if (match) return returnVals[i];
    }
    return ifNotFound;
}

QVariant FormulaEngine::funcINDEX(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int rowNum = static_cast<int>(toNumber(args[1]));
    int colNum = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) : 1;

    if (m_lastRangeArgs.empty()) return QVariant("#REF!");
    CellRange range = m_lastRangeArgs[0];
    auto table = getRangeValues2D(range);

    if (rowNum < 1 || rowNum > static_cast<int>(table.size())) return QVariant("#REF!");
    if (colNum < 1 || colNum > static_cast<int>(table[0].size())) return QVariant("#REF!");
    return table[rowNum - 1][colNum - 1];
}

QVariant FormulaEngine::funcMATCH(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QVariant lookupVal = args[0];
    int matchType = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) : 1;

    if (m_lastRangeArgs.empty()) return QVariant("#N/A");
    CellRange range = m_lastRangeArgs[0];
    auto values = getRangeValues(range);

    if (matchType == 0) {
        // Exact match
        for (size_t i = 0; i < values.size(); ++i) {
            if (values[i].toString().compare(lookupVal.toString(), Qt::CaseInsensitive) == 0)
                return static_cast<int>(i + 1);
            bool ok1, ok2;
            double d1 = values[i].toDouble(&ok1), d2 = lookupVal.toDouble(&ok2);
            if (ok1 && ok2 && d1 == d2) return static_cast<int>(i + 1);
        }
    } else if (matchType == 1) {
        // Largest value <= lookup (sorted ascending)
        int lastMatch = -1;
        for (size_t i = 0; i < values.size(); ++i) {
            if (toNumber(values[i]) <= toNumber(lookupVal)) lastMatch = static_cast<int>(i);
        }
        if (lastMatch >= 0) return lastMatch + 1;
    } else {
        // Smallest value >= lookup (sorted descending)
        int lastMatch = -1;
        for (size_t i = 0; i < values.size(); ++i) {
            if (toNumber(values[i]) >= toNumber(lookupVal)) lastMatch = static_cast<int>(i);
        }
        if (lastMatch >= 0) return lastMatch + 1;
    }
    return QVariant("#N/A");
}

// ---- Additional statistical functions ----

QVariant FormulaEngine::funcAVERAGEIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto range = flattenArgs({args[0]});
    QString criteria = toString(args[1]);
    auto avgRange = args.size() >= 3 ? flattenArgs({args[2]}) : range;
    double sum = 0; int count = 0;
    for (size_t i = 0; i < range.size() && i < avgRange.size(); ++i) {
        if (matchesCriteria(range[i], criteria)) { sum += toNumber(avgRange[i]); count++; }
    }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

QVariant FormulaEngine::funcCOUNTBLANK(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    int count = 0;
    for (const auto& v : flat) {
        if (v.isNull() || !v.isValid() || v.toString().isEmpty()) count++;
    }
    return count;
}

QVariant FormulaEngine::funcSUMPRODUCT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    std::vector<std::vector<QVariant>> arrays;
    for (const auto& arg : args) {
        if (arg.canConvert<std::vector<QVariant>>()) {
            arrays.push_back(arg.value<std::vector<QVariant>>());
        } else {
            arrays.push_back({arg});
        }
    }
    if (arrays.empty()) return 0.0;
    size_t len = arrays[0].size();
    for (const auto& arr : arrays) {
        if (arr.size() != len) return QVariant("#VALUE!");
    }
    double sum = 0;
    for (size_t i = 0; i < len; ++i) {
        double product = 1.0;
        for (const auto& arr : arrays) {
            product *= toNumber(arr[i]);
        }
        sum += product;
    }
    return sum;
}

QVariant FormulaEngine::funcMEDIAN(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    std::vector<double> nums;
    for (const auto& v : flat) {
        if (!v.isNull() && v.isValid()) {
            bool ok; double d = v.toDouble(&ok);
            if (ok) nums.push_back(d);
        }
    }
    if (nums.empty()) return QVariant("#NUM!");
    std::sort(nums.begin(), nums.end());
    size_t n = nums.size();
    if (n % 2 == 1) return nums[n / 2];
    return (nums[n / 2 - 1] + nums[n / 2]) / 2.0;
}

QVariant FormulaEngine::funcMODE(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    std::map<double, int> freq;
    for (const auto& v : flat) {
        if (!v.isNull() && v.isValid()) {
            bool ok; double d = v.toDouble(&ok);
            if (ok) freq[d]++;
        }
    }
    if (freq.empty()) return QVariant("#N/A");
    int maxCount = 0; double modeVal = 0;
    for (const auto& [val, cnt] : freq) {
        if (cnt > maxCount) { maxCount = cnt; modeVal = val; }
    }
    if (maxCount <= 1) return QVariant("#N/A");
    return modeVal;
}

QVariant FormulaEngine::funcSTDEV(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    std::vector<double> nums;
    for (const auto& v : flat) {
        if (!v.isNull() && v.isValid()) {
            bool ok; double d = v.toDouble(&ok);
            if (ok) nums.push_back(d);
        }
    }
    if (nums.size() < 2) return QVariant("#DIV/0!");
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / nums.size();
    double sq_sum = 0;
    for (double x : nums) sq_sum += (x - mean) * (x - mean);
    return std::sqrt(sq_sum / (nums.size() - 1)); // sample std dev
}

QVariant FormulaEngine::funcVAR(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    std::vector<double> nums;
    for (const auto& v : flat) {
        if (!v.isNull() && v.isValid()) {
            bool ok; double d = v.toDouble(&ok);
            if (ok) nums.push_back(d);
        }
    }
    if (nums.size() < 2) return QVariant("#DIV/0!");
    double mean = std::accumulate(nums.begin(), nums.end(), 0.0) / nums.size();
    double sq_sum = 0;
    for (double x : nums) sq_sum += (x - mean) * (x - mean);
    return sq_sum / (nums.size() - 1); // sample variance
}

QVariant FormulaEngine::funcLARGE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    int k = static_cast<int>(toNumber(args[1]));
    std::vector<double> nums;
    for (const auto& v : flat) {
        bool ok; double d = v.toDouble(&ok);
        if (ok) nums.push_back(d);
    }
    if (k < 1 || k > static_cast<int>(nums.size())) return QVariant("#NUM!");
    std::sort(nums.rbegin(), nums.rend());
    return nums[k - 1];
}

QVariant FormulaEngine::funcSMALL(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    int k = static_cast<int>(toNumber(args[1]));
    std::vector<double> nums;
    for (const auto& v : flat) {
        bool ok; double d = v.toDouble(&ok);
        if (ok) nums.push_back(d);
    }
    if (k < 1 || k > static_cast<int>(nums.size())) return QVariant("#NUM!");
    std::sort(nums.begin(), nums.end());
    return nums[k - 1];
}

QVariant FormulaEngine::funcRANK(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double number = toNumber(args[0]);
    auto flat = flattenArgs({args[1]});
    bool ascending = args.size() >= 3 ? toBoolean(args[2]) : false;
    std::vector<double> nums;
    for (const auto& v : flat) {
        bool ok; double d = v.toDouble(&ok);
        if (ok) nums.push_back(d);
    }
    if (ascending) std::sort(nums.begin(), nums.end());
    else std::sort(nums.rbegin(), nums.rend());
    for (size_t i = 0; i < nums.size(); ++i) {
        if (nums[i] == number) return static_cast<int>(i + 1);
    }
    return QVariant("#N/A");
}

QVariant FormulaEngine::funcPERCENTILE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    double k = toNumber(args[1]);
    if (k < 0 || k > 1) return QVariant("#NUM!");
    std::vector<double> nums;
    for (const auto& v : flat) {
        bool ok; double d = v.toDouble(&ok);
        if (ok) nums.push_back(d);
    }
    if (nums.empty()) return QVariant("#NUM!");
    std::sort(nums.begin(), nums.end());
    double idx = k * (nums.size() - 1);
    int lower = static_cast<int>(std::floor(idx));
    int upper = static_cast<int>(std::ceil(idx));
    if (lower == upper) return nums[lower];
    double frac = idx - lower;
    return nums[lower] + frac * (nums[upper] - nums[lower]);
}

// ---- Additional math functions ----

QVariant FormulaEngine::funcROUNDUP(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    int decimals = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 0;
    double factor = std::pow(10.0, decimals);
    return (val >= 0) ? std::ceil(val * factor) / factor : std::floor(val * factor) / factor;
}

QVariant FormulaEngine::funcROUNDDOWN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    int decimals = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 0;
    double factor = std::pow(10.0, decimals);
    return (val >= 0) ? std::floor(val * factor) / factor : std::ceil(val * factor) / factor;
}

QVariant FormulaEngine::funcLOG(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    if (val <= 0) return QVariant("#NUM!");
    double base = args.size() >= 2 ? toNumber(args[1]) : 10.0;
    if (base <= 0 || base == 1) return QVariant("#NUM!");
    return std::log(val) / std::log(base);
}

QVariant FormulaEngine::funcLN(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    return val <= 0 ? QVariant("#NUM!") : QVariant(std::log(val));
}

QVariant FormulaEngine::funcEXP(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    return std::exp(toNumber(args[0]));
}

QVariant FormulaEngine::funcRAND(const std::vector<QVariant>&) {
    static std::mt19937 gen(std::random_device{}());
    static std::uniform_real_distribution<double> dist(0.0, 1.0);
    return dist(gen);
}

QVariant FormulaEngine::funcRANDBETWEEN(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int low = static_cast<int>(toNumber(args[0]));
    int high = static_cast<int>(toNumber(args[1]));
    if (low > high) return QVariant("#VALUE!");
    static std::mt19937 gen(std::random_device{}());
    std::uniform_int_distribution<int> dist(low, high);
    return dist(gen);
}

// ---- Additional text functions ----

QVariant FormulaEngine::funcPROPER(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString str = toString(args[0]).toLower();
    bool capitalizeNext = true;
    for (int i = 0; i < str.length(); ++i) {
        if (capitalizeNext && str[i].isLetter()) {
            str[i] = str[i].toUpper();
            capitalizeNext = false;
        } else if (!str[i].isLetterOrNumber()) {
            capitalizeNext = true;
        }
    }
    return str;
}

QVariant FormulaEngine::funcSEARCH(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString search = toString(args[0]).toLower();
    QString text = toString(args[1]).toLower();
    int startPos = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) - 1 : 0;
    int idx = text.indexOf(search, startPos);
    return idx >= 0 ? QVariant(idx + 1) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcREPT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString str = toString(args[0]);
    int times = static_cast<int>(toNumber(args[1]));
    if (times < 0) return QVariant("#VALUE!");
    return str.repeated(times);
}

QVariant FormulaEngine::funcEXACT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    return QVariant(toString(args[0]) == toString(args[1]));
}

QVariant FormulaEngine::funcVALUE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QString str = toString(args[0]).trimmed();
    str.remove(',').remove('$').remove('%').remove(' ');
    bool ok; double d = str.toDouble(&ok);
    return ok ? QVariant(d) : QVariant("#VALUE!");
}

// ---- Additional logical/info functions ----

QVariant FormulaEngine::funcISBLANK(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant(true);
    return QVariant(args[0].isNull() || !args[0].isValid() || args[0].toString().isEmpty());
}

QVariant FormulaEngine::funcISERROR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant(false);
    return QVariant(toString(args[0]).startsWith('#'));
}

QVariant FormulaEngine::funcISNUMBER(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant(false);
    bool ok; args[0].toDouble(&ok);
    return QVariant(ok || args[0].typeId() == QMetaType::Int || args[0].typeId() == QMetaType::Double);
}

QVariant FormulaEngine::funcISTEXT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant(false);
    if (args[0].isNull() || !args[0].isValid()) return QVariant(false);
    bool ok; args[0].toDouble(&ok);
    return QVariant(!ok && args[0].typeId() != QMetaType::Bool);
}

QVariant FormulaEngine::funcCHOOSE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    int idx = static_cast<int>(toNumber(args[0]));
    if (idx < 1 || idx >= static_cast<int>(args.size())) return QVariant("#VALUE!");
    return args[idx];
}

QVariant FormulaEngine::funcSWITCH(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QVariant expr = args[0];
    // SWITCH(expr, val1, result1, val2, result2, ..., [default])
    for (size_t i = 1; i + 1 < args.size(); i += 2) {
        if (toString(expr) == toString(args[i])) return args[i + 1];
    }
    // If odd number of remaining args, last is default
    if (args.size() % 2 == 0) return args.back();
    return QVariant("#N/A");
}

// ---- Additional date functions ----

QVariant FormulaEngine::funcDATE(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    int year = static_cast<int>(toNumber(args[0]));
    int month = static_cast<int>(toNumber(args[1]));
    int day = static_cast<int>(toNumber(args[2]));
    QDate date(year, month, day);
    return date.isValid() ? QVariant(date.toString("yyyy-MM-dd")) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcHOUR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDateTime dt = QDateTime::fromString(toString(args[0]), "yyyy-MM-dd hh:mm:ss");
    if (!dt.isValid()) dt = QDateTime::fromString(toString(args[0]), "hh:mm:ss");
    return dt.isValid() ? QVariant(dt.time().hour()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcMINUTE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDateTime dt = QDateTime::fromString(toString(args[0]), "yyyy-MM-dd hh:mm:ss");
    if (!dt.isValid()) dt = QDateTime::fromString(toString(args[0]), "hh:mm:ss");
    return dt.isValid() ? QVariant(dt.time().minute()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcSECOND(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDateTime dt = QDateTime::fromString(toString(args[0]), "yyyy-MM-dd hh:mm:ss");
    if (!dt.isValid()) dt = QDateTime::fromString(toString(args[0]), "hh:mm:ss");
    return dt.isValid() ? QVariant(dt.time().second()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcDATEDIF(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QDate start = parseDate(args[0]);
    QDate end = parseDate(args[1]);
    QString unit = toString(args[2]).toUpper();
    if (!start.isValid() || !end.isValid()) return QVariant("#VALUE!");
    if (start > end) return QVariant("#NUM!");
    if (unit == "D") return static_cast<int>(start.daysTo(end));
    if (unit == "M") return (end.year() - start.year()) * 12 + end.month() - start.month();
    if (unit == "Y") return end.year() - start.year() - (end < QDate(end.year(), start.month(), start.day()) ? 1 : 0);
    return QVariant("#VALUE!");
}

QVariant FormulaEngine::funcNETWORKDAYS(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QDate start = parseDate(args[0]);
    QDate end = parseDate(args[1]);
    if (!start.isValid() || !end.isValid()) return QVariant("#VALUE!");
    int sign = 1;
    if (start > end) { std::swap(start, end); sign = -1; }
    int days = 0;
    QDate d = start;
    while (d <= end) {
        int dow = d.dayOfWeek();
        if (dow != 6 && dow != 7) days++;
        d = d.addDays(1);
    }
    return days * sign;
}

QVariant FormulaEngine::funcWEEKDAY(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = parseDate(args[0]);
    if (!date.isValid()) return QVariant("#VALUE!");
    int returnType = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    int dow = date.dayOfWeek(); // Qt: 1=Mon..7=Sun
    if (returnType == 1) return (dow % 7) + 1; // 1=Sun..7=Sat
    if (returnType == 2) return dow;             // 1=Mon..7=Sun
    if (returnType == 3) return dow - 1;         // 0=Mon..6=Sun
    return QVariant("#VALUE!");
}

QVariant FormulaEngine::funcEDATE(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QDate date = parseDate(args[0]);
    if (!date.isValid()) return QVariant("#VALUE!");
    int months = static_cast<int>(toNumber(args[1]));
    return date.addMonths(months).toString("yyyy-MM-dd");
}

QVariant FormulaEngine::funcEOMONTH(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QDate date = parseDate(args[0]);
    if (!date.isValid()) return QVariant("#VALUE!");
    int months = static_cast<int>(toNumber(args[1]));
    QDate result = date.addMonths(months);
    result = QDate(result.year(), result.month(), result.daysInMonth());
    return result.toString("yyyy-MM-dd");
}

QVariant FormulaEngine::funcDATEVALUE(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = parseDate(args[0]);
    if (!date.isValid()) return QVariant("#VALUE!");
    return date.toString("yyyy-MM-dd");
}
