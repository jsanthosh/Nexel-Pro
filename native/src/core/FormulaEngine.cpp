#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <cctype>
#include <limits>
#include <numeric>
#include <random>
#include <set>
#include <map>
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
    m_lastColumnDeps.clear();

    if (formula.isEmpty()) return QVariant();

    QString expr = formula.startsWith('=') ? formula.mid(1) : formula;

    // Use AST for all formulas except array constants ({})
    bool useAST = !expr.contains('{');

    if (useAST) {
        try {
            uint32_t root = FormulaASTPool::instance().parse(formula);
            QVariant result = evaluateAST(root);
            // If AST evaluation hit a parse error node, fall back to old parser
            // Only check string result for error types (avoid expensive toString on complex types)
            if (result.typeId() == QMetaType::QString) {
                QString resultStr = result.toString();
                if (resultStr == QStringLiteral("#PARSE!")) {
                    m_lastDependencies.clear();
                    m_lastRangeArgs.clear();
                    m_lastColumnDeps.clear();
                    return parseExpression(expr);
                }
            }
            return result;
        } catch (...) {
            // Fall back to old parser on any AST evaluation failure
            m_lastDependencies.clear();
            m_lastRangeArgs.clear();
            m_lastColumnDeps.clear();
        }
    }

    // Fallback: original string-parsing evaluation
    try {
        return parseExpression(expr);
    } catch (const std::exception& e) {
        m_lastError = QString::fromStdString(e.what());
        return QVariant("#ERROR!");
    }
}

// ============================================================================
// AST Evaluator — walks cached AST nodes, no string re-parsing
// ============================================================================
// Parse happens once per unique formula string (cached in FormulaASTPool).
// Evaluation walks the flat node array — ~10x faster than re-parsing.

QVariant FormulaEngine::evaluateAST(uint32_t nodeIndex) {
    auto& pool = FormulaASTPool::instance();
    const ASTNode& node = pool.getNode(nodeIndex);

    switch (node.type) {
        case ASTNodeType::Literal:
            return pool.getLiteral(node.literalIndex);

        case ASTNodeType::Boolean:
            return QVariant(node.boolValue);

        case ASTNodeType::Error:
            return pool.getLiteral(node.errorIndex);

        case ASTNodeType::CellRef: {
            CellAddress addr(node.cellRef.row, node.cellRef.col);
            m_lastDependencies.push_back(addr);
            return getCellValue(addr);
        }

        case ASTNodeType::RangeRef: {
            CellRange range(
                CellAddress(node.rangeRef.startRow, node.rangeRef.startCol),
                CellAddress(node.rangeRef.endRow, node.rangeRef.endCol));
            m_lastRangeArgs.push_back(range);

            // Track individual cell dependencies for small ranges
            long long cellCount = static_cast<long long>(range.getRowCount()) * range.getColumnCount();
            if (cellCount < 10000) {
                auto cells = range.getCells();
                for (const auto& c : cells)
                    m_lastDependencies.push_back(c);
            }

            // For large ranges, return lazy CellRange for streaming evaluation
            if (cellCount > 100000) {
                return QVariant::fromValue(range);
            }

            return QVariant::fromValue(getRangeValues(range));
        }

        case ASTNodeType::UnaryOp: {
            QVariant operand = evaluateAST(node.unary.operand);
            if (node.unary.op == UnaryOp::Negate)
                return QVariant(-toNumber(operand));
            return operand;
        }

        case ASTNodeType::BinaryOp: {
            QVariant left = evaluateAST(node.binary.left);
            QVariant right = evaluateAST(node.binary.right);

            switch (node.binary.op) {
                case BinaryOp::Add: return QVariant(toNumber(left) + toNumber(right));
                case BinaryOp::Sub: return QVariant(toNumber(left) - toNumber(right));
                case BinaryOp::Mul: return QVariant(toNumber(left) * toNumber(right));
                case BinaryOp::Div: {
                    double d = toNumber(right);
                    if (d == 0.0) return QVariant(QStringLiteral("#DIV/0!"));
                    return QVariant(toNumber(left) / d);
                }
                case BinaryOp::Pow:
                    return QVariant(std::pow(toNumber(left), toNumber(right)));
                case BinaryOp::Concat:
                    return QVariant(toString(left) + toString(right));
                case BinaryOp::Eq:
                    return QVariant(toString(left).compare(toString(right), Qt::CaseInsensitive) == 0);
                case BinaryOp::Neq:
                    return QVariant(toString(left).compare(toString(right), Qt::CaseInsensitive) != 0);
                case BinaryOp::Lt:  return QVariant(toNumber(left) < toNumber(right));
                case BinaryOp::Gt:  return QVariant(toNumber(left) > toNumber(right));
                case BinaryOp::Lte: return QVariant(toNumber(left) <= toNumber(right));
                case BinaryOp::Gte: return QVariant(toNumber(left) >= toNumber(right));
            }
            return QVariant(QStringLiteral("#ERROR!"));
        }

        case ASTNodeType::FunctionCall: {
            const QVariant& argData = pool.getLiteral(node.func.argStart);
            QVariantList argNodeIndices = argData.toList();
            return evaluateASTFunction(node.func.funcId, argNodeIndices);
        }

        case ASTNodeType::ColumnRef: {
            // Column reference D:D → expand to D1:D10000000 (sparse iteration makes this fast)
            int startCol = node.colRef.startCol;
            int endCol = node.colRef.endCol;
            for (int c = startCol; c <= endCol; c++) {
                m_lastColumnDeps.push_back(c);
            }
            CellRange range(CellAddress(0, startCol), CellAddress(9999999, endCol));
            m_lastRangeArgs.push_back(range);
            return QVariant::fromValue(range);  // lazy — streaming evaluation
        }

        case ASTNodeType::CrossSheetCell: {
            QString sheetName = pool.getLiteral(node.crossSheetCell.sheetNameIndex).toString();
            CellAddress addr(node.crossSheetCell.row, node.crossSheetCell.col);
            if (m_allSheets) {
                for (const auto& sheet : *m_allSheets) {
                    if (sheet->getSheetName().compare(sheetName, Qt::CaseInsensitive) == 0) {
                        return sheet->getCellValue(addr);
                    }
                }
            }
            return QVariant(QStringLiteral("#REF!"));
        }

        case ASTNodeType::CrossSheetRange: {
            QString sheetName = pool.getLiteral(node.crossSheetRange.sheetNameIndex).toString();
            CellRange range(
                CellAddress(node.crossSheetRange.startRow, node.crossSheetRange.startCol),
                CellAddress(node.crossSheetRange.endRow, node.crossSheetRange.endCol));
            m_lastRangeArgs.push_back(range);
            if (m_allSheets) {
                for (const auto& sheet : *m_allSheets) {
                    if (sheet->getSheetName().compare(sheetName, Qt::CaseInsensitive) == 0) {
                        std::vector<QVariant> values;
                        auto cells = range.getCells();
                        for (const auto& c : cells) values.push_back(sheet->getCellValue(c));
                        return QVariant::fromValue(values);
                    }
                }
            }
            return QVariant(QStringLiteral("#REF!"));
        }
    }

    return QVariant(QStringLiteral("#ERROR!"));
}

QVariant FormulaEngine::evaluateASTFunction(uint16_t funcId, const QVariantList& argNodeIndices) {
    auto& pool = FormulaASTPool::instance();
    QString funcName = pool.getFunctionName(funcId);

    // Evaluate all arguments by walking their AST subtrees
    std::vector<QVariant> args;
    args.reserve(argNodeIndices.size());
    for (const auto& idx : argNodeIndices) {
        args.push_back(evaluateAST(idx.toUInt()));
    }

    // Dispatch to existing function implementations (unchanged)
    return evaluateFunction(funcName, args);
}

QString FormulaEngine::getLastError() const { return m_lastError; }
bool FormulaEngine::hasError() const { return !m_lastError.isEmpty(); }
void FormulaEngine::clearCache() {
    m_cache.clear();
    FormulaASTPool::instance().clear();
}
void FormulaEngine::invalidateCell(const CellAddress& addr) { m_cache.erase(addr.toString().toStdString()); }

void FormulaEngine::skipWhitespace(const QString& expr, int& pos) {
    while (pos < expr.length() && expr[pos].isSpace()) pos++;
}

std::vector<QVariant> FormulaEngine::flattenArgs(const std::vector<QVariant>& args) {
    std::vector<QVariant> flat;
    for (const auto& arg : args) {
        if (arg.canConvert<CellRange>()) {
            // Lazy range: materialize values by streaming from spreadsheet
            streamRangeValues(arg.value<CellRange>(), [&](const QVariant& v) {
                flat.push_back(v);
            });
        } else if (arg.canConvert<std::vector<QVariant>>()) {
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
            if (d == 0.0) { m_lastError = "#DIV/0!"; return QVariant("#DIV/0!"); }
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

    // Cross-sheet reference with quoted sheet name: 'Sheet Name'!A1 or 'Sheet Name'!A1:B10
    if (pos < expr.length() && expr[pos] == '\'') {
        int savedPos = pos;
        pos++; // skip opening quote
        int nameStart = pos;
        while (pos < expr.length() && expr[pos] != '\'') pos++;
        QString sheetName = expr.mid(nameStart, pos - nameStart);
        if (pos < expr.length()) pos++; // skip closing quote
        if (pos < expr.length() && expr[pos] == '!') {
            pos++; // skip '!'
            int refStart = pos;
            while (pos < expr.length() && (expr[pos].isLetterOrNumber() || expr[pos] == ':' || expr[pos] == '$')) pos++;
            QString refToken = expr.mid(refStart, pos - refStart);

            if (m_allSheets) {
                for (const auto& sheet : *m_allSheets) {
                    if (sheet->getSheetName().compare(sheetName, Qt::CaseInsensitive) == 0) {
                        if (refToken.contains(':')) {
                            CellRange range(refToken);
                            m_lastRangeArgs.push_back(range);
                            std::vector<QVariant> values;
                            auto cells = range.getCells();
                            for (const auto& c : cells) values.push_back(sheet->getCellValue(c));
                            return QVariant::fromValue(values);
                        } else {
                            CellAddress addr = CellAddress::fromString(refToken);
                            return sheet->getCellValue(addr);
                        }
                    }
                }
            }
            m_lastError = "#REF!";
            return QVariant("#REF!");
        }
        // Not a valid cross-sheet ref, reset position
        pos = savedPos;
    }

    // Letter tokens: functions, cell refs, ranges, cross-sheet refs, named ranges
    if (pos < expr.length() && (expr[pos].isLetter() || expr[pos] == '$')) {
        int start = pos;
        while (pos < expr.length() && (expr[pos].isLetterOrNumber() || expr[pos] == ':' || expr[pos] == '$' || expr[pos] == '_')) pos++;
        QString token = expr.mid(start, pos - start);
        skipWhitespace(expr, pos);

        // Cross-sheet reference: SheetName!CellRef or SheetName!A1:B10
        if (pos < expr.length() && expr[pos] == '!') {
            pos++; // skip '!'
            QString sheetName = token;
            int refStart = pos;
            while (pos < expr.length() && (expr[pos].isLetterOrNumber() || expr[pos] == ':' || expr[pos] == '$')) pos++;
            QString refToken = expr.mid(refStart, pos - refStart);

            if (m_allSheets) {
                for (const auto& sheet : *m_allSheets) {
                    if (sheet->getSheetName().compare(sheetName, Qt::CaseInsensitive) == 0) {
                        if (refToken.contains(':')) {
                            CellRange range(refToken);
                            m_lastRangeArgs.push_back(range);
                            std::vector<QVariant> values;
                            auto cells = range.getCells();
                            for (const auto& c : cells) values.push_back(sheet->getCellValue(c));
                            return QVariant::fromValue(values);
                        } else {
                            CellAddress addr = CellAddress::fromString(refToken);
                            return sheet->getCellValue(addr);
                        }
                    }
                }
            }
            m_lastError = "#REF!";
            return QVariant("#REF!");
        }

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
            QString rangeToken = token;

            // Handle column references (D:D, $A:$Z, etc.) — letters only, no digits
            if (m_spreadsheet) {
                QString stripped = rangeToken;
                stripped.remove('$');
                QStringList parts = stripped.split(':');
                if (parts.size() == 2) {
                    auto allLetters = [](const QString& s) {
                        if (s.isEmpty()) return false;
                        for (QChar ch : s) if (!ch.isLetter()) return false;
                        return true;
                    };
                    if (allLetters(parts[0]) && allLetters(parts[1])) {
                        // Expand to large fixed range — sparse iteration makes this fast
                        rangeToken = parts[0] + "1:" + parts[1] + "10000000";
                        // Track column-level dependencies for recalculation
                        int startCol = CellAddress::fromString(parts[0] + "1").col;
                        int endCol = CellAddress::fromString(parts[1] + "1").col;
                        for (int c = startCol; c <= endCol; c++) {
                            m_lastColumnDeps.push_back(c);
                        }
                    }
                }
            }

            CellRange range(rangeToken);
            m_lastRangeArgs.push_back(range);

            // Only track individual cell dependencies for small ranges
            long long cellCount = static_cast<long long>(range.getRowCount()) * range.getColumnCount();
            if (cellCount < 10000) {
                auto cells = range.getCells();
                for (const auto& c : cells) m_lastDependencies.push_back(c);
            }

            // For very large ranges, store lazy reference to avoid materializing millions of values
            if (cellCount > 100000) {
                return QVariant::fromValue(range);
            }

            std::vector<QVariant> values = getRangeValues(range);
            return QVariant::fromValue(values);
        }

        // Named range lookup (case-insensitive) — check before treating as cell ref
        if (m_spreadsheet) {
            const auto* namedRange = m_spreadsheet->getNamedRange(token);
            if (namedRange) {
                CellRange range = namedRange->range;
                m_lastRangeArgs.push_back(range);

                long long cellCount = static_cast<long long>(range.getRowCount()) * range.getColumnCount();
                if (cellCount < 10000) {
                    auto cells = range.getCells();
                    for (const auto& c : cells) m_lastDependencies.push_back(c);
                }

                if (range.isSingleCell()) {
                    CellAddress addr = range.getStart();
                    m_lastDependencies.push_back(addr);
                    return getCellValue(addr);
                }

                if (cellCount > 100000) {
                    return QVariant::fromValue(range);
                }

                std::vector<QVariant> values = getRangeValues(range);
                return QVariant::fromValue(values);
            }
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
    if (fn == "IFS") return funcIFS(args);

    if (fn == "SUMIFS") return funcSUMIFS(args);
    if (fn == "COUNTIFS") return funcCOUNTIFS(args);
    if (fn == "AVERAGEIFS") return funcAVERAGEIFS(args);
    if (fn == "MINIFS") return funcMINIFS(args);
    if (fn == "MAXIFS") return funcMAXIFS(args);

    if (fn == "TEXTJOIN") return funcTEXTJOIN(args);
    if (fn == "REGEXMATCH") return funcREGEXMATCH(args);
    if (fn == "REGEXEXTRACT") return funcREGEXEXTRACT(args);
    if (fn == "REGEXREPLACE") return funcREGEXREPLACE(args);

    // Dynamic array functions
    if (fn == "FILTER") return funcFILTER(args);
    if (fn == "SORT") return funcSORT(args);
    if (fn == "UNIQUE") return funcUNIQUE(args);
    if (fn == "SEQUENCE") return funcSEQUENCE(args);

    m_lastError = "Unknown function: " + fn;
    return QVariant("#NAME?");
}

// ---- Aggregate functions ----

// Helper: check if any argument is a lazy CellRange (large range optimization)
static bool hasLazyRange(const std::vector<QVariant>& args) {
    for (const auto& a : args)
        if (a.canConvert<CellRange>()) return true;
    return false;
}

QVariant FormulaEngine::funcSUM(const std::vector<QVariant>& args) {
    if (hasLazyRange(args)) {
        double sum = 0;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                CellRange range = arg.value<CellRange>();
                // Fast path: use ColumnStore::sumColumn for direct chunk-level summation
                // Bypasses QVariant entirely — ~100x faster for millions of cells
                if (m_spreadsheet) {
                    auto& cs = m_spreadsheet->getColumnStore();
                    int startRow = range.getStart().row;
                    int endRow = range.getEnd().row;
                    for (int c = range.getStart().col; c <= range.getEnd().col; c++) {
                        sum += cs.sumColumn(c, startRow, endRow);
                    }
                }
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>())
                    if (!v.isNull() && v.isValid()) sum += toNumber(v);
            } else if (!arg.isNull() && arg.isValid()) {
                sum += toNumber(arg);
            }
        }
        return sum;
    }
    auto flat = flattenArgs(args); double sum = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) sum += toNumber(a);
    return sum;
}

QVariant FormulaEngine::funcAVERAGE(const std::vector<QVariant>& args) {
    if (hasLazyRange(args)) {
        double sum = 0; long long count = 0;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                CellRange range = arg.value<CellRange>();
                if (m_spreadsheet) {
                    auto& cs = m_spreadsheet->getColumnStore();
                    int startRow = range.getStart().row;
                    int endRow = range.getEnd().row;
                    for (int c = range.getStart().col; c <= range.getEnd().col; c++) {
                        sum += cs.sumColumn(c, startRow, endRow);
                        count += cs.countColumn(c, startRow, endRow);
                    }
                }
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>())
                    if (!v.isNull() && v.isValid()) { sum += toNumber(v); count++; }
            } else if (!arg.isNull() && arg.isValid()) {
                sum += toNumber(arg); count++;
            }
        }
        return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
    }
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0; int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { sum += toNumber(a); count++; }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

QVariant FormulaEngine::funcCOUNT(const std::vector<QVariant>& args) {
    if (hasLazyRange(args)) {
        long long count = 0;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                // Fast path: use ColumnStore::countColumn directly
                CellRange range = arg.value<CellRange>();
                if (m_spreadsheet) {
                    auto& cs = m_spreadsheet->getColumnStore();
                    int startRow = range.getStart().row;
                    int endRow = range.getEnd().row;
                    for (int c = range.getStart().col; c <= range.getEnd().col; c++) {
                        count += cs.countColumn(c, startRow, endRow);
                    }
                }
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>()) {
                    if (!v.isNull() && v.isValid()) {
                        bool ok = false; v.toDouble(&ok);
                        if (ok || v.typeId() == QMetaType::Int || v.typeId() == QMetaType::Double) count++;
                    }
                }
            } else if (!arg.isNull() && arg.isValid()) {
                bool ok = false; arg.toDouble(&ok);
                if (ok || arg.typeId() == QMetaType::Int || arg.typeId() == QMetaType::Double) count++;
            }
        }
        return static_cast<int>(count);
    }
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
    if (hasLazyRange(args)) {
        long long count = 0;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                streamRangeValues(arg.value<CellRange>(), [&](const QVariant& v) {
                    if (!v.isNull() && v.isValid() && !v.toString().isEmpty()) count++;
                });
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>())
                    if (!v.isNull() && v.isValid() && !v.toString().isEmpty()) count++;
            } else if (!arg.isNull() && arg.isValid() && !arg.toString().isEmpty()) {
                count++;
            }
        }
        return static_cast<int>(count);
    }
    auto flat = flattenArgs(args); int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid() && !a.toString().isEmpty()) count++;
    return count;
}

QVariant FormulaEngine::funcMIN(const std::vector<QVariant>& args) {
    if (hasLazyRange(args)) {
        double min = std::numeric_limits<double>::max(); bool found = false;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                streamRangeValues(arg.value<CellRange>(), [&](const QVariant& v) {
                    if (!v.isNull() && v.isValid()) { double d = toNumber(v); if (!found || d < min) { min = d; found = true; } }
                });
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>())
                    if (!v.isNull() && v.isValid()) { double d = toNumber(v); if (!found || d < min) { min = d; found = true; } }
            } else if (!arg.isNull() && arg.isValid()) {
                double d = toNumber(arg); if (!found || d < min) { min = d; found = true; }
            }
        }
        return found ? QVariant(min) : QVariant();
    }
    auto flat = flattenArgs(args); double min = std::numeric_limits<double>::max(); bool found = false;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { double v = toNumber(a); if (!found || v < min) { min = v; found = true; } }
    return found ? QVariant(min) : QVariant();
}

QVariant FormulaEngine::funcMAX(const std::vector<QVariant>& args) {
    if (hasLazyRange(args)) {
        double max = std::numeric_limits<double>::lowest(); bool found = false;
        for (const auto& arg : args) {
            if (arg.canConvert<CellRange>()) {
                streamRangeValues(arg.value<CellRange>(), [&](const QVariant& v) {
                    if (!v.isNull() && v.isValid()) { double d = toNumber(v); if (!found || d > max) { max = d; found = true; } }
                });
            } else if (arg.canConvert<std::vector<QVariant>>()) {
                for (const auto& v : arg.value<std::vector<QVariant>>())
                    if (!v.isNull() && v.isValid()) { double d = toNumber(v); if (!found || d > max) { max = d; found = true; } }
            } else if (!arg.isNull() && arg.isValid()) {
                double d = toNumber(arg); if (!found || d > max) { max = d; found = true; }
            }
        }
        return found ? QVariant(max) : QVariant();
    }
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

    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;

    long long totalCells = static_cast<long long>(endRow - startRow + 1) * (endCol - startCol + 1);

    if (totalCells > 10000) {
        // Sparse iteration — only visit occupied cells using nav index.
        // Correct for all aggregate functions (SUM, AVG, COUNT, MIN, MAX)
        // which skip null/empty values anyway.
        for (int c = startCol; c <= endCol; c++) {
            const auto& occupiedRows = m_spreadsheet->getOccupiedRowsInColumn(c);
            auto lo = std::lower_bound(occupiedRows.begin(), occupiedRows.end(), startRow);
            auto hi = std::upper_bound(occupiedRows.begin(), occupiedRows.end(), endRow);
            for (auto it = lo; it != hi; ++it) {
                values.push_back(m_spreadsheet->getCellValue(CellAddress(*it, c)));
            }
        }
    } else {
        // Small range: visit all cells to preserve positional ordering
        for (int r = startRow; r <= endRow; r++) {
            for (int c = startCol; c <= endCol; c++) {
                values.push_back(m_spreadsheet->getCellValue(CellAddress(r, c)));
            }
        }
    }

    return values;
}

void FormulaEngine::streamRangeValues(const CellRange& range, std::function<void(const QVariant&)> fn) {
    if (!m_spreadsheet) return;

    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;

    // Use Spreadsheet's direct streaming method
    for (int c = startCol; c <= endCol; c++) {
        m_spreadsheet->streamColumnValues(c, startRow, endRow, fn);
    }
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
        if (match) {
            if (c < table[rowIdx - 1].size()) return table[rowIdx - 1][c];
            return QVariant("#REF!");
        }
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

    if (table.empty()) return QVariant("#REF!");
    if (rowNum < 1 || rowNum > static_cast<int>(table.size())) return QVariant("#REF!");
    if (table[0].empty() || colNum < 1 || colNum > static_cast<int>(table[0].size())) return QVariant("#REF!");
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

// ---- Multi-criteria functions ----

QVariant FormulaEngine::funcSUMIFS(const std::vector<QVariant>& args) {
    // SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
    if (args.size() < 3 || (args.size() - 1) % 2 != 0) return QVariant("#VALUE!");
    auto sumRange = flattenArgs({args[0]});
    size_t numCriteria = (args.size() - 1) / 2;
    std::vector<std::vector<QVariant>> criteriaRanges;
    std::vector<QString> criteriaValues;
    for (size_t i = 0; i < numCriteria; ++i) {
        criteriaRanges.push_back(flattenArgs({args[1 + i * 2]}));
        criteriaValues.push_back(toString(args[2 + i * 2]));
    }
    double sum = 0;
    for (size_t i = 0; i < sumRange.size(); ++i) {
        bool allMatch = true;
        for (size_t c = 0; c < numCriteria; ++c) {
            if (i >= criteriaRanges[c].size() || !matchesCriteria(criteriaRanges[c][i], criteriaValues[c])) {
                allMatch = false;
                break;
            }
        }
        if (allMatch) sum += toNumber(sumRange[i]);
    }
    return sum;
}

QVariant FormulaEngine::funcCOUNTIFS(const std::vector<QVariant>& args) {
    // COUNTIFS(criteria_range1, criteria1, criteria_range2, criteria2, ...)
    if (args.size() < 2 || args.size() % 2 != 0) return QVariant("#VALUE!");
    size_t numCriteria = args.size() / 2;
    std::vector<std::vector<QVariant>> criteriaRanges;
    std::vector<QString> criteriaValues;
    for (size_t i = 0; i < numCriteria; ++i) {
        criteriaRanges.push_back(flattenArgs({args[i * 2]}));
        criteriaValues.push_back(toString(args[1 + i * 2]));
    }
    if (criteriaRanges.empty()) return 0;
    int count = 0;
    size_t len = criteriaRanges[0].size();
    for (size_t i = 0; i < len; ++i) {
        bool allMatch = true;
        for (size_t c = 0; c < numCriteria; ++c) {
            if (i >= criteriaRanges[c].size() || !matchesCriteria(criteriaRanges[c][i], criteriaValues[c])) {
                allMatch = false;
                break;
            }
        }
        if (allMatch) count++;
    }
    return count;
}

QVariant FormulaEngine::funcAVERAGEIFS(const std::vector<QVariant>& args) {
    // AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)
    if (args.size() < 3 || (args.size() - 1) % 2 != 0) return QVariant("#VALUE!");
    auto avgRange = flattenArgs({args[0]});
    size_t numCriteria = (args.size() - 1) / 2;
    std::vector<std::vector<QVariant>> criteriaRanges;
    std::vector<QString> criteriaValues;
    for (size_t i = 0; i < numCriteria; ++i) {
        criteriaRanges.push_back(flattenArgs({args[1 + i * 2]}));
        criteriaValues.push_back(toString(args[2 + i * 2]));
    }
    double sum = 0; int count = 0;
    for (size_t i = 0; i < avgRange.size(); ++i) {
        bool allMatch = true;
        for (size_t c = 0; c < numCriteria; ++c) {
            if (i >= criteriaRanges[c].size() || !matchesCriteria(criteriaRanges[c][i], criteriaValues[c])) {
                allMatch = false;
                break;
            }
        }
        if (allMatch) { sum += toNumber(avgRange[i]); count++; }
    }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

QVariant FormulaEngine::funcMINIFS(const std::vector<QVariant>& args) {
    // MINIFS(min_range, criteria_range1, criteria1, ...)
    if (args.size() < 3 || (args.size() - 1) % 2 != 0) return QVariant("#VALUE!");
    auto minRange = flattenArgs({args[0]});
    size_t numCriteria = (args.size() - 1) / 2;
    std::vector<std::vector<QVariant>> criteriaRanges;
    std::vector<QString> criteriaValues;
    for (size_t i = 0; i < numCriteria; ++i) {
        criteriaRanges.push_back(flattenArgs({args[1 + i * 2]}));
        criteriaValues.push_back(toString(args[2 + i * 2]));
    }
    double result = std::numeric_limits<double>::max();
    bool found = false;
    for (size_t i = 0; i < minRange.size(); ++i) {
        bool allMatch = true;
        for (size_t c = 0; c < numCriteria; ++c) {
            if (i >= criteriaRanges[c].size() || !matchesCriteria(criteriaRanges[c][i], criteriaValues[c])) {
                allMatch = false;
                break;
            }
        }
        if (allMatch) {
            double val = toNumber(minRange[i]);
            if (!found || val < result) result = val;
            found = true;
        }
    }
    return found ? QVariant(result) : QVariant(0.0);
}

QVariant FormulaEngine::funcMAXIFS(const std::vector<QVariant>& args) {
    // MAXIFS(max_range, criteria_range1, criteria1, ...)
    if (args.size() < 3 || (args.size() - 1) % 2 != 0) return QVariant("#VALUE!");
    auto maxRange = flattenArgs({args[0]});
    size_t numCriteria = (args.size() - 1) / 2;
    std::vector<std::vector<QVariant>> criteriaRanges;
    std::vector<QString> criteriaValues;
    for (size_t i = 0; i < numCriteria; ++i) {
        criteriaRanges.push_back(flattenArgs({args[1 + i * 2]}));
        criteriaValues.push_back(toString(args[2 + i * 2]));
    }
    double result = std::numeric_limits<double>::lowest();
    bool found = false;
    for (size_t i = 0; i < maxRange.size(); ++i) {
        bool allMatch = true;
        for (size_t c = 0; c < numCriteria; ++c) {
            if (i >= criteriaRanges[c].size() || !matchesCriteria(criteriaRanges[c][i], criteriaValues[c])) {
                allMatch = false;
                break;
            }
        }
        if (allMatch) {
            double val = toNumber(maxRange[i]);
            if (!found || val > result) result = val;
            found = true;
        }
    }
    return found ? QVariant(result) : QVariant(0.0);
}

// ---- IFS function ----

QVariant FormulaEngine::funcIFS(const std::vector<QVariant>& args) {
    // IFS(condition1, value1, condition2, value2, ...)
    if (args.size() < 2 || args.size() % 2 != 0) return QVariant("#VALUE!");
    for (size_t i = 0; i < args.size(); i += 2) {
        if (toBoolean(args[i])) return args[i + 1];
    }
    return QVariant("#N/A");
}

// ---- TEXTJOIN function ----

QVariant FormulaEngine::funcTEXTJOIN(const std::vector<QVariant>& args) {
    // TEXTJOIN(delimiter, ignore_empty, text1, ...)
    if (args.size() < 3) return QVariant("#VALUE!");
    QString delimiter = toString(args[0]);
    bool ignoreEmpty = toBoolean(args[1]);
    std::vector<QVariant> remaining(args.begin() + 2, args.end());
    auto flat = flattenArgs(remaining);
    QStringList parts;
    for (const auto& v : flat) {
        QString s = toString(v);
        if (ignoreEmpty && (s.isEmpty() || v.isNull() || !v.isValid())) continue;
        parts.append(s);
    }
    return parts.join(delimiter);
}

// ---- Regex functions ----

QVariant FormulaEngine::funcREGEXMATCH(const std::vector<QVariant>& args) {
    // REGEXMATCH(text, regex)
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QRegularExpression re(toString(args[1]));
    if (!re.isValid()) return QVariant("#VALUE!");
    return QVariant(re.match(text).hasMatch());
}

QVariant FormulaEngine::funcREGEXEXTRACT(const std::vector<QVariant>& args) {
    // REGEXEXTRACT(text, regex)
    if (args.size() < 2) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QRegularExpression re(toString(args[1]));
    if (!re.isValid()) return QVariant("#VALUE!");
    QRegularExpressionMatch match = re.match(text);
    if (!match.hasMatch()) return QVariant("#N/A");
    return match.captured(0);
}

QVariant FormulaEngine::funcREGEXREPLACE(const std::vector<QVariant>& args) {
    // REGEXREPLACE(text, regex, replacement)
    if (args.size() < 3) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QRegularExpression re(toString(args[1]));
    if (!re.isValid()) return QVariant("#VALUE!");
    return text.replace(re, toString(args[2]));
}

// ---- Dynamic Array Functions ----

// Helper: convert a 2D vector result to QVariantList of QVariantList
static QVariant make2DResult(const std::vector<std::vector<QVariant>>& data) {
    QVariantList outer;
    for (const auto& row : data) {
        QVariantList innerRow;
        for (const auto& val : row) {
            innerRow.append(val);
        }
        outer.append(QVariant(innerRow));
    }
    return QVariant(outer);
}

// Helper: get 2D data from an argument (CellRange or already-evaluated array)
static std::vector<std::vector<QVariant>> arg2D(FormulaEngine* engine,
                                                 const QVariant& arg,
                                                 const std::vector<CellRange>& rangeArgs,
                                                 int rangeIndex) {
    if (rangeIndex < static_cast<int>(rangeArgs.size())) {
        return engine->getRangeValues2D(rangeArgs[rangeIndex]);
    }
    // Fallback: try to interpret as a single value
    return {{arg}};
}

QVariant FormulaEngine::funcFILTER(const std::vector<QVariant>& args) {
    // FILTER(array, include, [if_empty])
    if (args.size() < 2) return QVariant("#VALUE!");

    auto array = arg2D(this, args[0], m_lastRangeArgs, 0);
    auto include = arg2D(this, args[1], m_lastRangeArgs, 1);

    if (array.empty()) return QVariant("#VALUE!");

    // Build list of row indices where include is TRUE
    std::vector<std::vector<QVariant>> filtered;
    int numRows = static_cast<int>(array.size());
    for (int r = 0; r < numRows; ++r) {
        bool shouldInclude = false;
        if (r < static_cast<int>(include.size()) && !include[r].empty()) {
            QVariant val = include[r][0];
            if (val.typeId() == QMetaType::Bool) {
                shouldInclude = val.toBool();
            } else {
                shouldInclude = toBoolean(val);
            }
        }
        if (shouldInclude) {
            filtered.push_back(array[r]);
        }
    }

    if (filtered.empty()) {
        if (args.size() >= 3) {
            return args[2]; // if_empty value
        }
        return QVariant("#CALC!");
    }

    return make2DResult(filtered);
}

QVariant FormulaEngine::funcSORT(const std::vector<QVariant>& args) {
    // SORT(array, [sort_index], [sort_order], [by_col])
    if (args.empty()) return QVariant("#VALUE!");

    auto array = arg2D(this, args[0], m_lastRangeArgs, 0);
    if (array.empty()) return QVariant("#VALUE!");

    int sortIndex = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    int sortOrder = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) : 1;
    bool byCol = args.size() >= 4 ? toBoolean(args[3]) : false;

    if (sortIndex < 1) sortIndex = 1;

    if (!byCol) {
        // Sort rows by column sortIndex (1-based)
        int colIdx = sortIndex - 1;
        std::stable_sort(array.begin(), array.end(),
            [colIdx, sortOrder, this](const std::vector<QVariant>& a, const std::vector<QVariant>& b) {
                QVariant va = (colIdx < static_cast<int>(a.size())) ? a[colIdx] : QVariant();
                QVariant vb = (colIdx < static_cast<int>(b.size())) ? b[colIdx] : QVariant();

                bool aEmpty = !va.isValid() || va.toString().isEmpty();
                bool bEmpty = !vb.isValid() || vb.toString().isEmpty();
                if (aEmpty && bEmpty) return false;
                if (aEmpty) return false; // empties sort last
                if (bEmpty) return true;

                bool aOk, bOk;
                double aNum = va.toString().toDouble(&aOk);
                double bNum = vb.toString().toDouble(&bOk);
                if (aOk && bOk) {
                    return sortOrder >= 0 ? (aNum < bNum) : (aNum > bNum);
                }
                int cmp = va.toString().compare(vb.toString(), Qt::CaseInsensitive);
                return sortOrder >= 0 ? (cmp < 0) : (cmp > 0);
            });
    } else {
        // Sort columns by row sortIndex (1-based)
        int rowIdx = sortIndex - 1;
        if (rowIdx >= static_cast<int>(array.size())) rowIdx = 0;

        int numCols = 0;
        for (const auto& row : array) {
            numCols = std::max(numCols, static_cast<int>(row.size()));
        }

        // Create column indices and sort them
        std::vector<int> colOrder(numCols);
        std::iota(colOrder.begin(), colOrder.end(), 0);
        std::stable_sort(colOrder.begin(), colOrder.end(),
            [&array, rowIdx, sortOrder, this](int a, int b) {
                QVariant va = (a < static_cast<int>(array[rowIdx].size())) ? array[rowIdx][a] : QVariant();
                QVariant vb = (b < static_cast<int>(array[rowIdx].size())) ? array[rowIdx][b] : QVariant();
                bool aOk, bOk;
                double aNum = va.toString().toDouble(&aOk);
                double bNum = vb.toString().toDouble(&bOk);
                if (aOk && bOk) return sortOrder >= 0 ? (aNum < bNum) : (aNum > bNum);
                int cmp = va.toString().compare(vb.toString(), Qt::CaseInsensitive);
                return sortOrder >= 0 ? (cmp < 0) : (cmp > 0);
            });

        // Rearrange columns
        std::vector<std::vector<QVariant>> result;
        for (const auto& row : array) {
            std::vector<QVariant> newRow;
            for (int ci : colOrder) {
                newRow.push_back(ci < static_cast<int>(row.size()) ? row[ci] : QVariant());
            }
            result.push_back(std::move(newRow));
        }
        array = std::move(result);
    }

    return make2DResult(array);
}

QVariant FormulaEngine::funcUNIQUE(const std::vector<QVariant>& args) {
    // UNIQUE(array, [by_col], [exactly_once])
    if (args.empty()) return QVariant("#VALUE!");

    auto array = arg2D(this, args[0], m_lastRangeArgs, 0);
    if (array.empty()) return QVariant("#VALUE!");

    bool byCol = args.size() >= 2 ? toBoolean(args[1]) : false;
    bool exactlyOnce = args.size() >= 3 ? toBoolean(args[2]) : false;

    if (!byCol) {
        // Unique rows: build a string key for each row
        auto rowKey = [](const std::vector<QVariant>& row) -> QString {
            QStringList parts;
            for (const auto& v : row) parts.append(v.toString());
            return parts.join(QChar(0x1F)); // unit separator
        };

        // Count occurrences
        std::map<QString, int> counts;
        std::vector<QString> keys;
        for (const auto& row : array) {
            QString key = rowKey(row);
            counts[key]++;
            keys.push_back(key);
        }

        std::vector<std::vector<QVariant>> result;
        std::set<QString> seen;
        for (size_t i = 0; i < array.size(); ++i) {
            const QString& key = keys[i];
            if (exactlyOnce) {
                if (counts[key] == 1) {
                    result.push_back(array[i]);
                }
            } else {
                if (seen.find(key) == seen.end()) {
                    seen.insert(key);
                    result.push_back(array[i]);
                }
            }
        }

        if (result.empty()) return QVariant("#CALC!");
        return make2DResult(result);
    } else {
        // Unique columns
        int numRows = static_cast<int>(array.size());
        int numCols = 0;
        for (const auto& row : array) numCols = std::max(numCols, static_cast<int>(row.size()));

        auto colKey = [&](int c) -> QString {
            QStringList parts;
            for (int r = 0; r < numRows; ++r) {
                parts.append(c < static_cast<int>(array[r].size()) ? array[r][c].toString() : "");
            }
            return parts.join(QChar(0x1F));
        };

        std::map<QString, int> counts;
        std::vector<QString> keys;
        for (int c = 0; c < numCols; ++c) {
            QString key = colKey(c);
            counts[key]++;
            keys.push_back(key);
        }

        std::vector<int> keepCols;
        std::set<QString> seen;
        for (int c = 0; c < numCols; ++c) {
            const QString& key = keys[c];
            if (exactlyOnce) {
                if (counts[key] == 1) keepCols.push_back(c);
            } else {
                if (seen.find(key) == seen.end()) {
                    seen.insert(key);
                    keepCols.push_back(c);
                }
            }
        }

        std::vector<std::vector<QVariant>> result;
        for (int r = 0; r < numRows; ++r) {
            std::vector<QVariant> newRow;
            for (int c : keepCols) {
                newRow.push_back(c < static_cast<int>(array[r].size()) ? array[r][c] : QVariant());
            }
            result.push_back(std::move(newRow));
        }

        if (result.empty()) return QVariant("#CALC!");
        return make2DResult(result);
    }
}

QVariant FormulaEngine::funcSEQUENCE(const std::vector<QVariant>& args) {
    // SEQUENCE(rows, [columns], [start], [step])
    if (args.empty()) return QVariant("#VALUE!");

    int rows = static_cast<int>(toNumber(args[0]));
    int cols = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    double start = args.size() >= 3 ? toNumber(args[2]) : 1.0;
    double step = args.size() >= 4 ? toNumber(args[3]) : 1.0;

    if (rows < 1) rows = 1;
    if (cols < 1) cols = 1;

    // Safety limit
    if (rows * cols > 1000000) return QVariant("#VALUE!");

    std::vector<std::vector<QVariant>> result;
    double current = start;
    for (int r = 0; r < rows; ++r) {
        std::vector<QVariant> row;
        for (int c = 0; c < cols; ++c) {
            row.push_back(QVariant(current));
            current += step;
        }
        result.push_back(std::move(row));
    }

    // Single cell: return scalar
    if (rows == 1 && cols == 1) return result[0][0];

    return make2DResult(result);
}
