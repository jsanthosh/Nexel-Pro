#ifndef FORMULAENGINE_H
#define FORMULAENGINE_H

#include <QString>
#include <QVariant>
#include <unordered_map>
#include <memory>
#include <functional>
#include "CellRange.h"

class Spreadsheet;

class FormulaEngine {
public:
    explicit FormulaEngine(Spreadsheet* spreadsheet = nullptr);
    ~FormulaEngine() = default;

    QVariant evaluate(const QString& formula);
    void setSpreadsheet(Spreadsheet* spreadsheet);

    void clearCache();
    void invalidateCell(const CellAddress& addr);

    QString getLastError() const;
    bool hasError() const;

    // Get cell references found during last evaluation
    const std::vector<CellAddress>& getLastDependencies() const { return m_lastDependencies; }

private:
    Spreadsheet* m_spreadsheet;
    QString m_lastError;
    std::unordered_map<std::string, QVariant> m_cache;
    std::vector<CellAddress> m_lastDependencies;

    // Parser hierarchy
    QVariant parseExpression(const QString& expr);
    QVariant evaluateComparison(const QString& expr, int& pos);
    QVariant evaluateTerm(const QString& expr, int& pos);
    QVariant evaluateMultiplicative(const QString& expr, int& pos);
    QVariant evaluateUnary(const QString& expr, int& pos);
    QVariant evaluatePower(const QString& expr, int& pos);
    QVariant evaluateFactor(const QString& expr, int& pos);

    QVariant evaluateFunction(const QString& funcName, const std::vector<QVariant>& args);

    // Aggregate functions
    QVariant funcSUM(const std::vector<QVariant>& args);
    QVariant funcAVERAGE(const std::vector<QVariant>& args);
    QVariant funcCOUNT(const std::vector<QVariant>& args);
    QVariant funcCOUNTA(const std::vector<QVariant>& args);
    QVariant funcMIN(const std::vector<QVariant>& args);
    QVariant funcMAX(const std::vector<QVariant>& args);
    QVariant funcIF(const std::vector<QVariant>& args);
    QVariant funcCONCAT(const std::vector<QVariant>& args);
    QVariant funcLEN(const std::vector<QVariant>& args);
    QVariant funcUPPER(const std::vector<QVariant>& args);
    QVariant funcLOWER(const std::vector<QVariant>& args);
    QVariant funcTRIM(const std::vector<QVariant>& args);

    // Math functions
    QVariant funcROUND(const std::vector<QVariant>& args);
    QVariant funcABS(const std::vector<QVariant>& args);
    QVariant funcSQRT(const std::vector<QVariant>& args);
    QVariant funcPOWER(const std::vector<QVariant>& args);
    QVariant funcMOD(const std::vector<QVariant>& args);
    QVariant funcINT(const std::vector<QVariant>& args);
    QVariant funcCEILING(const std::vector<QVariant>& args);
    QVariant funcFLOOR(const std::vector<QVariant>& args);

    // Logical functions
    QVariant funcAND(const std::vector<QVariant>& args);
    QVariant funcOR(const std::vector<QVariant>& args);
    QVariant funcNOT(const std::vector<QVariant>& args);
    QVariant funcIFERROR(const std::vector<QVariant>& args);

    // Text functions
    QVariant funcLEFT(const std::vector<QVariant>& args);
    QVariant funcRIGHT(const std::vector<QVariant>& args);
    QVariant funcMID(const std::vector<QVariant>& args);
    QVariant funcFIND(const std::vector<QVariant>& args);
    QVariant funcSUBSTITUTE(const std::vector<QVariant>& args);
    QVariant funcTEXT(const std::vector<QVariant>& args);

    // Statistical functions
    QVariant funcCOUNTIF(const std::vector<QVariant>& args);
    QVariant funcSUMIF(const std::vector<QVariant>& args);
    QVariant funcAVERAGEIF(const std::vector<QVariant>& args);
    QVariant funcCOUNTBLANK(const std::vector<QVariant>& args);
    QVariant funcSUMPRODUCT(const std::vector<QVariant>& args);
    QVariant funcMEDIAN(const std::vector<QVariant>& args);
    QVariant funcMODE(const std::vector<QVariant>& args);
    QVariant funcSTDEV(const std::vector<QVariant>& args);
    QVariant funcVAR(const std::vector<QVariant>& args);
    QVariant funcLARGE(const std::vector<QVariant>& args);
    QVariant funcSMALL(const std::vector<QVariant>& args);
    QVariant funcRANK(const std::vector<QVariant>& args);
    QVariant funcPERCENTILE(const std::vector<QVariant>& args);

    // Date functions
    QVariant funcNOW(const std::vector<QVariant>& args);
    QVariant funcTODAY(const std::vector<QVariant>& args);
    QVariant funcYEAR(const std::vector<QVariant>& args);
    QVariant funcMONTH(const std::vector<QVariant>& args);
    QVariant funcDAY(const std::vector<QVariant>& args);
    QVariant funcDATE(const std::vector<QVariant>& args);
    QVariant funcHOUR(const std::vector<QVariant>& args);
    QVariant funcMINUTE(const std::vector<QVariant>& args);
    QVariant funcSECOND(const std::vector<QVariant>& args);
    QVariant funcDATEDIF(const std::vector<QVariant>& args);
    QVariant funcNETWORKDAYS(const std::vector<QVariant>& args);
    QVariant funcWEEKDAY(const std::vector<QVariant>& args);
    QVariant funcEDATE(const std::vector<QVariant>& args);
    QVariant funcEOMONTH(const std::vector<QVariant>& args);
    QVariant funcDATEVALUE(const std::vector<QVariant>& args);

    // Lookup functions
    QVariant funcVLOOKUP(const std::vector<QVariant>& args);
    QVariant funcHLOOKUP(const std::vector<QVariant>& args);
    QVariant funcXLOOKUP(const std::vector<QVariant>& args);
    QVariant funcINDEX(const std::vector<QVariant>& args);
    QVariant funcMATCH(const std::vector<QVariant>& args);

    // Additional math functions
    QVariant funcROUNDUP(const std::vector<QVariant>& args);
    QVariant funcROUNDDOWN(const std::vector<QVariant>& args);
    QVariant funcLOG(const std::vector<QVariant>& args);
    QVariant funcLN(const std::vector<QVariant>& args);
    QVariant funcEXP(const std::vector<QVariant>& args);
    QVariant funcRAND(const std::vector<QVariant>& args);
    QVariant funcRANDBETWEEN(const std::vector<QVariant>& args);

    // Additional text functions
    QVariant funcPROPER(const std::vector<QVariant>& args);
    QVariant funcSEARCH(const std::vector<QVariant>& args);
    QVariant funcREPT(const std::vector<QVariant>& args);
    QVariant funcEXACT(const std::vector<QVariant>& args);
    QVariant funcVALUE(const std::vector<QVariant>& args);

    // Additional logical/info functions
    QVariant funcISBLANK(const std::vector<QVariant>& args);
    QVariant funcISERROR(const std::vector<QVariant>& args);
    QVariant funcISNUMBER(const std::vector<QVariant>& args);
    QVariant funcISTEXT(const std::vector<QVariant>& args);
    QVariant funcCHOOSE(const std::vector<QVariant>& args);
    QVariant funcSWITCH(const std::vector<QVariant>& args);

    // Helpers
    double toNumber(const QVariant& value);
    QString toString(const QVariant& value);
    bool toBoolean(const QVariant& value);
    QVariant getCellValue(const CellAddress& addr);
    std::vector<QVariant> getRangeValues(const CellRange& range);
    std::vector<std::vector<QVariant>> getRangeValues2D(const CellRange& range);
    std::vector<QVariant> flattenArgs(const std::vector<QVariant>& args);
    void skipWhitespace(const QString& expr, int& pos);
    bool matchesCriteria(const QVariant& value, const QString& criteria);
    QDate parseDate(const QVariant& value);

    // Range tracking for lookup functions
    std::vector<CellRange> m_lastRangeArgs;
};

#endif // FORMULAENGINE_H
