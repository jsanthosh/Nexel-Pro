#ifndef FORMULAENGINE_H
#define FORMULAENGINE_H

#include <QString>
#include <QVariant>
#include <unordered_map>
#include <memory>
#include <functional>
#include "CellRange.h"
#include "FormulaAST.h"

class Spreadsheet;

class FormulaEngine {
public:
    explicit FormulaEngine(Spreadsheet* spreadsheet = nullptr);
    ~FormulaEngine() = default;

    QVariant evaluate(const QString& formula);
    void setSpreadsheet(Spreadsheet* spreadsheet);
    void setAllSheets(const std::vector<std::shared_ptr<Spreadsheet>>* sheets) { m_allSheets = sheets; }
    const std::vector<std::shared_ptr<Spreadsheet>>* getAllSheets() const { return m_allSheets; }

    void clearCache();
    void invalidateCell(const CellAddress& addr);

    QString getLastError() const;
    bool hasError() const;

    // Get cell references found during last evaluation
    const std::vector<CellAddress>& getLastDependencies() const { return m_lastDependencies; }
    // Get column-level dependencies (from column references like D:D)
    const std::vector<int>& getLastColumnDeps() const { return m_lastColumnDeps; }
    // Get range arguments from last evaluation (for range-level dependency tracking)
    const std::vector<CellRange>& getLastRangeArgs() const { return m_lastRangeArgs; }

    // Public access to 2D range values (used by dynamic array helpers)
    std::vector<std::vector<QVariant>> getRangeValues2D(const CellRange& range);

private:
    Spreadsheet* m_spreadsheet;
    const std::vector<std::shared_ptr<Spreadsheet>>* m_allSheets = nullptr;
    QString m_lastError;
    std::unordered_map<std::string, QVariant> m_cache;
    std::vector<CellAddress> m_lastDependencies;
    std::vector<int> m_lastColumnDeps;

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
    QVariant funcSUMIFS(const std::vector<QVariant>& args);
    QVariant funcCOUNTIFS(const std::vector<QVariant>& args);
    QVariant funcAVERAGEIFS(const std::vector<QVariant>& args);
    QVariant funcMINIFS(const std::vector<QVariant>& args);
    QVariant funcMAXIFS(const std::vector<QVariant>& args);
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
    QVariant funcIFS(const std::vector<QVariant>& args);

    // Text join and regex functions
    QVariant funcTEXTJOIN(const std::vector<QVariant>& args);
    QVariant funcREGEXMATCH(const std::vector<QVariant>& args);
    QVariant funcREGEXEXTRACT(const std::vector<QVariant>& args);
    QVariant funcREGEXREPLACE(const std::vector<QVariant>& args);

    // Dynamic array functions
    QVariant funcFILTER(const std::vector<QVariant>& args);
    QVariant funcSORT(const std::vector<QVariant>& args);
    QVariant funcUNIQUE(const std::vector<QVariant>& args);
    QVariant funcSEQUENCE(const std::vector<QVariant>& args);

    // === Extended functions (FormulaFunctions.cpp) ===
    // Math & Trig
    QVariant funcPI(const std::vector<QVariant>& args);
    QVariant funcSIGN(const std::vector<QVariant>& args);
    QVariant funcTRUNC(const std::vector<QVariant>& args);
    QVariant funcPRODUCT(const std::vector<QVariant>& args);
    QVariant funcQUOTIENT(const std::vector<QVariant>& args);
    QVariant funcMROUND(const std::vector<QVariant>& args);
    QVariant funcCEILING_MATH(const std::vector<QVariant>& args);
    QVariant funcFLOOR_MATH(const std::vector<QVariant>& args);
    QVariant funcLOG10(const std::vector<QVariant>& args);
    QVariant funcSIN(const std::vector<QVariant>& args);
    QVariant funcCOS(const std::vector<QVariant>& args);
    QVariant funcTAN(const std::vector<QVariant>& args);
    QVariant funcASIN(const std::vector<QVariant>& args);
    QVariant funcACOS(const std::vector<QVariant>& args);
    QVariant funcATAN(const std::vector<QVariant>& args);
    QVariant funcATAN2(const std::vector<QVariant>& args);
    QVariant funcRADIANS(const std::vector<QVariant>& args);
    QVariant funcDEGREES(const std::vector<QVariant>& args);
    QVariant funcFACT(const std::vector<QVariant>& args);
    QVariant funcCOMBIN(const std::vector<QVariant>& args);
    QVariant funcPERMUT(const std::vector<QVariant>& args);
    QVariant funcGCD(const std::vector<QVariant>& args);
    QVariant funcLCM(const std::vector<QVariant>& args);
    QVariant funcEVEN(const std::vector<QVariant>& args);
    QVariant funcODD(const std::vector<QVariant>& args);
    QVariant funcSUMSQ(const std::vector<QVariant>& args);

    // Lookup & Reference
    QVariant funcROW(const std::vector<QVariant>& args);
    QVariant funcCOLUMN(const std::vector<QVariant>& args);
    QVariant funcROWS(const std::vector<QVariant>& args);
    QVariant funcCOLUMNS(const std::vector<QVariant>& args);
    QVariant funcADDRESS(const std::vector<QVariant>& args);
    QVariant funcTRANSPOSE(const std::vector<QVariant>& args);

    // Text
    QVariant funcCHAR(const std::vector<QVariant>& args);
    QVariant funcCODE(const std::vector<QVariant>& args);
    QVariant funcCLEAN(const std::vector<QVariant>& args);
    QVariant funcREPLACE(const std::vector<QVariant>& args);
    QVariant funcFIXED(const std::vector<QVariant>& args);
    QVariant funcT(const std::vector<QVariant>& args);
    QVariant funcN(const std::vector<QVariant>& args);
    QVariant funcNUMBERVALUE(const std::vector<QVariant>& args);
    QVariant funcUNICODE(const std::vector<QVariant>& args);
    QVariant funcUNICHAR(const std::vector<QVariant>& args);
    QVariant funcTEXTBEFORE(const std::vector<QVariant>& args);
    QVariant funcTEXTAFTER(const std::vector<QVariant>& args);

    // Logical
    QVariant funcXOR(const std::vector<QVariant>& args);
    QVariant funcIFNA(const std::vector<QVariant>& args);
    QVariant funcTRUE(const std::vector<QVariant>& args);
    QVariant funcFALSE(const std::vector<QVariant>& args);

    // Date & Time
    QVariant funcTIME(const std::vector<QVariant>& args);
    QVariant funcTIMEVALUE(const std::vector<QVariant>& args);
    QVariant funcDAYS(const std::vector<QVariant>& args);
    QVariant funcISOWEEKNUM(const std::vector<QVariant>& args);
    QVariant funcWEEKNUM(const std::vector<QVariant>& args);
    QVariant funcWORKDAY(const std::vector<QVariant>& args);

    // Financial
    QVariant funcPMT(const std::vector<QVariant>& args);
    QVariant funcFV(const std::vector<QVariant>& args);
    QVariant funcPV(const std::vector<QVariant>& args);
    QVariant funcNPV(const std::vector<QVariant>& args);
    QVariant funcNPER(const std::vector<QVariant>& args);
    QVariant funcIRR(const std::vector<QVariant>& args);
    QVariant funcEFFECT(const std::vector<QVariant>& args);
    QVariant funcNOMINAL(const std::vector<QVariant>& args);
    QVariant funcIPMT(const std::vector<QVariant>& args);
    QVariant funcPPMT(const std::vector<QVariant>& args);
    QVariant funcSLN(const std::vector<QVariant>& args);

    // Information
    QVariant funcISERR(const std::vector<QVariant>& args);
    QVariant funcISNA(const std::vector<QVariant>& args);
    QVariant funcISLOGICAL(const std::vector<QVariant>& args);
    QVariant funcISNONTEXT(const std::vector<QVariant>& args);
    QVariant funcISEVEN(const std::vector<QVariant>& args);
    QVariant funcISODD(const std::vector<QVariant>& args);
    QVariant funcTYPE(const std::vector<QVariant>& args);
    QVariant funcNA(const std::vector<QVariant>& args);
    QVariant funcERROR_TYPE(const std::vector<QVariant>& args);

    // Statistical
    QVariant funcAVERAGEA(const std::vector<QVariant>& args);
    QVariant funcMAXA(const std::vector<QVariant>& args);
    QVariant funcMINA(const std::vector<QVariant>& args);
    QVariant funcCORREL(const std::vector<QVariant>& args);
    QVariant funcSLOPE(const std::vector<QVariant>& args);
    QVariant funcINTERCEPT(const std::vector<QVariant>& args);
    QVariant funcFORECAST(const std::vector<QVariant>& args);

    // === Batch 2 (FormulaFunctions2.cpp) ===
    // Math
    QVariant funcSINH(const std::vector<QVariant>& args);
    QVariant funcCOSH(const std::vector<QVariant>& args);
    QVariant funcTANH(const std::vector<QVariant>& args);
    QVariant funcFACTDOUBLE(const std::vector<QVariant>& args);
    QVariant funcMULTINOMIAL(const std::vector<QVariant>& args);
    QVariant funcBASE(const std::vector<QVariant>& args);
    QVariant funcDECIMAL(const std::vector<QVariant>& args);
    QVariant funcROMAN(const std::vector<QVariant>& args);
    QVariant funcARABIC(const std::vector<QVariant>& args);
    QVariant funcSUBTOTAL(const std::vector<QVariant>& args);
    QVariant funcCOMBINA(const std::vector<QVariant>& args);
    QVariant funcSERIESSUM(const std::vector<QVariant>& args);
    // Lookup
    QVariant funcXMATCH(const std::vector<QVariant>& args);
    QVariant funcSORTBY(const std::vector<QVariant>& args);
    QVariant funcVSTACK(const std::vector<QVariant>& args);
    QVariant funcHSTACK(const std::vector<QVariant>& args);
    QVariant funcTAKE(const std::vector<QVariant>& args);
    QVariant funcDROP(const std::vector<QVariant>& args);
    // Text
    QVariant funcDOLLAR(const std::vector<QVariant>& args);
    QVariant funcLENB(const std::vector<QVariant>& args);
    QVariant funcFINDB(const std::vector<QVariant>& args);
    QVariant funcSEARCHB(const std::vector<QVariant>& args);
    QVariant funcVALUETOTEXT(const std::vector<QVariant>& args);
    // Statistical
    QVariant funcSTDEVP(const std::vector<QVariant>& args);
    QVariant funcVARP(const std::vector<QVariant>& args);
    QVariant funcPERCENTILE_INC(const std::vector<QVariant>& args);
    QVariant funcPERCENTILE_EXC(const std::vector<QVariant>& args);
    QVariant funcQUARTILE_INC(const std::vector<QVariant>& args);
    QVariant funcRANK_EQ(const std::vector<QVariant>& args);
    QVariant funcRANK_AVG(const std::vector<QVariant>& args);
    QVariant funcGEOMEAN(const std::vector<QVariant>& args);
    QVariant funcHARMEAN(const std::vector<QVariant>& args);
    QVariant funcTRIMMEAN(const std::vector<QVariant>& args);
    QVariant funcDEVSQ(const std::vector<QVariant>& args);
    QVariant funcAVEDEV(const std::vector<QVariant>& args);
    QVariant funcRSQ(const std::vector<QVariant>& args);
    QVariant funcFREQUENCY(const std::vector<QVariant>& args);
    // Financial
    QVariant funcRATE(const std::vector<QVariant>& args);
    QVariant funcXNPV(const std::vector<QVariant>& args);
    QVariant funcXIRR(const std::vector<QVariant>& args);
    QVariant funcCUMIPMT(const std::vector<QVariant>& args);
    QVariant funcCUMPRINC(const std::vector<QVariant>& args);
    QVariant funcPDURATION(const std::vector<QVariant>& args);
    QVariant funcRRI(const std::vector<QVariant>& args);
    QVariant funcFVSCHEDULE(const std::vector<QVariant>& args);
    // Date
    QVariant funcDAYS360(const std::vector<QVariant>& args);
    QVariant funcNETWORKDAYS_INTL(const std::vector<QVariant>& args);
    QVariant funcWORKDAY_INTL(const std::vector<QVariant>& args);
    // Info
    QVariant funcISFORMULA(const std::vector<QVariant>& args);
    QVariant funcISREF(const std::vector<QVariant>& args);

    // AST-based evaluation (parse once, evaluate many — 10x faster recalc)
    QVariant evaluateAST(uint32_t nodeIndex);
    QVariant evaluateASTFunction(uint16_t funcId, const QVariantList& argNodeIndices);

    // Helpers
    double toNumber(const QVariant& value);
    QString toString(const QVariant& value);
    bool toBoolean(const QVariant& value);
    QVariant getCellValue(const CellAddress& addr);
    std::vector<QVariant> getRangeValues(const CellRange& range);
    std::vector<QVariant> flattenArgs(const std::vector<QVariant>& args);
    void skipWhitespace(const QString& expr, int& pos);
    bool matchesCriteria(const QVariant& value, const QString& criteria);
    QDate parseDate(const QVariant& value);

    // Stream values from args (handles lazy CellRange + vector<QVariant> + scalar)
    template<typename Func>
    void forEachValue(const std::vector<QVariant>& args, Func fn);
    void streamRangeValues(const CellRange& range, std::function<void(const QVariant&)> fn);

    // Range tracking for lookup functions
    std::vector<CellRange> m_lastRangeArgs;
};

#endif // FORMULAENGINE_H
