#include "MacroEngine.h"
#include "../core/Spreadsheet.h"
#include <QTimer>
#include <QMessageBox>
#include <QEventLoop>

// ─── SpreadsheetAPI ──────────────────────────────────────────────────────────

SpreadsheetAPI::SpreadsheetAPI(QObject* parent)
    : QObject(parent)
{
}

void SpreadsheetAPI::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet)
{
    m_spreadsheet = std::move(spreadsheet);
}

CellAddress SpreadsheetAPI::parseCellRef(const QString& ref)
{
    return CellAddress::fromString(ref);
}

QVariant SpreadsheetAPI::getCellValue(const QString& cellRef)
{
    if (!m_spreadsheet) return {};
    CellAddress addr = parseCellRef(cellRef);
    return m_spreadsheet->getCellValue(addr);
}

void SpreadsheetAPI::setCellValue(const QString& cellRef, const QVariant& value)
{
    if (!m_spreadsheet) return;
    CellAddress addr = parseCellRef(cellRef);
    m_spreadsheet->setCellValue(addr, value);
    emit refreshRequested();
}

void SpreadsheetAPI::setCellFormula(const QString& cellRef, const QString& formula)
{
    if (!m_spreadsheet) return;
    CellAddress addr = parseCellRef(cellRef);
    m_spreadsheet->setCellFormula(addr, formula);
    emit refreshRequested();
}

QString SpreadsheetAPI::getCellFormula(const QString& cellRef)
{
    if (!m_spreadsheet) return {};
    CellAddress addr = parseCellRef(cellRef);
    auto cell = m_spreadsheet->getCellIfExists(addr);
    if (cell) {
        return cell->getFormula();
    }
    return {};
}

void SpreadsheetAPI::applyStyleChange(const QString& range,
                                      std::function<void(CellStyle&)> modifier)
{
    if (!m_spreadsheet) return;

    // Support both "A1:B5" range and single "A1" cell references
    if (range.contains(':')) {
        CellRange cellRange(range);
        auto cells = cellRange.getCells();
        for (const auto& addr : cells) {
            auto cell = m_spreadsheet->getCell(addr);
            if (cell) {
                CellStyle style = cell->getStyle();
                modifier(style);
                cell->setStyle(style);
            }
        }
    } else {
        CellAddress addr = parseCellRef(range);
        auto cell = m_spreadsheet->getCell(addr);
        if (cell) {
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);
        }
    }
    emit refreshRequested();
}

void SpreadsheetAPI::setBold(const QString& range, bool bold)
{
    applyStyleChange(range, [bold](CellStyle& style) {
        style.bold = bold;
    });
}

void SpreadsheetAPI::setItalic(const QString& range, bool italic)
{
    applyStyleChange(range, [italic](CellStyle& style) {
        style.italic = italic;
    });
}

void SpreadsheetAPI::setBackgroundColor(const QString& range, const QString& color)
{
    applyStyleChange(range, [&color](CellStyle& style) {
        style.backgroundColor = color;
    });
}

void SpreadsheetAPI::setForegroundColor(const QString& range, const QString& color)
{
    applyStyleChange(range, [&color](CellStyle& style) {
        style.foregroundColor = color;
    });
}

void SpreadsheetAPI::setFontSize(const QString& range, int size)
{
    applyStyleChange(range, [size](CellStyle& style) {
        style.fontSize = size;
    });
}

void SpreadsheetAPI::setNumberFormat(const QString& range, const QString& format)
{
    applyStyleChange(range, [&format](CellStyle& style) {
        style.numberFormat = format;
    });
}

void SpreadsheetAPI::mergeCells(const QString& range)
{
    if (!m_spreadsheet) return;
    CellRange cellRange(range);
    m_spreadsheet->mergeCells(cellRange);
    emit refreshRequested();
}

void SpreadsheetAPI::unmergeCells(const QString& range)
{
    if (!m_spreadsheet) return;
    CellRange cellRange(range);
    m_spreadsheet->unmergeCells(cellRange);
    emit refreshRequested();
}

void SpreadsheetAPI::setRowHeight(int row, int height)
{
    if (!m_spreadsheet) return;
    m_spreadsheet->setRowHeight(row, height);
    emit refreshRequested();
}

void SpreadsheetAPI::setColumnWidth(int col, int width)
{
    if (!m_spreadsheet) return;
    m_spreadsheet->setColumnWidth(col, width);
    emit refreshRequested();
}

int SpreadsheetAPI::getMaxRow()
{
    if (!m_spreadsheet) return 0;
    return m_spreadsheet->getMaxRow();
}

int SpreadsheetAPI::getMaxColumn()
{
    if (!m_spreadsheet) return 0;
    return m_spreadsheet->getMaxColumn();
}

QString SpreadsheetAPI::getSheetName()
{
    if (!m_spreadsheet) return {};
    return m_spreadsheet->getSheetName();
}

void SpreadsheetAPI::clearRange(const QString& range)
{
    if (!m_spreadsheet) return;
    CellRange cellRange(range);
    m_spreadsheet->clearRange(cellRange);
    emit refreshRequested();
}

void SpreadsheetAPI::alert(const QString& message)
{
    emit alertRequested(message);
}

void SpreadsheetAPI::log(const QString& message)
{
    emit logMessage(message);
}

// ─── MacroEngine ─────────────────────────────────────────────────────────────

MacroEngine::MacroEngine(QObject* parent)
    : QObject(parent)
    , m_api(nullptr)
{
    setupEngine();
}

void MacroEngine::setupEngine()
{
    // Create the spreadsheet API object
    m_api = new SpreadsheetAPI(this);

    // Inject the API as a global "sheet" object
    QJSValue apiValue = m_engine.newQObject(m_api);
    m_engine.globalObject().setProperty("sheet", apiValue);

    // Add convenience functions in JS
    m_engine.evaluate(QStringLiteral(
        "function getCellValue(ref) { return sheet.getCellValue(ref); }\n"
        "function setCellValue(ref, val) { sheet.setCellValue(ref, val); }\n"
        "function setCellFormula(ref, f) { sheet.setCellFormula(ref, f); }\n"
        "function alert(msg) { sheet.alert(msg); }\n"
        "function log(msg) { sheet.log(msg); }\n"
    ));

    // Connect API signals to engine signals
    connect(m_api, &SpreadsheetAPI::logMessage,
            this, &MacroEngine::logMessage);
    connect(m_api, &SpreadsheetAPI::alertRequested,
            this, [this](const QString& message) {
        QMessageBox::information(nullptr, tr("Macro Alert"), message);
    });
}

void MacroEngine::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet)
{
    m_api->setSpreadsheet(std::move(spreadsheet));
}

MacroEngine::MacroResult MacroEngine::execute(const QString& code)
{
    MacroResult result;
    result.success = true;

    // Set up a 10-second timeout to prevent infinite loops
    m_engine.setInterrupted(false);
    QTimer timeoutTimer;
    timeoutTimer.setSingleShot(true);
    connect(&timeoutTimer, &QTimer::timeout, this, [this]() {
        m_engine.setInterrupted(true);
    });
    timeoutTimer.start(10000);

    QJSValue jsResult = m_engine.evaluate(code);

    timeoutTimer.stop();
    m_engine.setInterrupted(false);

    if (jsResult.isError()) {
        result.success = false;
        result.error = QStringLiteral("Line %1: %2")
                           .arg(jsResult.property("lineNumber").toInt())
                           .arg(jsResult.toString());
    } else if (!jsResult.isUndefined()) {
        result.output = jsResult.toString();
    }

    // Request a UI refresh after execution
    emit m_api->refreshRequested();
    emit executionComplete(result.success, result.success ? result.output : result.error);

    return result;
}

void MacroEngine::startRecording()
{
    m_recording = true;
    m_recordedCode.clear();
    emit recordingStarted();
}

void MacroEngine::stopRecording()
{
    m_recording = false;
    emit recordingStopped(m_recordedCode);
}

void MacroEngine::recordAction(const QString& jsLine)
{
    if (!m_recording) return;
    if (!m_recordedCode.isEmpty()) {
        m_recordedCode += '\n';
    }
    m_recordedCode += jsLine;
}

void MacroEngine::saveMacro(const SavedMacro& macro)
{
    // Update existing or add new
    for (auto& existing : m_savedMacros) {
        if (existing.name == macro.name) {
            existing.code = macro.code;
            existing.shortcut = macro.shortcut;
            goto persist;
        }
    }
    m_savedMacros.append(macro);

persist:
    QSettings settings;
    settings.beginGroup("macros");
    settings.beginGroup(macro.name);
    settings.setValue("code", macro.code);
    settings.setValue("shortcut", macro.shortcut);
    settings.endGroup();
    settings.endGroup();
}

void MacroEngine::deleteMacro(const QString& name)
{
    m_savedMacros.removeIf([&name](const SavedMacro& m) {
        return m.name == name;
    });

    QSettings settings;
    settings.beginGroup("macros");
    settings.remove(name);
    settings.endGroup();
}

QList<SavedMacro> MacroEngine::getSavedMacros() const
{
    return m_savedMacros;
}

void MacroEngine::loadMacros()
{
    m_savedMacros.clear();

    QSettings settings;
    settings.beginGroup("macros");
    const QStringList macroNames = settings.childGroups();
    for (const QString& name : macroNames) {
        settings.beginGroup(name);
        SavedMacro macro;
        macro.name = name;
        macro.code = settings.value("code").toString();
        macro.shortcut = settings.value("shortcut").toString();
        m_savedMacros.append(macro);
        settings.endGroup();
    }
    settings.endGroup();
}
