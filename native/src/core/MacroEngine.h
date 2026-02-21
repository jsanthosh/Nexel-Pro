#ifndef MACROENGINE_H
#define MACROENGINE_H

#include <QObject>
#include <QJSEngine>
#include <QJSValue>
#include <QString>
#include <QStringList>
#include <QSettings>
#include <memory>
#include "CellRange.h"

class Spreadsheet;

class SpreadsheetAPI : public QObject {
    Q_OBJECT
public:
    explicit SpreadsheetAPI(QObject* parent = nullptr);

    void setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet);

    Q_INVOKABLE QVariant getCellValue(const QString& cellRef);
    Q_INVOKABLE void setCellValue(const QString& cellRef, const QVariant& value);
    Q_INVOKABLE void setCellFormula(const QString& cellRef, const QString& formula);
    Q_INVOKABLE QString getCellFormula(const QString& cellRef);

    Q_INVOKABLE void setBold(const QString& range, bool bold);
    Q_INVOKABLE void setItalic(const QString& range, bool italic);
    Q_INVOKABLE void setBackgroundColor(const QString& range, const QString& color);
    Q_INVOKABLE void setForegroundColor(const QString& range, const QString& color);
    Q_INVOKABLE void setFontSize(const QString& range, int size);
    Q_INVOKABLE void setNumberFormat(const QString& range, const QString& format);

    Q_INVOKABLE void mergeCells(const QString& range);
    Q_INVOKABLE void unmergeCells(const QString& range);

    Q_INVOKABLE void setRowHeight(int row, int height);
    Q_INVOKABLE void setColumnWidth(int col, int width);

    Q_INVOKABLE int getMaxRow();
    Q_INVOKABLE int getMaxColumn();
    Q_INVOKABLE QString getSheetName();

    Q_INVOKABLE void clearRange(const QString& range);

    Q_INVOKABLE void alert(const QString& message);
    Q_INVOKABLE void log(const QString& message);

signals:
    void logMessage(const QString& message);
    void alertRequested(const QString& message);
    void refreshRequested();

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    CellAddress parseCellRef(const QString& ref);
    void applyStyleChange(const QString& range, std::function<void(struct CellStyle&)> modifier);
};

struct SavedMacro {
    QString name;
    QString code;
    QString shortcut;
};

class MacroEngine : public QObject {
    Q_OBJECT
public:
    explicit MacroEngine(QObject* parent = nullptr);

    void setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet);

    struct MacroResult {
        bool success;
        QString output;
        QString error;
    };
    MacroResult execute(const QString& code);

    void startRecording();
    void stopRecording();
    bool isRecording() const { return m_recording; }
    void recordAction(const QString& jsLine);
    QString getRecordedCode() const { return m_recordedCode; }

    void saveMacro(const SavedMacro& macro);
    void deleteMacro(const QString& name);
    QList<SavedMacro> getSavedMacros() const;
    void loadMacros();

signals:
    void recordingStarted();
    void recordingStopped(const QString& code);
    void logMessage(const QString& message);
    void executionComplete(bool success, const QString& output);

private:
    void setupEngine();

    QJSEngine m_engine;
    SpreadsheetAPI* m_api;
    bool m_recording = false;
    QString m_recordedCode;
    QList<SavedMacro> m_savedMacros;
};

#endif // MACROENGINE_H
