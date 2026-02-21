#ifndef XLSXSERVICE_H
#define XLSXSERVICE_H

#include <QString>
#include <QStringList>
#include <QColor>
#include <QVector>
#include <memory>
#include <vector>
#include "../core/Spreadsheet.h"
#include "../core/Cell.h"

// Chart import data structures
struct ImportedChartSeries {
    QString name;
    QVector<double> values;      // numeric y values
    QVector<double> xNumeric;    // numeric x values (scatter charts)
    QVector<QString> categories; // string categories (bar/line/etc)
};

struct ImportedChart {
    int sheetIndex = 0;
    QString chartType;  // "column", "bar", "line", "area", "scatter", "pie", "donut"
    QString title;
    QString xAxisTitle;
    QString yAxisTitle;
    QVector<ImportedChartSeries> series;
    int x = 50, y = 50, width = 420, height = 320;
};

struct XlsxImportResult {
    std::vector<std::shared_ptr<Spreadsheet>> sheets;
    std::vector<ImportedChart> charts;
};

class XlsxService {
public:
    // Returns a vector of sheets (one per worksheet in the xlsx file)
    static XlsxImportResult importFromFile(const QString& filePath);

    // Export sheets to XLSX with all formatting
    static bool exportToFile(const std::vector<std::shared_ptr<Spreadsheet>>& sheets, const QString& filePath);

private:
    // Export helpers
    static QString columnIndexToLetter(int col);
    static QByteArray generateContentTypes(int sheetCount);
    static QByteArray generateRels();
    static QByteArray generateWorkbook(const std::vector<std::shared_ptr<Spreadsheet>>& sheets);
    static QByteArray generateWorkbookRels(int sheetCount);
    static QByteArray generateStyles(const std::vector<std::shared_ptr<Spreadsheet>>& sheets,
                                      std::map<QString, int>& styleIndexMap);
    static QByteArray generateSheet(Spreadsheet* sheet, const std::map<QString, int>& styleIndexMap,
                                     QStringList& sharedStrings);
    static QByteArray generateSharedStrings(const QStringList& sharedStrings);
    static QString cellStyleKey(const CellStyle& style);

    struct SheetInfo {
        QString name;
        QString rId;
        QString filePath;  // e.g. "worksheets/sheet1.xml"
    };

    struct XlsxFont {
        QString name = "Arial";
        int size = 11;
        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool strikethrough = false;
        QColor color = QColor("#000000");
    };

    struct XlsxFill {
        QColor fgColor = QColor("#FFFFFF");
        bool hasFg = false;
    };

    struct XlsxBorderSide {
        bool enabled = false;
        QString color = "#000000";
        int width = 1; // 1=thin, 2=medium, 3=thick
    };

    struct XlsxBorder {
        XlsxBorderSide left, right, top, bottom;
    };

    struct XlsxCellXf {
        int fontId = 0;
        int fillId = 0;
        int borderId = 0;
        int numFmtId = 0;
        HorizontalAlignment hAlign = HorizontalAlignment::General;
        VerticalAlignment vAlign = VerticalAlignment::Bottom;
        bool applyFont = false;
        bool applyFill = false;
        bool applyBorder = false;
        bool applyAlignment = false;
        bool applyNumberFormat = false;
    };

    static QStringList parseSharedStrings(const QByteArray& xmlData);
    static std::vector<SheetInfo> parseWorkbook(const QByteArray& workbookXml, const QByteArray& relsXml);
    static std::vector<XlsxFont> parseFonts(const QByteArray& stylesXml);
    static std::vector<XlsxFill> parseFills(const QByteArray& stylesXml);
    static std::vector<XlsxBorder> parseBorders(const QByteArray& stylesXml);
    static std::vector<XlsxCellXf> parseCellXfs(const QByteArray& stylesXml);
    static std::map<int, QString> parseNumFmts(const QByteArray& stylesXml);
    static CellStyle buildCellStyle(const XlsxCellXf& xf,
                                     const std::vector<XlsxFont>& fonts,
                                     const std::vector<XlsxFill>& fills,
                                     const std::vector<XlsxBorder>& borders,
                                     int numFmtId,
                                     const std::map<int, QString>& customNumFmts);
    static void parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                           const std::vector<CellStyle>& styles, Spreadsheet* sheet);
    static int columnLetterToIndex(const QString& letters);
    static QString mapNumFmtId(int id, const std::map<int, QString>& customNumFmts);
    static bool isDateFormatCode(const QString& formatCode);

    // Chart import helpers
    struct DrawingChartRef {
        QString chartRId;
        int fromCol = 0, fromRow = 0;
        int toCol = 10, toRow = 15;
    };

    static std::map<QString, QString> parseRels(const QByteArray& relsXml);
    static QString findDrawingRId(const QByteArray& sheetXml);
    static std::vector<DrawingChartRef> parseDrawing(const QByteArray& drawingXml);
    static ImportedChart parseChartXml(const QByteArray& chartXml);
    static QString resolveRelativePath(const QString& basePath, const QString& relativePath);
};

#endif // XLSXSERVICE_H
