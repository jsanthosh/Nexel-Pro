#include "XlsxService.h"
#include "../core/DocumentTheme.h"
#include <QFile>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QRegularExpression>
#include <QDate>
#include <QBuffer>
#include <QtCore/private/qzipreader_p.h>
#include <QtCore/private/qzipwriter_p.h>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <algorithm>
#include <cmath>

// Resolve a color string (theme or absolute) to "#RRGGBB" for XLSX export
static QString resolveColorForExport(const QString& colorStr) {
    if (DocumentTheme::isThemeColor(colorStr)) {
        QColor c = defaultDocumentTheme().resolveColor(colorStr);
        return c.isValid() ? c.name() : "#000000";
    }
    return colorStr;
}

XlsxImportResult XlsxService::importFromFile(const QString& filePath) {
    XlsxImportResult result;

    QZipReader zip(filePath);
    if (!zip.isReadable()) {
        return result;
    }

    // Read shared strings
    QStringList sharedStrings;
    QByteArray ssData = zip.fileData("xl/sharedStrings.xml");
    if (!ssData.isEmpty()) {
        sharedStrings = parseSharedStrings(ssData);
    }

    // Parse styles
    std::vector<CellStyle> styles;
    QByteArray stylesData = zip.fileData("xl/styles.xml");
    if (!stylesData.isEmpty()) {
        auto fonts = parseFonts(stylesData);
        auto fills = parseFills(stylesData);
        auto borders = parseBorders(stylesData);
        auto cellXfs = parseCellXfs(stylesData);
        auto customNumFmts = parseNumFmts(stylesData);
        for (const auto& xf : cellXfs) {
            styles.push_back(buildCellStyle(xf, fonts, fills, borders, xf.numFmtId, customNumFmts));
        }
    }

    // Parse workbook for sheet names and paths
    QByteArray workbookData = zip.fileData("xl/workbook.xml");
    QByteArray relsData = zip.fileData("xl/_rels/workbook.xml.rels");
    auto sheetInfos = parseWorkbook(workbookData, relsData);

    if (sheetInfos.empty()) {
        // Fallback: try sheet1.xml directly
        SheetInfo si;
        si.name = "Sheet1";
        si.filePath = "worksheets/sheet1.xml";
        sheetInfos.push_back(si);
    }

    // Parse each sheet
    for (int sheetIdx = 0; sheetIdx < static_cast<int>(sheetInfos.size()); ++sheetIdx) {
        const auto& info = sheetInfos[sheetIdx];
        QString path = "xl/" + info.filePath;
        QByteArray sheetData = zip.fileData(path);
        if (sheetData.isEmpty()) continue;

        auto spreadsheet = std::make_shared<Spreadsheet>();
        spreadsheet->setAutoRecalculate(false);
        spreadsheet->setSheetName(info.name);

        parseSheet(sheetData, sharedStrings, styles, spreadsheet.get());

        // Parse sheet rels (needed for hyperlinks and charts)
        QString fullSheetPath = "xl/" + info.filePath;
        int lastSlash = fullSheetPath.lastIndexOf('/');
        QString sheetDirPath = fullSheetPath.left(lastSlash);
        QString sheetFileName = fullSheetPath.mid(lastSlash + 1);
        QString sheetRelsPath = sheetDirPath + "/_rels/" + sheetFileName + ".rels";
        QByteArray sheetRelsData = zip.fileData(sheetRelsPath);
        auto sheetRels = parseRels(sheetRelsData);

        // ---- Hyperlink import: parse <hyperlink> elements from sheet XML ----
        {
            QXmlStreamReader hxml(sheetData);
            while (!hxml.atEnd()) {
                hxml.readNext();
                if (hxml.isStartElement() && hxml.name() == u"hyperlink") {
                    QString ref = hxml.attributes().value("ref").toString();
                    // External hyperlink via relationship
                    QString rId;
                    static const QString relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    rId = hxml.attributes().value(relNs, "id").toString();
                    if (rId.isEmpty()) rId = hxml.attributes().value("r:id").toString();
                    // Internal hyperlink (location within workbook)
                    QString location = hxml.attributes().value("location").toString();
                    QString display = hxml.attributes().value("display").toString();

                    QString url;
                    if (!rId.isEmpty()) {
                        auto it = sheetRels.find(rId);
                        if (it != sheetRels.end()) url = it->second;
                    } else if (!location.isEmpty()) {
                        url = location;
                    }

                    if (!url.isEmpty() && !ref.isEmpty()) {
                        // Parse cell reference
                        int row = -1, col = -1;
                        QString letters;
                        int numStart = -1;
                        for (int i = 0; i < ref.length(); ++i) {
                            if (ref[i].isLetter()) letters += ref[i];
                            else { numStart = i; break; }
                        }
                        if (numStart > 0) {
                            col = columnLetterToIndex(letters);
                            row = ref.mid(numStart).toInt() - 1;
                        }
                        if (row >= 0 && col >= 0) {
                            auto cell = spreadsheet->getCell(CellAddress(row, col));
                            cell->setHyperlink(url);
                        }
                    }
                }
            }
        }

        // Auto-expand row/col count
        int maxRow = spreadsheet->getMaxRow();
        int maxCol = spreadsheet->getMaxColumn();
        spreadsheet->setRowCount(std::max(1000, maxRow + 100));
        spreadsheet->setColumnCount(std::max(256, maxCol + 10));

        spreadsheet->setAutoRecalculate(true);
        result.sheets.push_back(spreadsheet);

        // ---- Chart import: scan for embedded charts ----
        QString drawingRId = findDrawingRId(sheetData);
        if (drawingRId.isEmpty()) continue;

        auto drawIt = sheetRels.find(drawingRId);
        if (drawIt == sheetRels.end()) continue;

        QString drawingPath = resolveRelativePath(fullSheetPath, drawIt->second);
        QByteArray drawingData = zip.fileData(drawingPath);
        if (drawingData.isEmpty()) continue;

        auto chartRefs = parseDrawing(drawingData);
        if (chartRefs.empty()) continue;

        // Parse drawing rels to find chart paths
        int drawLastSlash = drawingPath.lastIndexOf('/');
        QString drawingDir = drawingPath.left(drawLastSlash);
        QString drawingFileName = drawingPath.mid(drawLastSlash + 1);
        QString drawingRelsPath = drawingDir + "/_rels/" + drawingFileName + ".rels";
        QByteArray drawingRelsData = zip.fileData(drawingRelsPath);
        auto drawingRels = parseRels(drawingRelsData);

        for (const auto& ref : chartRefs) {
            auto chartIt = drawingRels.find(ref.chartRId);
            if (chartIt == drawingRels.end()) continue;

            QString chartPath = resolveRelativePath(drawingPath, chartIt->second);
            QByteArray chartData = zip.fileData(chartPath);
            if (chartData.isEmpty()) continue;

            ImportedChart chart = parseChartXml(chartData);
            chart.sheetIndex = sheetIdx;
            chart.x = ref.fromCol * 64;
            chart.y = ref.fromRow * 20;
            chart.width = qMax(200, (ref.toCol - ref.fromCol) * 64);
            chart.height = qMax(150, (ref.toRow - ref.fromRow) * 20);
            result.charts.push_back(chart);
        }
    }

    // Load Nexel-native metadata (picklist/checkbox definitions)
    QByteArray nexelMetaData = zip.fileData("docMetadata/nexel-metadata.json");
    if (nexelMetaData.isEmpty()) nexelMetaData = zip.fileData("xl/nexel-metadata.json"); // legacy
    if (!nexelMetaData.isEmpty()) {
        QJsonDocument metaDoc = QJsonDocument::fromJson(nexelMetaData);
        if (metaDoc.isObject()) {
            QJsonArray sheetsArray = metaDoc.object()["sheets"].toArray();
            for (const auto& sheetVal : sheetsArray) {
                QJsonObject sheetObj = sheetVal.toObject();
                int si = sheetObj["index"].toInt();
                if (si < 0 || si >= static_cast<int>(result.sheets.size())) continue;
                auto* sheet = result.sheets[si].get();

                // Restore picklist validation rules
                if (sheetObj.contains("picklists")) {
                    for (const auto& plVal : sheetObj["picklists"].toArray()) {
                        QJsonObject plObj = plVal.toObject();
                        Spreadsheet::DataValidationRule rule;
                        rule.range = CellRange(plObj["range"].toString());
                        rule.type = Spreadsheet::DataValidationRule::List;
                        rule.showErrorAlert = false;
                        for (const auto& item : plObj["options"].toArray()) {
                            rule.listItems.append(item.toString());
                        }
                        if (plObj.contains("sourceRange")) {
                            rule.listSourceRange = plObj["sourceRange"].toString();
                            rule.listSortMode = plObj["sortMode"].toInt(0);
                            rule.listIgnoreBlanks = plObj["ignoreBlanks"].toBool(true);
                        }
                        if (plObj.contains("optionColors")) {
                            for (const auto& c : plObj["optionColors"].toArray())
                                rule.listItemColors.append(c.toString());
                        }
                        sheet->addValidationRule(rule);
                        // Set Picklist numberFormat on cells in range
                        for (const auto& addr : rule.range.getCells()) {
                            auto cell = sheet->getCell(addr);
                            CellStyle style = cell->getStyle();
                            style.numberFormat = "Picklist";
                            cell->setStyle(style);
                        }
                    }
                }

                // Restore table styles
                if (sheetObj.contains("tables")) {
                    for (const auto& tblVal : sheetObj["tables"].toArray()) {
                        QJsonObject tblObj = tblVal.toObject();
                        SpreadsheetTable tbl;
                        tbl.name = tblObj["name"].toString();
                        tbl.range = CellRange(tblObj["range"].toString());
                        tbl.hasHeaderRow = tblObj["hasHeaderRow"].toBool(true);
                        tbl.bandedRows = tblObj["bandedRows"].toBool(true);
                        for (const auto& cn : tblObj["columnNames"].toArray())
                            tbl.columnNames.append(cn.toString());
                        QJsonObject themeObj = tblObj["theme"].toObject();
                        tbl.theme.name = themeObj["name"].toString();
                        tbl.theme.headerBg = QColor(themeObj["headerBg"].toString());
                        tbl.theme.headerFg = QColor(themeObj["headerFg"].toString());
                        tbl.theme.bandedRow1 = QColor(themeObj["bandedRow1"].toString());
                        tbl.theme.bandedRow2 = QColor(themeObj["bandedRow2"].toString());
                        tbl.theme.borderColor = QColor(themeObj["borderColor"].toString());
                        sheet->addTable(tbl);
                    }
                }

                // Restore checkboxes
                if (sheetObj.contains("checkboxes")) {
                    for (const auto& cbVal : sheetObj["checkboxes"].toArray()) {
                        QJsonObject cbObj = cbVal.toObject();
                        CellAddress addr = CellAddress::fromString(cbObj["cell"].toString());
                        auto cell = sheet->getCell(addr);
                        CellStyle style = cell->getStyle();
                        style.numberFormat = "Checkbox";
                        cell->setStyle(style);
                        cell->setValue(QVariant(cbObj["checked"].toBool()));
                    }
                }

                // Restore conditional formatting rules
                if (sheetObj.contains("conditionalFormats")) {
                    for (const auto& cfVal : sheetObj["conditionalFormats"].toArray()) {
                        QJsonObject cfObj = cfVal.toObject();
                        CellRange range(cfObj["range"].toString());
                        auto condType = static_cast<ConditionType>(cfObj["type"].toInt());
                        auto rule = std::make_shared<ConditionalFormat>(range, condType);
                        if (cfObj.contains("value1"))
                            rule->setValue1(cfObj["value1"].toVariant());
                        if (cfObj.contains("value2"))
                            rule->setValue2(cfObj["value2"].toVariant());
                        if (cfObj.contains("formula"))
                            rule->setFormula(cfObj["formula"].toString());
                        QJsonObject styleObj = cfObj["style"].toObject();
                        CellStyle st;
                        st.bold = styleObj["bold"].toBool();
                        st.italic = styleObj["italic"].toBool();
                        st.underline = styleObj["underline"].toBool();
                        if (styleObj.contains("fontColor"))
                            st.foregroundColor = styleObj["fontColor"].toString();
                        if (styleObj.contains("bgColor"))
                            st.backgroundColor = styleObj["bgColor"].toString();
                        rule->setStyle(st);

                        // Visual formatting configs
                        if (cfObj.contains("dataBar")) {
                            QJsonObject db = cfObj["dataBar"].toObject();
                            DataBarConfig cfg;
                            cfg.barColor = QColor(db["color"].toString());
                            cfg.minValue = db["minValue"].toDouble(0);
                            cfg.maxValue = db["maxValue"].toDouble(100);
                            cfg.autoRange = db["autoRange"].toBool(true);
                            cfg.showValue = db["showValue"].toBool(true);
                            rule->setDataBarConfig(cfg);
                        }
                        if (cfObj.contains("colorScale")) {
                            QJsonObject cs = cfObj["colorScale"].toObject();
                            ColorScaleConfig cfg;
                            cfg.minColor = QColor(cs["minColor"].toString());
                            cfg.midColor = QColor(cs["midColor"].toString());
                            cfg.maxColor = QColor(cs["maxColor"].toString());
                            cfg.minValue = cs["minValue"].toDouble(0);
                            cfg.maxValue = cs["maxValue"].toDouble(100);
                            cfg.autoRange = cs["autoRange"].toBool(true);
                            cfg.threeColor = cs["threeColor"].toBool(false);
                            rule->setColorScaleConfig(cfg);
                        }
                        if (cfObj.contains("iconSet")) {
                            QJsonObject is = cfObj["iconSet"].toObject();
                            IconSetConfig cfg;
                            cfg.iconType = static_cast<IconSetConfig::IconType>(is["iconType"].toInt(0));
                            cfg.threshold1 = is["threshold1"].toDouble(33.33);
                            cfg.threshold2 = is["threshold2"].toDouble(66.67);
                            cfg.showValue = is["showValue"].toBool(true);
                            cfg.reverseOrder = is["reverseOrder"].toBool(false);
                            rule->setIconSetConfig(cfg);
                        }

                        sheet->getConditionalFormatting().addRule(rule);
                    }
                }
            }
        }
    }

    // Load Nexel-native chart configs if present — these are the authoritative
    // source, so they replace any Excel chart entries parsed above.
    QByteArray nexelChartsData = zip.fileData("docMetadata/nexel-charts.json");
    if (nexelChartsData.isEmpty()) nexelChartsData = zip.fileData("xl/nexel-charts.json"); // legacy
    if (!nexelChartsData.isEmpty()) {
        QJsonDocument doc = QJsonDocument::fromJson(nexelChartsData);
        if (doc.isArray() && !doc.array().isEmpty()) {
            result.charts.clear();  // drop Excel chart duplicates
            for (const auto& val : doc.array()) {
                QJsonObject obj = val.toObject();
                ImportedChart chart;
                chart.sheetIndex = obj["sheetIndex"].toInt();
                chart.chartType = obj["chartType"].toString();
                chart.title = obj["title"].toString();
                chart.xAxisTitle = obj["xAxisTitle"].toString();
                chart.yAxisTitle = obj["yAxisTitle"].toString();
                chart.dataRange = obj["dataRange"].toString();
                chart.themeIndex = obj["themeIndex"].toInt();
                chart.showLegend = obj["showLegend"].toBool(true);
                chart.showGridLines = obj["showGridLines"].toBool(true);
                chart.x = obj["x"].toInt(50);
                chart.y = obj["y"].toInt(50);
                chart.width = obj["width"].toInt(420);
                chart.height = obj["height"].toInt(320);
                chart.isNexelNative = true;
                result.charts.push_back(chart);
            }
        }
    }

    zip.close();
    return result;
}

std::vector<XlsxService::SheetInfo> XlsxService::parseWorkbook(const QByteArray& workbookXml,
                                                                  const QByteArray& relsXml) {
    std::vector<SheetInfo> sheets;

    // Parse relationships: rId -> target path
    std::map<QString, QString> relMap;
    if (!relsXml.isEmpty()) {
        QXmlStreamReader relsReader(relsXml);
        while (!relsReader.atEnd()) {
            relsReader.readNext();
            if (relsReader.isStartElement() && relsReader.name() == u"Relationship") {
                QString id = relsReader.attributes().value("Id").toString();
                QString target = relsReader.attributes().value("Target").toString();
                relMap[id] = target;
            }
        }
    }

    // Parse workbook: sheet names and rIds
    if (!workbookXml.isEmpty()) {
        QXmlStreamReader xml(workbookXml);
        while (!xml.atEnd()) {
            xml.readNext();
            if (xml.isStartElement() && xml.name() == u"sheet") {
                SheetInfo info;
                info.name = xml.attributes().value("name").toString();
                info.rId = xml.attributes().value(u"r:id").toString();
                // Try r:id first, then Id
                if (info.rId.isEmpty()) {
                    // Some xlsx files use different namespace
                    for (const auto& attr : xml.attributes()) {
                        if (attr.name() == u"id" || attr.qualifiedName().toString().contains("id", Qt::CaseInsensitive)) {
                            info.rId = attr.value().toString();
                            break;
                        }
                    }
                }
                auto it = relMap.find(info.rId);
                if (it != relMap.end()) {
                    info.filePath = it->second;
                }
                if (!info.filePath.isEmpty()) {
                    sheets.push_back(info);
                }
            }
        }
    }

    return sheets;
}

QStringList XlsxService::parseSharedStrings(const QByteArray& xmlData) {
    QStringList strings;
    QXmlStreamReader xml(xmlData);

    QString currentString;
    bool inSi = false;
    bool inT = false;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"sst") {
                // Pre-reserve from uniqueCount to avoid QStringList reallocations
                int count = xml.attributes().value("uniqueCount").toInt();
                if (count > 0) strings.reserve(count);
            } else if (xml.name() == u"si") {
                inSi = true;
                currentString.clear();
            } else if (inSi && xml.name() == u"t") {
                inT = true;
            }
        } else if (xml.isCharacters() && inT) {
            // Only capture text inside <t> elements, not XML indentation whitespace
            currentString += xml.text().toString();
        } else if (xml.isEndElement()) {
            if (xml.name() == u"t") {
                inT = false;
            } else if (xml.name() == u"si") {
                strings.append(currentString);
                inSi = false;
            }
        }
    }

    return strings;
}

std::vector<XlsxService::XlsxFont> XlsxService::parseFonts(const QByteArray& stylesXml) {
    std::vector<XlsxFont> fonts;
    QXmlStreamReader xml(stylesXml);

    bool inFonts = false;
    bool inFont = false;
    XlsxFont currentFont;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"fonts") {
                inFonts = true;
            } else if (inFonts && xml.name() == u"font") {
                inFont = true;
                currentFont = XlsxFont();
            } else if (inFont) {
                if (xml.name() == u"name") {
                    currentFont.name = xml.attributes().value("val").toString();
                } else if (xml.name() == u"sz") {
                    currentFont.size = xml.attributes().value("val").toInt();
                } else if (xml.name() == u"b") {
                    QString val = xml.attributes().value("val").toString();
                    currentFont.bold = (val.isEmpty() || val != "0");
                } else if (xml.name() == u"i") {
                    QString val = xml.attributes().value("val").toString();
                    currentFont.italic = (val.isEmpty() || val != "0");
                } else if (xml.name() == u"u") {
                    currentFont.underline = true;
                } else if (xml.name() == u"strike") {
                    currentFont.strikethrough = true;
                } else if (xml.name() == u"color") {
                    QString rgb = xml.attributes().value("rgb").toString();
                    if (!rgb.isEmpty()) {
                        // ARGB format: strip alpha if 8 chars
                        if (rgb.length() == 8) rgb = rgb.mid(2);
                        currentFont.color = QColor("#" + rgb);
                    }
                }
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"font" && inFonts) {
                fonts.push_back(currentFont);
                inFont = false;
            } else if (xml.name() == u"fonts") {
                inFonts = false;
                break;  // done with fonts section
            }
        }
    }

    return fonts;
}

std::vector<XlsxService::XlsxFill> XlsxService::parseFills(const QByteArray& stylesXml) {
    std::vector<XlsxFill> fills;
    QXmlStreamReader xml(stylesXml);

    bool inFills = false;
    bool inFill = false;
    XlsxFill currentFill;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"fills") {
                inFills = true;
            } else if (inFills && xml.name() == u"fill") {
                inFill = true;
                currentFill = XlsxFill();
            } else if (inFill && xml.name() == u"fgColor") {
                QString rgb = xml.attributes().value("rgb").toString();
                if (!rgb.isEmpty()) {
                    if (rgb.length() == 8) rgb = rgb.mid(2);
                    currentFill.fgColor = QColor("#" + rgb);
                    currentFill.hasFg = true;
                }
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"fill" && inFills) {
                fills.push_back(currentFill);
                inFill = false;
            } else if (xml.name() == u"fills") {
                inFills = false;
                break;
            }
        }
    }

    return fills;
}

std::vector<XlsxService::XlsxBorder> XlsxService::parseBorders(const QByteArray& stylesXml) {
    std::vector<XlsxBorder> borders;
    QXmlStreamReader xml(stylesXml);

    bool inBorders = false;
    bool inBorder = false;
    XlsxBorder currentBorder;
    QString currentSide;

    auto parseBorderStyle = [](const QString& style) -> int {
        if (style == "thin" || style == "hair") return 1;
        if (style == "medium" || style == "dashed" || style == "dotted") return 2;
        if (style == "thick" || style == "double") return 3;
        if (!style.isEmpty()) return 1; // any non-empty style means border exists
        return 0;
    };

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"borders") {
                inBorders = true;
            } else if (inBorders && xml.name() == u"border") {
                inBorder = true;
                currentBorder = XlsxBorder();
            } else if (inBorder) {
                QString name = xml.name().toString();
                if (name == "left" || name == "right" || name == "top" || name == "bottom") {
                    currentSide = name;
                    QString style = xml.attributes().value("style").toString();
                    int width = parseBorderStyle(style);
                    if (width > 0) {
                        XlsxBorderSide side;
                        side.enabled = true;
                        side.width = width;
                        if (name == "left") currentBorder.left = side;
                        else if (name == "right") currentBorder.right = side;
                        else if (name == "top") currentBorder.top = side;
                        else if (name == "bottom") currentBorder.bottom = side;
                    }
                } else if (xml.name() == u"color" && !currentSide.isEmpty()) {
                    QString rgb = xml.attributes().value("rgb").toString();
                    if (!rgb.isEmpty()) {
                        if (rgb.length() == 8) rgb = rgb.mid(2);
                        QString color = "#" + rgb;
                        if (currentSide == "left") currentBorder.left.color = color;
                        else if (currentSide == "right") currentBorder.right.color = color;
                        else if (currentSide == "top") currentBorder.top.color = color;
                        else if (currentSide == "bottom") currentBorder.bottom.color = color;
                    }
                }
            }
        } else if (xml.isEndElement()) {
            if (inBorder) {
                QString name = xml.name().toString();
                if (name == "left" || name == "right" || name == "top" || name == "bottom") {
                    currentSide.clear();
                }
            }
            if (xml.name() == u"border" && inBorders) {
                borders.push_back(currentBorder);
                inBorder = false;
            } else if (xml.name() == u"borders") {
                inBorders = false;
                break;
            }
        }
    }

    return borders;
}

std::vector<XlsxService::XlsxCellXf> XlsxService::parseCellXfs(const QByteArray& stylesXml) {
    std::vector<XlsxCellXf> xfs;
    QXmlStreamReader xml(stylesXml);

    bool inCellXfs = false;
    bool inXf = false;
    XlsxCellXf currentXf;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"cellXfs") {
                inCellXfs = true;
            } else if (inCellXfs && xml.name() == u"xf") {
                inXf = true;
                currentXf = XlsxCellXf();
                currentXf.fontId = xml.attributes().value("fontId").toInt();
                currentXf.fillId = xml.attributes().value("fillId").toInt();
                currentXf.borderId = xml.attributes().value("borderId").toInt();
                currentXf.numFmtId = xml.attributes().value("numFmtId").toInt();
                currentXf.applyFont = (xml.attributes().value("applyFont") == u"1");
                currentXf.applyFill = (xml.attributes().value("applyFill") == u"1");
                currentXf.applyBorder = (xml.attributes().value("applyBorder") == u"1");
                currentXf.applyAlignment = (xml.attributes().value("applyAlignment") == u"1");
                currentXf.applyNumberFormat = (xml.attributes().value("applyNumberFormat") == u"1");
            } else if (inXf && xml.name() == u"alignment") {
                QString h = xml.attributes().value("horizontal").toString();
                QString v = xml.attributes().value("vertical").toString();
                if (h == "left") currentXf.hAlign = HorizontalAlignment::Left;
                else if (h == "center") currentXf.hAlign = HorizontalAlignment::Center;
                else if (h == "right") currentXf.hAlign = HorizontalAlignment::Right;
                else currentXf.hAlign = HorizontalAlignment::General;

                if (v == "top") currentXf.vAlign = VerticalAlignment::Top;
                else if (v == "center") currentXf.vAlign = VerticalAlignment::Middle;
                else currentXf.vAlign = VerticalAlignment::Bottom;
                currentXf.applyAlignment = true;
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"xf" && inCellXfs) {
                xfs.push_back(currentXf);
                inXf = false;
            } else if (xml.name() == u"cellXfs") {
                inCellXfs = false;
                break;
            }
        }
    }

    return xfs;
}

CellStyle XlsxService::buildCellStyle(const XlsxCellXf& xf,
                                        const std::vector<XlsxFont>& fonts,
                                        const std::vector<XlsxFill>& fills,
                                        const std::vector<XlsxBorder>& borders,
                                        int numFmtId,
                                        const std::map<int, QString>& customNumFmts) {
    CellStyle style;

    // Apply font
    if (xf.fontId >= 0 && xf.fontId < static_cast<int>(fonts.size())) {
        const auto& f = fonts[xf.fontId];
        style.fontName = f.name;
        style.fontSize = f.size;
        style.bold = f.bold;
        style.italic = f.italic;
        style.underline = f.underline;
        style.strikethrough = f.strikethrough;
        style.foregroundColor = f.color.name();
    }

    // Apply fill
    if (xf.fillId >= 0 && xf.fillId < static_cast<int>(fills.size())) {
        const auto& fl = fills[xf.fillId];
        if (fl.hasFg) {
            style.backgroundColor = fl.fgColor.name();
        }
    }

    // Apply borders
    if (xf.borderId >= 0 && xf.borderId < static_cast<int>(borders.size())) {
        const auto& b = borders[xf.borderId];
        if (b.left.enabled) {
            style.borderLeft.enabled = true;
            style.borderLeft.color = b.left.color;
            style.borderLeft.width = b.left.width;
        }
        if (b.right.enabled) {
            style.borderRight.enabled = true;
            style.borderRight.color = b.right.color;
            style.borderRight.width = b.right.width;
        }
        if (b.top.enabled) {
            style.borderTop.enabled = true;
            style.borderTop.color = b.top.color;
            style.borderTop.width = b.top.width;
        }
        if (b.bottom.enabled) {
            style.borderBottom.enabled = true;
            style.borderBottom.color = b.bottom.color;
            style.borderBottom.width = b.bottom.width;
        }
    }

    // Apply alignment
    style.hAlign = xf.hAlign;
    style.vAlign = xf.vAlign;

    // Apply number format
    style.numberFormat = mapNumFmtId(numFmtId, customNumFmts);

    return style;
}

QString XlsxService::mapNumFmtId(int id, const std::map<int, QString>& customNumFmts) {
    // XLSX built-in number format IDs
    if (id == 0) return "General";
    if (id >= 1 && id <= 4) return "Number";
    if (id >= 5 && id <= 8) return "Currency";
    if (id >= 9 && id <= 10) return "Percentage";
    if (id >= 14 && id <= 22) return "Date";
    if (id >= 45 && id <= 47) return "Time";
    if (id == 49) return "Text";

    // Check custom formats (id >= 164 typically)
    auto it = customNumFmts.find(id);
    if (it != customNumFmts.end()) {
        const QString& code = it->second;
        if (isDateFormatCode(code)) return "Date";
        // Check for time-only formats
        if (code.contains('h', Qt::CaseInsensitive) || code.contains("ss")) return "Time";
        // Check for percentage
        if (code.contains('%')) return "Percentage";
        // Check for currency symbols
        if (code.contains('$') || code.contains(QChar(0x20AC)) || code.contains(QChar(0x00A3))) return "Currency";
        // Numeric patterns
        if (code.contains('0') || code.contains('#')) return "Number";
    }

    return "General";
}

std::map<int, QString> XlsxService::parseNumFmts(const QByteArray& stylesXml) {
    std::map<int, QString> numFmts;
    QXmlStreamReader xml(stylesXml);

    bool inNumFmts = false;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"numFmts") {
                inNumFmts = true;
            } else if (inNumFmts && xml.name() == u"numFmt") {
                int id = xml.attributes().value("numFmtId").toInt();
                QString code = xml.attributes().value("formatCode").toString();
                numFmts[id] = code;
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"numFmts") {
                inNumFmts = false;
                break;
            }
        }
    }

    return numFmts;
}

bool XlsxService::isDateFormatCode(const QString& formatCode) {
    // Strip quoted literals (e.g. "AM/PM", text in brackets)
    QString cleaned = formatCode;
    // Remove quoted strings
    cleaned.replace(QRegularExpression("\"[^\"]*\""), "");
    // Remove bracketed color/locale codes like [$-409]
    cleaned.replace(QRegularExpression("\\[\\$?[^\\]]*\\]"), "");

    // Check for date/time tokens: y, m, d, h, s (but not standalone m which could be minutes)
    // Date patterns contain y or d; time patterns contain h or s
    bool hasYear = cleaned.contains('y', Qt::CaseInsensitive);
    bool hasDay = cleaned.contains('d', Qt::CaseInsensitive);
    bool hasMonth = cleaned.contains('m', Qt::CaseInsensitive);

    // If it has year or day tokens, it's a date format
    if (hasYear || hasDay) return true;

    // If it has month and also hour/second, it's a date-time
    bool hasHour = cleaned.contains('h', Qt::CaseInsensitive);
    bool hasSecond = cleaned.contains('s', Qt::CaseInsensitive);
    if (hasMonth && (hasHour || hasSecond)) return true;

    // Month alone (e.g. "mmm" or "mmmm") is also date
    if (hasMonth && !hasHour && !hasSecond) return true;

    return false;
}

// Fast inline cell reference parser — avoids regex overhead for millions of cells.
// Parses "A1", "AB123", "XFD1048576" etc. Returns false on failure.
static bool parseCellRef(const QStringView& ref, int& outRow, int& outCol) {
    const QChar* p = ref.data();
    const int len = ref.size();
    if (len == 0) return false;

    // Parse column letters (A-Z)
    int col = 0;
    int i = 0;
    while (i < len && p[i] >= u'A' && p[i] <= u'Z') {
        col = col * 26 + (p[i].unicode() - 'A' + 1);
        ++i;
    }
    if (i == 0 || i == len) return false; // No letters or no digits
    outCol = col - 1; // 0-based

    // Parse row digits
    int row = 0;
    while (i < len && p[i] >= u'0' && p[i] <= u'9') {
        row = row * 10 + (p[i].unicode() - '0');
        ++i;
    }
    if (i != len || row == 0) return false; // Trailing chars or row=0
    outRow = row - 1; // 0-based (XLSX is 1-based)
    return true;
}

// Fast inline range reference parser for "A1:B2" style strings
static bool parseRangeRef(const QStringView& ref,
                          int& startRow, int& startCol, int& endRow, int& endCol) {
    int colonPos = -1;
    for (int i = 0; i < ref.size(); ++i) {
        if (ref[i] == u':') { colonPos = i; break; }
    }
    if (colonPos < 0) return false;
    return parseCellRef(ref.left(colonPos), startRow, startCol) &&
           parseCellRef(ref.mid(colonPos + 1), endRow, endCol);
}

void XlsxService::parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                              const std::vector<CellStyle>& styles, Spreadsheet* sheet) {
    QXmlStreamReader xml(xmlData);

    // Pre-reserve cell storage based on XML size heuristic (~100 bytes per cell element)
    size_t estimatedCells = static_cast<size_t>(xmlData.size()) / 100;
    if (estimatedCells > 4096) {
        sheet->reserveCells(estimatedCells);
    }

    while (!xml.atEnd()) {
        xml.readNext();

        // Parse sheet view: <sheetView showGridLines="0" .../>
        if (xml.isStartElement() && xml.name() == u"sheetView") {
            QStringView showGrid = xml.attributes().value("showGridLines");
            if (showGrid == u"0") {
                sheet->setShowGridlines(false);
            }
            continue;
        }

        // Parse column widths: <col min="1" max="3" width="15.5" customWidth="1"/>
        if (xml.isStartElement() && xml.name() == u"col") {
            int minCol = xml.attributes().value("min").toInt() - 1;
            int maxCol = xml.attributes().value("max").toInt() - 1;
            double width = xml.attributes().value("width").toDouble();
            if (width > 0) {
                int pixelWidth = qMax(30, static_cast<int>(width * 7.5));
                int colLimit = qMin(maxCol, 16383); // XLSX max: XFD = 16384 cols
                for (int c = minCol; c <= colLimit; ++c) {
                    sheet->setColumnWidth(c, pixelWidth);
                }
            }
            continue;
        }

        // Parse row heights: <row r="1" ht="25.5" customHeight="1">
        if (xml.isStartElement() && xml.name() == u"row") {
            QString htStr = xml.attributes().value("ht").toString();
            if (!htStr.isEmpty()) {
                int rowIdx = xml.attributes().value("r").toInt() - 1;
                double ht = htStr.toDouble();
                if (ht > 0 && rowIdx >= 0) {
                    int pixelHeight = qMax(14, static_cast<int>(ht * 1.333));
                    sheet->setRowHeight(rowIdx, pixelHeight);
                }
            }
            continue; // row's child <c> elements will be hit next
        }

        // Parse merged cells: <mergeCell ref="A1:D1"/>
        if (xml.isStartElement() && xml.name() == u"mergeCell") {
            QStringView ref = xml.attributes().value("ref");
            int sr, sc, er, ec;
            if (parseRangeRef(ref, sr, sc, er, ec)) {
                sheet->mergeCells(CellRange(CellAddress(sr, sc), CellAddress(er, ec)));
            }
            continue;
        }

        // Parse XLSX conditional formatting: <conditionalFormatting sqref="A1:A10">
        if (xml.isStartElement() && xml.name() == u"conditionalFormatting") {
            QString sqref = xml.attributes().value("sqref").toString();
            CellRange cfRange(sqref);

            while (!xml.atEnd()) {
                xml.readNext();
                if (xml.isEndElement() && xml.name() == u"conditionalFormatting") break;
                if (!xml.isStartElement() || xml.name() != u"cfRule") continue;

                QString ruleType = xml.attributes().value("type").toString();

                if (ruleType == "dataBar") {
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, ConditionType::DataBar);
                    DataBarConfig cfg;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"color") {
                            QString rgb = xml.attributes().value("rgb").toString();
                            if (rgb.length() == 8) rgb = "#" + rgb.mid(2);
                            else if (!rgb.startsWith("#")) rgb = "#" + rgb;
                            cfg.barColor = QColor(rgb);
                        }
                        if (xml.isStartElement() && xml.name() == u"cfvo") {
                            QString cfvoType = xml.attributes().value("type").toString();
                            if (cfvoType == "min") cfg.autoRange = true;
                        }
                    }
                    rule->setDataBarConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else if (ruleType == "colorScale") {
                    QVector<QColor> colors;
                    int cfvoCount = 0;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"cfvo") cfvoCount++;
                        if (xml.isStartElement() && xml.name() == u"color") {
                            QString rgb = xml.attributes().value("rgb").toString();
                            if (rgb.length() == 8) rgb = "#" + rgb.mid(2);
                            else if (!rgb.startsWith("#")) rgb = "#" + rgb;
                            colors.append(QColor(rgb));
                        }
                    }
                    bool isThreeColor = (cfvoCount >= 3 && colors.size() >= 3);
                    auto condType = isThreeColor ? ConditionType::ColorScale3 : ConditionType::ColorScale2;
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, condType);
                    ColorScaleConfig cfg;
                    cfg.autoRange = true;
                    if (colors.size() >= 2) {
                        cfg.minColor = colors[0];
                        cfg.maxColor = colors[colors.size() - 1];
                    }
                    if (isThreeColor && colors.size() >= 3) {
                        cfg.midColor = colors[1];
                        cfg.threeColor = true;
                    }
                    rule->setColorScaleConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else if (ruleType == "iconSet") {
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, ConditionType::IconSet);
                    IconSetConfig cfg;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"iconSet") {
                            QString iconSetType = xml.attributes().value("iconSet").toString();
                            if (iconSetType.contains("3Arrows")) cfg.iconType = IconSetConfig::Arrows3;
                            else if (iconSetType.contains("3Flags")) cfg.iconType = IconSetConfig::Flags3;
                            else if (iconSetType.contains("3Stars")) cfg.iconType = IconSetConfig::Stars3;
                            else if (iconSetType.contains("3Symbols")) cfg.iconType = IconSetConfig::Checkmarks3;
                            else cfg.iconType = IconSetConfig::TrafficLights;
                            QString reverse = xml.attributes().value("reverse").toString();
                            cfg.reverseOrder = (reverse == "1");
                            QString showVal = xml.attributes().value("showValue").toString();
                            if (showVal == "0") cfg.showValue = false;
                        }
                    }
                    rule->setIconSetConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else {
                    xml.skipCurrentElement();
                }
            }
            continue;
        }

        if (xml.isStartElement() && xml.name() == u"c") {
            QStringView ref = xml.attributes().value("r");
            QString type = xml.attributes().value("t").toString();  // Must own string — QStringView invalidated by readNext()
            int styleIdx = xml.attributes().value("s").toInt();

            int row, col;
            if (!parseCellRef(ref, row, col)) continue;

            // Read child elements: <f> (formula), <v> (value), <is> (inline string)
            QString value;
            QString formula;
            QString inlineStr;
            bool hasInlineStr = false;
            int depth = 1;

            while (!xml.atEnd() && depth > 0) {
                xml.readNext();
                if (xml.isStartElement()) {
                    if (xml.name() == u"v") {
                        value = xml.readElementText();
                    } else if (xml.name() == u"f") {
                        formula = xml.readElementText();
                    } else if (xml.name() == u"is") {
                        hasInlineStr = true;
                    } else if (hasInlineStr && xml.name() == u"t") {
                        inlineStr += xml.readElementText();
                    } else {
                        depth++;
                    }
                } else if (xml.isEndElement()) {
                    if (xml.name() == u"c") {
                        depth = 0;
                    } else if (xml.name() == u"is") {
                        hasInlineStr = false;
                    }
                }
            }

            // Use fast path: getOrCreateCellFast returns CellProxy (no shared_ptr overhead)
            auto cell = CellProxy();
            bool cellSet = false;

            if (!formula.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                if (!formula.startsWith(u'=')) formula.prepend(u'=');
                cell->setFormula(formula);
                cellSet = true;
                if (!value.isEmpty() && type != u"s") {
                    bool ok;
                    double num = value.toDouble(&ok);
                    if (ok) cell->setComputedValue(num);
                }
            }
            else if (type == u"inlineStr" && !inlineStr.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(inlineStr); // known string — skip toDouble
                cellSet = true;
            }
            else if (type == u"s" && !value.isEmpty()) {
                int ssIdx = value.toInt();
                if (ssIdx >= 0 && ssIdx < sharedStrings.size()) {
                    cell = sheet->getOrCreateCellFast(row, col);
                    cell->setValueDirect(sharedStrings[ssIdx]); // known string — skip toDouble
                    cellSet = true;
                }
            }
            else if (type == u"b" && !value.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(value == u"1" ? QStringLiteral("TRUE") : QStringLiteral("FALSE"));
                cellSet = true;
            }
            else if (type == u"str" && !value.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(value); // known string — skip toDouble
                cellSet = true;
            }
            else if (!value.isEmpty()) {
                bool ok;
                double num = value.toDouble(&ok);
                if (ok) {
                    bool isDateFmt = false;
                    if (styleIdx > 0 && styleIdx < static_cast<int>(styles.size())) {
                        const auto& fmt = styles[styleIdx].numberFormat;
                        isDateFmt = (fmt == "Date" || fmt == "Time");
                    }
                    cell = sheet->getOrCreateCellFast(row, col);
                    if (isDateFmt && num > 0 && num < 2958466) {
                        QDate epoch(1899, 12, 30);
                        QDate date = epoch.addDays(static_cast<qint64>(num));
                        if (date.isValid()) {
                            cell->setValueDirect(date.toString("MM/dd/yyyy"));
                        } else {
                            cell->setValueDirect(num);
                        }
                    } else {
                        cell->setValueDirect(num);
                    }
                } else {
                    cell = sheet->getOrCreateCellFast(row, col);
                    cell->setValueDirect(value); // untyped fallback as string
                }
                cellSet = true;
            }

            // Apply style
            if ((cellSet || styleIdx > 0) && styleIdx < static_cast<int>(styles.size())) {
                if (!cell) cell = sheet->getOrCreateCellFast(row, col);
                CellStyle cellStyle = styles[styleIdx];
                if ((cellStyle.numberFormat == "Date" || cellStyle.numberFormat == "Time")) {
                    bool ok;
                    double num = value.toDouble(&ok);
                    if (ok && (num < 0 || num >= 2958466)) {
                        cellStyle.numberFormat = "General";
                    }
                }
                cell->setStyle(cellStyle);
            }
        }
    }

    sheet->finishBulkImport();
}

// ============================================================================
// Streaming sheet parser with progress callback
// ============================================================================
void XlsxService::parseSheetStreaming(const QByteArray& xmlData, const QStringList& sharedStrings,
                                       const std::vector<CellStyle>& styles, Spreadsheet* sheet,
                                       ImportProgressCallback progress, int sheetIndex) {
    QXmlStreamReader xml(xmlData);

    // Pre-reserve based on XML size heuristic
    size_t estimatedCells = static_cast<size_t>(xmlData.size()) / 100;
    if (estimatedCells > 4096) {
        sheet->reserveCells(estimatedCells);
    }

    int rowCount = 0;
    static constexpr int PROGRESS_INTERVAL = 50000; // Report every 50K rows

    while (!xml.atEnd()) {
        xml.readNext();

        // Parse sheet view
        if (xml.isStartElement() && xml.name() == u"sheetView") {
            QStringView showGrid = xml.attributes().value("showGridLines");
            if (showGrid == u"0") sheet->setShowGridlines(false);
            continue;
        }

        // Parse column widths
        if (xml.isStartElement() && xml.name() == u"col") {
            int minCol = xml.attributes().value("min").toInt() - 1;
            int maxCol = xml.attributes().value("max").toInt() - 1;
            double width = xml.attributes().value("width").toDouble();
            if (width > 0) {
                int pixelWidth = qMax(30, static_cast<int>(width * 7.5));
                int colLimit = qMin(maxCol, 16383);
                for (int c = minCol; c <= colLimit; ++c) {
                    sheet->setColumnWidth(c, pixelWidth);
                }
            }
            continue;
        }

        // Parse row heights
        if (xml.isStartElement() && xml.name() == u"row") {
            QString htStr = xml.attributes().value("ht").toString();
            if (!htStr.isEmpty()) {
                int rowIdx = xml.attributes().value("r").toInt() - 1;
                double ht = htStr.toDouble();
                if (ht > 0 && rowIdx >= 0) {
                    int pixelHeight = qMax(14, static_cast<int>(ht * 1.333));
                    sheet->setRowHeight(rowIdx, pixelHeight);
                }
            }
            // Track row count for progress
            rowCount++;
            if (progress && (rowCount % PROGRESS_INTERVAL == 0)) {
                progress(rowCount, sheetIndex);
            }
            continue;
        }

        // Parse merged cells
        if (xml.isStartElement() && xml.name() == u"mergeCell") {
            QStringView ref = xml.attributes().value("ref");
            int sr, sc, er, ec;
            if (parseRangeRef(ref, sr, sc, er, ec)) {
                sheet->mergeCells(CellRange(CellAddress(sr, sc), CellAddress(er, ec)));
            }
            continue;
        }

        // Parse conditional formatting (skip detail — same as non-streaming)
        if (xml.isStartElement() && xml.name() == u"conditionalFormatting") {
            QString sqref = xml.attributes().value("sqref").toString();
            CellRange cfRange(sqref);

            while (!xml.atEnd()) {
                xml.readNext();
                if (xml.isEndElement() && xml.name() == u"conditionalFormatting") break;
                if (!xml.isStartElement() || xml.name() != u"cfRule") continue;

                QString ruleType = xml.attributes().value("type").toString();

                if (ruleType == "dataBar") {
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, ConditionType::DataBar);
                    DataBarConfig cfg;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"color") {
                            QString rgb = xml.attributes().value("rgb").toString();
                            if (rgb.length() == 8) rgb = "#" + rgb.mid(2);
                            else if (!rgb.startsWith("#")) rgb = "#" + rgb;
                            cfg.barColor = QColor(rgb);
                        }
                        if (xml.isStartElement() && xml.name() == u"cfvo") {
                            QString cfvoType = xml.attributes().value("type").toString();
                            if (cfvoType == "min") cfg.autoRange = true;
                        }
                    }
                    rule->setDataBarConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else if (ruleType == "colorScale") {
                    QVector<QColor> colors;
                    int cfvoCount = 0;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"cfvo") cfvoCount++;
                        if (xml.isStartElement() && xml.name() == u"color") {
                            QString rgb = xml.attributes().value("rgb").toString();
                            if (rgb.length() == 8) rgb = "#" + rgb.mid(2);
                            else if (!rgb.startsWith("#")) rgb = "#" + rgb;
                            colors.append(QColor(rgb));
                        }
                    }
                    bool isThreeColor = (cfvoCount >= 3 && colors.size() >= 3);
                    auto condType = isThreeColor ? ConditionType::ColorScale3 : ConditionType::ColorScale2;
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, condType);
                    ColorScaleConfig cfg;
                    cfg.autoRange = true;
                    if (colors.size() >= 2) {
                        cfg.minColor = colors[0];
                        cfg.maxColor = colors[colors.size() - 1];
                    }
                    if (isThreeColor && colors.size() >= 3) {
                        cfg.midColor = colors[1];
                        cfg.threeColor = true;
                    }
                    rule->setColorScaleConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else if (ruleType == "iconSet") {
                    auto rule = std::make_shared<ConditionalFormat>(cfRange, ConditionType::IconSet);
                    IconSetConfig cfg;
                    while (!xml.atEnd()) {
                        xml.readNext();
                        if (xml.isEndElement() && xml.name() == u"cfRule") break;
                        if (xml.isStartElement() && xml.name() == u"iconSet") {
                            QString iconSetType = xml.attributes().value("iconSet").toString();
                            if (iconSetType.contains("3Arrows")) cfg.iconType = IconSetConfig::Arrows3;
                            else if (iconSetType.contains("3Flags")) cfg.iconType = IconSetConfig::Flags3;
                            else if (iconSetType.contains("3Stars")) cfg.iconType = IconSetConfig::Stars3;
                            else if (iconSetType.contains("3Symbols")) cfg.iconType = IconSetConfig::Checkmarks3;
                            else cfg.iconType = IconSetConfig::TrafficLights;
                            QString reverse = xml.attributes().value("reverse").toString();
                            cfg.reverseOrder = (reverse == "1");
                            QString showVal = xml.attributes().value("showValue").toString();
                            if (showVal == "0") cfg.showValue = false;
                        }
                    }
                    rule->setIconSetConfig(cfg);
                    sheet->getConditionalFormatting().addRule(rule);
                } else {
                    xml.skipCurrentElement();
                }
            }
            continue;
        }

        // Parse cell data
        if (xml.isStartElement() && xml.name() == u"c") {
            QStringView ref = xml.attributes().value("r");
            QString type = xml.attributes().value("t").toString();
            int styleIdx = xml.attributes().value("s").toInt();

            int row, col;
            if (!parseCellRef(ref, row, col)) continue;

            QString value;
            QString formula;
            QString inlineStr;
            bool hasInlineStr = false;
            int depth = 1;

            while (!xml.atEnd() && depth > 0) {
                xml.readNext();
                if (xml.isStartElement()) {
                    if (xml.name() == u"v") {
                        value = xml.readElementText();
                    } else if (xml.name() == u"f") {
                        formula = xml.readElementText();
                    } else if (xml.name() == u"is") {
                        hasInlineStr = true;
                    } else if (hasInlineStr && xml.name() == u"t") {
                        inlineStr += xml.readElementText();
                    } else {
                        depth++;
                    }
                } else if (xml.isEndElement()) {
                    if (xml.name() == u"c") {
                        depth = 0;
                    } else if (xml.name() == u"is") {
                        hasInlineStr = false;
                    }
                }
            }

            auto cell = CellProxy();
            bool cellSet = false;

            if (!formula.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                if (!formula.startsWith(u'=')) formula.prepend(u'=');
                cell->setFormula(formula);
                cellSet = true;
                if (!value.isEmpty() && type != u"s") {
                    bool ok;
                    double num = value.toDouble(&ok);
                    if (ok) cell->setComputedValue(num);
                }
            }
            else if (type == u"inlineStr" && !inlineStr.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(inlineStr);
                cellSet = true;
            }
            else if (type == u"s" && !value.isEmpty()) {
                int ssIdx = value.toInt();
                if (ssIdx >= 0 && ssIdx < sharedStrings.size()) {
                    cell = sheet->getOrCreateCellFast(row, col);
                    cell->setValueDirect(sharedStrings[ssIdx]);
                    cellSet = true;
                }
            }
            else if (type == u"b" && !value.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(value == u"1" ? QStringLiteral("TRUE") : QStringLiteral("FALSE"));
                cellSet = true;
            }
            else if (type == u"str" && !value.isEmpty()) {
                cell = sheet->getOrCreateCellFast(row, col);
                cell->setValueDirect(value);
                cellSet = true;
            }
            else if (!value.isEmpty()) {
                bool ok;
                double num = value.toDouble(&ok);
                if (ok) {
                    bool isDateFmt = false;
                    if (styleIdx > 0 && styleIdx < static_cast<int>(styles.size())) {
                        const auto& fmt = styles[styleIdx].numberFormat;
                        isDateFmt = (fmt == "Date" || fmt == "Time");
                    }
                    cell = sheet->getOrCreateCellFast(row, col);
                    if (isDateFmt && num > 0 && num < 2958466) {
                        QDate epoch(1899, 12, 30);
                        QDate date = epoch.addDays(static_cast<qint64>(num));
                        if (date.isValid()) {
                            cell->setValueDirect(date.toString("MM/dd/yyyy"));
                        } else {
                            cell->setValueDirect(num);
                        }
                    } else {
                        cell->setValueDirect(num);
                    }
                } else {
                    cell = sheet->getOrCreateCellFast(row, col);
                    cell->setValueDirect(value);
                }
                cellSet = true;
            }

            // Apply style
            if ((cellSet || styleIdx > 0) && styleIdx < static_cast<int>(styles.size())) {
                if (!cell) cell = sheet->getOrCreateCellFast(row, col);
                CellStyle cellStyle = styles[styleIdx];
                if ((cellStyle.numberFormat == "Date" || cellStyle.numberFormat == "Time")) {
                    bool ok;
                    double num = value.toDouble(&ok);
                    if (ok && (num < 0 || num >= 2958466)) {
                        cellStyle.numberFormat = "General";
                    }
                }
                cell->setStyle(cellStyle);
            }
        }
    }

    // Final progress report
    if (progress) {
        progress(rowCount, sheetIndex);
    }

    sheet->finishBulkImport();
}

// ============================================================================
// Streaming import entry point
// ============================================================================
XlsxImportResult XlsxService::importFromFileStreaming(const QString& filePath,
                                                       ImportProgressCallback progress) {
    XlsxImportResult result;

    QZipReader zip(filePath);
    if (!zip.isReadable()) return result;

    // Read shared strings
    QStringList sharedStrings;
    QByteArray ssData = zip.fileData("xl/sharedStrings.xml");
    if (!ssData.isEmpty()) {
        sharedStrings = parseSharedStrings(ssData);
    }

    // Parse styles
    std::vector<CellStyle> styles;
    QByteArray stylesData = zip.fileData("xl/styles.xml");
    if (!stylesData.isEmpty()) {
        auto fonts = parseFonts(stylesData);
        auto fills = parseFills(stylesData);
        auto borders = parseBorders(stylesData);
        auto cellXfs = parseCellXfs(stylesData);
        auto customNumFmts = parseNumFmts(stylesData);
        for (const auto& xf : cellXfs) {
            styles.push_back(buildCellStyle(xf, fonts, fills, borders, xf.numFmtId, customNumFmts));
        }
    }

    // Parse workbook
    QByteArray workbookData = zip.fileData("xl/workbook.xml");
    QByteArray relsData = zip.fileData("xl/_rels/workbook.xml.rels");
    auto sheetInfos = parseWorkbook(workbookData, relsData);

    if (sheetInfos.empty()) {
        SheetInfo si;
        si.name = "Sheet1";
        si.filePath = "worksheets/sheet1.xml";
        sheetInfos.push_back(si);
    }

    // Parse each sheet with streaming progress
    for (int sheetIdx = 0; sheetIdx < static_cast<int>(sheetInfos.size()); ++sheetIdx) {
        const auto& info = sheetInfos[sheetIdx];
        QString path = "xl/" + info.filePath;
        QByteArray sheetData = zip.fileData(path);
        if (sheetData.isEmpty()) continue;

        auto spreadsheet = std::make_shared<Spreadsheet>();
        spreadsheet->setAutoRecalculate(false);
        spreadsheet->setSheetName(info.name);

        // Use streaming parser with progress callback
        parseSheetStreaming(sheetData, sharedStrings, styles, spreadsheet.get(), progress, sheetIdx);

        // Parse sheet rels for charts (same as non-streaming)
        QString fullSheetPath = "xl/" + info.filePath;
        int lastSlash = fullSheetPath.lastIndexOf('/');
        QString sheetDirPath = fullSheetPath.left(lastSlash);
        QString sheetFileName = fullSheetPath.mid(lastSlash + 1);
        QString sheetRelsPath = sheetDirPath + "/_rels/" + sheetFileName + ".rels";
        QByteArray sheetRelsData = zip.fileData(sheetRelsPath);
        auto sheetRels = parseRels(sheetRelsData);

        // Chart import
        QByteArray sheetXmlForDrawing = sheetData;
        QString drawingRId = findDrawingRId(sheetXmlForDrawing);
        if (!drawingRId.isEmpty() && sheetRels.count(drawingRId)) {
            QString drawingPath = resolveRelativePath(fullSheetPath, sheetRels[drawingRId]);
            QByteArray drawingData = zip.fileData(drawingPath);
            if (!drawingData.isEmpty()) {
                auto chartRefs = parseDrawing(drawingData);
                int dLastSlash = drawingPath.lastIndexOf('/');
                QString drawingDir = drawingPath.left(dLastSlash);
                QString drawingFile = drawingPath.mid(dLastSlash + 1);
                QString drawingRelsPath = drawingDir + "/_rels/" + drawingFile + ".rels";
                QByteArray drawingRelsData = zip.fileData(drawingRelsPath);
                auto drawingRels = parseRels(drawingRelsData);

                for (const auto& cref : chartRefs) {
                    if (drawingRels.count(cref.chartRId)) {
                        QString chartPath = resolveRelativePath(drawingPath, drawingRels[cref.chartRId]);
                        QByteArray chartData = zip.fileData(chartPath);
                        if (!chartData.isEmpty()) {
                            ImportedChart chart = parseChartXml(chartData);
                            chart.sheetIndex = sheetIdx;
                            chart.x = cref.fromCol * 70;
                            chart.y = cref.fromRow * 25;
                            chart.width = (cref.toCol - cref.fromCol) * 70;
                            chart.height = (cref.toRow - cref.fromRow) * 25;
                            if (chart.width < 200) chart.width = 420;
                            if (chart.height < 100) chart.height = 320;
                            result.charts.push_back(chart);
                        }
                    }
                }
            }
        }

        result.sheets.push_back(spreadsheet);
    }

    // Check for Nexel-native custom JSON
    QByteArray customJson = zip.fileData("xl/nexel/custom.json");
    if (!customJson.isEmpty()) {
        QJsonDocument doc = QJsonDocument::fromJson(customJson);
        if (doc.isObject()) {
            QJsonObject root = doc.object();
            QJsonArray chartsArr = root["charts"].toArray();
            for (const auto& cv : chartsArr) {
                QJsonObject co = cv.toObject();
                int si = co["sheetIndex"].toInt();
                if (si < 0 || si >= static_cast<int>(result.charts.size())) continue;
                auto& chart = result.charts[si];
                chart.dataRange = co["dataRange"].toString();
                chart.themeIndex = co["themeIndex"].toInt();
                chart.showLegend = co["showLegend"].toBool(true);
                chart.showGridLines = co["showGridLines"].toBool(true);
                chart.isNexelNative = true;
            }
        }
    }

    return result;
}

int XlsxService::columnLetterToIndex(const QString& letters) {
    int result = 0;
    for (int i = 0; i < letters.length(); ++i) {
        result = result * 26 + (letters[i].unicode() - 'A' + 1);
    }
    return result - 1;
}

// ============== XLSX EXPORT ==============

QString XlsxService::columnIndexToLetter(int col) {
    QString result;
    col++; // 1-based
    while (col > 0) {
        col--;
        result.prepend(QChar('A' + (col % 26)));
        col /= 26;
    }
    return result;
}

static QString borderSideKey(const BorderStyle& b) {
    if (!b.enabled) return "0";
    return QString("1_%1_%2").arg(b.color, QString::number(b.width));
}

QString XlsxService::cellStyleKey(const CellStyle& style) {
    return QString("%1|%2|%3|%4|%5|%6|%7|%8|%9|%10|%11|%12|%13|%14|%15")
        .arg(style.fontName).arg(style.fontSize)
        .arg(style.bold).arg(style.italic).arg(style.underline).arg(style.strikethrough)
        .arg(style.foregroundColor).arg(style.backgroundColor)
        .arg(static_cast<int>(style.hAlign)).arg(static_cast<int>(style.vAlign))
        .arg(style.numberFormat)
        .arg(borderSideKey(style.borderLeft))
        .arg(borderSideKey(style.borderRight))
        .arg(borderSideKey(style.borderTop))
        .arg(borderSideKey(style.borderBottom));
}

bool XlsxService::exportToFile(const std::vector<std::shared_ptr<Spreadsheet>>& sheets,
                                 const QString& filePath,
                                 const std::vector<NexelChartExport>& charts) {
    if (sheets.empty()) return false;

    // Collect unique styles and build style index map
    std::map<QString, int> styleIndexMap;
    // Index 0 is reserved for default style
    CellStyle defaultStyle;
    styleIndexMap[cellStyleKey(defaultStyle)] = 0;

    for (const auto& sheet : sheets) {
        sheet->forEachCell([&](int, int, const Cell& cell) {
            const CellStyle& s = cell.getStyle();
            QString key = cellStyleKey(s);
            if (styleIndexMap.find(key) == styleIndexMap.end()) {
                int idx = static_cast<int>(styleIndexMap.size());
                styleIndexMap[key] = idx;
            }
        });
    }

    // Collect shared strings from all sheets
    QStringList sharedStrings;
    std::map<QString, int> ssMap; // string -> index
    std::vector<QByteArray> sheetXmls;

    // Per-sheet hyperlink data for rels generation
    struct HyperlinkExport { QString url; int rId; };
    std::vector<std::vector<HyperlinkExport>> allSheetHyperlinks;

    for (const auto& sheet : sheets) {
        sheet->forEachCell([&](int, int, const Cell& cell) {
            if (cell.getType() == CellType::Text || cell.getType() == CellType::Empty) {
                QString val = cell.getValue().toString();
                if (!val.isEmpty() && !val.startsWith('=')) {
                    if (ssMap.find(val) == ssMap.end()) {
                        ssMap[val] = sharedStrings.size();
                        sharedStrings.append(val);
                    }
                }
            }
        });
    }

    // Pre-compute which sheets have charts (for drawing references)
    std::map<int, std::vector<int>> chartsPerSheet; // sheetIndex -> chart indices
    for (int i = 0; i < static_cast<int>(charts.size()); ++i) {
        if (!charts[i].dataRange.isEmpty()) {
            chartsPerSheet[charts[i].sheetIndex].push_back(i);
        }
    }

    // Second pass: generate sheets
    int sheetIdx = 0;
    for (const auto& sheet : sheets) {
        QByteArray sheetXml;
        QBuffer buf(&sheetXml);
        buf.open(QIODevice::WriteOnly);
        QXmlStreamWriter xml(&buf);
        xml.setAutoFormatting(true);

        xml.writeStartDocument();
        xml.writeStartElement("worksheet");
        xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        xml.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Dimension element (required by Excel for validation)
        {
            int dimMaxRow = sheet->getMaxRow();
            int dimMaxCol = sheet->getMaxColumn();
            if (dimMaxRow < 0) dimMaxRow = 0;
            if (dimMaxCol < 0) dimMaxCol = 0;
            QString dimRef = "A1:" + columnIndexToLetter(dimMaxCol) + QString::number(dimMaxRow + 1);
            xml.writeStartElement("dimension");
            xml.writeAttribute("ref", dimRef);
            xml.writeEndElement();
        }

        // Sheet views (gridline visibility)
        xml.writeStartElement("sheetViews");
        xml.writeStartElement("sheetView");
        xml.writeAttribute("tabSelected", sheetIdx == 0 ? "1" : "0");
        xml.writeAttribute("workbookViewId", "0");
        if (!sheet->showGridlines()) {
            xml.writeAttribute("showGridLines", "0");
        }
        xml.writeEndElement(); // sheetView
        xml.writeEndElement(); // sheetViews

        // Default row/col dimensions (improves Excel compatibility)
        xml.writeStartElement("sheetFormatPr");
        xml.writeAttribute("defaultRowHeight", "15");
        xml.writeAttribute("defaultColWidth", "10.71");
        xml.writeEndElement();

        // Column widths
        const auto& colWidths = sheet->getColumnWidths();
        if (!colWidths.empty()) {
            xml.writeStartElement("cols");
            for (const auto& [col, widthPx] : colWidths) {
                xml.writeStartElement("col");
                xml.writeAttribute("min", QString::number(col + 1));
                xml.writeAttribute("max", QString::number(col + 1));
                xml.writeAttribute("width", QString::number(widthPx / 7.5, 'f', 2));
                xml.writeAttribute("customWidth", "1");
                xml.writeEndElement();
            }
            xml.writeEndElement(); // cols
        }

        // sheetData
        xml.writeStartElement("sheetData");

        int maxRow = sheet->getMaxRow();
        int maxCol = sheet->getMaxColumn();

        // Collect rows that need custom heights
        const auto& rowHeights = sheet->getRowHeights();

        // Determine which rows to write (those with data OR custom height)
        int maxRowForHeights = maxRow;
        if (!rowHeights.empty()) {
            maxRowForHeights = std::max(maxRow, rowHeights.rbegin()->first);
        }

        for (int r = 0; r <= maxRowForHeights; ++r) {
            bool hasData = false;
            if (r <= maxRow) {
                for (int c = 0; c <= maxCol; ++c) {
                    auto val = sheet->getCellValue(CellAddress(r, c));
                    if (val.isValid() && !val.toString().isEmpty()) { hasData = true; break; }
                    auto cell = sheet->getCellIfExists(r, c);
                    if (cell) {
                        QString key = cellStyleKey(cell->getStyle());
                        if (styleIndexMap.count(key) && styleIndexMap[key] != 0) { hasData = true; break; }
                    }
                }
            }

            bool hasCustomHeight = rowHeights.count(r) > 0;
            if (!hasData && !hasCustomHeight) continue;

            xml.writeStartElement("row");
            xml.writeAttribute("r", QString::number(r + 1));

            if (hasCustomHeight) {
                double htPoints = rowHeights.at(r) / 1.333;
                xml.writeAttribute("ht", QString::number(htPoints, 'f', 2));
                xml.writeAttribute("customHeight", "1");
            }

            for (int c = 0; c <= maxCol; ++c) {
                auto cell = sheet->getCellIfExists(r, c);
                if (!cell) continue;

                QVariant val = cell->getValue();
                QString formula = cell->getFormula();
                bool hasValue = val.isValid() && !val.toString().isEmpty();
                bool hasFormula = !formula.isEmpty();

                // Get style index
                QString key = cellStyleKey(cell->getStyle());
                int styleIdx = 0;
                if (styleIndexMap.count(key)) styleIdx = styleIndexMap[key];

                if (!hasValue && !hasFormula && styleIdx == 0) continue;

                QString cellRef = columnIndexToLetter(c) + QString::number(r + 1);
                xml.writeStartElement("c");
                xml.writeAttribute("r", cellRef);

                if (styleIdx > 0) {
                    xml.writeAttribute("s", QString::number(styleIdx));
                }

                if (hasFormula) {
                    // Determine if cached value is string (must set t= before child elements)
                    if (hasValue) {
                        bool ok;
                        double d = val.toDouble(&ok);
                        if (!ok || std::isnan(d) || std::isinf(d)) {
                            xml.writeAttribute("t", "str");
                        }
                    }
                    xml.writeTextElement("f", formula.startsWith('=') ? formula.mid(1) : formula);
                    // Write cached value
                    if (hasValue) {
                        bool ok;
                        double num = val.toDouble(&ok);
                        if (ok && !std::isnan(num) && !std::isinf(num)) {
                            xml.writeTextElement("v", QString::number(num, 'g', 15));
                        } else {
                            xml.writeTextElement("v", val.toString());
                        }
                    }
                } else if (hasValue) {
                    bool ok;
                    double num = val.toDouble(&ok);
                    if (ok && !std::isnan(num) && !std::isinf(num) && cell->getType() == CellType::Number) {
                        xml.writeTextElement("v", QString::number(num, 'g', 15));
                    } else {
                        // Shared string reference
                        QString text = val.toString();
                        if (ssMap.count(text)) {
                            xml.writeAttribute("t", "s");
                            xml.writeTextElement("v", QString::number(ssMap[text]));
                        } else {
                            xml.writeAttribute("t", "inlineStr");
                            xml.writeStartElement("is");
                            xml.writeTextElement("t", text);
                            xml.writeEndElement(); // is
                        }
                    }
                }

                xml.writeEndElement(); // c
            }

            xml.writeEndElement(); // row
        }

        xml.writeEndElement(); // sheetData

        // Merge cells
        const auto& mergedRegions = sheet->getMergedRegions();
        if (!mergedRegions.empty()) {
            xml.writeStartElement("mergeCells");
            xml.writeAttribute("count", QString::number(mergedRegions.size()));
            for (const auto& mr : mergedRegions) {
                xml.writeStartElement("mergeCell");
                QString ref = columnIndexToLetter(mr.range.getStart().col) +
                              QString::number(mr.range.getStart().row + 1) + ":" +
                              columnIndexToLetter(mr.range.getEnd().col) +
                              QString::number(mr.range.getEnd().row + 1);
                xml.writeAttribute("ref", ref);
                xml.writeEndElement();
            }
            xml.writeEndElement(); // mergeCells
        }

        // Write XLSX conditional formatting for visual types
        const auto& allCfRules = sheet->getConditionalFormatting().getAllRules();
        if (!allCfRules.empty()) {
            for (const auto& cfRule : allCfRules) {
                if (!cfRule->isVisualType()) continue;

                QString sqref = columnIndexToLetter(cfRule->getRange().getStart().col) +
                                QString::number(cfRule->getRange().getStart().row + 1) + ":" +
                                columnIndexToLetter(cfRule->getRange().getEnd().col) +
                                QString::number(cfRule->getRange().getEnd().row + 1);

                xml.writeStartElement("conditionalFormatting");
                xml.writeAttribute("sqref", sqref);

                if (cfRule->getType() == ConditionType::DataBar) {
                    xml.writeStartElement("cfRule");
                    xml.writeAttribute("type", "dataBar");
                    xml.writeAttribute("priority", "1");
                    xml.writeStartElement("dataBar");
                    xml.writeStartElement("cfvo");
                    xml.writeAttribute("type", cfRule->getDataBarConfig().autoRange ? "min" : "num");
                    if (!cfRule->getDataBarConfig().autoRange)
                        xml.writeAttribute("val", QString::number(cfRule->getDataBarConfig().minValue));
                    xml.writeEndElement();
                    xml.writeStartElement("cfvo");
                    xml.writeAttribute("type", cfRule->getDataBarConfig().autoRange ? "max" : "num");
                    if (!cfRule->getDataBarConfig().autoRange)
                        xml.writeAttribute("val", QString::number(cfRule->getDataBarConfig().maxValue));
                    xml.writeEndElement();
                    xml.writeStartElement("color");
                    QString rgb = cfRule->getDataBarConfig().barColor.name().mid(1);
                    xml.writeAttribute("rgb", "FF" + rgb.toUpper());
                    xml.writeEndElement();
                    xml.writeEndElement(); // dataBar
                    xml.writeEndElement(); // cfRule
                } else if (cfRule->getType() == ConditionType::ColorScale2 || cfRule->getType() == ConditionType::ColorScale3) {
                    xml.writeStartElement("cfRule");
                    xml.writeAttribute("type", "colorScale");
                    xml.writeAttribute("priority", "1");
                    xml.writeStartElement("colorScale");
                    const auto& csCfg = cfRule->getColorScaleConfig();
                    // cfvo elements
                    xml.writeStartElement("cfvo"); xml.writeAttribute("type", "min"); xml.writeEndElement();
                    if (csCfg.threeColor) {
                        xml.writeStartElement("cfvo"); xml.writeAttribute("type", "percentile"); xml.writeAttribute("val", "50"); xml.writeEndElement();
                    }
                    xml.writeStartElement("cfvo"); xml.writeAttribute("type", "max"); xml.writeEndElement();
                    // color elements
                    auto writeColor = [&xml](const QColor& c) {
                        xml.writeStartElement("color");
                        xml.writeAttribute("rgb", "FF" + c.name().mid(1).toUpper());
                        xml.writeEndElement();
                    };
                    writeColor(csCfg.minColor);
                    if (csCfg.threeColor) writeColor(csCfg.midColor);
                    writeColor(csCfg.maxColor);
                    xml.writeEndElement(); // colorScale
                    xml.writeEndElement(); // cfRule
                } else if (cfRule->getType() == ConditionType::IconSet) {
                    xml.writeStartElement("cfRule");
                    xml.writeAttribute("type", "iconSet");
                    xml.writeAttribute("priority", "1");
                    xml.writeStartElement("iconSet");
                    const auto& isCfg = cfRule->getIconSetConfig();
                    QString iconTypeName;
                    switch (isCfg.iconType) {
                        case IconSetConfig::TrafficLights: iconTypeName = "3TrafficLights1"; break;
                        case IconSetConfig::Arrows3: iconTypeName = "3Arrows"; break;
                        case IconSetConfig::Flags3: iconTypeName = "3Flags"; break;
                        case IconSetConfig::Stars3: iconTypeName = "3Stars"; break;
                        case IconSetConfig::Checkmarks3: iconTypeName = "3Symbols"; break;
                    }
                    xml.writeAttribute("iconSet", iconTypeName);
                    if (isCfg.reverseOrder) xml.writeAttribute("reverse", "1");
                    if (!isCfg.showValue) xml.writeAttribute("showValue", "0");
                    // cfvo elements
                    xml.writeStartElement("cfvo"); xml.writeAttribute("type", "percent"); xml.writeAttribute("val", "0"); xml.writeEndElement();
                    xml.writeStartElement("cfvo"); xml.writeAttribute("type", "percent"); xml.writeAttribute("val", QString::number(isCfg.threshold1)); xml.writeEndElement();
                    xml.writeStartElement("cfvo"); xml.writeAttribute("type", "percent"); xml.writeAttribute("val", QString::number(isCfg.threshold2)); xml.writeEndElement();
                    xml.writeEndElement(); // iconSet
                    xml.writeEndElement(); // cfRule
                }

                xml.writeEndElement(); // conditionalFormatting
            }
        }

        // Write data validation rules
        const auto& validationRules = sheet->getValidationRules();
        if (!validationRules.empty()) {
            xml.writeStartElement("dataValidations");
            xml.writeAttribute("count", QString::number(validationRules.size()));
            for (const auto& rule : validationRules) {
                xml.writeStartElement("dataValidation");

                // Map validation type to OOXML type string
                QString typeStr;
                switch (rule.type) {
                    case Spreadsheet::DataValidationRule::WholeNumber: typeStr = "whole"; break;
                    case Spreadsheet::DataValidationRule::Decimal:     typeStr = "decimal"; break;
                    case Spreadsheet::DataValidationRule::List:        typeStr = "list"; break;
                    case Spreadsheet::DataValidationRule::TextLength:  typeStr = "textLength"; break;
                    case Spreadsheet::DataValidationRule::Date:        typeStr = "date"; break;
                    case Spreadsheet::DataValidationRule::Custom:      typeStr = "custom"; break;
                }
                xml.writeAttribute("type", typeStr);

                // Map operator to OOXML operator string (not used for List/Custom)
                if (rule.type != Spreadsheet::DataValidationRule::List &&
                    rule.type != Spreadsheet::DataValidationRule::Custom) {
                    QString opStr;
                    switch (rule.op) {
                        case Spreadsheet::DataValidationRule::Between:              opStr = "between"; break;
                        case Spreadsheet::DataValidationRule::NotBetween:           opStr = "notBetween"; break;
                        case Spreadsheet::DataValidationRule::EqualTo:              opStr = "equal"; break;
                        case Spreadsheet::DataValidationRule::NotEqualTo:           opStr = "notEqual"; break;
                        case Spreadsheet::DataValidationRule::GreaterThan:          opStr = "greaterThan"; break;
                        case Spreadsheet::DataValidationRule::LessThan:             opStr = "lessThan"; break;
                        case Spreadsheet::DataValidationRule::GreaterThanOrEqual:   opStr = "greaterThanOrEqual"; break;
                        case Spreadsheet::DataValidationRule::LessThanOrEqual:      opStr = "lessThanOrEqual"; break;
                    }
                    xml.writeAttribute("operator", opStr);
                }

                // allowBlank (for list type, use listIgnoreBlanks; others default to 1)
                if (rule.type == Spreadsheet::DataValidationRule::List) {
                    xml.writeAttribute("allowBlank", rule.listIgnoreBlanks ? "1" : "0");
                } else {
                    xml.writeAttribute("allowBlank", "1");
                }

                if (rule.showInputMessage) {
                    xml.writeAttribute("showInputMessage", "1");
                }
                if (rule.showErrorAlert) {
                    xml.writeAttribute("showErrorMessage", "1");
                }

                // Error style
                if (rule.errorStyle == Spreadsheet::DataValidationRule::Warning) {
                    xml.writeAttribute("errorStyle", "warning");
                } else if (rule.errorStyle == Spreadsheet::DataValidationRule::Information) {
                    xml.writeAttribute("errorStyle", "information");
                }
                // "stop" is the default, no need to write it

                // Input/error messages
                if (!rule.errorTitle.isEmpty()) {
                    xml.writeAttribute("errorTitle", rule.errorTitle);
                }
                if (!rule.errorMessage.isEmpty()) {
                    xml.writeAttribute("error", rule.errorMessage);
                }
                if (!rule.inputTitle.isEmpty()) {
                    xml.writeAttribute("promptTitle", rule.inputTitle);
                }
                if (!rule.inputMessage.isEmpty()) {
                    xml.writeAttribute("prompt", rule.inputMessage);
                }

                // sqref — cell range this validation applies to
                QString sqref = columnIndexToLetter(rule.range.getStart().col) +
                                QString::number(rule.range.getStart().row + 1) + ":" +
                                columnIndexToLetter(rule.range.getEnd().col) +
                                QString::number(rule.range.getEnd().row + 1);
                xml.writeAttribute("sqref", sqref);

                // formula1
                if (rule.type == Spreadsheet::DataValidationRule::List) {
                    if (!rule.listSourceRange.isEmpty()) {
                        // Range-based list (e.g. "Sheet2!A1:A10")
                        xml.writeTextElement("formula1", rule.listSourceRange);
                    } else if (!rule.listItems.isEmpty()) {
                        // Inline comma-separated list in quotes
                        xml.writeTextElement("formula1", "\"" + rule.listItems.join(",") + "\"");
                    }
                } else if (rule.type == Spreadsheet::DataValidationRule::Custom) {
                    if (!rule.customFormula.isEmpty()) {
                        QString f = rule.customFormula;
                        if (f.startsWith('=')) f = f.mid(1);
                        xml.writeTextElement("formula1", f);
                    }
                } else {
                    // Numeric/date/text-length validation
                    if (!rule.value1.isEmpty()) {
                        xml.writeTextElement("formula1", rule.value1);
                    }
                    if (!rule.value2.isEmpty()) {
                        xml.writeTextElement("formula2", rule.value2);
                    }
                }

                xml.writeEndElement(); // dataValidation
            }
            xml.writeEndElement(); // dataValidations
        }

        // Collect hyperlinks for this sheet
        struct HyperlinkEntry { int row; int col; QString url; };
        std::vector<HyperlinkEntry> sheetHyperlinks;
        sheet->forEachCell([&](int r, int c, const Cell& cell) {
            if (cell.hasHyperlink()) {
                sheetHyperlinks.push_back({r, c, cell.getHyperlink()});
            }
        });

        // rId numbering: rId1 may be used by drawing, hyperlinks start after
        int nextRId = chartsPerSheet.count(sheetIdx) ? 2 : 1;

        // Write hyperlinks element
        if (!sheetHyperlinks.empty()) {
            xml.writeStartElement("hyperlinks");
            for (size_t hi = 0; hi < sheetHyperlinks.size(); ++hi) {
                const auto& hl = sheetHyperlinks[hi];
                QString cellRef = columnIndexToLetter(hl.col) + QString::number(hl.row + 1);
                xml.writeStartElement("hyperlink");
                xml.writeAttribute("ref", cellRef);
                xml.writeAttribute("r:id", QString("rId%1").arg(nextRId + static_cast<int>(hi)));
                xml.writeEndElement();
            }
            xml.writeEndElement(); // hyperlinks
        }

        // Drawing reference (for sheets with charts)
        if (chartsPerSheet.count(sheetIdx)) {
            xml.writeStartElement("drawing");
            xml.writeAttribute("r:id", "rId1");
            xml.writeEndElement();
        }

        xml.writeEndElement(); // worksheet
        xml.writeEndDocument();
        buf.close();
        sheetXmls.push_back(sheetXml);

        // Store hyperlink rels data for this sheet
        std::vector<HyperlinkExport> hlExports;
        for (size_t hi = 0; hi < sheetHyperlinks.size(); ++hi) {
            hlExports.push_back({sheetHyperlinks[hi].url, nextRId + static_cast<int>(hi)});
        }
        allSheetHyperlinks.push_back(std::move(hlExports));

        sheetIdx++;
    }

    // Build the ZIP
    int sheetCount = static_cast<int>(sheets.size());

    // Compute chart/drawing info for content types
    int totalChartCount = 0;
    std::vector<int> drawingSheetNums; // 1-based sheet numbers with drawings
    for (auto& [si, indices] : chartsPerSheet) {
        if (si < sheetCount) {
            totalChartCount += static_cast<int>(indices.size());
            drawingSheetNums.push_back(si + 1);
        }
    }

    // Determine if we'll have custom JSON parts (for content type registration)
    bool hasCustomJson = !charts.empty(); // nexel-charts.json
    if (!hasCustomJson) {
        for (const auto& sheet : sheets) {
            if (!sheet->getValidationRules().empty() || !sheet->getTables().empty()) {
                hasCustomJson = true;
                break;
            }
            bool hasCheckbox = false;
            sheet->forEachCell([&](int, int, const Cell& cell) {
                if (cell.getStyle().numberFormat == "Checkbox") hasCheckbox = true;
            });
            if (hasCheckbox) { hasCustomJson = true; break; }
        }
    }

    QZipWriter zip(filePath);
    if (zip.status() != QZipWriter::NoError) return false;

    zip.addFile("[Content_Types].xml",
                generateContentTypes(sheetCount, !sharedStrings.isEmpty(),
                                     totalChartCount, drawingSheetNums, hasCustomJson));
    zip.addFile("_rels/.rels", generateRels());
    zip.addFile("xl/workbook.xml", generateWorkbook(sheets));
    zip.addFile("xl/_rels/workbook.xml.rels", generateWorkbookRels(sheetCount, !sharedStrings.isEmpty()));
    zip.addFile("xl/styles.xml", generateStyles(sheets, styleIndexMap));

    if (!sharedStrings.isEmpty()) {
        zip.addFile("xl/sharedStrings.xml", generateSharedStrings(sharedStrings));
    }

    for (int i = 0; i < sheetCount; ++i) {
        zip.addFile(QString("xl/worksheets/sheet%1.xml").arg(i + 1), sheetXmls[i]);
    }

    // Generate OOXML chart parts (so Excel can display them)
    int globalChartNum = 1;
    for (auto& [si, indices] : chartsPerSheet) {
        if (si >= sheetCount) continue;
        QString sheetName = sheets[si]->getSheetName();
        int startNum = globalChartNum;

        // Generate chart XML files
        for (int ci : indices) {
            zip.addFile(QString("xl/charts/chart%1.xml").arg(globalChartNum),
                        generateChartXml(charts[ci], sheetName));
            globalChartNum++;
        }

        int chartCountForSheet = static_cast<int>(indices.size());

        // Drawing XML
        zip.addFile(QString("xl/drawings/drawing%1.xml").arg(si + 1),
                    generateDrawingXml(charts, indices));

        // Drawing relationships (maps rIdN -> chartN.xml)
        zip.addFile(QString("xl/drawings/_rels/drawing%1.xml.rels").arg(si + 1),
                    generateDrawingRels(chartCountForSheet, startNum));

        // Sheet relationships (maps rId1 -> drawingN.xml, plus hyperlink rels)
        {
            QString relsXml;
            relsXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            relsXml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n";
            relsXml += QString("<Relationship Id=\"rId1\" "
                               "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
                               "Target=\"../drawings/drawing%1.xml\"/>\n").arg(si + 1);
            if (si < static_cast<int>(allSheetHyperlinks.size())) {
                for (const auto& hl : allSheetHyperlinks[si]) {
                    relsXml += QString("<Relationship Id=\"rId%1\" "
                                       "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
                                       "Target=\"%2\" TargetMode=\"External\"/>\n")
                               .arg(hl.rId).arg(hl.url.toHtmlEscaped());
                }
            }
            relsXml += "</Relationships>\n";
            zip.addFile(QString("xl/worksheets/_rels/sheet%1.xml.rels").arg(si + 1), relsXml.toUtf8());
        }
    }

    // Generate sheet rels for sheets with hyperlinks but no charts
    for (int i = 0; i < sheetCount; ++i) {
        if (chartsPerSheet.count(i)) continue; // already handled above
        if (i < static_cast<int>(allSheetHyperlinks.size()) && !allSheetHyperlinks[i].empty()) {
            QString relsXml;
            relsXml += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            relsXml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n";
            for (const auto& hl : allSheetHyperlinks[i]) {
                relsXml += QString("<Relationship Id=\"rId%1\" "
                                   "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
                                   "Target=\"%2\" TargetMode=\"External\"/>\n")
                           .arg(hl.rId).arg(hl.url.toHtmlEscaped());
            }
            relsXml += "</Relationships>\n";
            zip.addFile(QString("xl/worksheets/_rels/sheet%1.xml.rels").arg(i + 1), relsXml.toUtf8());
        }
    }

    // Save Nexel chart configs as custom JSON part (for Nexel round-trip)
    if (!charts.empty()) {
        QJsonArray chartsArray;
        for (const auto& c : charts) {
            QJsonObject obj;
            obj["sheetIndex"] = c.sheetIndex;
            obj["chartType"] = c.chartType;
            obj["title"] = c.title;
            obj["xAxisTitle"] = c.xAxisTitle;
            obj["yAxisTitle"] = c.yAxisTitle;
            obj["dataRange"] = c.dataRange;
            obj["themeIndex"] = c.themeIndex;
            obj["showLegend"] = c.showLegend;
            obj["showGridLines"] = c.showGridLines;
            obj["x"] = c.x;
            obj["y"] = c.y;
            obj["width"] = c.width;
            obj["height"] = c.height;
            chartsArray.append(obj);
        }
        QJsonDocument doc(chartsArray);
        zip.addFile("docMetadata/nexel-charts.json", doc.toJson(QJsonDocument::Compact));
    }

    // Save Nexel-native metadata (picklist/checkbox definitions)
    {
        QJsonObject metadata;
        QJsonArray sheetsArray;
        for (int si = 0; si < static_cast<int>(sheets.size()); ++si) {
            auto* sheet = sheets[si].get();
            QJsonObject sheetObj;
            sheetObj["index"] = si;

            // Picklist/Checkbox cells: scan for cells with these formats
            QJsonArray picklistRules;
            for (const auto& rule : sheet->getValidationRules()) {
                if (rule.type == Spreadsheet::DataValidationRule::List &&
                    (!rule.listItems.isEmpty() || !rule.listSourceRange.isEmpty())) {
                    QJsonObject ruleObj;
                    ruleObj["range"] = rule.range.toString();
                    QJsonArray items;
                    for (const auto& item : rule.listItems) items.append(item);
                    ruleObj["options"] = items;
                    if (!rule.listSourceRange.isEmpty()) {
                        ruleObj["sourceRange"] = rule.listSourceRange;
                        ruleObj["sortMode"] = rule.listSortMode;
                        ruleObj["ignoreBlanks"] = rule.listIgnoreBlanks;
                    }
                    if (!rule.listItemColors.isEmpty()) {
                        QJsonArray colorsArr;
                        for (const auto& c : rule.listItemColors) colorsArr.append(c);
                        ruleObj["optionColors"] = colorsArr;
                    }
                    picklistRules.append(ruleObj);
                }
            }
            if (!picklistRules.isEmpty()) sheetObj["picklists"] = picklistRules;

            // Track checkbox cells
            QJsonArray checkboxCells;
            sheet->forEachCell([&](int row, int col, const Cell& cell) {
                const auto& style = cell.getStyle();
                if (style.numberFormat == "Checkbox") {
                    QJsonObject cb;
                    cb["cell"] = CellAddress(row, col).toString();
                    cb["checked"] = cell.getValue().toBool() ||
                                    cell.getValue().toString().toLower() == "true" ||
                                    cell.getValue().toString() == "1";
                    checkboxCells.append(cb);
                } else if (style.numberFormat == "Picklist") {
                    // Picklist values stored in validation rules above; just mark the format
                }
            });
            if (!checkboxCells.isEmpty()) sheetObj["checkboxes"] = checkboxCells;

            // Table styles
            const auto& tables = sheet->getTables();
            if (!tables.empty()) {
                QJsonArray tablesArray;
                for (const auto& tbl : tables) {
                    QJsonObject tblObj;
                    tblObj["name"] = tbl.name;
                    tblObj["range"] = tbl.range.toString();
                    tblObj["hasHeaderRow"] = tbl.hasHeaderRow;
                    tblObj["bandedRows"] = tbl.bandedRows;
                    QJsonArray colNames;
                    for (const auto& cn : tbl.columnNames) colNames.append(cn);
                    tblObj["columnNames"] = colNames;
                    // Theme
                    QJsonObject themeObj;
                    themeObj["name"] = tbl.theme.name;
                    themeObj["headerBg"] = tbl.theme.headerBg.name();
                    themeObj["headerFg"] = tbl.theme.headerFg.name();
                    themeObj["bandedRow1"] = tbl.theme.bandedRow1.name();
                    themeObj["bandedRow2"] = tbl.theme.bandedRow2.name();
                    themeObj["borderColor"] = tbl.theme.borderColor.name();
                    tblObj["theme"] = themeObj;
                    tablesArray.append(tblObj);
                }
                sheetObj["tables"] = tablesArray;
            }

            // Conditional formatting rules
            const auto& cfRules = sheet->getConditionalFormatting().getAllRules();
            if (!cfRules.empty()) {
                QJsonArray cfArray;
                for (const auto& rule : cfRules) {
                    QJsonObject rObj;
                    rObj["range"] = rule->getRange().toString();
                    rObj["type"] = static_cast<int>(rule->getType());
                    if (!rule->getValue1().isNull())
                        rObj["value1"] = QJsonValue::fromVariant(rule->getValue1());
                    if (!rule->getValue2().isNull())
                        rObj["value2"] = QJsonValue::fromVariant(rule->getValue2());
                    if (!rule->getFormula().isEmpty())
                        rObj["formula"] = rule->getFormula();
                    const CellStyle& st = rule->getStyle();
                    QJsonObject styleObj;
                    if (st.bold) styleObj["bold"] = true;
                    if (st.italic) styleObj["italic"] = true;
                    if (st.underline) styleObj["underline"] = true;
                    if (st.foregroundColor != "#000000") styleObj["fontColor"] = st.foregroundColor;
                    if (st.backgroundColor != "#FFFFFF") styleObj["bgColor"] = st.backgroundColor;
                    rObj["style"] = styleObj;

                    // Visual formatting configs
                    if (rule->getType() == ConditionType::DataBar) {
                        const auto& cfg = rule->getDataBarConfig();
                        QJsonObject db;
                        db["color"] = cfg.barColor.name();
                        db["minValue"] = cfg.minValue;
                        db["maxValue"] = cfg.maxValue;
                        db["autoRange"] = cfg.autoRange;
                        db["showValue"] = cfg.showValue;
                        rObj["dataBar"] = db;
                    }
                    if (rule->getType() == ConditionType::ColorScale2 || rule->getType() == ConditionType::ColorScale3) {
                        const auto& cfg = rule->getColorScaleConfig();
                        QJsonObject cs;
                        cs["minColor"] = cfg.minColor.name();
                        cs["midColor"] = cfg.midColor.name();
                        cs["maxColor"] = cfg.maxColor.name();
                        cs["minValue"] = cfg.minValue;
                        cs["maxValue"] = cfg.maxValue;
                        cs["autoRange"] = cfg.autoRange;
                        cs["threeColor"] = cfg.threeColor;
                        rObj["colorScale"] = cs;
                    }
                    if (rule->getType() == ConditionType::IconSet) {
                        const auto& cfg = rule->getIconSetConfig();
                        QJsonObject is;
                        is["iconType"] = static_cast<int>(cfg.iconType);
                        is["threshold1"] = cfg.threshold1;
                        is["threshold2"] = cfg.threshold2;
                        is["showValue"] = cfg.showValue;
                        is["reverseOrder"] = cfg.reverseOrder;
                        rObj["iconSet"] = is;
                    }

                    cfArray.append(rObj);
                }
                sheetObj["conditionalFormats"] = cfArray;
            }

            if (sheetObj.contains("picklists") || sheetObj.contains("checkboxes") || sheetObj.contains("tables") || sheetObj.contains("conditionalFormats")) {
                sheetsArray.append(sheetObj);
            }
        }
        if (!sheetsArray.isEmpty()) {
            metadata["sheets"] = sheetsArray;
            QJsonDocument metaDoc(metadata);
            zip.addFile("docMetadata/nexel-metadata.json", metaDoc.toJson(QJsonDocument::Compact));
        }
    }

    zip.close();
    return true;
}

QByteArray XlsxService::generateContentTypes(int sheetCount, bool hasSharedStrings,
                                               int chartCount, const std::vector<int>& drawingSheetNums,
                                               bool hasCustomJson) {
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("Types");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

    xml.writeStartElement("Default");
    xml.writeAttribute("Extension", "rels");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    xml.writeEndElement();

    xml.writeStartElement("Default");
    xml.writeAttribute("Extension", "xml");
    xml.writeAttribute("ContentType", "application/xml");
    xml.writeEndElement();

    // Register JSON content type only when custom JSON parts are included
    if (hasCustomJson) {
        xml.writeStartElement("Default");
        xml.writeAttribute("Extension", "json");
        xml.writeAttribute("ContentType", "application/json");
        xml.writeEndElement();
    }

    xml.writeStartElement("Override");
    xml.writeAttribute("PartName", "/xl/workbook.xml");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    xml.writeEndElement();

    xml.writeStartElement("Override");
    xml.writeAttribute("PartName", "/xl/styles.xml");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    xml.writeEndElement();

    if (hasSharedStrings) {
        xml.writeStartElement("Override");
        xml.writeAttribute("PartName", "/xl/sharedStrings.xml");
        xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        xml.writeEndElement();
    }

    for (int i = 0; i < sheetCount; ++i) {
        xml.writeStartElement("Override");
        xml.writeAttribute("PartName", QString("/xl/worksheets/sheet%1.xml").arg(i + 1));
        xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        xml.writeEndElement();
    }

    // OOXML chart content types
    for (int i = 1; i <= chartCount; ++i) {
        xml.writeStartElement("Override");
        xml.writeAttribute("PartName", QString("/xl/charts/chart%1.xml").arg(i));
        xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
        xml.writeEndElement();
    }

    // Drawing content types
    for (int sheetNum : drawingSheetNums) {
        xml.writeStartElement("Override");
        xml.writeAttribute("PartName", QString("/xl/drawings/drawing%1.xml").arg(sheetNum));
        xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
        xml.writeEndElement();
    }

    xml.writeEndElement(); // Types
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateRels() {
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("Relationships");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    xml.writeStartElement("Relationship");
    xml.writeAttribute("Id", "rId1");
    xml.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    xml.writeAttribute("Target", "xl/workbook.xml");
    xml.writeEndElement();

    xml.writeEndElement();
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateWorkbook(const std::vector<std::shared_ptr<Spreadsheet>>& sheets) {
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("workbook");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    xml.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

    xml.writeStartElement("sheets");
    for (int i = 0; i < static_cast<int>(sheets.size()); ++i) {
        xml.writeStartElement("sheet");
        xml.writeAttribute("name", sheets[i]->getSheetName());
        xml.writeAttribute("sheetId", QString::number(i + 1));
        xml.writeAttribute("r:id", QString("rId%1").arg(i + 1));
        xml.writeEndElement();
    }
    xml.writeEndElement(); // sheets

    xml.writeEndElement(); // workbook
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateWorkbookRels(int sheetCount, bool hasSharedStrings) {
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("Relationships");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    for (int i = 0; i < sheetCount; ++i) {
        xml.writeStartElement("Relationship");
        xml.writeAttribute("Id", QString("rId%1").arg(i + 1));
        xml.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
        xml.writeAttribute("Target", QString("worksheets/sheet%1.xml").arg(i + 1));
        xml.writeEndElement();
    }

    xml.writeStartElement("Relationship");
    xml.writeAttribute("Id", QString("rId%1").arg(sheetCount + 1));
    xml.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    xml.writeAttribute("Target", "styles.xml");
    xml.writeEndElement();

    if (hasSharedStrings) {
        xml.writeStartElement("Relationship");
        xml.writeAttribute("Id", QString("rId%1").arg(sheetCount + 2));
        xml.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
        xml.writeAttribute("Target", "sharedStrings.xml");
        xml.writeEndElement();
    }

    xml.writeEndElement();
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateStyles(const std::vector<std::shared_ptr<Spreadsheet>>& sheets,
                                         std::map<QString, int>& styleIndexMap) {
    // Collect all unique fonts, fills, borders, numFmts from the style index map
    struct FontEntry { QString name; int size; bool bold, italic, underline, strikethrough; QString color; };
    struct FillEntry { QString bgColor; };
    struct BorderSideEntry { bool enabled = false; QString color = "#000000"; int width = 1; };
    struct BorderEntry { BorderSideEntry left, right, top, bottom; };

    // Build sorted style list by index
    std::vector<CellStyle> sortedStyles(styleIndexMap.size());

    // Collect all styles from sheets
    std::map<QString, CellStyle> keyToStyle;
    CellStyle defStyle;
    keyToStyle[cellStyleKey(defStyle)] = defStyle;

    for (const auto& sheet : sheets) {
        sheet->forEachCell([&](int, int, const Cell& cell) {
            const CellStyle& s = cell.getStyle();
            keyToStyle[cellStyleKey(s)] = s;
        });
    }

    for (const auto& [key, idx] : styleIndexMap) {
        if (keyToStyle.count(key)) {
            sortedStyles[idx] = keyToStyle[key];
        }
    }

    // Build unique fonts
    std::vector<FontEntry> fonts;
    std::map<QString, int> fontMap;
    auto getFontKey = [](const CellStyle& s) {
        return QString("%1|%2|%3|%4|%5|%6|%7")
            .arg(s.fontName).arg(s.fontSize).arg(s.bold).arg(s.italic)
            .arg(s.underline).arg(s.strikethrough).arg(s.foregroundColor);
    };

    for (const auto& s : sortedStyles) {
        QString fk = getFontKey(s);
        if (!fontMap.count(fk)) {
            fontMap[fk] = static_cast<int>(fonts.size());
            fonts.push_back({s.fontName, s.fontSize, bool(s.bold), bool(s.italic),
                             bool(s.underline), bool(s.strikethrough), resolveColorForExport(s.foregroundColor)});
        }
    }

    // Build unique fills
    std::vector<FillEntry> fills;
    std::map<QString, int> fillMap;
    // XLSX requires 2 default fills
    fills.push_back({"none"});
    fills.push_back({"gray125"});
    fillMap["none"] = 0;
    fillMap["gray125"] = 1;

    for (const auto& s : sortedStyles) {
        QString bg = resolveColorForExport(s.backgroundColor);
        if (bg != "#FFFFFF" && bg != "#ffffff") {
            if (!fillMap.count(bg)) {
                fillMap[bg] = static_cast<int>(fills.size());
                fills.push_back({bg});
            }
        }
    }

    // Number format map
    std::map<QString, int> numFmtMap;
    int customFmtId = 164;
    numFmtMap["General"] = 0;
    numFmtMap["Number"] = 2;
    numFmtMap["Currency"] = 7;
    numFmtMap["Percentage"] = 10;
    numFmtMap["Date"] = 14;
    numFmtMap["Time"] = 21;
    numFmtMap["Text"] = 49;

    // Write styles XML
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("styleSheet");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

    // fonts
    xml.writeStartElement("fonts");
    xml.writeAttribute("count", QString::number(fonts.size()));
    for (const auto& f : fonts) {
        xml.writeStartElement("font");
        if (f.bold) { xml.writeStartElement("b"); xml.writeEndElement(); }
        if (f.italic) { xml.writeStartElement("i"); xml.writeEndElement(); }
        if (f.underline) { xml.writeStartElement("u"); xml.writeEndElement(); }
        if (f.strikethrough) { xml.writeStartElement("strike"); xml.writeEndElement(); }
        xml.writeStartElement("sz"); xml.writeAttribute("val", QString::number(f.size)); xml.writeEndElement();
        xml.writeStartElement("color"); xml.writeAttribute("rgb", "FF" + f.color.mid(1)); xml.writeEndElement();
        xml.writeStartElement("name"); xml.writeAttribute("val", f.name); xml.writeEndElement();
        xml.writeEndElement(); // font
    }
    xml.writeEndElement(); // fonts

    // fills
    xml.writeStartElement("fills");
    xml.writeAttribute("count", QString::number(fills.size()));
    for (int i = 0; i < static_cast<int>(fills.size()); ++i) {
        xml.writeStartElement("fill");
        xml.writeStartElement("patternFill");
        if (i == 0) {
            xml.writeAttribute("patternType", "none");
        } else if (i == 1) {
            xml.writeAttribute("patternType", "gray125");
        } else {
            xml.writeAttribute("patternType", "solid");
            xml.writeStartElement("fgColor");
            xml.writeAttribute("rgb", "FF" + fills[i].bgColor.mid(1));
            xml.writeEndElement();
        }
        xml.writeEndElement(); // patternFill
        xml.writeEndElement(); // fill
    }
    xml.writeEndElement(); // fills

    // Build unique borders
    auto getBorderKey = [](const CellStyle& s) {
        return borderSideKey(s.borderLeft) + "|" + borderSideKey(s.borderRight) + "|" +
               borderSideKey(s.borderTop) + "|" + borderSideKey(s.borderBottom);
    };

    std::vector<BorderEntry> borderEntries;
    std::map<QString, int> borderMap;
    // Index 0 = no border (required default)
    borderEntries.push_back({});
    borderMap["0|0|0|0"] = 0;

    for (const auto& s : sortedStyles) {
        QString bk = getBorderKey(s);
        if (!borderMap.count(bk)) {
            borderMap[bk] = static_cast<int>(borderEntries.size());
            BorderEntry be;
            be.left = {bool(s.borderLeft.enabled), resolveColorForExport(s.borderLeft.color), int(s.borderLeft.width)};
            be.right = {bool(s.borderRight.enabled), resolveColorForExport(s.borderRight.color), int(s.borderRight.width)};
            be.top = {bool(s.borderTop.enabled), resolveColorForExport(s.borderTop.color), int(s.borderTop.width)};
            be.bottom = {bool(s.borderBottom.enabled), resolveColorForExport(s.borderBottom.color), int(s.borderBottom.width)};
            borderEntries.push_back(be);
        }
    }

    // borders
    xml.writeStartElement("borders");
    xml.writeAttribute("count", QString::number(borderEntries.size()));

    auto writeBorderSide = [&](const QString& side, const BorderSideEntry& bs) {
        if (bs.enabled) {
            xml.writeStartElement(side);
            xml.writeAttribute("style", bs.width >= 2 ? "medium" : "thin");
            xml.writeStartElement("color");
            if (bs.color.isEmpty() || bs.color == "#000000") {
                xml.writeAttribute("auto", "1");
            } else {
                xml.writeAttribute("rgb", "FF" + bs.color.mid(1));
            }
            xml.writeEndElement();
            xml.writeEndElement();
        } else {
            xml.writeEmptyElement(side);
        }
    };

    for (const auto& be : borderEntries) {
        xml.writeStartElement("border");
        writeBorderSide("left", be.left);
        writeBorderSide("right", be.right);
        writeBorderSide("top", be.top);
        writeBorderSide("bottom", be.bottom);
        xml.writeEmptyElement("diagonal");
        xml.writeEndElement();
    }
    xml.writeEndElement(); // borders

    // cellStyleXfs (required by XLSX spec — at least one default entry)
    xml.writeStartElement("cellStyleXfs");
    xml.writeAttribute("count", "1");
    xml.writeStartElement("xf");
    xml.writeAttribute("numFmtId", "0");
    xml.writeAttribute("fontId", "0");
    xml.writeAttribute("fillId", "0");
    xml.writeAttribute("borderId", "0");
    xml.writeEndElement(); // xf
    xml.writeEndElement(); // cellStyleXfs

    // cellXfs
    xml.writeStartElement("cellXfs");
    xml.writeAttribute("count", QString::number(sortedStyles.size()));
    for (const auto& s : sortedStyles) {
        xml.writeStartElement("xf");
        xml.writeAttribute("xfId", "0"); // Reference base style in cellStyleXfs
        // Font
        QString fk = getFontKey(s);
        xml.writeAttribute("fontId", QString::number(fontMap[fk]));
        // Fill
        QString bg = resolveColorForExport(s.backgroundColor);
        int fillId = 0;
        if (bg != "#FFFFFF" && bg != "#ffffff" && fillMap.count(bg)) {
            fillId = fillMap[bg];
        }
        xml.writeAttribute("fillId", QString::number(fillId));
        // Border
        QString bk = getBorderKey(s);
        int borderId = borderMap.count(bk) ? borderMap[bk] : 0;
        bool hasBorder = borderId != 0;
        xml.writeAttribute("borderId", QString::number(borderId));
        // NumFmt
        int numFmtId = 0;
        if (numFmtMap.count(s.numberFormat)) {
            numFmtId = numFmtMap[s.numberFormat];
        }
        xml.writeAttribute("numFmtId", QString::number(numFmtId));

        if (fontMap[fk] != 0) xml.writeAttribute("applyFont", "1");
        if (fillId != 0) xml.writeAttribute("applyFill", "1");
        if (hasBorder) xml.writeAttribute("applyBorder", "1");
        if (numFmtId != 0) xml.writeAttribute("applyNumberFormat", "1");

        // Alignment
        bool hasAlign = (s.hAlign != HorizontalAlignment::General || s.vAlign != VerticalAlignment::Bottom);
        if (hasAlign) {
            xml.writeAttribute("applyAlignment", "1");
            xml.writeStartElement("alignment");
            switch (s.hAlign) {
                case HorizontalAlignment::Left: xml.writeAttribute("horizontal", "left"); break;
                case HorizontalAlignment::Center: xml.writeAttribute("horizontal", "center"); break;
                case HorizontalAlignment::Right: xml.writeAttribute("horizontal", "right"); break;
                default: break;
            }
            switch (s.vAlign) {
                case VerticalAlignment::Top: xml.writeAttribute("vertical", "top"); break;
                case VerticalAlignment::Middle: xml.writeAttribute("vertical", "center"); break;
                default: break; // Bottom is default
            }
            xml.writeEndElement(); // alignment
        }

        xml.writeEndElement(); // xf
    }
    xml.writeEndElement(); // cellXfs

    // cellStyles (required by XLSX spec)
    xml.writeStartElement("cellStyles");
    xml.writeAttribute("count", "1");
    xml.writeStartElement("cellStyle");
    xml.writeAttribute("name", "Normal");
    xml.writeAttribute("xfId", "0");
    xml.writeAttribute("builtinId", "0");
    xml.writeEndElement(); // cellStyle
    xml.writeEndElement(); // cellStyles

    xml.writeEndElement(); // styleSheet
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateSharedStrings(const QStringList& sharedStrings) {
    QByteArray data;
    QBuffer buf(&data);
    buf.open(QIODevice::WriteOnly);
    QXmlStreamWriter xml(&buf);
    xml.setAutoFormatting(true);
    xml.writeStartDocument();
    xml.writeStartElement("sst");
    xml.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    xml.writeAttribute("uniqueCount", QString::number(sharedStrings.size()));

    for (const auto& s : sharedStrings) {
        xml.writeStartElement("si");
        xml.writeStartElement("t");
        if (s.startsWith(' ') || s.endsWith(' ') || s.startsWith('\t') || s.endsWith('\t')) {
            xml.writeAttribute("xml:space", "preserve");
        }
        xml.writeCharacters(s);
        xml.writeEndElement(); // t
        xml.writeEndElement(); // si
    }

    xml.writeEndElement(); // sst
    xml.writeEndDocument();
    buf.close();
    return data;
}

QByteArray XlsxService::generateSheet(Spreadsheet* sheet,
                                        const std::map<QString, int>& styleIndexMap,
                                        QStringList& sharedStrings) {
    Q_UNUSED(styleIndexMap);
    Q_UNUSED(sharedStrings);
    // This method is not used in the final export flow - sheets are generated inline
    return QByteArray();
}

// ============== XLSX CHART IMPORT HELPERS ==============

QString XlsxService::resolveRelativePath(const QString& basePath, const QString& relativePath) {
    if (relativePath.startsWith('/')) {
        return relativePath.mid(1); // absolute path within package
    }

    int lastSlash = basePath.lastIndexOf('/');
    QStringList baseParts;
    if (lastSlash >= 0) {
        baseParts = basePath.left(lastSlash).split('/');
    }

    QStringList relParts = relativePath.split('/');
    for (const auto& part : relParts) {
        if (part == "..") {
            if (!baseParts.isEmpty()) baseParts.removeLast();
        } else if (part != "." && !part.isEmpty()) {
            baseParts.append(part);
        }
    }

    return baseParts.join('/');
}

std::map<QString, QString> XlsxService::parseRels(const QByteArray& relsXml) {
    std::map<QString, QString> rels;
    if (relsXml.isEmpty()) return rels;

    QXmlStreamReader xml(relsXml);
    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement() && xml.name() == u"Relationship") {
            QString id = xml.attributes().value("Id").toString();
            QString target = xml.attributes().value("Target").toString();
            if (!id.isEmpty() && !target.isEmpty()) {
                rels[id] = target;
            }
        }
    }
    return rels;
}

QString XlsxService::findDrawingRId(const QByteArray& sheetXml) {
    QXmlStreamReader xml(sheetXml);
    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement() && xml.name() == u"drawing") {
            // Try namespace-based lookup first (most reliable)
            static const QString relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            QString rId = xml.attributes().value(relNs, "id").toString();
            // Fallback: try qualified name r:id
            if (rId.isEmpty()) {
                rId = xml.attributes().value("r:id").toString();
            }
            // Last resort: scan all attributes for any ":id"
            if (rId.isEmpty()) {
                for (const auto& attr : xml.attributes()) {
                    if (attr.qualifiedName().toString().endsWith(":id")) {
                        rId = attr.value().toString();
                        break;
                    }
                }
            }
            return rId;
        }
    }
    return "";
}

std::vector<XlsxService::DrawingChartRef> XlsxService::parseDrawing(const QByteArray& drawingXml) {
    std::vector<DrawingChartRef> refs;
    QXmlStreamReader xml(drawingXml);

    DrawingChartRef current;
    bool inAnchor = false;
    bool inFrom = false, inTo = false;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            QString n = xml.name().toString();
            if (n == "twoCellAnchor" || n == "oneCellAnchor") {
                inAnchor = true;
                current = DrawingChartRef();
            } else if (inAnchor && n == "from") {
                inFrom = true;
            } else if (inAnchor && n == "to") {
                inTo = true;
            } else if ((inFrom || inTo) && n == "col") {
                int val = xml.readElementText().toInt();
                if (inFrom) current.fromCol = val;
                else current.toCol = val;
            } else if ((inFrom || inTo) && n == "row") {
                int val = xml.readElementText().toInt();
                if (inFrom) current.fromRow = val;
                else current.toRow = val;
            } else if (inAnchor && n == "chart") {
                // <c:chart r:id="rId1"/>
                QString rId;
                for (const auto& attr : xml.attributes()) {
                    QString qn = attr.qualifiedName().toString();
                    if (qn.endsWith(":id") || attr.name() == u"id") {
                        rId = attr.value().toString();
                        break;
                    }
                }
                if (!rId.isEmpty()) {
                    current.chartRId = rId;
                }
            }
        } else if (xml.isEndElement()) {
            QString n = xml.name().toString();
            if (n == "twoCellAnchor" || n == "oneCellAnchor") {
                if (!current.chartRId.isEmpty()) {
                    refs.push_back(current);
                }
                inAnchor = false;
                inFrom = false;
                inTo = false;
            } else if (n == "from") {
                inFrom = false;
            } else if (n == "to") {
                inTo = false;
            }
        }
    }

    return refs;
}

ImportedChart XlsxService::parseChartXml(const QByteArray& chartXml) {
    ImportedChart chart;
    chart.chartType = "column"; // default

    QXmlStreamReader xml(chartXml);

    // Context flags
    bool inPlotArea = false;
    bool inSer = false;
    bool inSerTx = false;   // <tx> inside <ser>
    bool inCat = false;     // <cat>
    bool inVal = false;     // <val>
    bool inXVal = false;    // <xVal> (scatter)
    bool inYVal = false;    // <yVal> (scatter)
    bool inChartTitle = false;
    bool inCatAx = false;
    bool inValAx = false;
    bool inAxTitle = false;
    bool chartTypeSet = false;
    bool inNumRef = false;
    bool inStrRef = false;

    ImportedChartSeries currentSeries;
    QVector<QString> sharedCategories;

    while (!xml.atEnd()) {
        xml.readNext();

        if (xml.isStartElement()) {
            QString n = xml.name().toString();

            if (n == "plotArea") {
                inPlotArea = true;
            }
            // Chart type detection (only first chart type element in plotArea)
            else if (inPlotArea && !chartTypeSet) {
                if (n == "barChart" || n == "bar3DChart") {
                    chart.chartType = "column";
                    chartTypeSet = true;
                } else if (n == "lineChart" || n == "line3DChart") {
                    chart.chartType = "line";
                    chartTypeSet = true;
                } else if (n == "areaChart" || n == "area3DChart") {
                    chart.chartType = "area";
                    chartTypeSet = true;
                } else if (n == "scatterChart") {
                    chart.chartType = "scatter";
                    chartTypeSet = true;
                } else if (n == "pieChart" || n == "pie3DChart" || n == "ofPieChart") {
                    chart.chartType = "pie";
                    chartTypeSet = true;
                } else if (n == "doughnutChart") {
                    chart.chartType = "donut";
                    chartTypeSet = true;
                } else if (n == "radarChart") {
                    chart.chartType = "line";
                    chartTypeSet = true;
                } else if (n == "bubbleChart") {
                    chart.chartType = "scatter";
                    chartTypeSet = true;
                } else if (n == "stockChart") {
                    chart.chartType = "line";
                    chartTypeSet = true;
                }
            }

            // Bar direction: col=column, bar=horizontal bar
            if (n == "barDir") {
                if (xml.attributes().value("val") == u"bar") {
                    chart.chartType = "bar";
                }
            }

            // Series
            if (inPlotArea && n == "ser") {
                inSer = true;
                currentSeries = ImportedChartSeries();
            }
            if (inSer && n == "tx") inSerTx = true;
            if (inSer && n == "cat") inCat = true;
            if (inSer && n == "val") inVal = true;
            if (inSer && n == "xVal") inXVal = true;
            if (inSer && n == "yVal") inYVal = true;
            if (inSer && n == "numRef") inNumRef = true;
            if (inSer && n == "strRef") inStrRef = true;

            // Formula reference inside <numRef> or <strRef>
            if (n == "f" && inSer && (inNumRef || inStrRef)) {
                QString ref = xml.readElementText();
                if (inVal || inYVal) {
                    currentSeries.valRef = ref;
                } else if (inCat || inXVal) {
                    currentSeries.catRef = ref;
                }
            }

            // Value element inside series context
            if (n == "v" && inSer && (inVal || inYVal || inXVal || inCat || inSerTx)) {
                QString text = xml.readElementText();
                if (inVal || inYVal) {
                    bool ok;
                    double v = text.toDouble(&ok);
                    currentSeries.values.append(ok ? v : 0.0);
                } else if (inXVal) {
                    bool ok;
                    double v = text.toDouble(&ok);
                    currentSeries.xNumeric.append(ok ? v : 0.0);
                } else if (inCat) {
                    currentSeries.categories.append(text);
                } else if (inSerTx) {
                    currentSeries.name = text;
                }
            }

            // Chart title (not inside plotArea or axes)
            if (n == "title" && !inPlotArea && !inCatAx && !inValAx && !inSer) {
                inChartTitle = true;
            }

            // Axis elements
            if (inPlotArea && (n == "catAx" || n == "dateAx")) inCatAx = true;
            if (inPlotArea && n == "valAx") inValAx = true;

            // Axis title
            if ((inCatAx || inValAx) && n == "title") inAxTitle = true;

            // Text element <a:t> for titles
            if (n == "t" && !inSer && (inChartTitle || inAxTitle)) {
                QString text = xml.readElementText();
                if (inAxTitle && inCatAx) {
                    if (!chart.xAxisTitle.isEmpty()) chart.xAxisTitle += " ";
                    chart.xAxisTitle += text;
                } else if (inAxTitle && inValAx) {
                    if (!chart.yAxisTitle.isEmpty()) chart.yAxisTitle += " ";
                    chart.yAxisTitle += text;
                } else if (inChartTitle && !inAxTitle) {
                    if (!chart.title.isEmpty()) chart.title += " ";
                    chart.title += text;
                }
            }
        }
        else if (xml.isEndElement()) {
            QString n = xml.name().toString();

            if (n == "plotArea") inPlotArea = false;
            if (n == "ser" && inSer) {
                chart.series.append(currentSeries);
                if (sharedCategories.isEmpty() && !currentSeries.categories.isEmpty()) {
                    sharedCategories = currentSeries.categories;
                }
                inSer = false;
                inSerTx = false;
                inCat = false;
                inVal = false;
                inXVal = false;
                inYVal = false;
            }
            if (n == "tx" && inSer) inSerTx = false;
            if (n == "cat") inCat = false;
            if (n == "val") inVal = false;
            if (n == "xVal") inXVal = false;
            if (n == "yVal") inYVal = false;
            if (n == "numRef") inNumRef = false;
            if (n == "strRef") inStrRef = false;
            if (n == "title" && inChartTitle && !inCatAx && !inValAx) inChartTitle = false;
            if (n == "title" && inAxTitle) inAxTitle = false;
            if (n == "catAx" || n == "dateAx") { inCatAx = false; inAxTitle = false; }
            if (n == "valAx") { inValAx = false; inAxTitle = false; }
        }
    }

    // Copy shared categories to series that don't have them
    if (!sharedCategories.isEmpty()) {
        for (auto& s : chart.series) {
            if (s.categories.isEmpty()) {
                s.categories = sharedCategories;
            }
        }
    }

    // Construct combined dataRange from series formula references
    if (!chart.series.isEmpty() && chart.dataRange.isEmpty()) {
        int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
        bool hasRef = false;

        auto processRef = [&](const QString& ref) {
            if (ref.isEmpty()) return;
            // Strip sheet name: "Sheet1!$A$1:$B$10" -> "$A$1:$B$10"
            QString cellPart = ref;
            int bangIdx = ref.lastIndexOf('!');
            if (bangIdx >= 0) cellPart = ref.mid(bangIdx + 1);

            CellRange cr(cellPart);
            if (cr.isValid()) {
                minRow = std::min(minRow, cr.getStart().row);
                maxRow = std::max(maxRow, cr.getEnd().row);
                minCol = std::min(minCol, cr.getStart().col);
                maxCol = std::max(maxCol, cr.getEnd().col);
                hasRef = true;
            }
        };

        for (const auto& s : chart.series) {
            processRef(s.valRef);
            processRef(s.catRef);
        }

        if (hasRef) {
            // Include header row if the range doesn't start at row 0
            if (minRow > 0) minRow--;
            chart.dataRange = CellAddress(minRow, minCol).toString() + ":"
                            + CellAddress(maxRow, maxCol).toString();
        }
    }

    return chart;
}

// ========== OOXML Chart Export Helpers ==========

static QString xmlEsc(const QString& str) {
    QString s = str;
    s.replace("&", "&amp;");
    s.replace("<", "&lt;");
    s.replace(">", "&gt;");
    s.replace("\"", "&quot;");
    return s;
}

QVector<QColor> XlsxService::chartThemeColors(int themeIndex) {
    static const QVector<QVector<QColor>> palettes = {
        { QColor("#4472C4"), QColor("#ED7D31"), QColor("#A5A5A5"), QColor("#FFC000"), QColor("#5B9BD5"), QColor("#70AD47") },
        { QColor("#2196F3"), QColor("#FF5722"), QColor("#4CAF50"), QColor("#FFC107"), QColor("#9C27B0"), QColor("#00BCD4") },
        { QColor("#268BD2"), QColor("#DC322F"), QColor("#859900"), QColor("#B58900"), QColor("#6C71C4"), QColor("#2AA198") },
        { QColor("#00C8FF"), QColor("#FF6384"), QColor("#36A2EB"), QColor("#FFCE56"), QColor("#9966FF"), QColor("#FF9F40") },
        { QColor("#333333"), QColor("#666666"), QColor("#999999"), QColor("#BBBBBB"), QColor("#444444"), QColor("#777777") },
        { QColor("#A8D8EA"), QColor("#FFB7B2"), QColor("#B5EAD7"), QColor("#FFDAC1"), QColor("#C7CEEA"), QColor("#E2F0CB") },
    };
    int idx = qBound(0, themeIndex, static_cast<int>(palettes.size()) - 1);
    return palettes[idx];
}

QByteArray XlsxService::generateChartXml(const NexelChartExport& chart, const QString& sheetName) {
    // Parse data range (e.g. "A1:D10")
    QStringList parts = chart.dataRange.split(':');
    if (parts.size() != 2) return QByteArray();

    // Parse cell references
    auto parseCellRef = [](const QString& ref, int& row, int& col) {
        int i = 0;
        while (i < ref.length() && ref[i].isLetter()) i++;
        QString letters = ref.left(i);
        col = 0;
        for (int j = 0; j < letters.length(); ++j)
            col = col * 26 + (letters[j].toUpper().unicode() - 'A' + 1);
        col -= 1;
        row = ref.mid(i).toInt() - 1;
    };

    int startRow, startCol, endRow, endCol;
    parseCellRef(parts[0].trimmed(), startRow, startCol);
    parseCellRef(parts[1].trimmed(), endRow, endCol);
    if (startRow > endRow) std::swap(startRow, endRow);
    if (startCol > endCol) std::swap(startCol, endCol);

    int numSeries = endCol - startCol; // first column = categories
    if (numSeries < 1) return QByteArray();

    // Sheet reference (quote if contains spaces)
    QString sheetRef = sheetName.contains(' ') ? "'" + sheetName + "'" : sheetName;

    // Determine chart type XML
    QString chartElem, chartAttrs;
    bool hasAxes = true, isScatter = false;

    if (chart.chartType == "column") {
        chartElem = "c:barChart";
        chartAttrs = "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
    } else if (chart.chartType == "bar") {
        chartElem = "c:barChart";
        chartAttrs = "<c:barDir val=\"bar\"/><c:grouping val=\"clustered\"/>";
    } else if (chart.chartType == "line") {
        chartElem = "c:lineChart";
        chartAttrs = "<c:grouping val=\"standard\"/>";
    } else if (chart.chartType == "area") {
        chartElem = "c:areaChart";
        chartAttrs = "<c:grouping val=\"standard\"/>";
    } else if (chart.chartType == "scatter") {
        chartElem = "c:scatterChart";
        chartAttrs = "<c:scatterStyle val=\"lineMarker\"/>";
        isScatter = true;
    } else if (chart.chartType == "pie") {
        chartElem = "c:pieChart";
        hasAxes = false;
    } else if (chart.chartType == "donut") {
        chartElem = "c:doughnutChart";
        hasAxes = false;
    } else {
        chartElem = "c:barChart";
        chartAttrs = "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
    }

    auto colors = chartThemeColors(chart.themeIndex);

    QString x;
    x += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    x += "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n";
    x += "<c:chart>\n";

    // Title
    if (!chart.title.isEmpty()) {
        x += "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>"
           + xmlEsc(chart.title) + "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>\n";
    }
    x += "<c:autoTitleDeleted val=\"0\"/>\n";
    x += "<c:plotArea>\n<c:layout/>\n";

    // Chart type element
    x += "<" + chartElem + ">\n" + chartAttrs + "\n";

    // Series
    QString catCol = columnIndexToLetter(startCol);
    for (int i = 0; i < numSeries; ++i) {
        int colIdx = startCol + 1 + i;
        QString serCol = columnIndexToLetter(colIdx);
        QColor clr = colors[i % colors.size()];
        QString hex = clr.name().mid(1).toUpper();

        x += "<c:ser>\n";
        x += QString("<c:idx val=\"%1\"/><c:order val=\"%1\"/>\n").arg(i);

        // Series name from header
        x += "<c:tx><c:strRef><c:f>" + xmlEsc(sheetRef) + "!$" + serCol + "$"
           + QString::number(startRow + 1) + "</c:f></c:strRef></c:tx>\n";

        // Series color
        x += "<c:spPr><a:solidFill><a:srgbClr val=\"" + hex + "\"/></a:solidFill></c:spPr>\n";

        QString catRange = xmlEsc(sheetRef) + "!$" + catCol + "$" + QString::number(startRow + 2)
                         + ":$" + catCol + "$" + QString::number(endRow + 1);
        QString valRange = xmlEsc(sheetRef) + "!$" + serCol + "$" + QString::number(startRow + 2)
                         + ":$" + serCol + "$" + QString::number(endRow + 1);

        if (isScatter) {
            x += "<c:xVal><c:numRef><c:f>" + catRange + "</c:f></c:numRef></c:xVal>\n";
            x += "<c:yVal><c:numRef><c:f>" + valRange + "</c:f></c:numRef></c:yVal>\n";
        } else {
            x += "<c:cat><c:strRef><c:f>" + catRange + "</c:f></c:strRef></c:cat>\n";
            x += "<c:val><c:numRef><c:f>" + valRange + "</c:f></c:numRef></c:val>\n";
        }

        x += "</c:ser>\n";
    }

    if (hasAxes) {
        x += "<c:axId val=\"111111111\"/><c:axId val=\"222222222\"/>\n";
    }
    if (chart.chartType == "donut") {
        x += "<c:holeSize val=\"50\"/>\n";
    }
    x += "</" + chartElem + ">\n";

    // Axes
    if (hasAxes) {
        if (isScatter) {
            x += "<c:valAx><c:axId val=\"111111111\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                 "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"222222222\"/></c:valAx>\n";
            x += "<c:valAx><c:axId val=\"222222222\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                 "<c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"111111111\"/></c:valAx>\n";
        } else {
            x += "<c:catAx><c:axId val=\"111111111\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                 "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"222222222\"/></c:catAx>\n";
            x += "<c:valAx><c:axId val=\"222222222\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                 "<c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"111111111\"/></c:valAx>\n";
        }
    }
    x += "</c:plotArea>\n";

    if (chart.showLegend) {
        x += "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>\n";
    }
    x += "<c:plotVisOnly val=\"1\"/>\n";
    x += "</c:chart>\n</c:chartSpace>\n";

    return x.toUtf8();
}

QByteArray XlsxService::generateDrawingXml(const std::vector<NexelChartExport>& allCharts,
                                            const std::vector<int>& chartIndices) {
    QString x;
    x += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    x += "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n";

    for (int i = 0; i < static_cast<int>(chartIndices.size()); ++i) {
        const auto& c = allCharts[chartIndices[i]];
        // Convert pixel position to approximate column/row (default col ~64px, row ~20px)
        int fromCol = c.x / 64;
        int fromRow = c.y / 20;
        int toCol = (c.x + c.width) / 64;
        int toRow = (c.y + c.height) / 20;
        int fromColOff = (c.x % 64) * 9525; // EMU
        int fromRowOff = (c.y % 20) * 9525;
        int toColOff = ((c.x + c.width) % 64) * 9525;
        int toRowOff = ((c.y + c.height) % 20) * 9525;

        x += "<xdr:twoCellAnchor>\n";
        x += QString("<xdr:from><xdr:col>%1</xdr:col><xdr:colOff>%2</xdr:colOff>"
                     "<xdr:row>%3</xdr:row><xdr:rowOff>%4</xdr:rowOff></xdr:from>\n")
             .arg(fromCol).arg(fromColOff).arg(fromRow).arg(fromRowOff);
        x += QString("<xdr:to><xdr:col>%1</xdr:col><xdr:colOff>%2</xdr:colOff>"
                     "<xdr:row>%3</xdr:row><xdr:rowOff>%4</xdr:rowOff></xdr:to>\n")
             .arg(toCol).arg(toColOff).arg(toRow).arg(toRowOff);

        x += "<xdr:graphicFrame macro=\"\">\n";
        x += QString("<xdr:nvGraphicFramePr><xdr:cNvPr id=\"%1\" name=\"Chart %1\"/>"
                     "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>\n").arg(i + 1);
        x += "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>\n";
        x += "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
             "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
             "r:id=\"rId" + QString::number(i + 1) + "\"/>"
             "</a:graphicData></a:graphic>\n";
        x += "</xdr:graphicFrame>\n";
        x += "<xdr:clientData/>\n";
        x += "</xdr:twoCellAnchor>\n";
    }

    x += "</xdr:wsDr>\n";
    return x.toUtf8();
}

QByteArray XlsxService::generateDrawingRels(int chartCount, int startChartNum) {
    QString x;
    x += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    x += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n";
    for (int i = 0; i < chartCount; ++i) {
        x += QString("<Relationship Id=\"rId%1\" "
                     "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" "
                     "Target=\"../charts/chart%2.xml\"/>\n")
             .arg(i + 1).arg(startChartNum + i);
    }
    x += "</Relationships>\n";
    return x.toUtf8();
}

QByteArray XlsxService::generateSheetRels(int drawingNum) {
    QString x;
    x += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    x += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n";
    x += QString("<Relationship Id=\"rId1\" "
                 "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
                 "Target=\"../drawings/drawing%1.xml\"/>\n").arg(drawingNum);
    x += "</Relationships>\n";
    return x.toUtf8();
}
