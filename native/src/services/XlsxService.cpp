#include "XlsxService.h"
#include <QFile>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QRegularExpression>
#include <QDate>
#include <QBuffer>
#include <QtCore/private/qzipreader_p.h>
#include <QtCore/private/qzipwriter_p.h>
#include <algorithm>

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

        // Parse sheet rels to find drawing path
        QString fullSheetPath = "xl/" + info.filePath;
        int lastSlash = fullSheetPath.lastIndexOf('/');
        QString sheetDirPath = fullSheetPath.left(lastSlash);
        QString sheetFileName = fullSheetPath.mid(lastSlash + 1);
        QString sheetRelsPath = sheetDirPath + "/_rels/" + sheetFileName + ".rels";
        QByteArray sheetRelsData = zip.fileData(sheetRelsPath);
        auto sheetRels = parseRels(sheetRelsData);

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

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"si") {
                inSi = true;
                currentString.clear();
            }
        } else if (xml.isCharacters() && inSi) {
            currentString += xml.text().toString();
        } else if (xml.isEndElement()) {
            if (xml.name() == u"si") {
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

void XlsxService::parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                              const std::vector<CellStyle>& styles, Spreadsheet* sheet) {
    QXmlStreamReader xml(xmlData);

    static QRegularExpression cellRefRe("^([A-Z]+)(\\d+)$");
    static QRegularExpression rangeRefRe("^([A-Z]+)(\\d+):([A-Z]+)(\\d+)$");

    while (!xml.atEnd()) {
        xml.readNext();

        // Parse column widths: <col min="1" max="3" width="15.5" customWidth="1"/>
        if (xml.isStartElement() && xml.name() == u"col") {
            int minCol = xml.attributes().value("min").toInt() - 1; // XLSX is 1-based
            int maxCol = xml.attributes().value("max").toInt() - 1;
            double width = xml.attributes().value("width").toDouble();
            if (width > 0 && maxCol < 256) {
                // Excel width units ≈ character widths; convert to pixels (approx 7.5px per unit)
                int pixelWidth = qMax(30, static_cast<int>(width * 7.5));
                for (int c = minCol; c <= maxCol && c < 256; ++c) {
                    sheet->setColumnWidth(c, pixelWidth);
                }
            }
            continue;
        }

        // Parse row heights: <row r="1" ht="25.5" customHeight="1">
        if (xml.isStartElement() && xml.name() == u"row") {
            QString htStr = xml.attributes().value("ht").toString();
            if (!htStr.isEmpty()) {
                int rowIdx = xml.attributes().value("r").toInt() - 1; // XLSX is 1-based
                double ht = htStr.toDouble();
                if (ht > 0 && rowIdx >= 0) {
                    // Excel height is in points; convert to pixels (1pt ≈ 1.333px)
                    int pixelHeight = qMax(14, static_cast<int>(ht * 1.333));
                    sheet->setRowHeight(rowIdx, pixelHeight);
                }
            }
            // Don't continue — row contains child <c> elements
        }

        // Parse merged cells: <mergeCell ref="A1:D1"/>
        if (xml.isStartElement() && xml.name() == u"mergeCell") {
            QString ref = xml.attributes().value("ref").toString();
            auto rangeMatch = rangeRefRe.match(ref);
            if (rangeMatch.hasMatch()) {
                int startCol = columnLetterToIndex(rangeMatch.captured(1));
                int startRow = rangeMatch.captured(2).toInt() - 1;
                int endCol = columnLetterToIndex(rangeMatch.captured(3));
                int endRow = rangeMatch.captured(4).toInt() - 1;
                CellRange range(CellAddress(startRow, startCol), CellAddress(endRow, endCol));
                sheet->mergeCells(range);
            }
            continue;
        }

        if (xml.isStartElement() && xml.name() == u"c") {
            QString ref = xml.attributes().value("r").toString();
            QString type = xml.attributes().value("t").toString();
            int styleIdx = xml.attributes().value("s").toInt();

            auto match = cellRefRe.match(ref);
            if (!match.hasMatch()) continue;

            int col = columnLetterToIndex(match.captured(1));
            int row = match.captured(2).toInt() - 1; // XLSX is 1-based

            // Read child elements: <f> (formula), <v> (value), <is> (inline string)
            QString value;
            QString formula;
            QString inlineStr;
            bool hasInlineStr = false;
            int depth = 1; // track nesting depth within <c>

            while (!xml.atEnd() && depth > 0) {
                xml.readNext();
                if (xml.isStartElement()) {
                    if (xml.name() == u"v") {
                        value = xml.readElementText();
                    } else if (xml.name() == u"f") {
                        formula = xml.readElementText();
                    } else if (xml.name() == u"is") {
                        // Inline string - read <t> children
                        hasInlineStr = true;
                    } else if (hasInlineStr && xml.name() == u"t") {
                        inlineStr += xml.readElementText();
                    } else {
                        depth++;
                    }
                } else if (xml.isEndElement()) {
                    if (xml.name() == u"c") {
                        depth = 0; // exit
                    } else if (xml.name() == u"is") {
                        hasInlineStr = false;
                    }
                }
            }

            CellAddress addr(row, col);
            bool cellSet = false;

            // If the cell has a formula, set it via setCellFormula
            if (!formula.isEmpty()) {
                QString f = formula;
                if (!f.startsWith('=')) f = "=" + f;
                sheet->setCellFormula(addr, f);
                cellSet = true;
                // Also set the cached value if present
                if (!value.isEmpty() && type != "s") {
                    // Numeric cached value - set directly so it displays before recalc
                    bool ok;
                    double num = value.toDouble(&ok);
                    if (ok) {
                        sheet->setCellValue(addr, num);
                    }
                }
            }
            // Handle inline strings first (type="inlineStr")
            else if (type == "inlineStr" && !inlineStr.isEmpty()) {
                sheet->setCellValue(addr, inlineStr);
                cellSet = true;
            }
            // Handle shared string reference
            else if (type == "s" && !value.isEmpty()) {
                int ssIdx = value.toInt();
                if (ssIdx >= 0 && ssIdx < sharedStrings.size()) {
                    sheet->setCellValue(addr, sharedStrings[ssIdx]);
                    cellSet = true;
                }
            }
            // Handle boolean
            else if (type == "b" && !value.isEmpty()) {
                sheet->setCellValue(addr, value == "1" ? "TRUE" : "FALSE");
                cellSet = true;
            }
            // Handle string formula result (type="str")
            else if (type == "str" && !value.isEmpty()) {
                sheet->setCellValue(addr, value);
                cellSet = true;
            }
            // Handle numeric / date values
            else if (!value.isEmpty()) {
                bool ok;
                double num = value.toDouble(&ok);
                if (ok) {
                    // Check if this cell has a Date/Time format - convert serial to date string
                    bool isDateFmt = false;
                    if (styleIdx > 0 && styleIdx < static_cast<int>(styles.size())) {
                        const auto& fmt = styles[styleIdx].numberFormat;
                        isDateFmt = (fmt == "Date" || fmt == "Time");
                    }
                    if (isDateFmt && num > 0 && num < 2958466) {
                        QDate epoch(1899, 12, 30);
                        QDate date = epoch.addDays(static_cast<qint64>(num));
                        if (date.isValid()) {
                            sheet->setCellValue(addr, date.toString("MM/dd/yyyy"));
                        } else {
                            sheet->setCellValue(addr, num);
                        }
                    } else {
                        sheet->setCellValue(addr, num);
                    }
                } else {
                    sheet->setCellValue(addr, value);
                }
                cellSet = true;
            }

            // Apply style (even for cells without values, e.g. styled empty cells)
            if ((cellSet || styleIdx > 0) && styleIdx < static_cast<int>(styles.size())) {
                auto cell = sheet->getCell(addr);
                if (cell) {
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
    }
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

QString XlsxService::cellStyleKey(const CellStyle& style) {
    return QString("%1|%2|%3|%4|%5|%6|%7|%8|%9|%10|%11|%12")
        .arg(style.fontName).arg(style.fontSize)
        .arg(style.bold).arg(style.italic).arg(style.underline).arg(style.strikethrough)
        .arg(style.foregroundColor).arg(style.backgroundColor)
        .arg(static_cast<int>(style.hAlign)).arg(static_cast<int>(style.vAlign))
        .arg(style.numberFormat)
        .arg(style.borderTop.enabled || style.borderBottom.enabled ||
             style.borderLeft.enabled || style.borderRight.enabled ? 1 : 0);
}

bool XlsxService::exportToFile(const std::vector<std::shared_ptr<Spreadsheet>>& sheets,
                                 const QString& filePath) {
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

    // Generate shared strings
    QStringList sharedStrings;
    std::map<QString, int> ssMap; // string -> index

    // Generate sheet data
    std::vector<QByteArray> sheetXmls;
    for (const auto& sheet : sheets) {
        QStringList localSS;
        QByteArray sheetXml = generateSheet(sheet.get(), styleIndexMap, localSS);
        // Merge local shared strings into global
        for (const auto& s : localSS) {
            if (ssMap.find(s) == ssMap.end()) {
                ssMap[s] = sharedStrings.size();
                sharedStrings.append(s);
            }
        }
        sheetXmls.push_back(sheetXml);
    }

    // Now regenerate sheets with correct shared string indices
    sharedStrings.clear();
    ssMap.clear();
    sheetXmls.clear();
    for (const auto& sheet : sheets) {
        // First pass: collect strings
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

    // Second pass: generate sheets
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

        // sheetData
        xml.writeStartElement("sheetData");

        int maxRow = sheet->getMaxRow();
        int maxCol = sheet->getMaxColumn();

        for (int r = 0; r <= maxRow; ++r) {
            bool hasData = false;
            for (int c = 0; c <= maxCol; ++c) {
                auto val = sheet->getCellValue(CellAddress(r, c));
                if (val.isValid() && !val.toString().isEmpty()) { hasData = true; break; }
                auto cell = sheet->getCell(CellAddress(r, c));
                if (cell) {
                    QString key = cellStyleKey(cell->getStyle());
                    if (styleIndexMap.count(key) && styleIndexMap[key] != 0) { hasData = true; break; }
                }
            }
            if (!hasData) continue;

            xml.writeStartElement("row");
            xml.writeAttribute("r", QString::number(r + 1));

            for (int c = 0; c <= maxCol; ++c) {
                CellAddress addr(r, c);
                auto cell = sheet->getCell(addr);
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
                    xml.writeTextElement("f", formula.startsWith('=') ? formula.mid(1) : formula);
                    // Write cached value
                    if (hasValue) {
                        bool ok;
                        double num = val.toDouble(&ok);
                        if (ok) {
                            xml.writeTextElement("v", QString::number(num, 'g', 15));
                        } else {
                            xml.writeAttribute("t", "str");
                            xml.writeTextElement("v", val.toString());
                        }
                    }
                } else if (hasValue) {
                    bool ok;
                    double num = val.toDouble(&ok);
                    if (ok && cell->getType() == CellType::Number) {
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

        xml.writeEndElement(); // worksheet
        xml.writeEndDocument();
        buf.close();
        sheetXmls.push_back(sheetXml);
    }

    // Build the ZIP
    int sheetCount = static_cast<int>(sheets.size());

    QZipWriter zip(filePath);
    if (zip.status() != QZipWriter::NoError) return false;

    zip.addFile("[Content_Types].xml", generateContentTypes(sheetCount));
    zip.addFile("_rels/.rels", generateRels());
    zip.addFile("xl/workbook.xml", generateWorkbook(sheets));
    zip.addFile("xl/_rels/workbook.xml.rels", generateWorkbookRels(sheetCount));
    zip.addFile("xl/styles.xml", generateStyles(sheets, styleIndexMap));

    if (!sharedStrings.isEmpty()) {
        zip.addFile("xl/sharedStrings.xml", generateSharedStrings(sharedStrings));
    }

    for (int i = 0; i < sheetCount; ++i) {
        zip.addFile(QString("xl/worksheets/sheet%1.xml").arg(i + 1), sheetXmls[i]);
    }

    zip.close();
    return true;
}

QByteArray XlsxService::generateContentTypes(int sheetCount) {
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

    xml.writeStartElement("Override");
    xml.writeAttribute("PartName", "/xl/workbook.xml");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    xml.writeEndElement();

    xml.writeStartElement("Override");
    xml.writeAttribute("PartName", "/xl/styles.xml");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    xml.writeEndElement();

    xml.writeStartElement("Override");
    xml.writeAttribute("PartName", "/xl/sharedStrings.xml");
    xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    xml.writeEndElement();

    for (int i = 0; i < sheetCount; ++i) {
        xml.writeStartElement("Override");
        xml.writeAttribute("PartName", QString("/xl/worksheets/sheet%1.xml").arg(i + 1));
        xml.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
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

QByteArray XlsxService::generateWorkbookRels(int sheetCount) {
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

    xml.writeStartElement("Relationship");
    xml.writeAttribute("Id", QString("rId%1").arg(sheetCount + 2));
    xml.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
    xml.writeAttribute("Target", "sharedStrings.xml");
    xml.writeEndElement();

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
    struct BorderEntry { bool hasAny; };

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
            fonts.push_back({s.fontName, s.fontSize, s.bold, s.italic,
                             s.underline, s.strikethrough, s.foregroundColor});
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
        QString bg = s.backgroundColor;
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

    // borders
    xml.writeStartElement("borders");
    xml.writeAttribute("count", "2");
    // Default border (none)
    xml.writeStartElement("border");
    xml.writeEmptyElement("left");
    xml.writeEmptyElement("right");
    xml.writeEmptyElement("top");
    xml.writeEmptyElement("bottom");
    xml.writeEmptyElement("diagonal");
    xml.writeEndElement();
    // Thin border (all sides)
    xml.writeStartElement("border");
    auto writeBorderSide = [&](const QString& side) {
        xml.writeStartElement(side);
        xml.writeAttribute("style", "thin");
        xml.writeStartElement("color"); xml.writeAttribute("auto", "1"); xml.writeEndElement();
        xml.writeEndElement();
    };
    writeBorderSide("left");
    writeBorderSide("right");
    writeBorderSide("top");
    writeBorderSide("bottom");
    xml.writeEmptyElement("diagonal");
    xml.writeEndElement();
    xml.writeEndElement(); // borders

    // cellXfs
    xml.writeStartElement("cellXfs");
    xml.writeAttribute("count", QString::number(sortedStyles.size()));
    for (const auto& s : sortedStyles) {
        xml.writeStartElement("xf");
        // Font
        QString fk = getFontKey(s);
        xml.writeAttribute("fontId", QString::number(fontMap[fk]));
        // Fill
        QString bg = s.backgroundColor;
        int fillId = 0;
        if (bg != "#FFFFFF" && bg != "#ffffff" && fillMap.count(bg)) {
            fillId = fillMap[bg];
        }
        xml.writeAttribute("fillId", QString::number(fillId));
        // Border
        bool hasBorder = s.borderTop.enabled || s.borderBottom.enabled ||
                         s.borderLeft.enabled || s.borderRight.enabled;
        xml.writeAttribute("borderId", hasBorder ? "1" : "0");
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
    xml.writeAttribute("count", QString::number(sharedStrings.size()));
    xml.writeAttribute("uniqueCount", QString::number(sharedStrings.size()));

    for (const auto& s : sharedStrings) {
        xml.writeStartElement("si");
        xml.writeTextElement("t", s);
        xml.writeEndElement();
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

    return chart;
}
