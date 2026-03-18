#include "DocumentRepository.h"
#include "DatabaseManager.h"
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QUuid>
#include <QDateTime>
#include <sqlite3.h>

DocumentRepository& DocumentRepository::instance() {
    static DocumentRepository s_instance;
    return s_instance;
}

DocumentRepository::DocumentRepository() {
}

bool DocumentRepository::createDocument(const QString& name, std::shared_ptr<Spreadsheet> spreadsheet) {
    QString id = QUuid::createUuid().toString();
    
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "INSERT INTO documents (id, name, content) VALUES (?, ?, ?)";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return false;
    }

    QJsonObject jsonContent = serializeSpreadsheet(spreadsheet);
    QJsonDocument doc(jsonContent);
    QByteArray jsonData = doc.toJson(QJsonDocument::Compact);

    sqlite3_bind_text(stmt, 1, id.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 2, name.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_blob(stmt, 3, jsonData.constData(), jsonData.size(), SQLITE_TRANSIENT);

    if (sqlite3_step(stmt) != SQLITE_DONE) {
        m_lastError = sqlite3_errmsg(db);
        sqlite3_finalize(stmt);
        return false;
    }

    sqlite3_finalize(stmt);
    return true;
}

std::shared_ptr<Document> DocumentRepository::getDocument(const QString& id) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return nullptr;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "SELECT id, name, createdAt, updatedAt, content FROM documents WHERE id = ?";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return nullptr;
    }

    sqlite3_bind_text(stmt, 1, id.toStdString().c_str(), -1, SQLITE_TRANSIENT);

    if (sqlite3_step(stmt) != SQLITE_ROW) {
        sqlite3_finalize(stmt);
        return nullptr;
    }

    auto doc = std::make_shared<Document>();
    doc->id = id;
    doc->name = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1)));
    doc->createdAt = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2)));
    doc->updatedAt = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 3)));

    const void* blobData = sqlite3_column_blob(stmt, 4);
    int blobSize = sqlite3_column_bytes(stmt, 4);
    QByteArray jsonData(static_cast<const char*>(blobData), blobSize);
    
    QJsonDocument jsonDoc = QJsonDocument::fromJson(jsonData);
    doc->spreadsheet = deserializeSpreadsheet(jsonDoc.object());

    sqlite3_finalize(stmt);
    return doc;
}

QVector<std::shared_ptr<Document>> DocumentRepository::getAllDocuments() {
    QVector<std::shared_ptr<Document>> documents;

    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return documents;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "SELECT id, name, createdAt, updatedAt FROM documents ORDER BY updatedAt DESC";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return documents;
    }

    while (sqlite3_step(stmt) == SQLITE_ROW) {
        auto doc = std::make_shared<Document>();
        doc->id = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 0)));
        doc->name = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1)));
        doc->createdAt = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 2)));
        doc->updatedAt = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 3)));
        documents.append(doc);
    }

    sqlite3_finalize(stmt);
    return documents;
}

bool DocumentRepository::updateDocument(const QString& id, const QString& name, std::shared_ptr<Spreadsheet> spreadsheet) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "UPDATE documents SET name = ?, content = ?, updatedAt = CURRENT_TIMESTAMP WHERE id = ?";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return false;
    }

    QJsonObject jsonContent = serializeSpreadsheet(spreadsheet);
    QJsonDocument doc(jsonContent);
    QByteArray jsonData = doc.toJson(QJsonDocument::Compact);

    sqlite3_bind_text(stmt, 1, name.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_blob(stmt, 2, jsonData.constData(), jsonData.size(), SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 3, id.toStdString().c_str(), -1, SQLITE_TRANSIENT);

    bool success = sqlite3_step(stmt) == SQLITE_DONE;
    if (!success) {
        m_lastError = sqlite3_errmsg(db);
    }

    sqlite3_finalize(stmt);
    return success;
}

bool DocumentRepository::deleteDocument(const QString& id) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "DELETE FROM documents WHERE id = ?";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return false;
    }

    sqlite3_bind_text(stmt, 1, id.toStdString().c_str(), -1, SQLITE_TRANSIENT);

    bool success = sqlite3_step(stmt) == SQLITE_DONE;
    if (!success) {
        m_lastError = sqlite3_errmsg(db);
    }

    sqlite3_finalize(stmt);
    return success;
}

bool DocumentRepository::addSheet(const QString& documentId, const QString& sheetName, int index) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    std::string id = QUuid::createUuid().toString().toStdString();
    
    const char* sql = "INSERT INTO sheets (id, documentId, name, index) VALUES (?, ?, ?, ?)";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return false;
    }

    sqlite3_bind_text(stmt, 1, id.c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 2, documentId.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 3, sheetName.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 4, index);

    bool success = sqlite3_step(stmt) == SQLITE_DONE;
    if (!success) {
        m_lastError = sqlite3_errmsg(db);
    }

    sqlite3_finalize(stmt);
    return success;
}

bool DocumentRepository::removeSheet(const QString& documentId, int index) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = "DELETE FROM sheets WHERE documentId = ? AND index = ?";
    
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = sqlite3_errmsg(db);
        return false;
    }

    sqlite3_bind_text(stmt, 1, documentId.toStdString().c_str(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, index);

    bool success = sqlite3_step(stmt) == SQLITE_DONE;
    if (!success) {
        m_lastError = sqlite3_errmsg(db);
    }

    sqlite3_finalize(stmt);
    return success;
}

bool DocumentRepository::saveDocument(const QString& id) {
    // Implementation would save to database
    return true;
}

bool DocumentRepository::loadDocument(const QString& id) {
    // Implementation would load from database
    return true;
}

bool DocumentRepository::saveVersion(const QString& documentId) {
    // Implementation would save version history
    return true;
}

QVector<std::shared_ptr<Document>> DocumentRepository::getVersionHistory(const QString& documentId) {
    // Implementation would retrieve version history
    return QVector<std::shared_ptr<Document>>();
}

bool DocumentRepository::restoreVersion(const QString& documentId, const QString& versionId) {
    // Implementation would restore a specific version
    return true;
}

QString DocumentRepository::getLastError() const {
    return m_lastError;
}

QJsonObject DocumentRepository::serializeSpreadsheet(const std::shared_ptr<Spreadsheet>& spreadsheet) {
    QJsonObject json;
    json["name"] = spreadsheet->getSheetName();
    json["maxRow"] = spreadsheet->getMaxRow();
    json["maxCol"] = spreadsheet->getMaxColumn();

    QJsonArray cellsArray;
    spreadsheet->forEachCell([&](int row, int col, const Cell& cell) {
        QJsonObject cellObj;
        cellObj["r"] = row;
        cellObj["c"] = col;
        cellObj["t"] = static_cast<int>(cell.getType());

        if (cell.getType() == CellType::Formula) {
            cellObj["f"] = cell.getFormula();
        } else {
            QVariant val = cell.getValue();
            if (val.typeId() == QMetaType::Double) {
                cellObj["v"] = val.toDouble();
            } else if (val.typeId() == QMetaType::Int) {
                cellObj["v"] = val.toInt();
            } else if (val.typeId() == QMetaType::Bool) {
                cellObj["v"] = val.toBool();
            } else {
                cellObj["v"] = val.toString();
            }
        }

        // Serialize style if non-default
        const CellStyle& style = cell.getStyle();
        QJsonObject styleObj;
        bool hasCustomStyle = false;

        if (style.bold) { styleObj["b"] = true; hasCustomStyle = true; }
        if (style.italic) { styleObj["i"] = true; hasCustomStyle = true; }
        if (style.underline) { styleObj["u"] = true; hasCustomStyle = true; }
        if (style.fontName != "Arial") { styleObj["fn"] = style.fontName; hasCustomStyle = true; }
        if (style.fontSize != 11) { styleObj["fs"] = style.fontSize; hasCustomStyle = true; }
        if (style.foregroundColor != "#000000") { styleObj["fg"] = style.foregroundColor; hasCustomStyle = true; }
        if (style.backgroundColor != "#FFFFFF") { styleObj["bg"] = style.backgroundColor; hasCustomStyle = true; }
        if (style.hAlign != HorizontalAlignment::Left) { styleObj["ha"] = static_cast<int>(style.hAlign); hasCustomStyle = true; }
        if (style.vAlign != VerticalAlignment::Middle) { styleObj["va"] = static_cast<int>(style.vAlign); hasCustomStyle = true; }
        if (style.numberFormat != "General") { styleObj["nf"] = style.numberFormat; hasCustomStyle = true; }

        if (hasCustomStyle) {
            cellObj["s"] = styleObj;
        }

        cellsArray.append(cellObj);
    });
    json["cells"] = cellsArray;

    return json;
}

std::shared_ptr<Spreadsheet> DocumentRepository::deserializeSpreadsheet(const QJsonObject& json) {
    auto spreadsheet = std::make_shared<Spreadsheet>();

    if (json.contains("name")) {
        spreadsheet->setSheetName(json["name"].toString());
    }

    spreadsheet->setAutoRecalculate(false);

    if (json.contains("cells")) {
        QJsonArray cellsArray = json["cells"].toArray();
        for (const auto& cellVal : cellsArray) {
            QJsonObject cellObj = cellVal.toObject();
            int row = cellObj["r"].toInt();
            int col = cellObj["c"].toInt();
            CellAddress addr(row, col);

            int typeInt = cellObj["t"].toInt();
            CellType type = static_cast<CellType>(typeInt);

            if (type == CellType::Formula) {
                spreadsheet->setCellFormula(addr, cellObj["f"].toString());
            } else {
                QJsonValue v = cellObj["v"];
                if (v.isDouble()) {
                    spreadsheet->setCellValue(addr, v.toDouble());
                } else if (v.isBool()) {
                    spreadsheet->setCellValue(addr, v.toBool());
                } else {
                    spreadsheet->setCellValue(addr, v.toString());
                }
            }

            // Restore style
            if (cellObj.contains("s")) {
                QJsonObject styleObj = cellObj["s"].toObject();
                auto cell = spreadsheet->getCell(addr);
                CellStyle style = cell->getStyle();

                if (styleObj.contains("b")) style.bold = styleObj["b"].toBool();
                if (styleObj.contains("i")) style.italic = styleObj["i"].toBool();
                if (styleObj.contains("u")) style.underline = styleObj["u"].toBool();
                if (styleObj.contains("fn")) style.fontName = styleObj["fn"].toString();
                if (styleObj.contains("fs")) style.fontSize = styleObj["fs"].toInt();
                if (styleObj.contains("fg")) style.foregroundColor = styleObj["fg"].toString();
                if (styleObj.contains("bg")) style.backgroundColor = styleObj["bg"].toString();
                if (styleObj.contains("ha")) style.hAlign = static_cast<HorizontalAlignment>(styleObj["ha"].toInt());
                if (styleObj.contains("va")) style.vAlign = static_cast<VerticalAlignment>(styleObj["va"].toInt());
                if (styleObj.contains("nf")) style.numberFormat = styleObj["nf"].toString();

                cell->setStyle(style);
            }
        }
    }

    spreadsheet->setAutoRecalculate(true);

    return spreadsheet;
}

// ============================================================================
// Chunk-level storage methods
// ============================================================================

bool DocumentRepository::saveChunk(const QString& docId, int sheetIdx, int col, int chunkId,
                                    const QByteArray& chunkData, int rowCount) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        INSERT OR REPLACE INTO cell_chunks (doc_id, sheet_idx, col, chunk_id, chunk_data, row_count)
        VALUES (?, ?, ?, ?, ?, ?)
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
        return false;
    }

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, sheetIdx);
    sqlite3_bind_int(stmt, 3, col);
    sqlite3_bind_int(stmt, 4, chunkId);
    sqlite3_bind_blob(stmt, 5, chunkData.constData(), chunkData.size(), SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 6, rowCount);

    bool success = (sqlite3_step(stmt) == SQLITE_DONE);
    if (!success) m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
    sqlite3_finalize(stmt);
    return success;
}

QByteArray DocumentRepository::loadChunk(const QString& docId, int sheetIdx, int col, int chunkId) {
    if (!DatabaseManager::instance().isInitialized()) return {};

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        SELECT chunk_data FROM cell_chunks
        WHERE doc_id = ? AND sheet_idx = ? AND col = ? AND chunk_id = ?
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) return {};

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, sheetIdx);
    sqlite3_bind_int(stmt, 3, col);
    sqlite3_bind_int(stmt, 4, chunkId);

    QByteArray result;
    if (sqlite3_step(stmt) == SQLITE_ROW) {
        const void* blob = sqlite3_column_blob(stmt, 0);
        int blobSize = sqlite3_column_bytes(stmt, 0);
        result = QByteArray(static_cast<const char*>(blob), blobSize);
    }
    sqlite3_finalize(stmt);
    return result;
}

bool DocumentRepository::saveSheetMeta(const QString& docId, int sheetIdx, const QString& name,
                                        int rowCount, int colCount, const QByteArray& stylesBlob,
                                        const QString& settingsJson) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        INSERT OR REPLACE INTO sheet_meta (doc_id, sheet_idx, name, row_count, col_count, styles_blob, settings_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
        return false;
    }

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, sheetIdx);
    sqlite3_bind_text(stmt, 3, name.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 4, rowCount);
    sqlite3_bind_int(stmt, 5, colCount);
    sqlite3_bind_blob(stmt, 6, stylesBlob.constData(), stylesBlob.size(), SQLITE_TRANSIENT);
    sqlite3_bind_text(stmt, 7, settingsJson.toUtf8().constData(), -1, SQLITE_TRANSIENT);

    bool success = (sqlite3_step(stmt) == SQLITE_DONE);
    if (!success) m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
    sqlite3_finalize(stmt);
    return success;
}

DocumentRepository::SheetMeta DocumentRepository::loadSheetMeta(const QString& docId, int sheetIdx) {
    SheetMeta meta;
    if (!DatabaseManager::instance().isInitialized()) return meta;

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        SELECT name, row_count, col_count, styles_blob, settings_json FROM sheet_meta
        WHERE doc_id = ? AND sheet_idx = ?
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) return meta;

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, sheetIdx);

    if (sqlite3_step(stmt) == SQLITE_ROW) {
        meta.name = QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_column_text(stmt, 0)));
        meta.rowCount = sqlite3_column_int(stmt, 1);
        meta.colCount = sqlite3_column_int(stmt, 2);
        const void* blob = sqlite3_column_blob(stmt, 3);
        int blobSize = sqlite3_column_bytes(stmt, 3);
        meta.stylesBlob = QByteArray(static_cast<const char*>(blob), blobSize);
        const char* json = reinterpret_cast<const char*>(sqlite3_column_text(stmt, 4));
        if (json) meta.settingsJson = QString::fromUtf8(json);
    }
    sqlite3_finalize(stmt);
    return meta;
}

QVector<DocumentRepository::ChunkInfo> DocumentRepository::getChunkList(const QString& docId) {
    QVector<ChunkInfo> chunks;
    if (!DatabaseManager::instance().isInitialized()) return chunks;

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        SELECT sheet_idx, col, chunk_id, row_count FROM cell_chunks
        WHERE doc_id = ? ORDER BY sheet_idx, col, chunk_id
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) return chunks;

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);

    while (sqlite3_step(stmt) == SQLITE_ROW) {
        ChunkInfo ci;
        ci.sheetIdx = sqlite3_column_int(stmt, 0);
        ci.col = sqlite3_column_int(stmt, 1);
        ci.chunkId = sqlite3_column_int(stmt, 2);
        ci.rowCount = sqlite3_column_int(stmt, 3);
        chunks.push_back(ci);
    }
    sqlite3_finalize(stmt);
    return chunks;
}

bool DocumentRepository::saveVersionDelta(const QString& docId, int versionId, int sheetIdx,
                                            int col, int chunkId, const QByteArray& oldChunkData) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();
    const char* sql = R"(
        INSERT INTO version_deltas (doc_id, version_id, sheet_idx, col, chunk_id, old_chunk_data)
        VALUES (?, ?, ?, ?, ?, ?)
    )";

    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
        return false;
    }

    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(stmt, 2, versionId);
    sqlite3_bind_int(stmt, 3, sheetIdx);
    sqlite3_bind_int(stmt, 4, col);
    sqlite3_bind_int(stmt, 5, chunkId);
    sqlite3_bind_blob(stmt, 6, oldChunkData.constData(), oldChunkData.size(), SQLITE_TRANSIENT);

    bool success = (sqlite3_step(stmt) == SQLITE_DONE);
    if (!success) m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
    sqlite3_finalize(stmt);
    return success;
}

bool DocumentRepository::deleteChunks(const QString& docId) {
    if (!DatabaseManager::instance().isInitialized()) {
        m_lastError = "Database not initialized";
        return false;
    }

    sqlite3* db = DatabaseManager::instance().getDatabase();

    // Delete chunks
    const char* sql1 = "DELETE FROM cell_chunks WHERE doc_id = ?";
    sqlite3_stmt* stmt;
    if (sqlite3_prepare_v2(db, sql1, -1, &stmt, nullptr) != SQLITE_OK) {
        m_lastError = QString::fromUtf8(sqlite3_errmsg(db));
        return false;
    }
    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_step(stmt);
    sqlite3_finalize(stmt);

    // Delete sheet meta
    const char* sql2 = "DELETE FROM sheet_meta WHERE doc_id = ?";
    if (sqlite3_prepare_v2(db, sql2, -1, &stmt, nullptr) != SQLITE_OK) return false;
    sqlite3_bind_text(stmt, 1, docId.toUtf8().constData(), -1, SQLITE_TRANSIENT);
    sqlite3_step(stmt);
    sqlite3_finalize(stmt);

    return true;
}
