#ifndef DOCUMENTREPOSITORY_H
#define DOCUMENTREPOSITORY_H

#include <QString>
#include <QVector>
#include <QJsonObject>
#include <memory>
#include "../core/Spreadsheet.h"

struct Document {
    QString id;
    QString name;
    QString createdAt;
    QString updatedAt;
    std::shared_ptr<Spreadsheet> spreadsheet;
};

class DocumentRepository {
public:
    static DocumentRepository& instance();

    // CRUD operations
    bool createDocument(const QString& name, std::shared_ptr<Spreadsheet> spreadsheet);
    std::shared_ptr<Document> getDocument(const QString& id);
    QVector<std::shared_ptr<Document>> getAllDocuments();
    bool updateDocument(const QString& id, const QString& name, std::shared_ptr<Spreadsheet> spreadsheet);
    bool deleteDocument(const QString& id);

    // Sheet operations
    bool addSheet(const QString& documentId, const QString& sheetName, int index);
    bool removeSheet(const QString& documentId, int index);

    // Save/Load operations
    bool saveDocument(const QString& id);
    bool loadDocument(const QString& id);

    // Version control
    bool saveVersion(const QString& documentId);
    QVector<std::shared_ptr<Document>> getVersionHistory(const QString& documentId);
    bool restoreVersion(const QString& documentId, const QString& versionId);

    // ---- Chunk-level storage (for 20M+ row scalability) ----

    // Save a single column chunk as a compressed blob
    bool saveChunk(const QString& docId, int sheetIdx, int col, int chunkId,
                   const QByteArray& chunkData, int rowCount);

    // Load a single column chunk
    QByteArray loadChunk(const QString& docId, int sheetIdx, int col, int chunkId);

    // Save sheet metadata (name, dimensions, styles, settings)
    bool saveSheetMeta(const QString& docId, int sheetIdx, const QString& name,
                       int rowCount, int colCount, const QByteArray& stylesBlob,
                       const QString& settingsJson);

    // Load sheet metadata
    struct SheetMeta {
        QString name;
        int rowCount = 0;
        int colCount = 0;
        QByteArray stylesBlob;
        QString settingsJson;
    };
    SheetMeta loadSheetMeta(const QString& docId, int sheetIdx);

    // Get list of all chunks for a document (for loading)
    struct ChunkInfo {
        int sheetIdx;
        int col;
        int chunkId;
        int rowCount;
    };
    QVector<ChunkInfo> getChunkList(const QString& docId);

    // Save a version delta (old chunk data before modification)
    bool saveVersionDelta(const QString& docId, int versionId, int sheetIdx,
                          int col, int chunkId, const QByteArray& oldChunkData);

    // Delete all chunks for a document
    bool deleteChunks(const QString& docId);

    QString getLastError() const;

private:
    DocumentRepository();
    ~DocumentRepository() = default;

    QString m_lastError;

    QJsonObject serializeSpreadsheet(const std::shared_ptr<Spreadsheet>& spreadsheet);
    std::shared_ptr<Spreadsheet> deserializeSpreadsheet(const QJsonObject& json);
};

#endif // DOCUMENTREPOSITORY_H
