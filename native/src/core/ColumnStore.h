#ifndef COLUMNSTORE_H
#define COLUMNSTORE_H

#include <cstdint>
#include <cstring>
#include <vector>
#include <memory>
#include <functional>
#include <algorithm>
#include <mutex>
#include <shared_mutex>
#include <QString>
#include <QVariant>
#include "Cell.h"
#include "CellRange.h"

// Data type stored in a column chunk cell
enum class CellDataType : uint8_t {
    Empty    = 0,
    Double   = 1,
    String   = 2,  // stored as StringPool ID
    Formula  = 3,  // stored as formula string index
    Boolean  = 4,
    Error    = 5,
    Date     = 6   // stored as double (serial number)
};

// ============================================================================
// ColumnChunk — 64K-row block of a single column
// ============================================================================
// Each chunk represents a contiguous range of rows [baseRow, baseRow + CHUNK_SIZE).
// Uses a presence bitmap to track which rows have data (sparse within the chunk).
// Values are stored in typed contiguous arrays for cache locality and SIMD.
struct ColumnChunk {
    static constexpr int CHUNK_SIZE = 65536;  // 2^16 rows per chunk
    static constexpr int BITMAP_WORDS = CHUNK_SIZE / 64;  // 1024 uint64_t = 8 KB

    int baseRow = 0;

    // Presence bitmap: bit i = 1 means row (baseRow + i) has data
    uint64_t presence[BITMAP_WORDS] = {};

    // Per-populated-cell type (indexed by dense index, not by row offset)
    std::vector<uint8_t> types;

    // Unified value storage — ONE slot per populated cell, indexed by denseIdx (1:1 with types).
    // - Double/Date: stored directly as double
    // - Boolean: 0.0 or 1.0
    // - String: StringPool ID packed via packId()
    // - Error: StringPool ID packed via packId()
    // - Formula: 0.0 (formula text stored in separate sparse map)
    std::vector<double> values;

    // Formula text stored sparsely (keyed by rowOffset, not denseIdx)
    std::unique_ptr<std::unordered_map<int, QString>> formulas;

    // Style index per populated cell: 0 = default, >0 = StyleTable index
    // Only allocated if any cell in this chunk has non-default style (lazy)
    std::unique_ptr<std::vector<uint16_t>> styleIndices;

    // Flags bitmap (3 bits per row): dirty | hasComment | hasHyperlink
    // Only allocated on demand
    std::unique_ptr<uint64_t[]> flagsBitmap;  // 3 * BITMAP_WORDS uint64_t

    // Comment/hyperlink storage (sparse — only for cells that have them)
    std::unique_ptr<std::unordered_map<int, QString>> comments;   // rowOffset → comment
    std::unique_ptr<std::unordered_map<int, QString>> hyperlinks; // rowOffset → url

    // Spill parent tracking (sparse)
    std::unique_ptr<std::unordered_map<int, CellAddress>> spillParents; // rowOffset → parent addr

    int populatedCount = 0;

    // ---- Presence bitmap operations ----
    bool hasData(int rowOffset) const {
        return (presence[rowOffset / 64] >> (rowOffset % 64)) & 1ULL;
    }

    void setPresence(int rowOffset) {
        presence[rowOffset / 64] |= (1ULL << (rowOffset % 64));
    }

    void clearPresence(int rowOffset) {
        presence[rowOffset / 64] &= ~(1ULL << (rowOffset % 64));
    }

    // Dense index: count of set bits before this position
    // This maps a sparse row offset to a contiguous array index
    int denseIndex(int rowOffset) const;

    // ---- Cell operations ----

    // Set a numeric value at the given row offset
    void setNumeric(int rowOffset, double value, uint16_t styleIdx = 0);

    // Set a string value (via StringPool ID)
    void setString(int rowOffset, uint32_t stringId, uint16_t styleIdx = 0);

    // Set a formula
    void setFormula(int rowOffset, const QString& formula, uint16_t styleIdx = 0);

    // Set a boolean value
    void setBoolean(int rowOffset, bool value, uint16_t styleIdx = 0);

    // Set an error value (via StringPool ID)
    void setError(int rowOffset, uint32_t errorId, uint16_t styleIdx = 0);

    // Set a date value (as serial number)
    void setDate(int rowOffset, double serialDate, uint16_t styleIdx = 0);

    // Get the type at a row offset (Empty if not present)
    CellDataType getType(int rowOffset) const;

    // Get numeric value (Double, Date, Boolean)
    double getNumeric(int rowOffset) const;

    // Get string ID
    uint32_t getStringId(int rowOffset) const;

    // Get formula string
    const QString& getFormulaString(int rowOffset) const;

    // Get error ID
    uint32_t getErrorId(int rowOffset) const;

    // Get QVariant value (convenience, slightly slower)
    QVariant getValue(int rowOffset) const;

    // Get style index
    uint16_t getStyleIndex(int rowOffset) const;

    // Set style index
    void setStyleIndex(int rowOffset, uint16_t styleIdx);

    // Remove cell at row offset
    void removeCell(int rowOffset);

    // Clear entire chunk
    void clear();

    // Copy-on-write support: make a deep copy of this chunk
    std::unique_ptr<ColumnChunk> clone() const;

    // Pack/unpack uint32_t ID into double for unified value storage
    static double packId(uint32_t id) {
        double d;
        uint64_t u = id;
        std::memcpy(&d, &u, sizeof(d));
        return d;
    }
    static uint32_t unpackId(double d) {
        uint64_t u;
        std::memcpy(&u, &d, sizeof(u));
        return static_cast<uint32_t>(u);
    }

private:
    // Insert a new slot at the given dense index, shifting subsequent elements
    void insertSlot(int denseIdx, CellDataType type, uint16_t styleIdx);

    // Remove the slot at the given dense index
    void removeSlot(int denseIdx);
};


// ============================================================================
// Column — manages chunks for one column
// ============================================================================
class Column {
public:
    Column() = default;

    // Get chunk containing the given row (nullptr if not created)
    ColumnChunk* getChunk(int row) const;

    // Get or create chunk for the given row
    ColumnChunk* getOrCreateChunk(int row);

    // Get all chunks overlapping [startRow, endRow]
    void getChunksInRange(int startRow, int endRow,
                          std::vector<ColumnChunk*>& out) const;

    // Remove chunk containing the given row (if empty)
    void removeChunkIfEmpty(int row);

    // Total populated cells across all chunks
    int totalPopulated() const;

    // Iterate all chunks
    const std::vector<std::unique_ptr<ColumnChunk>>& chunks() const { return m_chunks; }
    std::vector<std::unique_ptr<ColumnChunk>>& chunks() { return m_chunks; }

    // Reserve capacity
    void reserveChunks(int estimatedRows);

    // Apply permutation (for sorting): reorder values in [startRow, startRow+perm.size())
    void applyPermutation(int startRow, const std::vector<int>& perm);

    // Clear all data
    void clear();

    // Find chunk containing row (range-based, handles non-aligned baseRows)
    int findChunkIndex(int row) const;

private:
    // Chunks sorted by baseRow. Typically sparse (most columns have few chunks).
    std::vector<std::unique_ptr<ColumnChunk>> m_chunks;
};


// ============================================================================
// ColumnStore — the top-level columnar storage engine
// ============================================================================
// Replaces Spreadsheet's `unordered_map<CellKey, shared_ptr<Cell>>`.
// Designed for 20M+ rows with <500 MB memory budget.
class ColumnStore {
public:
    ColumnStore();
    ~ColumnStore();

    // ---- Cell access (single cell) ----

    // Set a cell value (auto-detects type from QVariant)
    void setCellValue(int row, int col, const QVariant& value);

    // Set a numeric value directly (fastest path)
    void setCellNumeric(int row, int col, double value);

    // Set a string value directly (via StringPool)
    void setCellString(int row, int col, const QString& value);

    // Set a formula
    void setCellFormula(int row, int col, const QString& formula);

    // Get cell value as QVariant
    QVariant getCellValue(int row, int col) const;

    // Get cell type
    CellDataType getCellType(int row, int col) const;

    // Check if cell exists (has data)
    bool hasCell(int row, int col) const;

    // Remove a cell
    void removeCell(int row, int col);

    // ---- Style ----

    void setCellStyle(int row, int col, uint16_t styleIndex);
    uint16_t getCellStyleIndex(int row, int col) const;

    // ---- Formula ----

    QString getCellFormula(int row, int col) const;

    // ---- Comments ----

    void setCellComment(int row, int col, const QString& comment);
    QString getCellComment(int row, int col) const;
    bool hasCellComment(int row, int col) const;

    // ---- Hyperlinks ----

    void setCellHyperlink(int row, int col, const QString& url);
    QString getCellHyperlink(int row, int col) const;
    bool hasCellHyperlink(int row, int col) const;

    // ---- Spill tracking ----

    void setSpillParent(int row, int col, const CellAddress& parent);
    CellAddress getSpillParent(int row, int col) const;
    bool isSpillCell(int row, int col) const;
    void clearSpillParent(int row, int col);

    // ---- Computed values (formula results) ----

    void setComputedValue(int row, int col, const QVariant& value);
    QVariant getComputedValue(int row, int col) const;

    // ---- Bulk / streaming operations (the performance-critical path) ----

    // Scan a column range WITHOUT materializing addresses.
    // Callback receives (row, value) for each populated cell in [startRow, endRow].
    // This is the primary API for formula evaluation (SUM, AVERAGE, etc.)
    template<typename Fn>
    void scanColumnValues(int col, int startRow, int endRow, Fn&& fn) const;

    // Scan with type information
    template<typename Fn>
    void scanColumnTyped(int col, int startRow, int endRow, Fn&& fn) const;

    // Fast numeric sum over a column range (SIMD-friendly)
    double sumColumn(int col, int startRow, int endRow) const;

    // Fast count of populated cells in range
    int countColumn(int col, int startRow, int endRow) const;

    // ---- Iteration ----

    // Iterate all populated cells (for serialization)
    void forEachCell(std::function<void(int row, int col, CellDataType type,
                                        const QVariant& value)> callback) const;

    // ---- Navigation (O(1) per chunk, no index building) ----

    // First occupied row >= startRow in column, or -1 if none
    int nextOccupiedRow(int col, int startRow) const;
    // Last occupied row <= startRow in column, or -1 if none
    int prevOccupiedRow(int col, int startRow) const;
    // First empty row >= startRow in column (always succeeds)
    int nextEmptyRow(int col, int startRow) const;
    // Last empty row <= startRow in column (always succeeds, min 0)
    int prevEmptyRow(int col, int startRow) const;

    // ---- Column access ----

    Column* getColumn(int col) const;
    Column* getOrCreateColumn(int col);
    int columnCount() const { return static_cast<int>(m_columns.size()); }

    // ---- Bulk operations ----

    // Reserve capacity for estimated cell count
    void reserve(int estimatedRows, int estimatedCols);

    // Apply sort permutation to all columns in range
    void applyPermutation(int startRow, int startCol, int endCol,
                          const std::vector<int>& perm);

    // Clear all data
    void clear();

    // ---- Dimensions ----

    // Find the extent of data (max row/col with data)
    int maxRow() const;
    int maxCol() const;

    // ---- Thread safety ----

    // Lock for exclusive write access to a specific column
    std::shared_mutex& columnMutex(int col);

    // Global read lock (for operations that span multiple columns)
    mutable std::shared_mutex m_globalMutex;

private:
    std::vector<std::unique_ptr<Column>> m_columns;

    // Per-column mutexes for fine-grained locking
    std::vector<std::unique_ptr<std::shared_mutex>> m_columnMutexes;

    // Ensure column exists at index
    void ensureColumn(int col);

    // Computed value storage (separate from source data)
    // Key: (row << 20) | col — supports up to 1M columns
    std::unordered_map<uint64_t, QVariant> m_computedValues;
    mutable std::shared_mutex m_computedMutex;

    static uint64_t computedKey(int row, int col) {
        return (static_cast<uint64_t>(row) << 20) | static_cast<uint64_t>(col);
    }
};


// ============================================================================
// Template implementations
// ============================================================================

template<typename Fn>
void ColumnStore::scanColumnValues(int col, int startRow, int endRow, Fn&& fn) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col])
        return;

    const auto& column = *m_columns[col];
    std::vector<ColumnChunk*> chunks;
    column.getChunksInRange(startRow, endRow, chunks);

    for (const auto* chunk : chunks) {
        int chunkStart = std::max(0, startRow - chunk->baseRow);
        int chunkEnd = std::min(ColumnChunk::CHUNK_SIZE - 1, endRow - chunk->baseRow);

        for (int offset = chunkStart; offset <= chunkEnd; ++offset) {
            if (chunk->hasData(offset)) {
                int row = chunk->baseRow + offset;
                QVariant val = chunk->getValue(offset);
                fn(row, val);
            }
        }
    }
}

template<typename Fn>
void ColumnStore::scanColumnTyped(int col, int startRow, int endRow, Fn&& fn) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col])
        return;

    const auto& column = *m_columns[col];
    std::vector<ColumnChunk*> chunks;
    column.getChunksInRange(startRow, endRow, chunks);

    for (const auto* chunk : chunks) {
        int chunkStart = std::max(0, startRow - chunk->baseRow);
        int chunkEnd = std::min(ColumnChunk::CHUNK_SIZE - 1, endRow - chunk->baseRow);

        for (int offset = chunkStart; offset <= chunkEnd; ++offset) {
            if (chunk->hasData(offset)) {
                int row = chunk->baseRow + offset;
                CellDataType type = chunk->getType(offset);
                fn(row, type, chunk->getValue(offset));
            }
        }
    }
}

#endif // COLUMNSTORE_H
