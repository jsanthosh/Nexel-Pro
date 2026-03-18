#include "ColumnStore.h"
#include "StringPool.h"
#include <numeric>   // std::iota
#include <cstring>

// Cross-platform bit intrinsics
#ifdef _MSC_VER
#include <intrin.h>
static inline int ctzll(uint64_t x) {
    unsigned long idx;
    _BitScanForward64(&idx, x);
    return static_cast<int>(idx);
}
static inline int clzll(uint64_t x) {
    unsigned long idx;
    _BitScanReverse64(&idx, x);
    return 63 - static_cast<int>(idx);
}
static inline int popcountll(uint64_t x) {
    return static_cast<int>(__popcnt64(x));
}
#else
static inline int ctzll(uint64_t x) { return __builtin_ctzll(x); }
static inline int clzll(uint64_t x) { return __builtin_clzll(x); }
static inline int popcountll(uint64_t x) { return __builtin_popcountll(x); }
#endif

// ============================================================================
// ColumnChunk — popcount-based dense indexing
// ============================================================================

int ColumnChunk::denseIndex(int rowOffset) const {
    int wordIdx = rowOffset / 64;
    int bitIdx  = rowOffset % 64;

    int count = 0;
    for (int w = 0; w < wordIdx; ++w) {
        count += popcountll(presence[w]);
    }
    if (bitIdx > 0) {
        uint64_t mask = (1ULL << bitIdx) - 1;
        count += popcountll(presence[wordIdx] & mask);
    }
    return count;
}

// ---- Insert / Remove slots (O(1) index, amortized O(1) append) ----

void ColumnChunk::insertSlot(int denseIdx, CellDataType type, uint16_t styleIdx) {
    types.insert(types.begin() + denseIdx, static_cast<uint8_t>(type));
    values.insert(values.begin() + denseIdx, 0.0);

    if (styleIdx != 0 || styleIndices) {
        if (!styleIndices) {
            styleIndices = std::make_unique<std::vector<uint16_t>>(populatedCount, 0);
        }
        styleIndices->insert(styleIndices->begin() + denseIdx, styleIdx);
    }

    ++populatedCount;
}

void ColumnChunk::removeSlot(int denseIdx) {
    if (denseIdx < 0 || denseIdx >= populatedCount) return;

    // Clean up formula if this was a formula cell
    auto type = static_cast<CellDataType>(types[denseIdx]);
    if (type == CellDataType::Formula && formulas) {
        // We need to find the rowOffset for this denseIdx to remove from formula map
        // Scan presence bitmap to find the rowOffset corresponding to denseIdx
        int count = 0;
        for (int offset = 0; offset < CHUNK_SIZE && count <= denseIdx; ++offset) {
            if (hasData(offset)) {
                if (count == denseIdx) {
                    formulas->erase(offset);
                    break;
                }
                ++count;
            }
        }
    }

    types.erase(types.begin() + denseIdx);
    values.erase(values.begin() + denseIdx);

    if (styleIndices && denseIdx < static_cast<int>(styleIndices->size())) {
        styleIndices->erase(styleIndices->begin() + denseIdx);
    }

    --populatedCount;
}

// ---- Setters ----

void ColumnChunk::setNumeric(int rowOffset, double value, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        auto oldType = static_cast<CellDataType>(types[idx]);
        if (oldType == CellDataType::Formula && formulas) {
            formulas->erase(rowOffset);
        }
        types[idx] = static_cast<uint8_t>(CellDataType::Double);
        values[idx] = value;
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::Double, styleIdx);
        values[idx] = value;
    }
}

void ColumnChunk::setString(int rowOffset, uint32_t stringId, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        auto oldType = static_cast<CellDataType>(types[idx]);
        if (oldType == CellDataType::Formula && formulas) {
            formulas->erase(rowOffset);
        }
        types[idx] = static_cast<uint8_t>(CellDataType::String);
        values[idx] = packId(stringId);
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::String, styleIdx);
        values[idx] = packId(stringId);
    }
}

void ColumnChunk::setFormula(int rowOffset, const QString& formula, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        types[idx] = static_cast<uint8_t>(CellDataType::Formula);
        values[idx] = 0.0;
        if (!formulas) {
            formulas = std::make_unique<std::unordered_map<int, QString>>();
        }
        (*formulas)[rowOffset] = formula;
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::Formula, styleIdx);
        if (!formulas) {
            formulas = std::make_unique<std::unordered_map<int, QString>>();
        }
        (*formulas)[rowOffset] = formula;
    }
}

void ColumnChunk::setBoolean(int rowOffset, bool value, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        auto oldType = static_cast<CellDataType>(types[idx]);
        if (oldType == CellDataType::Formula && formulas) {
            formulas->erase(rowOffset);
        }
        types[idx] = static_cast<uint8_t>(CellDataType::Boolean);
        values[idx] = value ? 1.0 : 0.0;
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::Boolean, styleIdx);
        values[idx] = value ? 1.0 : 0.0;
    }
}

void ColumnChunk::setError(int rowOffset, uint32_t errorId, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        auto oldType = static_cast<CellDataType>(types[idx]);
        if (oldType == CellDataType::Formula && formulas) {
            formulas->erase(rowOffset);
        }
        types[idx] = static_cast<uint8_t>(CellDataType::Error);
        values[idx] = packId(errorId);
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::Error, styleIdx);
        values[idx] = packId(errorId);
    }
}

void ColumnChunk::setDate(int rowOffset, double serialDate, uint16_t styleIdx) {
    if (hasData(rowOffset)) {
        int idx = denseIndex(rowOffset);
        auto oldType = static_cast<CellDataType>(types[idx]);
        if (oldType == CellDataType::Formula && formulas) {
            formulas->erase(rowOffset);
        }
        types[idx] = static_cast<uint8_t>(CellDataType::Date);
        values[idx] = serialDate;
        setStyleIndex(rowOffset, styleIdx);
    } else {
        setPresence(rowOffset);
        int idx = denseIndex(rowOffset);
        insertSlot(idx, CellDataType::Date, styleIdx);
        values[idx] = serialDate;
    }
}

// ---- Getters ----

CellDataType ColumnChunk::getType(int rowOffset) const {
    if (!hasData(rowOffset)) return CellDataType::Empty;
    int idx = denseIndex(rowOffset);
    return static_cast<CellDataType>(types[idx]);
}

double ColumnChunk::getNumeric(int rowOffset) const {
    if (!hasData(rowOffset)) return 0.0;
    int idx = denseIndex(rowOffset);
    return values[idx];
}

uint32_t ColumnChunk::getStringId(int rowOffset) const {
    if (!hasData(rowOffset)) return 0;
    int idx = denseIndex(rowOffset);
    return unpackId(values[idx]);
}

const QString& ColumnChunk::getFormulaString(int rowOffset) const {
    static const QString empty;
    if (!formulas) return empty;
    auto it = formulas->find(rowOffset);
    return (it != formulas->end()) ? it->second : empty;
}

uint32_t ColumnChunk::getErrorId(int rowOffset) const {
    if (!hasData(rowOffset)) return 0;
    int idx = denseIndex(rowOffset);
    return unpackId(values[idx]);
}

QVariant ColumnChunk::getValue(int rowOffset) const {
    if (!hasData(rowOffset)) return QVariant();
    int idx = denseIndex(rowOffset);
    auto type = static_cast<CellDataType>(types[idx]);

    switch (type) {
        case CellDataType::Double:
        case CellDataType::Date:
            return QVariant(values[idx]);
        case CellDataType::Boolean:
            return QVariant(values[idx] != 0.0);
        case CellDataType::String:
            return QVariant(StringPool::instance().get(unpackId(values[idx])));
        case CellDataType::Formula: {
            if (formulas) {
                auto it = formulas->find(rowOffset);
                if (it != formulas->end()) return QVariant(it->second);
            }
            return QVariant();
        }
        case CellDataType::Error:
            return QVariant(StringPool::instance().get(unpackId(values[idx])));
        default:
            return QVariant();
    }
}

uint16_t ColumnChunk::getStyleIndex(int rowOffset) const {
    if (!hasData(rowOffset) || !styleIndices) return 0;
    int idx = denseIndex(rowOffset);
    if (idx >= static_cast<int>(styleIndices->size())) return 0;
    return (*styleIndices)[idx];
}

void ColumnChunk::setStyleIndex(int rowOffset, uint16_t styleIdx) {
    if (!hasData(rowOffset)) return;
    int idx = denseIndex(rowOffset);

    if (styleIdx == 0 && !styleIndices) return;

    if (!styleIndices) {
        styleIndices = std::make_unique<std::vector<uint16_t>>(populatedCount, 0);
    }
    if (idx < static_cast<int>(styleIndices->size())) {
        (*styleIndices)[idx] = styleIdx;
    }
}

void ColumnChunk::removeCell(int rowOffset) {
    if (!hasData(rowOffset)) return;
    int idx = denseIndex(rowOffset);

    // Clean up formula
    if (static_cast<CellDataType>(types[idx]) == CellDataType::Formula && formulas) {
        formulas->erase(rowOffset);
    }

    // Remove from dense arrays
    types.erase(types.begin() + idx);
    values.erase(values.begin() + idx);

    if (styleIndices && idx < static_cast<int>(styleIndices->size())) {
        styleIndices->erase(styleIndices->begin() + idx);
    }

    --populatedCount;
    clearPresence(rowOffset);

    // Clean up optional data
    if (comments) comments->erase(rowOffset);
    if (hyperlinks) hyperlinks->erase(rowOffset);
    if (spillParents) spillParents->erase(rowOffset);
}

void ColumnChunk::clear() {
    std::memset(presence, 0, sizeof(presence));
    types.clear();
    values.clear();
    formulas.reset();
    styleIndices.reset();
    flagsBitmap.reset();
    comments.reset();
    hyperlinks.reset();
    spillParents.reset();
    populatedCount = 0;
}

std::unique_ptr<ColumnChunk> ColumnChunk::clone() const {
    auto copy = std::make_unique<ColumnChunk>();
    copy->baseRow = baseRow;
    std::memcpy(copy->presence, presence, sizeof(presence));
    copy->types = types;
    copy->values = values;
    copy->populatedCount = populatedCount;

    if (formulas)
        copy->formulas = std::make_unique<std::unordered_map<int, QString>>(*formulas);
    if (styleIndices)
        copy->styleIndices = std::make_unique<std::vector<uint16_t>>(*styleIndices);
    if (flagsBitmap) {
        size_t sz = 3 * BITMAP_WORDS;
        copy->flagsBitmap = std::make_unique<uint64_t[]>(sz);
        std::memcpy(copy->flagsBitmap.get(), flagsBitmap.get(), sz * sizeof(uint64_t));
    }
    if (comments)
        copy->comments = std::make_unique<std::unordered_map<int, QString>>(*comments);
    if (hyperlinks)
        copy->hyperlinks = std::make_unique<std::unordered_map<int, QString>>(*hyperlinks);
    if (spillParents)
        copy->spillParents = std::make_unique<std::unordered_map<int, CellAddress>>(*spillParents);

    return copy;
}


// ============================================================================
// Column
// ============================================================================

int Column::findChunkIndex(int row) const {
    int chunkBase = (row / ColumnChunk::CHUNK_SIZE) * ColumnChunk::CHUNK_SIZE;

    int lo = 0, hi = static_cast<int>(m_chunks.size()) - 1;
    while (lo <= hi) {
        int mid = lo + (hi - lo) / 2;
        if (m_chunks[mid]->baseRow == chunkBase) return mid;
        if (m_chunks[mid]->baseRow < chunkBase) lo = mid + 1;
        else hi = mid - 1;
    }
    return -1;
}

ColumnChunk* Column::getChunk(int row) const {
    int idx = findChunkIndex(row);
    return (idx >= 0) ? m_chunks[idx].get() : nullptr;
}

ColumnChunk* Column::getOrCreateChunk(int row) {
    int chunkBase = (row / ColumnChunk::CHUNK_SIZE) * ColumnChunk::CHUNK_SIZE;

    auto it = std::lower_bound(m_chunks.begin(), m_chunks.end(), chunkBase,
        [](const std::unique_ptr<ColumnChunk>& chunk, int base) {
            return chunk->baseRow < base;
        });

    if (it != m_chunks.end() && (*it)->baseRow == chunkBase) {
        return it->get();
    }

    auto chunk = std::make_unique<ColumnChunk>();
    chunk->baseRow = chunkBase;
    auto* ptr = chunk.get();
    m_chunks.insert(it, std::move(chunk));
    return ptr;
}

void Column::getChunksInRange(int startRow, int endRow,
                               std::vector<ColumnChunk*>& out) const {
    out.clear();
    int startBase = (startRow / ColumnChunk::CHUNK_SIZE) * ColumnChunk::CHUNK_SIZE;

    auto it = std::lower_bound(m_chunks.begin(), m_chunks.end(), startBase,
        [](const std::unique_ptr<ColumnChunk>& chunk, int base) {
            return chunk->baseRow < base;
        });

    for (; it != m_chunks.end(); ++it) {
        if ((*it)->baseRow > endRow) break;
        out.push_back(it->get());
    }
}

void Column::removeChunkIfEmpty(int row) {
    int idx = findChunkIndex(row);
    if (idx >= 0 && m_chunks[idx]->populatedCount == 0) {
        m_chunks.erase(m_chunks.begin() + idx);
    }
}

int Column::totalPopulated() const {
    int total = 0;
    for (const auto& chunk : m_chunks) {
        total += chunk->populatedCount;
    }
    return total;
}

void Column::reserveChunks(int estimatedRows) {
    int estimatedChunks = (estimatedRows + ColumnChunk::CHUNK_SIZE - 1) / ColumnChunk::CHUNK_SIZE;
    m_chunks.reserve(estimatedChunks);
}

void Column::applyPermutation(int startRow, const std::vector<int>& perm) {
    int n = static_cast<int>(perm.size());
    if (n == 0) return;

    struct CellData {
        CellDataType type = CellDataType::Empty;
        double val = 0.0;
        QString formula;
        uint16_t styleIdx = 0;
    };

    std::vector<CellData> original(n);
    for (int i = 0; i < n; ++i) {
        int row = startRow + i;
        auto* chunk = getChunk(row);
        if (!chunk) continue;
        int offset = row - chunk->baseRow;
        if (!chunk->hasData(offset)) continue;

        auto& cell = original[i];
        cell.type = chunk->getType(offset);
        cell.styleIdx = chunk->getStyleIndex(offset);
        cell.val = chunk->getNumeric(offset);  // works for all types stored in values[]

        if (cell.type == CellDataType::Formula) {
            cell.formula = chunk->getFormulaString(offset);
        }
    }

    // Clear the range
    for (int i = 0; i < n; ++i) {
        int row = startRow + i;
        auto* chunk = getChunk(row);
        if (chunk) {
            int offset = row - chunk->baseRow;
            chunk->removeCell(offset);
        }
    }

    // Write back in permuted order
    for (int i = 0; i < n; ++i) {
        const auto& src = original[perm[i]];
        if (src.type == CellDataType::Empty) continue;

        int row = startRow + i;
        auto* chunk = getOrCreateChunk(row);
        int offset = row - chunk->baseRow;

        switch (src.type) {
            case CellDataType::Double:
                chunk->setNumeric(offset, src.val, src.styleIdx);
                break;
            case CellDataType::Date:
                chunk->setDate(offset, src.val, src.styleIdx);
                break;
            case CellDataType::Boolean:
                chunk->setBoolean(offset, src.val != 0.0, src.styleIdx);
                break;
            case CellDataType::String:
                chunk->setString(offset, ColumnChunk::unpackId(src.val), src.styleIdx);
                break;
            case CellDataType::Formula:
                chunk->setFormula(offset, src.formula, src.styleIdx);
                break;
            case CellDataType::Error:
                chunk->setError(offset, ColumnChunk::unpackId(src.val), src.styleIdx);
                break;
            default:
                break;
        }
    }
}

void Column::clear() {
    m_chunks.clear();
}


// ============================================================================
// ColumnStore
// ============================================================================

ColumnStore::ColumnStore() {}
ColumnStore::~ColumnStore() {}

void ColumnStore::ensureColumn(int col) {
    if (col < 0) return;
    if (col >= static_cast<int>(m_columns.size())) {
        int newSize = col + 1;
        m_columns.resize(newSize);
        m_columnMutexes.resize(newSize);
        for (int i = 0; i < newSize; ++i) {
            if (!m_columnMutexes[i]) {
                m_columnMutexes[i] = std::make_unique<std::shared_mutex>();
            }
        }
    }
    if (!m_columns[col]) {
        m_columns[col] = std::make_unique<Column>();
    }
}

Column* ColumnStore::getColumn(int col) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size())) return nullptr;
    return m_columns[col].get();
}

Column* ColumnStore::getOrCreateColumn(int col) {
    ensureColumn(col);
    return m_columns[col].get();
}

void ColumnStore::setCellValue(int row, int col, const QVariant& value) {
    ensureColumn(col);

    if (value.isNull() || !value.isValid()) {
        removeCell(row, col);
        return;
    }

    auto* column = m_columns[col].get();
    auto* chunk = column->getOrCreateChunk(row);
    int offset = row - chunk->baseRow;

    switch (value.typeId()) {
        case QMetaType::Double:
        case QMetaType::Float:
        case QMetaType::Int:
        case QMetaType::LongLong:
        case QMetaType::UInt:
        case QMetaType::ULongLong:
            chunk->setNumeric(offset, value.toDouble());
            break;
        case QMetaType::Bool:
            chunk->setBoolean(offset, value.toBool());
            break;
        case QMetaType::QString: {
            const QString& str = value.toString();
            if (str.startsWith('=')) {
                chunk->setFormula(offset, str);
            } else {
                bool ok;
                double d = str.toDouble(&ok);
                if (ok) {
                    chunk->setNumeric(offset, d);
                } else {
                    uint32_t id = StringPool::instance().intern(str);
                    chunk->setString(offset, id);
                }
            }
            break;
        }
        default:
            uint32_t id = StringPool::instance().intern(value.toString());
            chunk->setString(offset, id);
            break;
    }
}

void ColumnStore::setCellNumeric(int row, int col, double value) {
    ensureColumn(col);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    chunk->setNumeric(row - chunk->baseRow, value);
}

void ColumnStore::setCellString(int row, int col, const QString& value) {
    ensureColumn(col);
    uint32_t id = StringPool::instance().intern(value);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    chunk->setString(row - chunk->baseRow, id);
}

void ColumnStore::setCellFormula(int row, int col, const QString& formula) {
    ensureColumn(col);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    chunk->setFormula(row - chunk->baseRow, formula);
}

QVariant ColumnStore::getCellValue(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return QVariant();
    auto* chunk = column->getChunk(row);
    if (!chunk) return QVariant();
    return chunk->getValue(row - chunk->baseRow);
}

CellDataType ColumnStore::getCellType(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return CellDataType::Empty;
    auto* chunk = column->getChunk(row);
    if (!chunk) return CellDataType::Empty;
    return chunk->getType(row - chunk->baseRow);
}

bool ColumnStore::hasCell(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return false;
    auto* chunk = column->getChunk(row);
    if (!chunk) return false;
    return chunk->hasData(row - chunk->baseRow);
}

void ColumnStore::removeCell(int row, int col) {
    auto* column = getColumn(col);
    if (!column) return;
    auto* chunk = column->getChunk(row);
    if (!chunk) return;
    chunk->removeCell(row - chunk->baseRow);
    column->removeChunkIfEmpty(row);
}

void ColumnStore::setCellStyle(int row, int col, uint16_t styleIndex) {
    auto* column = getColumn(col);
    if (!column) return;
    auto* chunk = column->getChunk(row);
    if (!chunk) return;
    chunk->setStyleIndex(row - chunk->baseRow, styleIndex);
}

uint16_t ColumnStore::getCellStyleIndex(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return 0;
    auto* chunk = column->getChunk(row);
    if (!chunk) return 0;
    return chunk->getStyleIndex(row - chunk->baseRow);
}

QString ColumnStore::getCellFormula(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return QString();
    auto* chunk = column->getChunk(row);
    if (!chunk) return QString();
    int offset = row - chunk->baseRow;
    if (chunk->getType(offset) != CellDataType::Formula) return QString();
    return chunk->getFormulaString(offset);
}

// ---- Comments ----

void ColumnStore::setCellComment(int row, int col, const QString& comment) {
    ensureColumn(col);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    int offset = row - chunk->baseRow;
    if (!chunk->comments) {
        chunk->comments = std::make_unique<std::unordered_map<int, QString>>();
    }
    if (comment.isEmpty()) {
        chunk->comments->erase(offset);
    } else {
        (*chunk->comments)[offset] = comment;
    }
}

QString ColumnStore::getCellComment(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return QString();
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->comments) return QString();
    auto it = chunk->comments->find(row - chunk->baseRow);
    return (it != chunk->comments->end()) ? it->second : QString();
}

bool ColumnStore::hasCellComment(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return false;
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->comments) return false;
    return chunk->comments->count(row - chunk->baseRow) > 0;
}

// ---- Hyperlinks ----

void ColumnStore::setCellHyperlink(int row, int col, const QString& url) {
    ensureColumn(col);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    int offset = row - chunk->baseRow;
    if (!chunk->hyperlinks) {
        chunk->hyperlinks = std::make_unique<std::unordered_map<int, QString>>();
    }
    if (url.isEmpty()) {
        chunk->hyperlinks->erase(offset);
    } else {
        (*chunk->hyperlinks)[offset] = url;
    }
}

QString ColumnStore::getCellHyperlink(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return QString();
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->hyperlinks) return QString();
    auto it = chunk->hyperlinks->find(row - chunk->baseRow);
    return (it != chunk->hyperlinks->end()) ? it->second : QString();
}

bool ColumnStore::hasCellHyperlink(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return false;
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->hyperlinks) return false;
    return chunk->hyperlinks->count(row - chunk->baseRow) > 0;
}

// ---- Spill tracking ----

void ColumnStore::setSpillParent(int row, int col, const CellAddress& parent) {
    ensureColumn(col);
    auto* chunk = m_columns[col]->getOrCreateChunk(row);
    int offset = row - chunk->baseRow;
    if (!chunk->spillParents) {
        chunk->spillParents = std::make_unique<std::unordered_map<int, CellAddress>>();
    }
    (*chunk->spillParents)[offset] = parent;
}

CellAddress ColumnStore::getSpillParent(int row, int col) const {
    auto* column = getColumn(col);
    if (!column) return CellAddress(-1, -1);
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->spillParents) return CellAddress(-1, -1);
    auto it = chunk->spillParents->find(row - chunk->baseRow);
    return (it != chunk->spillParents->end()) ? it->second : CellAddress(-1, -1);
}

bool ColumnStore::isSpillCell(int row, int col) const {
    auto parent = getSpillParent(row, col);
    return parent.row >= 0;
}

void ColumnStore::clearSpillParent(int row, int col) {
    auto* column = getColumn(col);
    if (!column) return;
    auto* chunk = column->getChunk(row);
    if (!chunk || !chunk->spillParents) return;
    chunk->spillParents->erase(row - chunk->baseRow);
}

// ---- Computed values ----

void ColumnStore::setComputedValue(int row, int col, const QVariant& value) {
    std::unique_lock lock(m_computedMutex);
    m_computedValues[computedKey(row, col)] = value;
}

QVariant ColumnStore::getComputedValue(int row, int col) const {
    std::shared_lock lock(m_computedMutex);
    auto it = m_computedValues.find(computedKey(row, col));
    return (it != m_computedValues.end()) ? it->second : QVariant();
}

// ---- Bulk operations ----

double ColumnStore::sumColumn(int col, int startRow, int endRow) const {
    double sum = 0.0;
    auto* column = getColumn(col);
    if (!column) return sum;

    std::vector<ColumnChunk*> chunks;
    column->getChunksInRange(startRow, endRow, chunks);

    for (const auto* chunk : chunks) {
        int chunkStart = std::max(0, startRow - chunk->baseRow);
        int chunkEnd = std::min(ColumnChunk::CHUNK_SIZE - 1, endRow - chunk->baseRow);

        for (int offset = chunkStart; offset <= chunkEnd; ++offset) {
            if (chunk->hasData(offset)) {
                auto type = chunk->getType(offset);
                if (type == CellDataType::Double || type == CellDataType::Date) {
                    int idx = chunk->denseIndex(offset);
                    sum += chunk->values[idx];
                }
            }
        }
    }
    return sum;
}

int ColumnStore::countColumn(int col, int startRow, int endRow) const {
    int count = 0;
    auto* column = getColumn(col);
    if (!column) return count;

    std::vector<ColumnChunk*> chunks;
    column->getChunksInRange(startRow, endRow, chunks);

    for (const auto* chunk : chunks) {
        int chunkStart = std::max(0, startRow - chunk->baseRow);
        int chunkEnd = std::min(ColumnChunk::CHUNK_SIZE - 1, endRow - chunk->baseRow);

        for (int offset = chunkStart; offset <= chunkEnd; ++offset) {
            if (chunk->hasData(offset)) ++count;
        }
    }
    return count;
}

// ---- Navigation: direct bitmap queries, O(chunks) not O(cells) ----

int ColumnStore::nextOccupiedRow(int col, int startRow) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col]) return -1;
    auto* column = m_columns[col].get();

    for (const auto& chunk : column->chunks()) {
        int chunkEnd = chunk->baseRow + ColumnChunk::CHUNK_SIZE;
        if (chunkEnd <= startRow) continue;  // chunk entirely before startRow
        if (chunk->populatedCount == 0) continue;

        int firstOffset = std::max(0, startRow - chunk->baseRow);
        int firstWord = firstOffset / 64;
        int firstBit = firstOffset % 64;

        // Check first partial word
        uint64_t word = chunk->presence[firstWord] & (~0ULL << firstBit);
        if (word) {
            int bit = ctzll(word);
            return chunk->baseRow + firstWord * 64 + bit;
        }
        // Check remaining words in this chunk
        for (int w = firstWord + 1; w < ColumnChunk::BITMAP_WORDS; ++w) {
            word = chunk->presence[w];
            if (word) {
                int bit = ctzll(word);
                return chunk->baseRow + w * 64 + bit;
            }
        }
    }
    return -1;
}

int ColumnStore::prevOccupiedRow(int col, int startRow) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col]) return -1;
    auto* column = m_columns[col].get();
    auto& chunks = column->chunks();

    // Iterate chunks in reverse
    for (int ci = static_cast<int>(chunks.size()) - 1; ci >= 0; --ci) {
        auto* chunk = chunks[ci].get();
        if (chunk->baseRow > startRow) continue;  // chunk entirely after startRow
        if (chunk->populatedCount == 0) continue;

        int lastOffset = std::min(ColumnChunk::CHUNK_SIZE - 1, startRow - chunk->baseRow);
        int lastWord = lastOffset / 64;
        int lastBit = lastOffset % 64;

        // Check last partial word (mask off bits above lastBit)
        uint64_t word = chunk->presence[lastWord] & ((2ULL << lastBit) - 1);
        if (word) {
            int bit = 63 - clzll(word);  // highest set bit
            return chunk->baseRow + lastWord * 64 + bit;
        }
        // Check remaining words in reverse
        for (int w = lastWord - 1; w >= 0; --w) {
            word = chunk->presence[w];
            if (word) {
                int bit = 63 - clzll(word);
                return chunk->baseRow + w * 64 + bit;
            }
        }
    }
    return -1;
}

int ColumnStore::nextEmptyRow(int col, int startRow) const {
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col]) return startRow;
    auto* column = m_columns[col].get();

    for (const auto& chunk : column->chunks()) {
        int chunkEnd = chunk->baseRow + ColumnChunk::CHUNK_SIZE;
        if (chunkEnd <= startRow) continue;
        if (chunk->populatedCount == 0) return startRow;  // whole chunk empty

        // If startRow is before this chunk, the gap is empty
        if (startRow < chunk->baseRow) return startRow;

        int firstOffset = startRow - chunk->baseRow;
        int firstWord = firstOffset / 64;
        int firstBit = firstOffset % 64;

        // Check first partial word for zero bits
        uint64_t word = chunk->presence[firstWord] | ((1ULL << firstBit) - 1);  // fill lower bits
        if (~word) {  // has at least one zero bit at or above firstBit
            uint64_t inv = ~word;
            int bit = ctzll(inv);
            if (bit < 64) return chunk->baseRow + firstWord * 64 + bit;
        }
        for (int w = firstWord + 1; w < ColumnChunk::BITMAP_WORDS; ++w) {
            if (~chunk->presence[w]) {
                int bit = ctzll(~chunk->presence[w]);
                return chunk->baseRow + w * 64 + bit;
            }
        }
        // Entire rest of chunk is full — continue to next chunk gap
        if (chunkEnd > startRow) {
            startRow = chunkEnd;  // skip to end of this chunk
        }
    }
    return startRow;
}

int ColumnStore::prevEmptyRow(int col, int startRow) const {
    if (startRow < 0) return 0;
    if (col < 0 || col >= static_cast<int>(m_columns.size()) || !m_columns[col]) return startRow;
    auto* column = m_columns[col].get();
    auto& chunks = column->chunks();

    for (int ci = static_cast<int>(chunks.size()) - 1; ci >= 0; --ci) {
        auto* chunk = chunks[ci].get();
        if (chunk->baseRow > startRow) continue;
        if (chunk->populatedCount == 0) return startRow;

        int chunkEnd = chunk->baseRow + ColumnChunk::CHUNK_SIZE;
        // If startRow is past this chunk, the gap is empty
        if (startRow >= chunkEnd) return startRow;

        int lastOffset = startRow - chunk->baseRow;
        int lastWord = lastOffset / 64;
        int lastBit = lastOffset % 64;

        // Check last partial word for zero bits
        uint64_t word = chunk->presence[lastWord] | (~((2ULL << lastBit) - 1));  // fill upper bits
        if (~word) {
            uint64_t inv = ~word;
            int bit = 63 - clzll(inv);  // highest zero bit
            return chunk->baseRow + lastWord * 64 + bit;
        }
        for (int w = lastWord - 1; w >= 0; --w) {
            if (~chunk->presence[w]) {
                int bit = 63 - clzll(~chunk->presence[w]);
                return chunk->baseRow + w * 64 + bit;
            }
        }
        // Entire chunk is full up to here — check before chunk
        if (chunk->baseRow > 0) {
            startRow = chunk->baseRow - 1;
        } else {
            return -1;  // row 0 is occupied, no empty row before it
        }
    }
    return startRow;
}

void ColumnStore::forEachCell(
    std::function<void(int row, int col, CellDataType type, const QVariant& value)> callback) const
{
    for (int c = 0; c < static_cast<int>(m_columns.size()); ++c) {
        if (!m_columns[c]) continue;
        for (const auto& chunk : m_columns[c]->chunks()) {
            for (int offset = 0; offset < ColumnChunk::CHUNK_SIZE; ++offset) {
                if (chunk->hasData(offset)) {
                    int row = chunk->baseRow + offset;
                    CellDataType type = chunk->getType(offset);
                    QVariant val = chunk->getValue(offset);
                    callback(row, c, type, val);
                }
            }
        }
    }
}

void ColumnStore::reserve(int estimatedRows, int estimatedCols) {
    ensureColumn(estimatedCols - 1);
    for (int c = 0; c < estimatedCols; ++c) {
        if (m_columns[c]) {
            m_columns[c]->reserveChunks(estimatedRows);
        }
    }
}

void ColumnStore::applyPermutation(int startRow, int startCol, int endCol,
                                    const std::vector<int>& perm) {
    for (int c = startCol; c <= endCol; ++c) {
        if (c < static_cast<int>(m_columns.size()) && m_columns[c]) {
            m_columns[c]->applyPermutation(startRow, perm);
        }
    }
}

void ColumnStore::clear() {
    m_columns.clear();
    m_columnMutexes.clear();
    std::unique_lock lock(m_computedMutex);
    m_computedValues.clear();
}

int ColumnStore::maxRow() const {
    int maxR = -1;
    for (const auto& col : m_columns) {
        if (!col) continue;
        const auto& chunks = col->chunks();
        if (chunks.empty()) continue;
        const auto& lastChunk = chunks.back();
        for (int offset = ColumnChunk::CHUNK_SIZE - 1; offset >= 0; --offset) {
            if (lastChunk->hasData(offset)) {
                maxR = std::max(maxR, lastChunk->baseRow + offset);
                break;
            }
        }
    }
    return maxR;
}

int ColumnStore::maxCol() const {
    for (int c = static_cast<int>(m_columns.size()) - 1; c >= 0; --c) {
        if (m_columns[c] && m_columns[c]->totalPopulated() > 0) return c;
    }
    return -1;
}

std::shared_mutex& ColumnStore::columnMutex(int col) {
    ensureColumn(col);
    return *m_columnMutexes[col];
}
