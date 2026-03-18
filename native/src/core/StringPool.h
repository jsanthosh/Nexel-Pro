#ifndef STRINGPOOL_H
#define STRINGPOOL_H

#include <QString>
#include <QHash>
#include <unordered_map>
#include <vector>
#include <shared_mutex>
#include <atomic>
#include <cstdint>

// ============================================================================
// StringPool — Sharded global string deduplication
// ============================================================================
// 16 shards keyed by hash for concurrent intern() from multiple threads.
// Append-only global ID→string vector for O(1) get().
// Typical spreadsheet: 100K unique strings, avg 50 bytes = 5 MB total.
//
class StringPool {
public:
    static constexpr int SHARD_COUNT = 16;
    static constexpr int SHARD_MASK = SHARD_COUNT - 1;

    static StringPool& instance();

    // Returns an interned string ID. Same content always returns same ID.
    // Thread-safe: locks only the relevant shard + brief global append lock.
    uint32_t intern(const QString& str);

    // Bulk intern — groups strings by shard to minimize lock transitions.
    // Returns vector of IDs in same order as input.
    std::vector<uint32_t> internBatch(const std::vector<QString>& strings);

    // Get string by ID. Lock-free for valid IDs below committed count.
    const QString& get(uint32_t id) const;

    // Stats
    size_t uniqueCount() const;
    size_t totalInterned() const;

    void clear();

private:
    StringPool();

    struct QStringHash {
        size_t operator()(const QString& s) const noexcept { return qHash(s); }
    };

    int shardIndex(const QString& s) const {
        return static_cast<int>(qHash(s) & SHARD_MASK);
    }

    // Per-shard: own mutex + string→id map
    struct Shard {
        mutable std::shared_mutex mutex;
        std::unordered_map<QString, uint32_t, QStringHash> stringToId;
    };

    Shard m_shards[SHARD_COUNT];

    // Global ID→string vector (append-only)
    mutable std::shared_mutex m_globalMutex;
    std::vector<QString> m_idToString;

    std::atomic<size_t> m_totalInterned{0};
};

#endif // STRINGPOOL_H
