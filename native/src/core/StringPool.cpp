#include "StringPool.h"
#include <algorithm>
#include <mutex>

StringPool& StringPool::instance() {
    static StringPool s_instance;
    return s_instance;
}

StringPool::StringPool() {
    // Reserve space in global vector to reduce early reallocations
    m_idToString.reserve(4096);
}

uint32_t StringPool::intern(const QString& str) {
    m_totalInterned.fetch_add(1, std::memory_order_relaxed);

    int shard = shardIndex(str);
    auto& s = m_shards[shard];

    // Fast path: read lock on shard — check if already interned
    {
        std::shared_lock readLock(s.mutex);
        auto it = s.stringToId.find(str);
        if (it != s.stringToId.end()) {
            return it->second;
        }
    }

    // Slow path: write lock on shard + global
    std::unique_lock shardLock(s.mutex);

    // Double-check after acquiring write lock
    auto it = s.stringToId.find(str);
    if (it != s.stringToId.end()) {
        return it->second;
    }

    // Allocate global ID
    uint32_t id;
    {
        std::unique_lock globalLock(m_globalMutex);
        id = static_cast<uint32_t>(m_idToString.size());
        m_idToString.push_back(str);
    }

    s.stringToId.emplace(str, id);
    return id;
}

std::vector<uint32_t> StringPool::internBatch(const std::vector<QString>& strings) {
    if (strings.empty()) return {};

    std::vector<uint32_t> ids(strings.size());

    // Group indices by shard to minimize lock transitions
    std::vector<std::vector<size_t>> shardBuckets(SHARD_COUNT);
    for (size_t i = 0; i < strings.size(); ++i) {
        shardBuckets[shardIndex(strings[i])].push_back(i);
    }

    // Process each shard's batch
    for (int shard = 0; shard < SHARD_COUNT; ++shard) {
        const auto& bucket = shardBuckets[shard];
        if (bucket.empty()) continue;

        auto& s = m_shards[shard];

        // First pass: resolve already-interned strings with read lock
        std::vector<size_t> needInsert;
        {
            std::shared_lock readLock(s.mutex);
            for (size_t idx : bucket) {
                auto it = s.stringToId.find(strings[idx]);
                if (it != s.stringToId.end()) {
                    ids[idx] = it->second;
                } else {
                    needInsert.push_back(idx);
                }
            }
        }

        if (needInsert.empty()) {
            m_totalInterned.fetch_add(bucket.size(), std::memory_order_relaxed);
            continue;
        }

        // Second pass: insert new strings with write lock
        std::unique_lock shardLock(s.mutex);
        std::unique_lock globalLock(m_globalMutex);

        for (size_t idx : needInsert) {
            // Double-check (another thread may have inserted between read and write lock)
            auto it = s.stringToId.find(strings[idx]);
            if (it != s.stringToId.end()) {
                ids[idx] = it->second;
            } else {
                uint32_t id = static_cast<uint32_t>(m_idToString.size());
                m_idToString.push_back(strings[idx]);
                s.stringToId.emplace(strings[idx], id);
                ids[idx] = id;
            }
        }

        m_totalInterned.fetch_add(bucket.size(), std::memory_order_relaxed);
    }

    return ids;
}

const QString& StringPool::get(uint32_t id) const {
    static const QString s_empty;
    std::shared_lock readLock(m_globalMutex);
    if (id >= m_idToString.size()) {
        return s_empty;
    }
    return m_idToString[id];
}

size_t StringPool::uniqueCount() const {
    std::shared_lock readLock(m_globalMutex);
    return m_idToString.size();
}

size_t StringPool::totalInterned() const {
    return m_totalInterned.load(std::memory_order_relaxed);
}

void StringPool::clear() {
    // Lock all shards first, then global
    for (int i = 0; i < SHARD_COUNT; ++i) {
        m_shards[i].mutex.lock();
    }
    std::unique_lock globalLock(m_globalMutex);

    for (int i = 0; i < SHARD_COUNT; ++i) {
        m_shards[i].stringToId.clear();
    }
    m_idToString.clear();
    m_totalInterned.store(0, std::memory_order_relaxed);

    // Unlock shards
    for (int i = SHARD_COUNT - 1; i >= 0; --i) {
        m_shards[i].mutex.unlock();
    }
}
