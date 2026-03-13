#ifndef STRINGPOOL_H
#define STRINGPOOL_H

#include <QString>
#include <QHash>
#include <unordered_map>
#include <vector>
#include <mutex>
#include <cstdint>

// Global string pool for deduplicating repeated cell values (e.g., "Yes"/"No", category names)
// Reduces memory when many cells contain the same string
class StringPool {
public:
    static StringPool& instance();

    // Returns an interned string ID. Same content always returns same ID.
    uint32_t intern(const QString& str);

    // Get string by ID
    const QString& get(uint32_t id) const;

    // Stats
    size_t uniqueCount() const;
    size_t totalInterned() const;

    void clear();

private:
    StringPool() = default;
    struct QStringHash {
        size_t operator()(const QString& s) const noexcept { return qHash(s); }
    };
    mutable std::mutex m_mutex;
    std::unordered_map<QString, uint32_t, QStringHash> m_stringToId;
    std::vector<QString> m_idToString;
    size_t m_totalInterned = 0;
};

#endif // STRINGPOOL_H
