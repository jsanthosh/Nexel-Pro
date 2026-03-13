#include "StringPool.h"

StringPool& StringPool::instance() {
    static StringPool s_instance;
    return s_instance;
}

uint32_t StringPool::intern(const QString& str) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_totalInterned++;

    auto it = m_stringToId.find(str);
    if (it != m_stringToId.end()) {
        return it->second;
    }

    uint32_t id = static_cast<uint32_t>(m_idToString.size());
    m_idToString.push_back(str);
    m_stringToId.emplace(str, id);
    return id;
}

const QString& StringPool::get(uint32_t id) const {
    std::lock_guard<std::mutex> lock(m_mutex);
    static const QString s_empty;
    if (id >= m_idToString.size()) {
        return s_empty;
    }
    return m_idToString[id];
}

size_t StringPool::uniqueCount() const {
    std::lock_guard<std::mutex> lock(m_mutex);
    return m_idToString.size();
}

size_t StringPool::totalInterned() const {
    std::lock_guard<std::mutex> lock(m_mutex);
    return m_totalInterned;
}

void StringPool::clear() {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_stringToId.clear();
    m_idToString.clear();
    m_totalInterned = 0;
}
