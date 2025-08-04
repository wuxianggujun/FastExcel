/**
 * @file CacheSystem.cpp
 * @brief 缓存系统实现
 */

#include "CacheSystem.hpp"
#include "Format.hpp"
#include <iostream>

namespace fastexcel {
namespace core {

// CacheManager 实现
CacheManager::CacheManager()
    : string_cache_(std::make_unique<StringCache>(1000))
    , format_cache_(std::make_unique<FormatCache>(500)) {
}

CacheManager& CacheManager::getInstance() {
    static CacheManager instance;
    return instance;
}

StringCache& CacheManager::getStringCache() {
    return *string_cache_;
}

FormatCache& CacheManager::getFormatCache() {
    return *format_cache_;
}

void CacheManager::clearAll() {
    std::lock_guard<std::mutex> lock(mutex_);
    if (string_cache_) {
        string_cache_->clear();
    }
    if (format_cache_) {
        format_cache_->clear();
    }
}

CacheManager::GlobalCacheStats CacheManager::getGlobalStatistics() const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    GlobalCacheStats stats{};
    
    if (string_cache_) {
        stats.string_cache = string_cache_->getStatistics();
    }
    
    if (format_cache_) {
        stats.format_cache = format_cache_->getStatistics();
    }
    
    // 估算内存使用量
    stats.total_memory_usage = 
        (string_cache_ ? string_cache_->size() * 64 : 0) +  // 假设每个字符串平均64字节
        (format_cache_ ? format_cache_->size() * 128 : 0);  // 假设每个格式对象128字节
    
    return stats;
}

void CacheManager::setStringCacheSize(size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    string_cache_ = std::make_unique<StringCache>(size);
}

void CacheManager::setFormatCacheSize(size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    format_cache_ = std::make_unique<FormatCache>(size);
}

} // namespace core
} // namespace fastexcel