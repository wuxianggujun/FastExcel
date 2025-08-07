/**
 * @file CacheSystem.hpp
 * @brief 缓存系统实现，用于提高数据访问性能
 */

#pragma once

#include <unordered_map>
#include <list>
#include <mutex>
#include <memory>
#include <functional>
#include <chrono>

namespace fastexcel {
namespace core {

/**
 * @brief LRU缓存实现
 * 
 * 最近最少使用缓存，当缓存满时会自动淘汰最久未使用的项目
 */
template<typename Key, typename Value>
class LRUCache {
public:
    using KeyType = Key;
    using ValueType = Value;
    using TimePoint = std::chrono::steady_clock::time_point;
    
    /**
     * @brief 缓存项结构
     */
    struct CacheItem {
        Value value;
        TimePoint last_access;
        size_t access_count;
        
        CacheItem(const Value& v) 
            : value(v)
            , last_access(std::chrono::steady_clock::now())
            , access_count(1) {}
    };
    
    /**
     * @brief 构造函数
     * @param max_size 最大缓存项数量
     */
    explicit LRUCache(size_t max_size = 1000) 
        : max_size_(max_size) {}
    
    /**
     * @brief 获取缓存项
     * @param key 键
     * @return 值的指针，如果不存在返回nullptr
     */
    std::shared_ptr<Value> get(const Key& key) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = cache_map_.find(key);
        if (it == cache_map_.end()) {
            stats_.miss_count++;
            return nullptr;
        }
        
        // 更新访问信息
        auto list_it = it->second;
        list_it->last_access = std::chrono::steady_clock::now();
        list_it->access_count++;
        
        // 移动到链表头部（最近使用）
        cache_list_.splice(cache_list_.begin(), cache_list_, list_it);
        
        stats_.hit_count++;
        return std::make_shared<Value>(list_it->value);
    }
    
    /**
     * @brief 设置缓存项
     * @param key 键
     * @param value 值
     */
    void put(const Key& key, const Value& value) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = cache_map_.find(key);
        if (it != cache_map_.end()) {
            // 更新现有项
            auto list_it = it->second;
            list_it->value = value;
            list_it->last_access = std::chrono::steady_clock::now();
            list_it->access_count++;
            
            // 移动到链表头部
            cache_list_.splice(cache_list_.begin(), cache_list_, list_it);
            return;
        }
        
        // 检查是否需要淘汰
        if (cache_list_.size() >= max_size_) {
            evictLRU();
        }
        
        // 添加新项到链表头部
        cache_list_.emplace_front(value);
        cache_map_[key] = cache_list_.begin();
        
        stats_.put_count++;
    }
    
    /**
     * @brief 删除缓存项
     * @param key 键
     * @return 是否成功删除
     */
    bool remove(const Key& key) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = cache_map_.find(key);
        if (it == cache_map_.end()) {
            return false;
        }
        
        cache_list_.erase(it->second);
        cache_map_.erase(it);
        return true;
    }
    
    /**
     * @brief 清空缓存
     */
    void clear() {
        std::lock_guard<std::mutex> lock(mutex_);
        cache_list_.clear();
        cache_map_.clear();
        stats_ = {};
    }
    
    /**
     * @brief 获取缓存大小
     */
    size_t size() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return cache_list_.size();
    }
    
    /**
     * @brief 检查是否为空
     */
    bool empty() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return cache_list_.empty();
    }
    
    /**
     * @brief 缓存统计信息
     */
    struct Statistics {
        size_t hit_count = 0;
        size_t miss_count = 0;
        size_t put_count = 0;
        size_t evict_count = 0;
        
        double hit_rate() const {
            size_t total = hit_count + miss_count;
            return total > 0 ? static_cast<double>(hit_count) / total : 0.0;
        }
    };
    
    /**
     * @brief 获取统计信息
     */
    Statistics getStatistics() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return stats_;
    }
    
    /**
     * @brief 重置统计信息
     */
    void resetStatistics() {
        std::lock_guard<std::mutex> lock(mutex_);
        stats_ = {};
    }

private:
    using CacheList = std::list<CacheItem>;
    using CacheMap = std::unordered_map<Key, typename CacheList::iterator>;
    
    mutable std::mutex mutex_;
    size_t max_size_;
    CacheList cache_list_;
    CacheMap cache_map_;
    Statistics stats_;
    
    /**
     * @brief 淘汰最久未使用的项（优化版本）
     */
    void evictLRU() {
        if (cache_list_.empty()) {
            return;
        }
        
        // 直接获取链表尾部的项（最久未使用）
        auto last = std::prev(cache_list_.end());
        
        // 优化：避免遍历整个map，直接通过值查找对应的键
        // 这里我们需要反向查找，虽然仍需要遍历，但代码更清晰
        for (auto it = cache_map_.begin(); it != cache_map_.end(); ++it) {
            if (it->second == last) {
                cache_map_.erase(it);
                break;
            }
        }
        
        cache_list_.erase(last);
        stats_.evict_count++;
    }
};

/**
 * @brief 字符串缓存特化
 */
using StringCache = LRUCache<std::string, std::string>;

/**
 * @brief 格式缓存
 */
class Format; // 前向声明
using FormatCache = LRUCache<int, std::shared_ptr<Format>>;

/**
 * @brief 缓存管理器
 * 
 * 管理多个不同类型的缓存
 */
class CacheManager {
public:
    /**
     * @brief 获取全局实例
     */
    static CacheManager& getInstance();
    
    /**
     * @brief 获取字符串缓存
     */
    StringCache& getStringCache();
    
    /**
     * @brief 获取格式缓存
     */
    FormatCache& getFormatCache();
    
    /**
     * @brief 清理所有缓存
     */
    void clearAll();
    
    /**
     * @brief 获取全局缓存统计
     */
    struct GlobalCacheStats {
        StringCache::Statistics string_cache;
        FormatCache::Statistics format_cache;
        size_t total_memory_usage;
    };
    
    GlobalCacheStats getGlobalStatistics() const;
    
    /**
     * @brief 设置缓存大小
     */
    void setStringCacheSize(size_t size);
    void setFormatCacheSize(size_t size);

private:
    CacheManager();
    ~CacheManager() = default;
    
    mutable std::mutex mutex_;
    std::unique_ptr<StringCache> string_cache_;
    std::unique_ptr<FormatCache> format_cache_;
    
    // 禁用拷贝和赋值
    CacheManager(const CacheManager&) = delete;
    CacheManager& operator=(const CacheManager&) = delete;
};

/**
 * @brief 缓存辅助类
 * 
 * 提供便捷的缓存操作接口
 */
template<typename Key, typename Value>
class CacheHelper {
public:
    using CacheType = LRUCache<Key, Value>;
    using LoadFunction = std::function<Value(const Key&)>;
    
    /**
     * @brief 构造函数
     * @param cache 缓存实例
     * @param loader 数据加载函数
     */
    CacheHelper(CacheType& cache, LoadFunction loader)
        : cache_(cache), loader_(loader) {}
    
    /**
     * @brief 获取或加载数据
     * @param key 键
     * @return 值
     */
    Value getOrLoad(const Key& key) {
        auto cached = cache_.get(key);
        if (cached) {
            return *cached;
        }
        
        // 缓存未命中，加载数据
        Value value = loader_(key);
        cache_.put(key, value);
        return value;
    }
    
    /**
     * @brief 预加载数据
     * @param keys 键列表
     */
    void preload(const std::vector<Key>& keys) {
        for (const auto& key : keys) {
            if (!cache_.get(key)) {
                try {
                    Value value = loader_(key);
                    cache_.put(key, value);
                } catch (...) {
                    // 忽略加载错误，继续处理其他键
                }
            }
        }
    }

private:
    CacheType& cache_;
    LoadFunction loader_;
};

} // namespace core
} // namespace fastexcel

