#pragma once

#include "FormatDescriptor.hpp"
#include <memory>
#include <unordered_map>
#include <vector>
#include <shared_mutex>
#include <atomic>

namespace fastexcel {
namespace core {

/**
 * @brief 格式仓储 - 线程安全的格式去重存储
 * 
 * 使用Repository模式实现格式的去重存储和检索。
 * 提供线程安全的操作，支持高并发访问。
 */
class FormatRepository {
private:
    // 线程安全保护
    mutable std::shared_mutex mutex_;
    
    // 格式存储：不可变的格式描述符
    std::vector<std::shared_ptr<const FormatDescriptor>> formats_;
    
    // 哈希到ID的映射，用于快速查找
    std::unordered_map<size_t, int> hash_to_id_;
    
    // 统计信息
    mutable std::atomic<size_t> total_requests_{0};
    mutable std::atomic<size_t> cache_hits_{0};
    
    // 默认格式ID（总是0）
    static constexpr int DEFAULT_FORMAT_ID = 0;

public:
    FormatRepository();
    ~FormatRepository() = default;
    
    // 禁用拷贝，允许移动
    FormatRepository(const FormatRepository&) = delete;
    FormatRepository& operator=(const FormatRepository&) = delete;
    FormatRepository(FormatRepository&&) = default;
    FormatRepository& operator=(FormatRepository&&) = default;
    
    /**
     * @brief 添加格式到仓储（幂等操作）
     * @param format 格式描述符
     * @return 格式ID，如果已存在则返回现有ID
     */
    int addFormat(const FormatDescriptor& format);
    
    /**
     * @brief 根据ID获取格式
     * @param id 格式ID
     * @return 格式描述符的共享指针，如果ID无效则返回默认格式
     */
    std::shared_ptr<const FormatDescriptor> getFormat(int id) const;
    
    /**
     * @brief 获取默认格式ID
     * @return 默认格式ID（总是0）
     */
    int getDefaultFormatId() const { return DEFAULT_FORMAT_ID; }
    
    /**
     * @brief 获取默认格式
     * @return 默认格式描述符
     */
    std::shared_ptr<const FormatDescriptor> getDefaultFormat() const;
    
    /**
     * @brief 获取格式数量
     * @return 存储的格式数量
     */
    size_t getFormatCount() const {
        std::shared_lock lock(mutex_);
        return formats_.size();
    }
    
    /**
     * @brief 检查格式ID是否有效
     * @param id 格式ID
     * @return 是否有效
     */
    bool isValidFormatId(int id) const {
        std::shared_lock lock(mutex_);
        return id >= 0 && static_cast<size_t>(id) < formats_.size();
    }
    
    /**
     * @brief 清空仓储（保留默认格式）
     */
    void clear();
    
    /**
     * @brief 获取缓存命中率
     * @return 缓存命中率（0.0-1.0）
     */
    double getCacheHitRate() const;
    
    /**
     * @brief 获取去重统计
     */
    struct DeduplicationStats {
        size_t total_requests;
        size_t unique_formats;
        double deduplication_ratio;
    };
    DeduplicationStats getDeduplicationStats() const;
    
    /**
     * @brief 获取内存使用估算（字节）
     * @return 估算的内存使用量
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 批量导入格式（用于跨工作簿复制）
     * @param source_repo 源仓储
     * @param id_mapping 返回的ID映射（源ID -> 目标ID）
     */
    void importFormats(const FormatRepository& source_repo, 
                      std::unordered_map<int, int>& id_mapping);
    
    // ========== 迭代器支持 ==========
    
    class const_iterator {
    private:
        std::vector<std::shared_ptr<const FormatDescriptor>>::const_iterator iter_;
        int id_;
        
    public:
        const_iterator(const std::vector<std::shared_ptr<const FormatDescriptor>>::const_iterator& iter, int id)
            : iter_(iter), id_(id) {}
        
        const_iterator& operator++() { ++iter_; ++id_; return *this; }
        bool operator!=(const const_iterator& other) const { return iter_ != other.iter_; }
        
        struct value_type {
            int id;
            std::shared_ptr<const FormatDescriptor> format;
        };
        
        value_type operator*() const { return {id_, *iter_}; }
    };
    
    const_iterator begin() const {
        std::shared_lock lock(mutex_);
        return const_iterator(formats_.begin(), 0);
    }
    
    const_iterator end() const {
        std::shared_lock lock(mutex_);
        return const_iterator(formats_.end(), static_cast<int>(formats_.size()));
    }

private:
    /**
     * @brief 内部添加格式方法（需要调用者持有锁）
     * @param format 格式描述符
     * @return 格式ID
     */
    int addFormatInternal(const FormatDescriptor& format);
};

}} // namespace fastexcel::core