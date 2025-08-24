/**
 * @file StringMemoryPool.hpp
 * @brief 专门用于管理字符串的内存池
 */

#pragma once

#include "StringPool.hpp"
#include <string>
#include <string_view>
#include <atomic>

namespace fastexcel {
namespace memory {

/**
 * @brief 字符串专用内存池管理器
 * 
 * 提供高效的字符串内存分配和去重，针对Excel工作簿中
 * 大量重复字符串的存储和共享场景进行优化
 */
class StringMemoryPool {
private:
    mutable StringPool string_pool_;
    
    // 统计信息
    mutable std::atomic<size_t> total_interns_{0};
    mutable std::atomic<size_t> duplicate_saves_{0};
    
public:
    StringMemoryPool() = default;
    ~StringMemoryPool() = default;
    
    // 禁止拷贝，允许移动
    StringMemoryPool(const StringMemoryPool&) = delete;
    StringMemoryPool& operator=(const StringMemoryPool&) = delete;
    StringMemoryPool(StringMemoryPool&&) = default;
    StringMemoryPool& operator=(StringMemoryPool&&) = default;
    
    /**
     * @brief 将字符串加入内存池并返回池化的指针
     * @param value 要池化的字符串
     * @return 指向池化字符串的指针
     */
    const std::string* intern(const std::string& value) const {
        ++total_interns_;
        
        // 检查是否已经存在相同的字符串
        const std::string* pooled_string = string_pool_.intern(value);
        
        // 如果返回的指针不是新分配的，说明节省了重复存储
        if (pooled_string != &value) {
            ++duplicate_saves_;
        }
        
        return pooled_string;
    }
    
    /**
     * @brief 将string_view加入内存池并返回池化的指针
     * @param value 要池化的字符串视图
     * @return 指向池化字符串的指针
     */
    const std::string* intern(std::string_view value) const {
        ++total_interns_;
        const std::string* pooled_string = string_pool_.intern(std::string(value));
        return pooled_string;
    }
    
    /**
     * @brief 将C风格字符串加入内存池并返回池化的指针
     * @param value 要池化的C字符串
     * @return 指向池化字符串的指针
     */
    const std::string* intern(const char* value) const {
        if (!value) return nullptr;
        return intern(std::string(value));
    }
    
    /**
     * @brief 检查字符串是否已经在池中
     * @param value 要检查的字符串
     * @return 如果在池中返回true
     */
    bool contains(const std::string& value) const {
        return string_pool_.contains(value);
    }
    
    /**
     * @brief 查找池中的字符串指针
     * @param value 要查找的字符串
     * @return 如果找到返回指针，否则返回nullptr
     */
    const std::string* find(const std::string& value) const {
        return string_pool_.find(value);
    }
    
    /**
     * @brief 获取字符串池统计信息
     */
    struct Statistics {
        size_t total_unique_strings = 0;
        size_t total_interns = 0;
        size_t duplicate_saves = 0;
        size_t memory_saved_bytes = 0;
        double deduplication_ratio = 0.0;
    };
    
    Statistics getStatistics() const {
        Statistics stats;
        stats.total_unique_strings = string_pool_.size();
        stats.total_interns = total_interns_.load();
        stats.duplicate_saves = duplicate_saves_.load();
        
        // 计算节省的内存（估算）
        stats.memory_saved_bytes = stats.duplicate_saves * 
            (stats.total_unique_strings > 0 ? 
             string_pool_.getTotalMemoryUsage() / stats.total_unique_strings : 0);
        
        // 计算去重比率
        if (stats.total_interns > 0) {
            stats.deduplication_ratio = 
                static_cast<double>(stats.duplicate_saves) / stats.total_interns;
        }
        
        return stats;
    }
    
    /**
     * @brief 获取池中字符串总数
     * @return 唯一字符串数量
     */
    size_t size() const {
        return string_pool_.size();
    }
    
    /**
     * @brief 获取内存使用情况
     * @return 总内存使用字节数
     */
    size_t getMemoryUsage() const {
        return string_pool_.getTotalMemoryUsage();
    }
    
    /**
     * @brief 收缩内存池，释放未使用的内存
     */
    void shrink() {
        string_pool_.shrink();
    }
    
    /**
     * @brief 清空字符串池
     */
    void clear() {
        string_pool_.clear();
        total_interns_ = 0;
        duplicate_saves_ = 0;
    }
    
    /**
     * @brief 预留内存空间
     * @param capacity 预期的字符串数量
     */
    void reserve(size_t capacity) {
        string_pool_.reserve(capacity);
    }
};

}} // namespace fastexcel::memory