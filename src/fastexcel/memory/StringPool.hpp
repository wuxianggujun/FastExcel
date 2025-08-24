/**
 * @file StringPool.hpp  
 * @brief 内存管理专用的字符串池
 */

#pragma once

#include <string>
#include <string_view>
#include <unordered_set>
#include <unordered_map>
#include <atomic>

namespace fastexcel {
namespace memory {

/**
 * @brief 字符串池 - 内存管理专用版本
 * 
 * 高效的字符串去重和内存管理，专门用于Excel工作簿中的字符串优化
 */
class StringPool {
public:
    StringPool() = default;
    ~StringPool() = default;
    
    // 禁止拷贝，允许移动
    StringPool(const StringPool&) = delete;
    StringPool& operator=(const StringPool&) = delete;
    
    StringPool(StringPool&& other) noexcept
        : pool_(std::move(other.pool_))
        , memory_usage_(other.memory_usage_.load()) {
        other.memory_usage_ = 0;
    }
    
    StringPool& operator=(StringPool&& other) noexcept {
        if (this != &other) {
            pool_ = std::move(other.pool_);
            memory_usage_ = other.memory_usage_.load();
            other.memory_usage_ = 0;
        }
        return *this;
    }
    
    /**
     * @brief 将字符串加入池中并返回指针
     * @param str 输入字符串
     * @return 指向池中字符串的指针
     */
    const std::string* intern(const std::string& str) {
        auto it = pool_.find(str);
        if (it != pool_.end()) {
            return &(*it);
        }
        
        // 插入新字符串并更新内存使用统计
        auto [inserted_it, success] = pool_.insert(str);
        if (success) {
            memory_usage_ += str.capacity() + sizeof(std::string);
        }
        return &(*inserted_it);
    }
    
    /**
     * @brief 将字符串视图加入池中并返回指针
     * @param str 输入字符串视图
     * @return 指向池中字符串的指针
     */
    const std::string* intern(std::string_view str) {
        // 先查找是否已存在
        std::string str_copy(str);
        auto it = pool_.find(str_copy);
        if (it != pool_.end()) {
            return &(*it);
        }
        
        // 插入新字符串
        auto [inserted_it, success] = pool_.insert(std::move(str_copy));
        if (success) {
            memory_usage_ += str.size() + sizeof(std::string);
        }
        return &(*inserted_it);
    }
    
    /**
     * @brief 检查字符串是否在池中
     * @param str 要检查的字符串
     * @return 如果在池中返回true
     */
    bool contains(const std::string& str) const {
        return pool_.find(str) != pool_.end();
    }
    
    /**
     * @brief 查找池中的字符串指针
     * @param str 要查找的字符串
     * @return 如果找到返回指针，否则返回nullptr
     */
    const std::string* find(const std::string& str) const {
        auto it = pool_.find(str);
        return (it != pool_.end()) ? &(*it) : nullptr;
    }
    
    /**
     * @brief 获取池中字符串数量
     * @return 唯一字符串数量
     */
    size_t size() const noexcept {
        return pool_.size();
    }
    
    /**
     * @brief 获取内存使用量（估算）
     * @return 内存使用字节数
     */
    size_t getMemoryUsage() const noexcept {
        return memory_usage_.load();
    }
    
    /**
     * @brief 获取总内存使用量
     * @return 包含容器开销的总内存使用量
     */
    size_t getTotalMemoryUsage() const noexcept {
        size_t base_memory = memory_usage_.load();
        size_t container_overhead = pool_.bucket_count() * sizeof(void*);
        return base_memory + container_overhead;
    }
    
    /**
     * @brief 清空字符串池
     */
    void clear() {
        pool_.clear();
        memory_usage_ = 0;
    }
    
    /**
     * @brief 收缩内存池
     */
    void shrink() {
        // 对于std::unordered_set，我们可以重建来收缩
        if (pool_.empty()) {
            return;
        }
        
        std::unordered_set<std::string> new_pool;
        new_pool.reserve(pool_.size());
        
        for (auto&& str : pool_) {
            new_pool.insert(std::move(const_cast<std::string&>(str)));
        }
        
        pool_ = std::move(new_pool);
    }
    
    /**
     * @brief 预留空间
     * @param capacity 预期的字符串数量
     */
    void reserve(size_t capacity) {
        pool_.reserve(capacity);
    }
    
    /**
     * @brief 获取负载因子
     * @return 当前负载因子
     */
    double load_factor() const noexcept {
        return pool_.load_factor();
    }
    
    /**
     * @brief 获取最大负载因子
     * @return 最大负载因子
     */
    double max_load_factor() const noexcept {
        return pool_.max_load_factor();
    }
    
    /**
     * @brief 设置最大负载因子
     * @param factor 新的最大负载因子
     */
    void max_load_factor(double factor) {
        pool_.max_load_factor(static_cast<float>(factor));
    }
    
    /**
     * @brief 重新分配哈希表大小
     * @param bucket_count 新的桶数量
     */
    void rehash(size_t bucket_count) {
        pool_.rehash(bucket_count);
    }

private:
    std::unordered_set<std::string> pool_;
    mutable std::atomic<size_t> memory_usage_{0};
};

/**
 * @brief 字符串构建器
 * 
 * 高效的字符串拼接工具，避免频繁的内存分配
 */
class StringBuilder {
public:
    explicit StringBuilder(size_t initial_capacity = 256)
        : result_() {
        result_.reserve(initial_capacity);
    }
    
    /**
     * @brief 添加字符串
     */
    StringBuilder& append(const std::string& str) {
        result_ += str;
        return *this;
    }
    
    StringBuilder& append(std::string_view str) {
        result_ += str;
        return *this;
    }
    
    StringBuilder& append(const char* str) {
        if (str) {
            result_ += str;
        }
        return *this;
    }
    
    StringBuilder& append(char c) {
        result_ += c;
        return *this;
    }
    
    /**
     * @brief 添加数字
     */
    StringBuilder& append(int value) {
        result_ += std::to_string(value);
        return *this;
    }
    
    StringBuilder& append(double value) {
        result_ += std::to_string(value);
        return *this;
    }
    
    /**
     * @brief 重载操作符
     */
    StringBuilder& operator<<(const std::string& str) {
        return append(str);
    }
    
    StringBuilder& operator<<(std::string_view str) {
        return append(str);
    }
    
    StringBuilder& operator<<(const char* str) {
        return append(str);
    }
    
    StringBuilder& operator<<(char c) {
        return append(c);
    }
    
    StringBuilder& operator<<(int value) {
        return append(value);
    }
    
    StringBuilder& operator<<(double value) {
        return append(value);
    }
    
    /**
     * @brief 构建最终字符串
     */
    std::string build() {
        return std::move(result_);
    }
    
    /**
     * @brief 获取当前内容
     */
    const std::string& str() const noexcept {
        return result_;
    }
    
    /**
     * @brief 清空内容
     */
    void clear() {
        result_.clear();
    }
    
    /**
     * @brief 获取当前长度
     */
    size_t size() const noexcept {
        return result_.size();
    }
    
    /**
     * @brief 获取容量
     */
    size_t capacity() const noexcept {
        return result_.capacity();
    }
    
    /**
     * @brief 预留空间
     */
    void reserve(size_t capacity) {
        result_.reserve(capacity);
    }

private:
    std::string result_;
};

}} // namespace fastexcel::memory