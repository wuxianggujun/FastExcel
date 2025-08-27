/**
 * @file StringPool.hpp
 * @brief 字符串内存池优化
 */

#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <cstddef>

namespace fastexcel {
namespace core {

/**
 * @brief 高性能字符串池
 * 
 * 专门优化频繁的小字符串分配，特别适合解析场景：
 * - 预分配大块内存
 * - 避免频繁的 malloc/free
 * - 提供 string_view 接口避免复制
 */
class StringPool {
public:
    explicit StringPool(size_t initial_capacity = 1024 * 1024); // 1MB 初始容量
    ~StringPool() = default;
    
    /**
     * @brief 添加字符串到池中
     * @param str 要添加的字符串
     * @return 指向池中字符串的 string_view
     */
    std::string_view addString(std::string_view str);
    
    /**
     * @brief 添加字符串到池中（移动语义）
     * @param str 要移动的字符串
     * @return 指向池中字符串的 string_view
     */
    std::string_view addString(std::string&& str);
    
    /**
     * @brief 清空池
     */
    void clear();
    
    /**
     * @brief 获取当前使用的内存大小
     */
    size_t getUsedMemory() const { return buffer_.size(); }
    
    /**
     * @brief 获取字符串数量
     */
    size_t getStringCount() const { return string_count_; }
    
private:
    std::string buffer_;       // 主缓冲区
    size_t string_count_;      // 字符串数量统计
    
    // 确保有足够空间
    void ensureCapacity(size_t needed_size);
};

}} // namespace fastexcel::core