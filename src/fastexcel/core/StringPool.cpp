/**
 * @file StringPool.cpp
 * @brief 字符串内存池实现
 */

#include "StringPool.hpp"
#include <algorithm>
#include <cstring>

namespace fastexcel {
namespace core {

StringPool::StringPool(size_t initial_capacity) 
    : string_count_(0) {
    buffer_.reserve(initial_capacity);
}

std::string_view StringPool::addString(std::string_view str) {
    if (str.empty()) {
        return std::string_view{};
    }
    
    // 确保有足够空间
    ensureCapacity(str.size());
    
    // 记录起始位置
    size_t start_pos = buffer_.size();
    
    // 添加字符串到缓冲区
    buffer_.append(str.data(), str.size());
    
    // 增加计数
    ++string_count_;
    
    // 返回指向缓冲区中字符串的 string_view
    return std::string_view(buffer_.data() + start_pos, str.size());
}

std::string_view StringPool::addString(std::string&& str) {
    if (str.empty()) {
        return std::string_view{};
    }
    
    // 确保有足够空间
    ensureCapacity(str.size());
    
    // 记录起始位置
    size_t start_pos = buffer_.size();
    
    // 移动字符串内容到缓冲区
    buffer_.append(std::move(str));
    
    // 增加计数
    ++string_count_;
    
    // 返回指向缓冲区中字符串的 string_view
    return std::string_view(buffer_.data() + start_pos, str.size());
}

void StringPool::clear() {
    buffer_.clear();
    string_count_ = 0;
}

void StringPool::ensureCapacity(size_t needed_size) {
    size_t required_capacity = buffer_.size() + needed_size;
    if (required_capacity > buffer_.capacity()) {
        // 以 1.5x 的速度增长
        size_t new_capacity = std::max(required_capacity, buffer_.capacity() * 3 / 2);
        buffer_.reserve(new_capacity);
    }
}

}} // namespace fastexcel::core