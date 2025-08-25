/**
 * @file SafeBuffer.hpp
 * @brief 内存安全的缓冲区管理器
 */

#pragma once

#include <array>
#include <cstring>
#include <stdexcept>
#include <functional>
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Constants.hpp"

namespace fastexcel {
namespace utils {

/**
 * @brief 内存安全的固定大小缓冲区
 * 
 * 提供边界检查、自动刷新等安全特性
 * @tparam BufferSize 缓冲区大小
 */
template<size_t BufferSize = fastexcel::core::Constants::kIOBufferSize>
class SafeBuffer {
public:
    using FlushCallback = std::function<void(const char* data, size_t length)>;
    
    /**
     * @brief 构造函数
     * @param flush_callback 缓冲区满时的刷新回调
     * @param auto_flush 是否自动刷新
     */
    explicit SafeBuffer(FlushCallback flush_callback = nullptr, bool auto_flush = true)
        : flush_callback_(std::move(flush_callback))
        , auto_flush_(auto_flush) {
        clear();
    }
    
    /**
     * @brief 析构函数，自动刷新剩余数据
     */
    ~SafeBuffer() {
        if (pos_ > 0 && flush_callback_) {
            try {
                flush();
            } catch (...) {
                // 析构函数中忽略异常
            }
        }
    }
    
    // 禁用拷贝，允许移动
    SafeBuffer(const SafeBuffer&) = delete;
    SafeBuffer& operator=(const SafeBuffer&) = delete;
    
    SafeBuffer(SafeBuffer&& other) noexcept
        : buffer_(std::move(other.buffer_))
        , pos_(other.pos_)
        , flush_callback_(std::move(other.flush_callback_))
        , auto_flush_(other.auto_flush_) {
        other.pos_ = 0;
    }
    
    SafeBuffer& operator=(SafeBuffer&& other) noexcept {
        if (this != &other) {
            flush(); // 刷新当前缓冲区
            buffer_ = std::move(other.buffer_);
            pos_ = other.pos_;
            flush_callback_ = std::move(other.flush_callback_);
            auto_flush_ = other.auto_flush_;
            other.pos_ = 0;
        }
        return *this;
    }
    
    /**
     * @brief 追加数据到缓冲区
     * @param data 数据指针
     * @param length 数据长度
     * @return 是否成功追加
     */
    bool append(const char* data, size_t length) {
        if (!data || length == 0) {
            return true;
        }
        
        // 检查单次写入是否超过缓冲区大小
        if (length > BufferSize) {
            if (auto_flush_ && flush_callback_) {
                // 先刷新现有数据
                flush();
                // 直接写入大数据
                flush_callback_(data, length);
                return true;
            } else {
                throw core::MemoryException(
                    "Data size exceeds buffer capacity",
                    length,
                    __FILE__, __LINE__
                );
            }
        }
        
        // 检查是否需要刷新
        if (pos_ + length > BufferSize) {
            if (auto_flush_ && flush_callback_) {
                flush();
            } else {
                return false; // 缓冲区满
            }
        }
        
        // 复制数据到缓冲区
        std::memcpy(buffer_.data() + pos_, data, length);
        pos_ += length;
        
        return true;
    }
    
    /**
     * @brief 追加字符串到缓冲区
     */
    bool append(const std::string& str) {
        return append(str.c_str(), str.size());
    }
    
    /**
     * @brief 追加单个字符到缓冲区
     */
    bool append(char c) {
        return append(&c, 1);
    }
    
    /**
     * @brief 强制刷新缓冲区
     */
    void flush() {
        if (pos_ > 0 && flush_callback_) {
            flush_callback_(buffer_.data(), pos_);
            pos_ = 0;
        }
    }
    
    /**
     * @brief 清空缓冲区（不刷新）
     */
    void clear() noexcept {
        pos_ = 0;
        // 移除不必要的清零操作，性能优化
    }
    
    /**
     * @brief 获取当前数据大小
     */
    size_t size() const noexcept {
        return pos_;
    }
    
    /**
     * @brief 获取剩余空间
     */
    size_t remaining() const noexcept {
        return BufferSize - pos_;
    }
    
    /**
     * @brief 检查缓冲区是否为空
     */
    bool empty() const noexcept {
        return pos_ == 0;
    }
    
    /**
     * @brief 检查缓冲区是否已满
     */
    bool full() const noexcept {
        return pos_ >= BufferSize;
    }
    
    /**
     * @brief 获取缓冲区容量
     */
    static constexpr size_t capacity() noexcept {
        return BufferSize;
    }
    
    /**
     * @brief 设置刷新回调
     */
    void setFlushCallback(FlushCallback callback) {
        flush_callback_ = std::move(callback);
    }
    
    /**
     * @brief 设置自动刷新模式
     */
    void setAutoFlush(bool auto_flush) {
        auto_flush_ = auto_flush;
    }
    
    /**
     * @brief 获取缓冲区数据（只读）
     */
    const char* data() const noexcept {
        return buffer_.data();
    }

private:
    std::array<char, BufferSize> buffer_;
    size_t pos_ = 0;
    FlushCallback flush_callback_;
    bool auto_flush_ = true;
};


} // namespace utils
} // namespace fastexcel