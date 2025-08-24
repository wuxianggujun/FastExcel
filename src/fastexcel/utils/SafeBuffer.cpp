#include "SafeBuffer.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace utils {

DynamicSafeBuffer::DynamicSafeBuffer(
    size_t initial_capacity,
    size_t max_capacity,
    FlushCallback flush_callback
) : max_capacity_(max_capacity)
  , flush_callback_(std::move(flush_callback)) {
    
    if (initial_capacity > max_capacity) {
        initial_capacity = max_capacity;
    }
    
    buffer_.reserve(initial_capacity);
    pos_ = 0;
    
    FASTEXCEL_LOG_DEBUG("DynamicSafeBuffer created with initial_capacity={}, max_capacity={}", 
                       initial_capacity, max_capacity);
}

bool DynamicSafeBuffer::append(const char* data, size_t length) {
    if (!data || length == 0) {
        return true;
    }
    
    // 检查是否超过最大容量限制
    if (pos_ + length > max_capacity_) {
        if (flush_callback_) {
            // 尝试刷新后再添加
            flush();
            if (length > max_capacity_) {
                FASTEXCEL_LOG_ERROR("Data size {} exceeds maximum buffer capacity {}", 
                                   length, max_capacity_);
                return false;
            }
        } else {
            FASTEXCEL_LOG_ERROR("Buffer overflow: current={}, length={}, max={}", 
                               pos_, length, max_capacity_);
            return false;
        }
    }
    
    // 确保有足够容量
    if (!ensureCapacity(pos_ + length)) {
        return false;
    }
    
    // 复制数据
    std::memcpy(buffer_.data() + pos_, data, length);
    pos_ += length;
    
    return true;
}

void DynamicSafeBuffer::flush() {
    if (pos_ > 0 && flush_callback_) {
        flush_callback_(buffer_.data(), pos_);
        pos_ = 0;
    }
}

void DynamicSafeBuffer::clear() noexcept {
    pos_ = 0;
    // 不清理buffer_内容，只重置位置
}

bool DynamicSafeBuffer::ensureCapacity(size_t required_size) {
    if (required_size <= buffer_.size()) {
        return true;
    }
    
    if (required_size > max_capacity_) {
        FASTEXCEL_LOG_ERROR("Required size {} exceeds maximum capacity {}", 
                           required_size, max_capacity_);
        return false;
    }
    
    // 计算新容量：至少是当前容量的1.5倍，但不超过最大容量
    size_t new_capacity = std::max(required_size, buffer_.size() * 3 / 2);
    new_capacity = std::min(new_capacity, max_capacity_);
    
    try {
        buffer_.resize(new_capacity);
        FASTEXCEL_LOG_DEBUG("Buffer resized from {} to {}", buffer_.size(), new_capacity);
        return true;
    } catch (const std::bad_alloc& e) {
        FASTEXCEL_LOG_ERROR("Failed to resize buffer to {}: {}", new_capacity, e.what());
        return false;
    }
}

} // namespace utils
} // namespace fastexcel