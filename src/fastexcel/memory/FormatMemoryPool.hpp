/**
 * @file FormatMemoryPool.hpp
 * @brief 专门用于管理FormatDescriptor对象的内存池
 */

#pragma once

#include "MemoryPoolOptimized.hpp"
#include "LazyInitializer.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include <memory>
#include <atomic>

namespace fastexcel {
namespace memory {

/**
 * @brief FormatDescriptor对象专用内存池管理器
 * 
 * 提供高效的FormatDescriptor对象分配和回收，针对Excel工作簿中
 * 格式描述符对象的创建和共享场景进行优化
 */
class FormatMemoryPool {
private:
    static constexpr size_t POOL_SIZE = 512;
    
    // 使用LazyInitializer延迟初始化内存池
    LazyInitializer<FixedSizePool<core::FormatDescriptor, POOL_SIZE>> pool_;
    
    // 统计信息
    std::atomic<size_t> total_allocations_{0};
    std::atomic<size_t> total_deallocations_{0};
    
public:
    FormatMemoryPool() = default;
    ~FormatMemoryPool() = default;
    
    // 禁止拷贝和移动
    FormatMemoryPool(const FormatMemoryPool&) = delete;
    FormatMemoryPool& operator=(const FormatMemoryPool&) = delete;
    FormatMemoryPool(FormatMemoryPool&&) = delete;
    FormatMemoryPool& operator=(FormatMemoryPool&&) = delete;
    
    /**
     * @brief 分配一个FormatDescriptor对象
     * @param args FormatDescriptor构造函数参数
     * @return 分配的FormatDescriptor对象指针
     */
    template<typename... Args>
    core::FormatDescriptor* allocate(Args&&... args) {
        if (!pool_.isInitialized()) {
            pool_.initialize();
        }
        
        core::FormatDescriptor* format = pool_.get().allocate(std::forward<Args>(args)...);
        ++total_allocations_;
        return format;
    }
    
    /**
     * @brief 创建智能指针管理的FormatDescriptor对象
     * @param args FormatDescriptor构造函数参数
     * @return unique_ptr<FormatDescriptor>
     */
    template<typename... Args>
    std::unique_ptr<core::FormatDescriptor> createFormat(Args&&... args) {
        core::FormatDescriptor* format = allocate(std::forward<Args>(args)...);
        return std::unique_ptr<core::FormatDescriptor>(format);
    }
    
    /**
     * @brief 基于默认格式创建新的格式描述符
     * @return unique_ptr<FormatDescriptor>
     */
    std::unique_ptr<core::FormatDescriptor> createDefaultFormat() {
        const auto& defaultFormat = core::FormatDescriptor::getDefault();
        return createFormat(
            defaultFormat.getFontName(),
            defaultFormat.getFontSize(),
            defaultFormat.isBold(),
            defaultFormat.isItalic(),
            defaultFormat.getUnderline(),
            defaultFormat.isStrikeout(),
            defaultFormat.getFontScript(),
            defaultFormat.getFontColor(),
            defaultFormat.getFontFamily(),
            defaultFormat.getFontCharset(),
            defaultFormat.getHorizontalAlign(),
            defaultFormat.getVerticalAlign(),
            defaultFormat.isTextWrap(),
            defaultFormat.getRotation(),
            defaultFormat.getIndent(),
            defaultFormat.isShrink(),
            defaultFormat.getLeftBorder(),
            defaultFormat.getRightBorder(),
            defaultFormat.getTopBorder(),
            defaultFormat.getBottomBorder(),
            defaultFormat.getDiagBorder(),
            defaultFormat.getDiagType(),
            defaultFormat.getLeftBorderColor(),
            defaultFormat.getRightBorderColor(),
            defaultFormat.getTopBorderColor(),
            defaultFormat.getBottomBorderColor(),
            defaultFormat.getDiagBorderColor(),
            defaultFormat.getPattern(),
            defaultFormat.getBackgroundColor(),
            defaultFormat.getForegroundColor(),
            defaultFormat.getNumberFormat(),
            defaultFormat.getNumberFormatIndex(),
            defaultFormat.isLocked(),
            defaultFormat.isHidden()
        );
    }
    
    /**
     * @brief 释放FormatDescriptor对象
     * @param format 要释放的FormatDescriptor对象指针
     */
    void deallocate(core::FormatDescriptor* format) {
        if (pool_.isInitialized() && format) {
            pool_.get().deallocate(format);
            ++total_deallocations_;
        }
    }
    
    /**
     * @brief 获取当前内存使用统计
     */
    struct Statistics {
        size_t current_usage = 0;
        size_t peak_usage = 0;
        size_t total_allocations = 0;
        size_t total_deallocations = 0;
        size_t active_objects = 0;
    };
    
    Statistics getStatistics() const {
        Statistics stats;
        stats.total_allocations = total_allocations_.load();
        stats.total_deallocations = total_deallocations_.load();
        stats.active_objects = stats.total_allocations - stats.total_deallocations;
        
        if (pool_.isInitialized()) {
            stats.current_usage = pool_.get().getCurrentUsage();
            stats.peak_usage = pool_.get().getPeakUsage();
        }
        
        return stats;
    }
    
    /**
     * @brief 收缩内存池，释放未使用的内存
     */
    void shrink() {
        if (pool_.isInitialized()) {
            pool_.get().shrink();
        }
    }
    
    /**
     * @brief 预分配指定数量的对象空间
     * @param count 要预分配的对象数量
     */
    void reserve(size_t /*count*/) {
        if (!pool_.isInitialized()) {
            pool_.initialize();
        }
        // 内存池会自动扩展，这里可以预触发扩展
    }
    
    /**
     * @brief 清理所有分配的内存
     */
    void clear() {
        // 重置统计信息
        total_allocations_ = 0;
        total_deallocations_ = 0;
        
        if (pool_.isInitialized()) {
            pool_.get().clear();
        }
    }
};

}} // namespace fastexcel::memory