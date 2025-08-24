/**
 * @file CellMemoryPool.hpp
 * @brief 专门用于管理Cell对象的内存池
 */

#pragma once

#include "MemoryPoolOptimized.hpp"
#include "LazyInitializer.hpp"
#include "fastexcel/core/Cell.hpp"
#include <memory>
#include <atomic>

namespace fastexcel {
namespace memory {

/**
 * @brief Cell对象专用内存池管理器
 * 
 * 提供高效的Cell对象分配和回收，针对Excel工作簿中
 * 大量Cell对象的创建和销毁场景进行优化
 */
class CellMemoryPool {
private:
    static constexpr size_t POOL_SIZE = 2048;
    
    // 使用LazyInitializer延迟初始化内存池
    LazyInitializer<FixedSizePool<core::Cell, POOL_SIZE>> pool_;
    
    // 统计信息
    std::atomic<size_t> total_allocations_{0};
    std::atomic<size_t> total_deallocations_{0};
    
public:
    CellMemoryPool() = default;
    ~CellMemoryPool() = default;
    
    // 禁止拷贝和移动
    CellMemoryPool(const CellMemoryPool&) = delete;
    CellMemoryPool& operator=(const CellMemoryPool&) = delete;
    CellMemoryPool(CellMemoryPool&&) = delete;
    CellMemoryPool& operator=(CellMemoryPool&&) = delete;
    
    /**
     * @brief 分配一个Cell对象
     * @param args Cell构造函数参数
     * @return 分配的Cell对象指针
     */
    template<typename... Args>
    core::Cell* allocate(Args&&... args) {
        if (!pool_.isInitialized()) {
            pool_.initialize();
        }
        
        core::Cell* cell = pool_.get().allocate(std::forward<Args>(args)...);
        ++total_allocations_;
        return cell;
    }
    
    /**
     * @brief 创建智能指针管理的Cell对象
     * @param args Cell构造函数参数
     * @return unique_ptr<Cell>
     */
    template<typename... Args>
    std::unique_ptr<core::Cell> createCell(Args&&... args) {
        core::Cell* cell = allocate(std::forward<Args>(args)...);
        return std::unique_ptr<core::Cell>(cell);
    }
    
    /**
     * @brief 释放Cell对象（由智能指针或手动调用）
     * @param cell 要释放的Cell对象指针
     */
    void deallocate(core::Cell* cell) {
        if (pool_.isInitialized() && cell) {
            pool_.get().deallocate(cell);
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
        // 具体实现依据FixedSizePool的接口
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