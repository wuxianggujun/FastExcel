/**
 * @file FixedSizePool.hpp
 * @brief 固定大小对象内存池实现
 */

#pragma once

#include "AlignedAllocator.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Exception.hpp"
#include <memory>
#include <vector>
#include <mutex>
#include <atomic>
#include <type_traits>

namespace fastexcel {
namespace memory {

/**
 * @brief 固定大小对象的内存池
 * 
 * 专为频繁分配/释放相同大小对象设计，提供高性能的内存管理
 */
template<typename T, size_t PoolSize = 1024>
class FixedSizePool {
    static_assert(sizeof(T) >= sizeof(void*), "Object size too small for pool");
    static_assert(std::is_trivially_destructible_v<T> || std::has_virtual_destructor_v<T>,
                 "Type must be trivially destructible or have virtual destructor");

private:
    // 内存块结构
    struct Block {
        alignas(T) char data[sizeof(T)];
        Block* next = nullptr;
    };
    
    // 内存页结构  
    struct Page {
        std::unique_ptr<Block[]> blocks;
        Page* next = nullptr;
        size_t used_count = 0;
        
        explicit Page(size_t block_count) {
            blocks = std::make_unique<Block[]>(block_count);
            
            // 初始化空闲链表
            for (size_t i = 0; i < block_count - 1; ++i) {
                blocks[i].next = &blocks[i + 1];
            }
            blocks[block_count - 1].next = nullptr;
        }
    };

public:
    /**
     * @brief 构造函数
     */
    FixedSizePool() {
        allocateNewPage();
        FASTEXCEL_LOG_DEBUG("Created FixedSizePool for type {} with pool size {}", 
                           typeid(T).name(), PoolSize);
    }
    
    /**
     * @brief 析构函数
     */
    ~FixedSizePool() {
        cleanup();
        FASTEXCEL_LOG_DEBUG("Destroyed FixedSizePool for type {}. "
                           "Total allocated: {}, Peak usage: {}", 
                           typeid(T).name(), total_allocated_, peak_usage_);
    }
    
    // 禁用拷贝，允许移动
    FixedSizePool(const FixedSizePool&) = delete;
    FixedSizePool& operator=(const FixedSizePool&) = delete;
    
    FixedSizePool(FixedSizePool&& other) noexcept {
        *this = std::move(other);
    }
    
    FixedSizePool& operator=(FixedSizePool&& other) noexcept {
        if (this != &other) {
            cleanup();
            
            std::lock_guard<std::mutex> lock1(mutex_);
            std::lock_guard<std::mutex> lock2(other.mutex_);
            
            pages_ = std::move(other.pages_);
            free_list_ = other.free_list_;
            current_usage_ = other.current_usage_.load();
            peak_usage_ = other.peak_usage_;
            total_allocated_ = other.total_allocated_;
            
            other.free_list_ = nullptr;
            other.current_usage_ = 0;
            other.peak_usage_ = 0;
            other.total_allocated_ = 0;
        }
        return *this;
    }
    
    /**
     * @brief 分配对象
     * @return 新分配的对象指针
     */
    template<typename... Args>
    T* allocate(Args&&... args) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        Block* block = getFreeBlock();
        if (!block) {
            allocateNewPage();
            block = getFreeBlock();
            
            if (!block) {
                throw core::MemoryException(
                    "Failed to allocate memory from pool",
                    sizeof(T),
                    __FILE__, __LINE__
                );
            }
        }
        
        // 构造对象
        T* obj = reinterpret_cast<T*>(block);
        if constexpr (sizeof...(args) > 0) {
            new (obj) T(std::forward<Args>(args)...);
        } else {
            new (obj) T();
        }
        
        // 更新统计信息
        ++current_usage_;
        ++total_allocated_;
        if (current_usage_ > peak_usage_) {
            peak_usage_ = current_usage_.load();
        }
        
        FASTEXCEL_LOG_TRACE("Allocated object from pool, current usage: {}", current_usage_.load());
        
        return obj;
    }
    
    /**
     * @brief 释放对象
     * @param obj 要释放的对象指针
     */
    void deallocate(T* obj) {
        if (!obj) return;
        
        std::lock_guard<std::mutex> lock(mutex_);
        
        // 调用析构函数
        if constexpr (!std::is_trivially_destructible_v<T>) {
            obj->~T();
        }
        
        // 添加到空闲链表
        Block* block = reinterpret_cast<Block*>(obj);
        block->next = free_list_;
        free_list_ = block;
        
        // 更新统计信息
        --current_usage_;
        
        FASTEXCEL_LOG_TRACE("Deallocated object to pool, current usage: {}", current_usage_.load());
    }
    
    /**
     * @brief 获取当前使用量
     */
    size_t getCurrentUsage() const noexcept {
        return current_usage_.load();
    }
    
    /**
     * @brief 获取峰值使用量
     */
    size_t getPeakUsage() const noexcept {
        return peak_usage_;
    }
    
    /**
     * @brief 获取总分配数
     */
    size_t getTotalAllocated() const noexcept {
        return total_allocated_;
    }
    
    /**
     * @brief 预分配内存页
     */
    void preAllocate(size_t page_count = 1) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        for (size_t i = 0; i < page_count; ++i) {
            allocateNewPage();
        }
        
        FASTEXCEL_LOG_DEBUG("Pre-allocated {} pages for pool", page_count);
    }
    
    /**
     * @brief 收缩内存（释放未使用的页面）
     */
    void shrink() {
        std::lock_guard<std::mutex> lock(mutex_);
        
        // 简单实现：只有当使用率很低时才收缩
        if (current_usage_ == 0 && pages_.size() > 1) {
            // 保留一个页面，释放其他页面
            auto first_page = std::move(pages_.front());
            pages_.clear();
            pages_.push_back(std::move(first_page));
            
            // 重建空闲链表
            rebuildFreeList();
            
            FASTEXCEL_LOG_DEBUG("Pool shrunk to 1 page");
        }
    }
    
    /**
     * @brief 清理所有内存
     */
    void clear() {
        cleanup();
    }

private:
    mutable std::mutex mutex_;
    std::vector<std::unique_ptr<Page>> pages_;
    Block* free_list_ = nullptr;
    std::atomic<size_t> current_usage_{0};
    size_t peak_usage_ = 0;
    size_t total_allocated_ = 0;
    
    Block* getFreeBlock() {
        if (free_list_) {
            Block* block = free_list_;
            free_list_ = free_list_->next;
            return block;
        }
        return nullptr;
    }
    
    void allocateNewPage() {
        auto new_page = std::make_unique<Page>(PoolSize);
        
        // 将新页面的空闲块添加到总的空闲链表中
        if (PoolSize > 0) {
            new_page->blocks[PoolSize - 1].next = free_list_;
            free_list_ = &new_page->blocks[0];
        }
        
        pages_.push_back(std::move(new_page));
        
        FASTEXCEL_LOG_DEBUG("Allocated new page for pool, total pages: {}", pages_.size());
    }
    
    void rebuildFreeList() {
        free_list_ = nullptr;
        
        for (auto& page : pages_) {
            // 重建页面内的空闲链表
            for (size_t i = 0; i < PoolSize - 1; ++i) {
                page->blocks[i].next = &page->blocks[i + 1];
            }
            page->blocks[PoolSize - 1].next = free_list_;
            free_list_ = &page->blocks[0];
        }
    }
    
    void cleanup() noexcept {
        try {
            std::lock_guard<std::mutex> lock(mutex_);
            pages_.clear();
            free_list_ = nullptr;
        } catch (...) {
            // 忽略清理时的异常
        }
    }
};

}} // namespace fastexcel::memory