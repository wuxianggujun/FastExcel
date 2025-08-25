/**
 * @file FixedSizePool.hpp
 * @brief 固定大小对象内存池实现
 */

#pragma once

#include "AlignedAllocator.hpp"
#include "IMemoryPool.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Exception.hpp"
#include <memory>
#include <vector>
#include <mutex>
#include <atomic>
#include <type_traits>
#include <thread>

namespace fastexcel {
namespace memory {

/**
 * @brief 固定大小对象的内存池
 * 
 * 专为频繁分配/释放相同大小对象设计，提供高性能的内存管理
 */
template<typename T, size_t PoolSize = 1024>
class FixedSizePool : public IMemoryPool {
    static_assert(sizeof(T) >= sizeof(void*), "Object size too small for pool");
    static_assert(std::is_trivially_destructible_v<T> || std::has_virtual_destructor_v<T>,
                 "Type must be trivially destructible or have virtual destructor");

private:
    // 前向声明
    struct Block;
    
    // 内存块结构（必须先定义）
    struct Block {
        alignas(T) char data[sizeof(T)];
        std::atomic<Block*> next{nullptr};
        
        // 转换为原始指针用于链表操作
        Block* getNext() const { return next.load(); }
        void setNext(Block* n) { next.store(n); }
    };
    
    // 无锁的原子栈结构
    struct AtomicStack {
        std::atomic<Block*> head{nullptr};
        
        void push(Block* block) {
            Block* old_head = head.load();
            do {
                block->setNext(old_head);
            } while (!head.compare_exchange_weak(old_head, block));
        }
        
        Block* pop() {
            Block* old_head = head.load();
            while (old_head && !head.compare_exchange_weak(old_head, old_head->getNext())) {
                // 重试直到成功或栈为空
            }
            return old_head;
        }
        
        bool empty() const {
            return head.load() == nullptr;
        }
    };
    
    // 线程本地缓存
    struct ThreadLocalCache {
        static constexpr size_t CACHE_SIZE = 64; // 每个线程缓存的块数量
        Block* local_free_list = nullptr;
        size_t cache_count = 0;
        
        Block* getBlock() {
            if (local_free_list) {
                Block* block = local_free_list;
                local_free_list = local_free_list->getNext();
                --cache_count;
                return block;
            }
            return nullptr;
        }
        
        bool returnBlock(Block* block) {
            if (cache_count < CACHE_SIZE) {
                block->setNext(local_free_list);
                local_free_list = block;
                ++cache_count;
                return true;
            }
            return false; // 缓存已满
        }
        
        void flushToGlobal(AtomicStack& global_stack) {
            while (local_free_list && cache_count > CACHE_SIZE / 2) {
                Block* block = local_free_list;
                local_free_list = local_free_list->getNext();
                global_stack.push(block);
                --cache_count;
            }
        }
    };
    
    // 内存页结构  
    struct Page {
        std::unique_ptr<Block[]> blocks;
        Page* next = nullptr;
        std::atomic<size_t> used_count{0};
        
        explicit Page(size_t block_count) {
            blocks = std::make_unique<Block[]>(block_count);
            
            // 初始化空闲链表（非原子版本用于初始化）
            for (size_t i = 0; i < block_count - 1; ++i) {
                blocks[i].setNext(&blocks[i + 1]);
            }
            blocks[block_count - 1].setNext(nullptr);
        }
        
        // 将页面的所有块添加到全局栈
        void addToGlobalStack(AtomicStack& global_stack) {
            for (size_t i = 0; i < PoolSize; ++i) {
                global_stack.push(&blocks[i]);
            }
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
                           typeid(T).name(), total_allocated_.load(), peak_usage_.load());
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
            
            std::lock_guard<std::mutex> lock1(pages_mutex_);
            std::lock_guard<std::mutex> lock2(other.pages_mutex_);
            
            // 移动页面和统计数据
            pages_ = std::move(other.pages_);
            current_usage_ = other.current_usage_.load();
            peak_usage_ = other.peak_usage_.load();
            total_allocated_ = other.total_allocated_.load();
            
            // 移动全局栈内容（简单方式：重建）
            while (Block* block = other.global_free_stack_.pop()) {
                global_free_stack_.push(block);
            }
            
            // 重置源对象
            other.current_usage_ = 0;
            other.peak_usage_ = 0;
            other.total_allocated_ = 0;
        }
        return *this;
    }
    
    /**
     * @brief 分配对象（无锁优化版本）
     * @return 新分配的对象指针
     */
    template<typename... Args>
    T* allocate(Args&&... args) {
        // 1. 尝试从线程本地缓存获取
        Block* block = thread_cache_.getBlock();
        if (block) {
            cache_hits_.fetch_add(1); // 缓存命中
        } else {
            cache_misses_.fetch_add(1); // 缓存未命中
            
            // 2. 如果本地缓存为空，从全局栈获取
            block = global_free_stack_.pop();
        }
        
        // 3. 如果全局栈也为空，分配新页面
        if (!block) {
            allocateNewPage();
            block = global_free_stack_.pop();
            page_allocations_.fetch_add(1); // 记录页面分配
            
            if (!block) {
                throw core::MemoryException(
                    "Failed to allocate memory from pool",
                    sizeof(T),
                    __FILE__, __LINE__
                );
            }
        }
        
        // 4. 构造对象
        T* obj = reinterpret_cast<T*>(block);
        if constexpr (sizeof...(args) > 0) {
            new (obj) T(std::forward<Args>(args)...);
        } else {
            new (obj) T();
        }
        
        // 5. 更新统计信息（原子操作）
        size_t current = current_usage_.fetch_add(1) + 1;
        size_t total_allocs = total_allocated_.fetch_add(1) + 1;
        
        // 更新峰值使用量（无锁CAS）
        size_t expected_peak = peak_usage_.load();
        while (current > expected_peak && 
               !peak_usage_.compare_exchange_weak(expected_peak, current)) {
            // 重试直到成功
        }
        
        // 6. 动态调整检查（定期执行）
        if (total_allocs % SHRINK_CHECK_INTERVAL == 0) {
            performDynamicAdjustment();
        }
        
        return obj;
    }

    
    /**
     * @brief 释放对象（无锁优化版本）
     * @param obj 要释放的对象指针
     */
    void deallocate(T* obj) {
        if (!obj) return;
        
        // 1. 调用析构函数
        if constexpr (!std::is_trivially_destructible_v<T>) {
            obj->~T();
        }
        
        Block* block = reinterpret_cast<Block*>(obj);
        
        // 2. 尝试放入线程本地缓存
        if (!thread_cache_.returnBlock(block)) {
            // 3. 本地缓存已满，刷新部分到全局栈，然后放入全局栈
            thread_cache_.flushToGlobal(global_free_stack_);
            global_free_stack_.push(block);
        }
        
        // 4. 更新统计信息
        current_usage_.fetch_sub(1);
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
        std::lock_guard<std::mutex> lock(pages_mutex_);
        
        for (size_t i = 0; i < page_count; ++i) {
            allocateNewPage();
        }
        
        FASTEXCEL_LOG_DEBUG("Pre-allocated {} pages for pool", page_count);
    }
    
    /**
     * @brief 收缩内存（释放未使用的页面）
     */
    void shrink() {
        std::lock_guard<std::mutex> lock(pages_mutex_);
        
        // 简单实现：只有当使用率很低时才收缩
        if (current_usage_ == 0 && pages_.size() > 1) {
            // 清空全局栈（因为要释放页面）
            while (Block* block = global_free_stack_.pop()) {
                // 清空栈中的所有块
            }
            
            // 保留一个页面，释放其他页面
            auto first_page = std::move(pages_.front());
            pages_.clear();
            pages_.push_back(std::move(first_page));
            
            // 重建全局空闲栈 - 将保留页面的所有块重新添加到栈中
            first_page->addToGlobalStack(global_free_stack_);
            
            FASTEXCEL_LOG_DEBUG("Pool shrunk to 1 page");
        }
    }
    
    /**
     * @brief 清理所有内存
     */
    void clear() {
        cleanup();
    }

    /**
     * @brief 执行动态调整策略
     */
    void performDynamicAdjustment() {
        size_t current = current_usage_.load();
        size_t total_capacity = pages_.size() * PoolSize;
        
        if (total_capacity == 0) return;
        
        double usage_ratio = static_cast<double>(current) / total_capacity;
        
        // 低使用率时收缩
        if (usage_ratio < LOW_USAGE_THRESHOLD && pages_.size() > 1) {
            FASTEXCEL_LOG_DEBUG("Pool usage low ({:.2f}%), attempting shrink", usage_ratio * 100);
            shrink();
        }
        // 高使用率时预分配
        else if (usage_ratio > HIGH_USAGE_THRESHOLD) {
            FASTEXCEL_LOG_DEBUG("Pool usage high ({:.2f}%), pre-allocating page", usage_ratio * 100);
            preAllocate(1);
        }
        
        last_shrink_check_.store(total_allocated_.load());
    }

    /**
     * @brief 获取详细的性能统计信息
     */
    struct DetailedStatistics {
        size_t current_usage = 0;
        size_t peak_usage = 0;
        size_t total_allocated = 0;
        size_t total_deallocated = 0;
        size_t active_objects = 0;
        size_t cache_hits = 0;
        size_t cache_misses = 0;
        size_t cache_hit_rate_percent = 0;
        size_t page_allocations = 0;
        size_t contention_count = 0;
        size_t pages_count = 0;
        size_t total_capacity = 0;
        size_t usage_percent = 0;
        size_t memory_overhead_bytes = 0;
    };

    DetailedStatistics getDetailedStatistics() const {
        DetailedStatistics stats;
        
        stats.current_usage = current_usage_.load();
        stats.peak_usage = peak_usage_.load();
        stats.total_allocated = total_allocated_.load();
        stats.total_deallocated = stats.total_allocated - stats.current_usage;
        stats.active_objects = stats.current_usage;
        
        stats.cache_hits = cache_hits_.load();
        stats.cache_misses = cache_misses_.load();
        size_t total_accesses = stats.cache_hits + stats.cache_misses;
        stats.cache_hit_rate_percent = total_accesses > 0 ? 
            (stats.cache_hits * 100) / total_accesses : 0;
        
        stats.page_allocations = page_allocations_.load();
        stats.contention_count = contention_count_.load();
        
        {
            std::lock_guard<std::mutex> lock(pages_mutex_);
            stats.pages_count = pages_.size();
        }
        
        stats.total_capacity = stats.pages_count * PoolSize;
        stats.usage_percent = stats.total_capacity > 0 ? 
            (stats.current_usage * 100) / stats.total_capacity : 0;
        
        // 计算内存开销（元数据、管理结构等）
        stats.memory_overhead_bytes = sizeof(FixedSizePool) + 
            stats.pages_count * sizeof(Page) + 
            stats.pages_count * sizeof(std::unique_ptr<Page>);
        
        return stats;
    }

    /**
     * @brief 打印详细的性能报告
     */
    void printPerformanceReport() const {
        auto stats = getDetailedStatistics();
        
        FASTEXCEL_LOG_INFO("=== FixedSizePool Performance Report ===");
        FASTEXCEL_LOG_INFO("Object type: {}", typeid(T).name());
        FASTEXCEL_LOG_INFO("Object size: {} bytes", sizeof(T));
        FASTEXCEL_LOG_INFO("Pool size per page: {}", PoolSize);
        
        FASTEXCEL_LOG_INFO("Memory Usage:");
        FASTEXCEL_LOG_INFO("  Current usage: {} objects ({} bytes)", 
                          stats.current_usage, stats.current_usage * sizeof(T));
        FASTEXCEL_LOG_INFO("  Peak usage: {} objects ({} bytes)", 
                          stats.peak_usage, stats.peak_usage * sizeof(T));
        FASTEXCEL_LOG_INFO("  Total capacity: {} objects ({} bytes)", 
                          stats.total_capacity, stats.total_capacity * sizeof(T));
        FASTEXCEL_LOG_INFO("  Usage ratio: {}%", stats.usage_percent);
        FASTEXCEL_LOG_INFO("  Memory overhead: {} bytes", stats.memory_overhead_bytes);
        
        FASTEXCEL_LOG_INFO("Allocation Statistics:");
        FASTEXCEL_LOG_INFO("  Total allocated: {} objects", stats.total_allocated);
        FASTEXCEL_LOG_INFO("  Total deallocated: {} objects", stats.total_deallocated);
        FASTEXCEL_LOG_INFO("  Active objects: {} objects", stats.active_objects);
        FASTEXCEL_LOG_INFO("  Pages allocated: {} pages", stats.page_allocations);
        
        FASTEXCEL_LOG_INFO("Cache Performance:");
        FASTEXCEL_LOG_INFO("  Cache hits: {}", stats.cache_hits);
        FASTEXCEL_LOG_INFO("  Cache misses: {}", stats.cache_misses);
        FASTEXCEL_LOG_INFO("  Cache hit rate: {}%", stats.cache_hit_rate_percent);
        
        FASTEXCEL_LOG_INFO("Threading:");
        FASTEXCEL_LOG_INFO("  Lock contention count: {}", stats.contention_count);
        FASTEXCEL_LOG_INFO("==========================================");
    }

    /**
     * @brief 内存预热 - 预分配并初始化指定数量的对象
     */
    void warmUp(size_t object_count = PoolSize / 2) {
        FASTEXCEL_LOG_INFO("Warming up memory pool with {} objects", object_count);
        
        // 预分配足够的页面
        size_t pages_needed = (object_count + PoolSize - 1) / PoolSize;
        preAllocate(pages_needed);
        
        // 对于不能默认构造的类型，仅预分配页面即可
        // 不进行对象的实际构造和析构操作
        if constexpr (std::is_default_constructible_v<T>) {
            // 预分配对象然后立即释放，以填充缓存
            std::vector<T*> temp_objects;
            temp_objects.reserve(object_count);
            
            try {
                for (size_t i = 0; i < object_count; ++i) {
                    temp_objects.push_back(allocate());
                }
                
                // 释放所有对象，填充各级缓存
                for (auto* obj : temp_objects) {
                    deallocate(obj);
                }
                
                FASTEXCEL_LOG_INFO("Memory pool warm-up completed with object construction");
            } catch (const std::exception& e) {
                FASTEXCEL_LOG_ERROR("Memory pool warm-up failed: {}", e.what());
                // 清理已分配的对象
                for (auto* obj : temp_objects) {
                    if (obj) deallocate(obj);
                }
            }
        } else {
            FASTEXCEL_LOG_INFO("Memory pool warm-up completed (page pre-allocation only)");
        }
    }

    // IMemoryPool 原始接口实现
    void* allocate(std::size_t size, std::size_t alignment = alignof(std::max_align_t)) override {
        // 仅支持不超过 T 大小的分配，且对齐不超过 T 的对齐
        if (size > sizeof(T) || alignment > alignof(T)) {
            // 回退到通用分配，交由 deallocate 判断是否来自池
            return AlignedAllocator::allocate(alignment, size);
        }

        // 使用新的无锁分配逻辑
        Block* block = global_free_stack_.pop();
        
        // 如果全局栈为空，分配新页面
        if (!block) {
            allocateNewPage();
            block = global_free_stack_.pop();
            if (!block) return nullptr;
        }
        
        // 更新统计信息（原子操作）
        current_usage_.fetch_add(1);
        total_allocated_.fetch_add(1);
        
        // 更新峰值使用量（无锁CAS）
        size_t current = current_usage_.load();
        size_t expected_peak = peak_usage_.load();
        while (current > expected_peak && 
               !peak_usage_.compare_exchange_weak(expected_peak, current)) {
            // 重试直到成功
        }
        
        return reinterpret_cast<void*>(block);
    }

    void deallocate(void* ptr, std::size_t /*size*/, std::size_t /*alignment*/ = alignof(std::max_align_t)) override {
        if (!ptr) return;

        // 判断指针是否来自本池的页面
        if (isFromThisPool(ptr)) {
            Block* block = reinterpret_cast<Block*>(ptr);
            global_free_stack_.push(block);
            current_usage_.fetch_sub(1);
        } else {
            // 外部分配，直接释放
            AlignedAllocator::deallocate(ptr);
        }
    }

    IMemoryPool::Statistics getStatistics() const override {
        IMemoryPool::Statistics s{};
        s.current_usage = current_usage_.load() * sizeof(T);
        s.peak_usage = peak_usage_ * sizeof(T);
        s.total_allocations = total_allocated_ * sizeof(T);
        s.total_deallocations = (total_allocated_ - current_usage_.load()) * sizeof(T);
        return s;
    }

private:
    // 全局无锁空闲块栈
    AtomicStack global_free_stack_;
    
    // 页面管理（仍需要锁保护）
    mutable std::mutex pages_mutex_;
    std::vector<std::unique_ptr<Page>> pages_;
    
    // 线程本地缓存
    thread_local static ThreadLocalCache thread_cache_;
    
    // 统计信息（原子操作）
    std::atomic<size_t> current_usage_{0};
    std::atomic<size_t> peak_usage_{0};
    std::atomic<size_t> total_allocated_{0};
    
    // 高级性能监控
    std::atomic<size_t> cache_hits_{0};        // 线程本地缓存命中数
    std::atomic<size_t> cache_misses_{0};      // 缓存未命中数
    std::atomic<size_t> page_allocations_{0};  // 页面分配次数
    std::atomic<size_t> contention_count_{0};  // 锁竞争次数
    
    // 动态调整参数
    std::atomic<size_t> last_shrink_check_{0}; // 上次收缩检查时间
    static constexpr size_t SHRINK_CHECK_INTERVAL = 10000; // 检查间隔（分配次数）
    static constexpr double LOW_USAGE_THRESHOLD = 0.1;     // 低使用率阈值
    static constexpr double HIGH_USAGE_THRESHOLD = 0.8;    // 高使用率阈值
    
    // 分配新页面的方法
    void allocateNewPage() {
        std::lock_guard<std::mutex> lock(pages_mutex_);
        
        auto new_page = std::make_unique<Page>(PoolSize);
        
        // 将新页面的所有块添加到全局栈中
        new_page->addToGlobalStack(global_free_stack_);
        
        pages_.push_back(std::move(new_page));
        
        FASTEXCEL_LOG_DEBUG("Allocated new page for pool, total pages: {}", pages_.size());
    }
    
    // 清理所有资源的内部方法
    void cleanup() noexcept {
        try {
            // 清空线程本地缓存
            while (Block* block = thread_cache_.getBlock()) {
                // 这些块会自动随页面释放而释放
            }
            
            // 清空全局栈
            while (Block* block = global_free_stack_.pop()) {
                // 这些块会自动随页面释放而释放  
            }
            
            // 释放所有页面
            std::lock_guard<std::mutex> lock(pages_mutex_);
            pages_.clear();
            
        } catch (const std::exception& e) {
            // 记录其他异常但不抛出（析构函数中调用）
            FASTEXCEL_LOG_ERROR("Exception during FixedSizePool cleanup: {}", e.what());
        }
    }

    bool isFromThisPool(void* ptr) const noexcept {
        std::lock_guard<std::mutex> lock(pages_mutex_);
        for (const auto& page : pages_) {
            const Block* base = page->blocks.get();
            if (!base) continue;
            const Block* end = base + PoolSize;
            if (reinterpret_cast<const Block*>(ptr) >= base &&
                reinterpret_cast<const Block*>(ptr) < end) {
                return true;
            }
        }
        return false;
    }
};

// 线程本地缓存的定义
template<typename T, size_t PoolSize>
thread_local typename FixedSizePool<T, PoolSize>::ThreadLocalCache 
    FixedSizePool<T, PoolSize>::thread_cache_;

}} // namespace fastexcel::memory
