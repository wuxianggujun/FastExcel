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
#include <thread>
#include <type_traits>
#include <unordered_set>
#include <unordered_map>

namespace fastexcel {
namespace memory {

/**
 * @brief 内存池配置结构
 */
struct PoolConfig {
    size_t initial_pages = 1;           // 初始页面数
    size_t max_pages = 1000;            // 最大页面数
    double shrink_threshold = 0.1;      // 收缩阈值
    size_t thread_cache_size = 32;      // 线程本地缓存大小
    bool enable_statistics = true;      // 启用统计信息
    bool enable_debug_tracking = false; // 启用调试跟踪（强制开启，即使非调试模式）
    size_t batch_stats_size = 64;       // 批量统计更新大小
    size_t shrink_check_interval = 10000; // 收缩检查间隔
    double high_usage_threshold = 0.8;  // 高使用率阈值
    
    // 验证配置有效性
    bool isValid() const {
        return initial_pages > 0 && 
               max_pages >= initial_pages &&
               shrink_threshold > 0.0 && shrink_threshold < 1.0 &&
               thread_cache_size > 0 &&
               batch_stats_size > 0 &&
               shrink_check_interval > 0 &&
               high_usage_threshold > shrink_threshold &&
               high_usage_threshold < 1.0;
    }
};

/**
 * @brief 固定大小对象的内存池
 * 
 * 专为频繁分配/释放相同大小对象设计，提供高性能的内存管理
 */
template<typename T, size_t PoolSize = 1024>
class FixedSizePool : public IMemoryPool {
    // 放宽约束条件，允许更多类型使用内存池
    static_assert(sizeof(T) >= 1, "Type must have non-zero size");
    // 对于小类型或复杂类型，我们仍然支持，但会有特殊处理

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
    FixedSizePool(const PoolConfig& config = {}) : config_(config) {
        if (!config_.isValid()) {
            throw std::invalid_argument("Invalid pool configuration");
        }
        
        // 预分配初始页面
        preAllocate(config_.initial_pages);
        
        FASTEXCEL_LOG_DEBUG("Created FixedSizePool for type {} with pool size {}, config: "
                           "initial_pages={}, max_pages={}, cache_size={}", 
                           typeid(T).name(), PoolSize, 
                           config_.initial_pages, config_.max_pages, config_.thread_cache_size);
    }
    
    /**
     * @brief 析构函数
     */
    ~FixedSizePool() {
        // 设置析构状态标志，阻止新的异步操作
        is_destroying_.store(true, std::memory_order_release);
        
        // 稍等一下，让正在进行的异步操作完成
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
        
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
        // 如果正在析构，拒绝新的分配
        if (is_destroying_.load(std::memory_order_acquire)) {
            throw std::runtime_error("Pool is being destroyed");
        }
        
        // 1. 尝试从线程本地缓存获取（每实例TLS缓存）
        auto& thread_cache = getThreadCache();
        Block* block = thread_cache.getBlock();
        // 防御性检查：若缓存中遗留了已失效页面的块（如收缩后），丢弃该块
        while (block && !isFromThisPool(block)) {
            block = thread_cache.getBlock();
        }
        if (block) {
            updateBatchStats(true); // 缓存命中
        } else {
            updateBatchStats(false); // 缓存未命中
            
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
        
        // 6. 动态调整检查（定期执行，避免在持有锁时调用）
        if (total_allocs % config_.shrink_check_interval == 0 && !is_destroying_.load(std::memory_order_acquire)) {
            // 使用异步方式执行动态调整，避免死锁
            std::thread([this]() {
                // 再次检查析构状态
                if (is_destroying_.load(std::memory_order_acquire)) {
                    return; // 已在析构，直接退出
                }
                try {
                    performDynamicAdjustment();
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_WARN("Dynamic adjustment failed: {}", e.what());
                }
            }).detach();
        }
        
#ifdef _DEBUG
        // 调试模式下跟踪分配
        trackAllocation(obj);
#endif
        
        return obj;
    }

    
    /**
     * @brief 释放对象（无锁优化版本）
     * @param obj 要释放的对象指针
     */
    void deallocate(T* obj) {
        if (!obj) return;
        
        // 先验证指针是否来自这个池
        if (!isFromThisPool(obj)) {
            // 不是来自池的指针，抛出异常让调用者知道
            throw std::invalid_argument("Pointer not allocated from this pool");
        }
        
#ifdef _DEBUG
        // 调试模式下跟踪释放（只对池内指针）
        trackDeallocation(obj);
#endif
        
        Block* block = reinterpret_cast<Block*>(obj);
        
        // 1. 调用析构函数（只有在确认属于本池后再析构）
        if constexpr (!std::is_trivially_destructible_v<T>) {
            obj->~T();
        }

        // 2. 尝试放入线程本地缓存
        auto& thread_cache = getThreadCache();
        if (!thread_cache.returnBlock(block)) {
            // 3. 本地缓存已满，刷新部分到全局栈，然后放入全局栈
            thread_cache.flushToGlobal(global_free_stack_);
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
            allocateNewPageInternal();
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
        // 检查析构状态
        if (is_destroying_.load(std::memory_order_acquire)) {
            return; // 已在析构，直接退出
        }
        
        size_t current = current_usage_.load();
        size_t total_capacity = 0;
        {
            std::lock_guard<std::mutex> lock(pages_mutex_);
            total_capacity = pages_.size() * PoolSize;
        }
        
        if (total_capacity == 0) return;
        
        double usage_ratio = static_cast<double>(current) / total_capacity;
        
        // 低使用率时收缩
        if (usage_ratio < config_.shrink_threshold && !is_destroying_.load()) {
            FASTEXCEL_LOG_DEBUG("Pool usage low ({:.2f}%), attempting shrink", usage_ratio * 100);
            shrink();
        }
        // 高使用率时预分配
        else if (usage_ratio > config_.high_usage_threshold && !is_destroying_.load()) {
            FASTEXCEL_LOG_DEBUG("Pool usage high ({:.2f}%), pre-allocating page", usage_ratio * 100);
            // 检查是否达到最大页面限制
            size_t current_pages = 0;
            {
                std::lock_guard<std::mutex> lock(pages_mutex_);
                current_pages = pages_.size();
            }
            
            if (current_pages < config_.max_pages) {
                preAllocate(1);
            } else {
                FASTEXCEL_LOG_WARN("Pool reached maximum pages limit ({}), cannot pre-allocate", 
                                   config_.max_pages);
            }
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
    
#ifdef _DEBUG
    /**
     * @brief 获取内存泄漏报告
     */
    std::vector<void*> getLeakedPointers() const {
        std::lock_guard<std::mutex> lock(debug_mutex_);
        return std::vector<void*>(allocated_pointers_.begin(), allocated_pointers_.end());
    }
    
    /**
     * @brief 打印内存泄漏报告
     */
    void printLeakReport() const {
        std::lock_guard<std::mutex> lock(debug_mutex_);
        if (allocated_pointers_.empty()) {
            FASTEXCEL_LOG_INFO("No memory leaks detected in pool");
            return;
        }
        
        FASTEXCEL_LOG_ERROR("Memory leak detected! {} objects not freed:", allocated_pointers_.size());
        FASTEXCEL_LOG_ERROR("Total allocations: {}, Total deallocations: {}", 
                           total_debug_allocations_.load(), total_debug_deallocations_.load());
        
        size_t count = 0;
        for (void* ptr : allocated_pointers_) {
            if (++count <= 10) { // 只显示前10个泄漏的指针
                FASTEXCEL_LOG_ERROR("  Leaked pointer: {}", ptr);
            }
        }
        
        if (allocated_pointers_.size() > 10) {
            FASTEXCEL_LOG_ERROR("  ... and {} more leaked pointers", 
                               allocated_pointers_.size() - 10);
        }
    }
#endif
    
    /**
     * @brief 获取当前配置
     */
    const PoolConfig& getConfig() const { return config_; }
    
    /**
     * @brief 更新配置（某些配置需要重启才能生效）
     */
    void updateConfig(const PoolConfig& new_config) {
        if (!new_config.isValid()) {
            throw std::invalid_argument("Invalid pool configuration");
        }
        
        PoolConfig old_config = config_;
        config_ = new_config;
        
        FASTEXCEL_LOG_INFO("Pool configuration updated. Some changes may require restart to take effect.");
        
        // 记录配置变化
        if (old_config.max_pages != new_config.max_pages) {
            FASTEXCEL_LOG_INFO("Max pages changed: {} -> {}", old_config.max_pages, new_config.max_pages);
        }
        if (old_config.shrink_threshold != new_config.shrink_threshold) {
            FASTEXCEL_LOG_INFO("Shrink threshold changed: {:.2f} -> {:.2f}", 
                               old_config.shrink_threshold, new_config.shrink_threshold);
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

#ifdef _DEBUG
        // 调试模式下跟踪分配（通过IMemoryPool接口的分配也要跟踪）
        trackAllocation(reinterpret_cast<void*>(block));
#endif

        return reinterpret_cast<void*>(block);
    }

    void deallocate(void* ptr, std::size_t /*size*/, std::size_t /*alignment*/ = alignof(std::max_align_t)) override {
        if (!ptr) return;

        // 判断指针是否来自本池的页面
        if (isFromThisPool(ptr)) {
            // 先调用析构（IMemoryPool 接口不知道T，只能按原始指针释放；对于固定大小池，T已知，因此无此路径）
            Block* block = reinterpret_cast<Block*>(ptr);
#ifdef _DEBUG
            // 调试模式下对IMemoryPool接口的释放也进行跟踪
            trackDeallocation(ptr);
#endif
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
    
    // 配置对象
    PoolConfig config_;
    
    // 页面管理（仍需要锁保护）
    mutable std::mutex pages_mutex_;
    std::vector<std::unique_ptr<Page>> pages_;
    
    // 获取每实例的线程本地缓存（以 this 作为键）
    ThreadLocalCache& getThreadCache() const {
        static thread_local std::unordered_map<const void*, ThreadLocalCache> tls_caches;
        return tls_caches[this];
    }
    
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
    std::atomic<bool> is_destroying_{false};   // 析构状态标志
    
#ifdef _DEBUG
    // 调试模式下的内存泄漏检测
    mutable std::mutex debug_mutex_;
    std::unordered_set<void*> allocated_pointers_;
    std::atomic<size_t> total_debug_allocations_{0};
    std::atomic<size_t> total_debug_deallocations_{0};
    
    void trackAllocation(void* ptr) {
        if (!ptr) return;
        std::lock_guard<std::mutex> lock(debug_mutex_);
        allocated_pointers_.insert(ptr);
        total_debug_allocations_.fetch_add(1);
        FASTEXCEL_LOG_DEBUG("Tracked allocation: {}, total active: {}", 
                           ptr, allocated_pointers_.size());
    }
    
    void trackDeallocation(void* ptr) {
        if (!ptr) return;
        std::lock_guard<std::mutex> lock(debug_mutex_);
        auto it = allocated_pointers_.find(ptr);
        if (it == allocated_pointers_.end()) {
            // 指针未被跟踪，可能是:
            // 1. 在调试模式启用前分配的
            // 2. 来自其他池或系统的分配
            // 3. warmup期间的对象
            // 我们只记录调试信息，不抛出异常
            FASTEXCEL_LOG_DEBUG("Deallocating untracked pointer: {} (likely pre-debug allocation)", ptr);
            return;  // 静默处理，让释放继续进行
        }
        allocated_pointers_.erase(it);
        total_debug_deallocations_.fetch_add(1);
        FASTEXCEL_LOG_DEBUG("Tracked deallocation: {}, total active: {}", 
                           ptr, allocated_pointers_.size());
    }
#endif
    
    // 分配新页面的方法
    void allocateNewPage() {
        std::lock_guard<std::mutex> lock(pages_mutex_);
        allocateNewPageInternal();
    }
    
    // 分配新页面的内部方法（假设已经持有锁）
    void allocateNewPageInternal() {
        auto new_page = std::make_unique<Page>(PoolSize);
        
        // 将新页面的所有块添加到全局栈中
        new_page->addToGlobalStack(global_free_stack_);
        
        pages_.push_back(std::move(new_page));
        
        FASTEXCEL_LOG_DEBUG("Allocated new page for pool, total pages: {}", pages_.size());
    }
    
    // 清理所有资源的内部方法
    void cleanup() noexcept {
        try {
            // 在清理其他资源前，先清空全局栈和释放页面
            // 避免在析构过程中访问thread_local缓存
            
            // 清空全局栈
            while (Block* block = global_free_stack_.pop()) {
                // 这些块会自动随页面释放而释放  
            }
            
            // 释放所有页面
            {
                std::lock_guard<std::mutex> lock(pages_mutex_);
                pages_.clear();
            }
            
            // 注意：我们不再在cleanup中访问线程本地缓存
            // 因为在析构时访问thread_local可能导致未定义行为
            // thread_local缓存会随着线程结束而自动清理
            
        } catch (const std::exception& e) {
            // 记录其他异常但不抛出（析构函数中调用）
            FASTEXCEL_LOG_ERROR("Exception during FixedSizePool cleanup: {}", e.what());
        }
    }

    // 批量统计更新 - 减少原子操作开销
    void updateBatchStats(bool cache_hit) {
        // 如果正在析构，跳过统计更新以避免访问thread_local
        if (is_destroying_.load(std::memory_order_acquire)) {
            return;
        }
        
        static thread_local size_t local_hits = 0;
        static thread_local size_t local_misses = 0;
        const size_t batch_size = config_.batch_stats_size; // 使用配置值
        
        if (cache_hit) {
            ++local_hits;
            if (local_hits >= batch_size) {
                cache_hits_.fetch_add(batch_size, std::memory_order_relaxed);
                local_hits = 0;
            }
        } else {
            ++local_misses;
            if (local_misses >= batch_size) {
                cache_misses_.fetch_add(batch_size, std::memory_order_relaxed);
                local_misses = 0;
            }
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

// 线程本地缓存：已改为每实例TLS映射，不再需要单一的模板级TLS实例

}} // namespace fastexcel::memory
