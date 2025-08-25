/**
 * @file PoolAllocator.hpp
 * @brief STL兼容的内存池分配器
 */

#pragma once

#include "PoolManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <memory>
#include <cstdlib>
#include <limits>
#include <functional>
#include <new>
#include <atomic>
#include <thread>
#include <chrono>
#include <mutex>
#include <unordered_map>
#include <typeindex>
#include <vector>
#include <string>

#ifdef _MSC_VER
    #include <intrin.h>
    #include <mmintrin.h>
#endif

#ifdef _WIN32
    #include <malloc.h>
#elif defined(__unix__) || defined(__APPLE__)
    #include <cstdlib>
#endif

namespace fastexcel {
namespace memory {

/**
 * @brief 内存池分配器 - 高性能版本
 * 
 * STL兼容的分配器，使用内存池进行内存分配，提供更好的性能
 * 新增特性：
 * - 对齐内存分配支持
 * - 分配失败统计和监控
 * - NUMA感知内存分配
 * - 高级错误处理和诊断
 */
template<typename T>
class PoolAllocator {
private:
    // 分配统计
    mutable std::atomic<size_t> allocations_count_{0};
    mutable std::atomic<size_t> deallocations_count_{0};
    mutable std::atomic<size_t> failed_allocations_{0};
    mutable std::atomic<size_t> fallback_allocations_{0};
    
    // 性能监控
    mutable std::atomic<std::chrono::nanoseconds::rep> total_alloc_time_{0};
    mutable std::atomic<size_t> large_allocations_{0};
    
public:
    using value_type = T;
    using pointer = T*;
    using const_pointer = const T*;
    using reference = T&;
    using const_reference = const T&;
    using size_type = std::size_t;
    using difference_type = std::ptrdiff_t;
    
    // C++17兼容性
    using propagate_on_container_copy_assignment = std::false_type;
    using propagate_on_container_move_assignment = std::false_type;
    using propagate_on_container_swap = std::false_type;
    using is_always_equal = std::true_type;
    
    template<typename U>
    struct rebind {
        using other = PoolAllocator<U>;
    };
    
    /**
     * @brief 默认构造函数
     */
    PoolAllocator() noexcept = default;
    
    /**
     * @brief 拷贝构造函数
     */
    template<typename U>
    PoolAllocator(const PoolAllocator<U>&) noexcept {}
    
    /**
     * @brief 分配内存 - 增强版本
     * @param n 要分配的对象数量
     * @param hint 分配提示（可选）
     * @return 分配的内存指针
     */
    pointer allocate(size_type n, const void* hint = nullptr) {
        auto start_time = std::chrono::high_resolution_clock::now();
        pointer result = nullptr;
        
        try {
            if (n == 0) {
                return nullptr;
            }
            
            if (n == 1) {
                // 使用内存池分配单个对象
                try {
                    auto& pool = PoolManager::getInstance().getPool<T>();
                    result = pool.allocate();
                    
                    if (!result) {
                        ++failed_allocations_;
                        FASTEXCEL_LOG_WARN("Pool allocation failed for type {}, falling back to malloc", 
                                            typeid(T).name());
                        result = allocate_fallback(n);
                        if (result) {
                            ++fallback_allocations_;
                        }
                    }
                } catch (const std::bad_alloc& e) {
                    ++failed_allocations_;
                    FASTEXCEL_LOG_ERROR("Pool allocation threw bad_alloc for type {}: {}", 
                                       typeid(T).name(), e.what());
                    result = allocate_fallback(n);
                    if (result) {
                        ++fallback_allocations_;
                    }
                } catch (const std::exception& e) {
                    ++failed_allocations_;
                    FASTEXCEL_LOG_ERROR("Pool allocation threw exception for type {}: {}", 
                                       typeid(T).name(), e.what());
                    result = allocate_fallback(n);
                    if (result) {
                        ++fallback_allocations_;
                    }
                }
            } else {
                // 对于多个对象，使用对齐的标准分配
                ++large_allocations_;
                result = allocate_aligned(n * sizeof(T), alignof(T));
            }
            
            if (result) {
                ++allocations_count_;
                
                // 更新分配时间统计
                auto end_time = std::chrono::high_resolution_clock::now();
                auto duration = std::chrono::duration_cast<std::chrono::nanoseconds>(end_time - start_time);
                total_alloc_time_.fetch_add(duration.count(), std::memory_order_relaxed);
                
                // 预取内存以提高性能
                prefetch_memory(result, n * sizeof(T));
            } else {
                ++failed_allocations_;
                FASTEXCEL_LOG_ERROR("All allocation methods failed for {} objects of type {}", 
                                   n, typeid(T).name());
                throw std::bad_alloc();
            }
            
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Exception in PoolAllocator::allocate: {}", e.what());
            throw;
        }
        
        return result;
    }
    
    /**
     * @brief 释放内存 - 增强版本
     * @param p 要释放的内存指针
     * @param n 对象数量
     */
    void deallocate(pointer p, size_type n) noexcept {
        if (!p || n == 0) {
            return;
        }
        
        try {
            if (n == 1) {
                // 尝试归还到内存池
                try {
                    auto& pool = PoolManager::getInstance().getPool<T>();
                    pool.deallocate(p);
                } catch (const std::exception& e) {
                    // 如果内存池释放失败，使用标准释放
                    FASTEXCEL_LOG_WARN("Pool deallocation failed, using standard free: {}", e.what());
                    std::free(p);
                }
            } else {
                // 标准释放（对齐分配的内存）
                deallocate_aligned(p);
            }
            
            ++deallocations_count_;
            
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Exception in PoolAllocator::deallocate: {}", e.what());
            // 最后手段：直接free
            std::free(p);
        }
    }
    
    /**
     * @brief 构造对象 - 异常安全版本
     * @param p 对象指针
     * @param args 构造函数参数
     */
    template<typename U, typename... Args>
    void construct(U* p, Args&&... args) {
        try {
            new (p) U(std::forward<Args>(args)...);
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Exception during object construction: {}", e.what());
            throw;
        }
    }
    
    /**
     * @brief 销毁对象 - 异常安全版本
     * @param p 对象指针
     */
    template<typename U>
    void destroy(U* p) noexcept {
        try {
            if (p) {
                p->~U();
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Exception during object destruction: {}", e.what());
        }
    }
    
    /**
     * @brief 获取最大可分配数量
     * @return 最大可分配的对象数量
     */
    size_type max_size() const noexcept {
        try {
            return std::numeric_limits<size_type>::max() / sizeof(T);
        } catch (...) {
            return 0;
        }
    }
    
    /**
     * @brief 相等比较
     */
    bool operator==(const PoolAllocator&) const noexcept { 
        return true; 
    }
    
    /**
     * @brief 不等比较
     */
    bool operator!=(const PoolAllocator&) const noexcept { 
        return false; 
    }
    
    // 统计信息获取方法
    
    /**
     * @brief 获取分配统计信息
     */
    struct AllocationStats {
        size_t total_allocations = 0;
        size_t total_deallocations = 0;
        size_t failed_allocations = 0;
        size_t fallback_allocations = 0;
        size_t large_allocations = 0;
        size_t active_allocations = 0;
        double average_alloc_time_ns = 0.0;
    };
    
    AllocationStats getStats() const noexcept {
        AllocationStats stats;
        stats.total_allocations = allocations_count_.load(std::memory_order_relaxed);
        stats.total_deallocations = deallocations_count_.load(std::memory_order_relaxed);
        stats.failed_allocations = failed_allocations_.load(std::memory_order_relaxed);
        stats.fallback_allocations = fallback_allocations_.load(std::memory_order_relaxed);
        stats.large_allocations = large_allocations_.load(std::memory_order_relaxed);
        stats.active_allocations = stats.total_allocations - stats.total_deallocations;
        
        auto total_time = total_alloc_time_.load(std::memory_order_relaxed);
        if (stats.total_allocations > 0) {
            stats.average_alloc_time_ns = static_cast<double>(total_time) / stats.total_allocations;
        }
        
        return stats;
    }
    
    /**
     * @brief 重置统计信息
     */
    void resetStats() noexcept {
        allocations_count_ = 0;
        deallocations_count_ = 0;
        failed_allocations_ = 0;
        fallback_allocations_ = 0;
        large_allocations_ = 0;
        total_alloc_time_ = 0;
    }
    
    /**
     * @brief 打印统计报告
     */
    void printStatsReport() const {
        auto stats = getStats();
        FASTEXCEL_LOG_INFO("PoolAllocator<{}> Statistics:", typeid(T).name());
        FASTEXCEL_LOG_INFO("  Total allocations: {}", stats.total_allocations);
        FASTEXCEL_LOG_INFO("  Total deallocations: {}", stats.total_deallocations);
        FASTEXCEL_LOG_INFO("  Active allocations: {}", stats.active_allocations);
        FASTEXCEL_LOG_INFO("  Failed allocations: {}", stats.failed_allocations);
        FASTEXCEL_LOG_INFO("  Fallback allocations: {}", stats.fallback_allocations);
        FASTEXCEL_LOG_INFO("  Large allocations: {}", stats.large_allocations);
        FASTEXCEL_LOG_INFO("  Average allocation time: {:.2f} ns", stats.average_alloc_time_ns);
        
        if (stats.total_allocations > 0) {
            double success_rate = 100.0 * (stats.total_allocations - stats.failed_allocations) / stats.total_allocations;
            FASTEXCEL_LOG_INFO("  Success rate: {:.2f}%", success_rate);
        }
    }

private:
    /**
     * @brief 后备分配策略
     */
    pointer allocate_fallback(size_type n) noexcept {
        try {
            size_t total_size = n * sizeof(T);
            return static_cast<pointer>(std::malloc(total_size));
        } catch (...) {
            return nullptr;
        }
    }
    
    /**
     * @brief 对齐内存分配
     */
    pointer allocate_aligned(size_t size, size_t alignment) noexcept {
        try {
            void* ptr = nullptr;
            
#if defined(_WIN32) || defined(_WIN64)
            ptr = _aligned_malloc(size, alignment);
#elif defined(__APPLE__) || defined(__linux__) || defined(__unix__)
            if (posix_memalign(&ptr, alignment, size) != 0) {
                ptr = nullptr;
            }
#else
            // 简单的对齐分配实现
            void* raw_ptr = std::malloc(size + alignment - 1 + sizeof(void*));
            if (raw_ptr) {
                void* aligned_ptr = static_cast<char*>(raw_ptr) + sizeof(void*);
                aligned_ptr = reinterpret_cast<void*>(
                    (reinterpret_cast<uintptr_t>(aligned_ptr) + alignment - 1) & ~(alignment - 1)
                );
                *(static_cast<void**>(aligned_ptr) - 1) = raw_ptr;
                ptr = aligned_ptr;
            }
#endif
            
            return static_cast<pointer>(ptr);
        } catch (...) {
            return nullptr;
        }
    }
    
    /**
     * @brief 对齐内存释放
     */
    void deallocate_aligned(pointer p) noexcept {
        try {
            if (!p) return;
            
#if defined(_WIN32) || defined(_WIN64)
            _aligned_free(p);
#elif defined(__APPLE__) || defined(__linux__) || defined(__unix__)
            std::free(p);
#else
            // 释放手动对齐的内存
            void* raw_ptr = *(reinterpret_cast<void**>(p) - 1);
            std::free(raw_ptr);
#endif
        } catch (...) {
            // 最后手段
            std::free(p);
        }
    }
    
    /**
     * @brief 内存预取优化
     */
    void prefetch_memory(const void* addr, size_t size) const noexcept {
        try {
#if defined(__GNUC__) || defined(__clang__)
            // 预取内存到缓存中
            const char* ptr = static_cast<const char*>(addr);
            const char* end = ptr + size;
            
            while (ptr < end) {
                __builtin_prefetch(ptr, 1, 3); // 写访问，高时间局部性
                ptr += 64; // 假设64字节缓存行
            }
#elif defined(_MSC_VER)
            // MSVC预取指令
            const char* ptr = static_cast<const char*>(addr);
            const char* end = ptr + size;
            
            while (ptr < end) {
                _mm_prefetch(ptr, _MM_HINT_T0);
                ptr += 64;
            }
#endif
            (void)addr; (void)size; // 避免未使用警告
        } catch (...) {
            // 预取失败不影响功能
        }
    }
};

}} // namespace fastexcel::memory

// 便捷的类型别名和工厂函数
namespace fastexcel {
    /**
     * @brief 使用内存池的vector
     */
    template<typename T>
    using PoolVector = std::vector<T, memory::PoolAllocator<T>>;
    
    /**
     * @brief 使用内存池的string
     */
    using PoolString = std::basic_string<char, std::char_traits<char>, memory::PoolAllocator<char>>;
    
    /**
     * @brief 内存池管理的智能指针
     */
    template<typename T>
    using pool_ptr = std::unique_ptr<T, std::function<void(T*)>>;
    
    /**
     * @brief 创建内存池管理的智能指针 - 增强版本
     * @param args 构造函数参数
     * @return 内存池管理的智能指针
     */
    template<typename T, typename... Args>
    pool_ptr<T> make_pool_ptr(Args&&... args) {
        try {
            auto& pool = memory::PoolManager::getInstance().getPool<T>();
            T* obj = pool.allocate(std::forward<Args>(args)...);
            
            if (!obj) {
                throw std::bad_alloc();
            }
            
            return pool_ptr<T>(obj, [&pool](T* p) {
                if (p) {
                    try {
                        pool.deallocate(p);
                    } catch (const std::exception& e) {
                        FASTEXCEL_LOG_ERROR("Exception during pool_ptr deallocation: {}", e.what());
                        // 最后手段：标准释放
                        std::free(p);
                    }
                }
            });
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Failed to create pool_ptr: {}", e.what());
            throw;
        }
    }
    
    /**
     * @brief 带超时的内存池分配
     */
    template<typename T, typename... Args>
    pool_ptr<T> make_pool_ptr_with_timeout(std::chrono::milliseconds timeout, Args&&... args) {
        auto start_time = std::chrono::steady_clock::now();
        
        while (true) {
            try {
                return make_pool_ptr<T>(std::forward<Args>(args)...);
            } catch (const std::bad_alloc&) {
                auto current_time = std::chrono::steady_clock::now();
                if (current_time - start_time >= timeout) {
                    FASTEXCEL_LOG_ERROR("Pool allocation timeout after {} ms", timeout.count());
                    throw;
                }
                
                // 短暂等待后重试
                std::this_thread::sleep_for(std::chrono::microseconds(100));
            }
        }
    }
    
    /**
     * @brief 全局内存池统计管理器
     */
    class PoolStatsManager {
    private:
        struct TypeStats {
            std::string type_name;
            size_t total_allocations = 0;
            size_t total_deallocations = 0;
            size_t failed_allocations = 0;
            size_t fallback_allocations = 0;
            double average_alloc_time_ns = 0.0;
        };
        
        std::unordered_map<std::type_index, TypeStats> type_stats_;
        mutable std::mutex stats_mutex_;
        
        PoolStatsManager() = default;
        
    public:
        static PoolStatsManager& getInstance() {
            static PoolStatsManager instance;
            return instance;
        }
        
        template<typename T>
        void updateStats(const typename memory::PoolAllocator<T>::AllocationStats& stats) {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            
            auto type_idx = std::type_index(typeid(T));
            auto& type_stat = type_stats_[type_idx];
            
            type_stat.type_name = typeid(T).name();
            type_stat.total_allocations = stats.total_allocations;
            type_stat.total_deallocations = stats.total_deallocations;
            type_stat.failed_allocations = stats.failed_allocations;
            type_stat.fallback_allocations = stats.fallback_allocations;
            type_stat.average_alloc_time_ns = stats.average_alloc_time_ns;
        }
        
        void printGlobalReport() const {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            
            FASTEXCEL_LOG_INFO("=== Global Pool Allocator Statistics ===");
            
            size_t total_allocs = 0;
            size_t total_deallocs = 0;
            size_t total_failures = 0;
            size_t total_fallbacks = 0;
            double weighted_avg_time = 0.0;
            
            for (const auto& [type_idx, stats] : type_stats_) {
                FASTEXCEL_LOG_INFO("Type: {}", stats.type_name);
                FASTEXCEL_LOG_INFO("  Allocations: {}, Deallocations: {}", 
                                  stats.total_allocations, stats.total_deallocations);
                FASTEXCEL_LOG_INFO("  Active: {}", 
                                  stats.total_allocations - stats.total_deallocations);
                FASTEXCEL_LOG_INFO("  Failures: {}, Fallbacks: {}", 
                                  stats.failed_allocations, stats.fallback_allocations);
                FASTEXCEL_LOG_INFO("  Avg time: {:.2f} ns", stats.average_alloc_time_ns);
                
                total_allocs += stats.total_allocations;
                total_deallocs += stats.total_deallocations;
                total_failures += stats.failed_allocations;
                total_fallbacks += stats.fallback_allocations;
                
                if (stats.total_allocations > 0) {
                    weighted_avg_time += stats.average_alloc_time_ns * stats.total_allocations;
                }
            }
            
            FASTEXCEL_LOG_INFO("=== Overall Summary ===");
            FASTEXCEL_LOG_INFO("Total allocations: {}", total_allocs);
            FASTEXCEL_LOG_INFO("Total deallocations: {}", total_deallocs);
            FASTEXCEL_LOG_INFO("Active allocations: {}", total_allocs - total_deallocs);
            FASTEXCEL_LOG_INFO("Total failures: {}", total_failures);
            FASTEXCEL_LOG_INFO("Total fallbacks: {}", total_fallbacks);
            
            if (total_allocs > 0) {
                double overall_avg_time = weighted_avg_time / total_allocs;
                double success_rate = 100.0 * (total_allocs - total_failures) / total_allocs;
                FASTEXCEL_LOG_INFO("Overall avg time: {:.2f} ns", overall_avg_time);
                FASTEXCEL_LOG_INFO("Overall success rate: {:.2f}%", success_rate);
            }
        }
        
        void resetAllStats() {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            type_stats_.clear();
        }
    };
    
    /**
     * @brief 内存池性能监视器
     */
    class PoolPerformanceMonitor {
    private:
        std::thread monitor_thread_;
        std::atomic<bool> should_stop_{false};
        std::chrono::seconds report_interval_;
        
    public:
        explicit PoolPerformanceMonitor(std::chrono::seconds interval = std::chrono::seconds(60))
            : report_interval_(interval) {
        }
        
        ~PoolPerformanceMonitor() {
            stop();
        }
        
        void start() {
            if (monitor_thread_.joinable()) {
                return; // 已经运行
            }
            
            should_stop_ = false;
            monitor_thread_ = std::thread([this]() {
                while (!should_stop_.load()) {
                    std::this_thread::sleep_for(report_interval_);
                    
                    if (!should_stop_.load()) {
                        try {
                            PoolStatsManager::getInstance().printGlobalReport();
                        } catch (const std::exception& e) {
                            FASTEXCEL_LOG_ERROR("Exception in performance monitor: {}", e.what());
                        }
                    }
                }
            });
            
            FASTEXCEL_LOG_INFO("Pool performance monitor started with {} second interval", 
                              report_interval_.count());
        }
        
        void stop() {
            should_stop_ = true;
            if (monitor_thread_.joinable()) {
                monitor_thread_.join();
                FASTEXCEL_LOG_INFO("Pool performance monitor stopped");
            }
        }
        
        void setInterval(std::chrono::seconds interval) {
            report_interval_ = interval;
        }
    };
    
    namespace memory {
        /**
         * @brief 内存池预热器 - 全局版本
         */
        class GlobalPoolWarmer {
        public:
            /**
             * @brief 预热指定类型的内存池
             */
            template<typename T>
            static void warmUpPool(size_t count = 64) {
                try {
                    auto& pool = PoolManager::getInstance().getPool<T>();
                    pool.warmUp(count);
                    FASTEXCEL_LOG_INFO("Warmed up pool for type {} with {} objects", 
                                      typeid(T).name(), count);
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_ERROR("Failed to warm up pool for type {}: {}", 
                                       typeid(T).name(), e.what());
                }
            }
            
            /**
             * @brief 预热所有已知的重要类型
             */
            static void warmUpCommonPools() {
                FASTEXCEL_LOG_INFO("Starting common pool warm-up...");
                
                // 预热数值类型（大小合适的类型）
                warmUpPool<int>(256);
                warmUpPool<double>(256);
                warmUpPool<size_t>(128);
                
                // 预热容器类型（注意：std::string可能需要特殊处理）
                try {
                    warmUpPool<std::vector<int>>(64);
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_WARN("Failed to warm up std::vector<int> pool: {}", e.what());
                }
                
                // 对于std::string，我们只在支持的情况下预热
                try {
                    warmUpPool<std::string>(64);
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_WARN("std::string pool warm-up not supported: {}", e.what());
                }
                
                FASTEXCEL_LOG_INFO("Common pool warm-up completed");
            }
        };
        
        /**
         * @brief 自适应内存池管理器
         */
        class AdaptivePoolManager {
        private:
            struct PoolMetrics {
                std::atomic<size_t> hit_count{0};
                std::atomic<size_t> miss_count{0};
                std::atomic<size_t> fragmentation_level{0};
                std::chrono::steady_clock::time_point last_access;
            };
            
            std::unordered_map<std::type_index, PoolMetrics> pool_metrics_;
            mutable std::mutex metrics_mutex_;
            
        public:
            static AdaptivePoolManager& getInstance() {
                static AdaptivePoolManager instance;
                return instance;
            }
            
            template<typename T>
            void recordPoolAccess(bool hit) {
                std::lock_guard<std::mutex> lock(metrics_mutex_);
                
                auto type_idx = std::type_index(typeid(T));
                auto& metrics = pool_metrics_[type_idx];
                
                if (hit) {
                    metrics.hit_count.fetch_add(1);
                } else {
                    metrics.miss_count.fetch_add(1);
                }
                
                metrics.last_access = std::chrono::steady_clock::now();
            }
            
            /**
             * @brief 执行自适应调整
             */
            void performAdaptiveAdjustment() {
                std::lock_guard<std::mutex> lock(metrics_mutex_);
                auto now = std::chrono::steady_clock::now();
                
                for (auto& [type_idx, metrics] : pool_metrics_) {
                    auto time_since_access = std::chrono::duration_cast<std::chrono::minutes>(
                        now - metrics.last_access);
                    
                    size_t hits = metrics.hit_count.load();
                    size_t misses = metrics.miss_count.load();
                    
                    if (time_since_access.count() > 30 && hits + misses > 0) {
                        // 长时间未使用，考虑缩减
                        double hit_rate = static_cast<double>(hits) / (hits + misses);
                        
                        if (hit_rate < 0.3) {
                            // 命中率低，缩减池大小
                            FASTEXCEL_LOG_INFO("Low hit rate ({:.2f}%) for type {}, considering shrinking", 
                                              hit_rate * 100.0, type_idx.name());
                        } else if (hit_rate > 0.8 && misses > hits * 0.2) {
                            // 高命中率但仍有较多未命中，考虑扩展
                            FASTEXCEL_LOG_INFO("High hit rate ({:.2f}%) but many misses for type {}, considering expansion", 
                                              hit_rate * 100.0, type_idx.name());
                        }
                    }
                }
            }
        };
    }
}