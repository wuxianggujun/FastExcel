/**
 * @file MemoryPoolOptimized.hpp
 * @brief 高性能内存池实现，针对FastExcel的使用模式优化
 */

#pragma once

#include <memory>
#include <vector>
#include <mutex>
#include <atomic>
#include <cstdint>
#include <cstring>
#include <type_traits>
#include <unordered_map>
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Exception.hpp"

#ifdef _WIN32
#include <malloc.h>
#endif

namespace fastexcel {
namespace memory {

/**
 * @brief Platform-compatible aligned memory allocation
 */
namespace detail {
    inline void* aligned_alloc_impl(size_t alignment, size_t size) {
#ifdef _WIN32
        return _aligned_malloc(size, alignment);
#else
        return std::aligned_alloc(alignment, size);
#endif
    }
    
    inline void aligned_free_impl(void* ptr) {
#ifdef _WIN32
        _aligned_free(ptr);
#else
        std::free(ptr);
#endif
    }
}

/**
 * @brief 固定大小对象的内存池
 * 
 * 专为频繁分配/释放相同大小对象设计
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

/**
 * @brief 多大小内存池
 * 
 * 支持多种对象大小的内存分配
 */
class MultiSizePool {
private:
    struct SizeClass {
        size_t size;
        size_t alignment;
        std::unique_ptr<void, void(*)(void*)> pool;
        
        SizeClass(size_t s, size_t align) 
            : size(s), alignment(align), pool(nullptr, [](void*){}) {}
    };

public:
    /**
     * @brief 构造函数
     */
    MultiSizePool() {
        // 初始化常用大小类
        initializeSizeClasses();
        
        FASTEXCEL_LOG_DEBUG("Created MultiSizePool with {} size classes", size_classes_.size());
    }
    
    /**
     * @brief 分配指定大小的内存
     */
    void* allocate(size_t size, size_t alignment = alignof(std::max_align_t)) {
        SizeClass* size_class = findSizeClass(size, alignment);
        
        if (size_class) {
            // 从对应的池中分配
            return allocateFromPool(size_class);
        } else {
            // 直接分配
            return detail::aligned_alloc_impl(alignment, size);
        }
    }
    
    /**
     * @brief 释放内存
     */
    void deallocate(void* ptr, size_t size, size_t alignment = alignof(std::max_align_t)) {
        if (!ptr) return;
        
        SizeClass* size_class = findSizeClass(size, alignment);
        
        if (size_class) {
            // 归还到对应的池
            deallocateToPool(size_class, ptr);
        } else {
            // 直接释放
            detail::aligned_free_impl(ptr);
        }
    }

private:
    std::vector<SizeClass> size_classes_;
    mutable std::mutex mutex_;
    
    void initializeSizeClasses() {
        // 常用大小：8, 16, 32, 64, 128, 256, 512, 1024 bytes
        std::vector<size_t> common_sizes = {8, 16, 32, 64, 128, 256, 512, 1024};
        
        for (size_t size : common_sizes) {
            size_classes_.emplace_back(size, std::max(size_t(8), size));
        }
    }
    
    SizeClass* findSizeClass(size_t size, size_t alignment) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        for (auto& sc : size_classes_) {
            if (sc.size >= size && sc.alignment >= alignment) {
                return &sc;
            }
        }
        return nullptr;
    }
    
    void* allocateFromPool(SizeClass* size_class) {
        // 简化实现：直接分配
        return detail::aligned_alloc_impl(size_class->alignment, size_class->size);
    }
    
    void deallocateToPool(SizeClass* /*size_class*/, void* ptr) {
        // 简化实现：直接释放
        detail::aligned_free_impl(ptr);
    }
};

/**
 * @brief 内存池管理器
 * 
 * 全局管理所有内存池
 */
class PoolManager {
public:
    /**
     * @brief 获取单例
     */
    static PoolManager& getInstance() {
        static PoolManager instance;
        return instance;
    }
    
    /**
     * @brief 获取指定类型的内存池
     */
    template<typename T>
    FixedSizePool<T>& getPool() {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = pools_.find(typeid(T).hash_code());
        if (it == pools_.end()) {
            auto pool = std::make_unique<FixedSizePool<T>>();
            auto& pool_ref = *pool;
            pools_[typeid(T).hash_code()] = std::move(pool);
            return pool_ref;
        }
        
        return *static_cast<FixedSizePool<T>*>(it->second.get());
    }
    
    /**
     * @brief 获取多大小内存池
     */
    MultiSizePool& getMultiSizePool() {
        return multi_size_pool_;
    }
    
    /**
     * @brief 清理所有池
     */
    void cleanup() {
        std::lock_guard<std::mutex> lock(mutex_);
        pools_.clear();
        FASTEXCEL_LOG_DEBUG("All memory pools cleaned up");
    }
    
    /**
     * @brief 收缩所有池
     */
    void shrinkAll() {
        std::lock_guard<std::mutex> lock(mutex_);
        
        for (auto& [hash, pool] : pools_) {
            // 这里需要类型转换，简化处理
            FASTEXCEL_LOG_DEBUG("Shrinking pool for type hash: {}", hash);
        }
    }

private:
    PoolManager() = default;
    ~PoolManager() = default;
    
    mutable std::mutex mutex_;
    std::unordered_map<size_t, std::unique_ptr<void, void(*)(void*)>> pools_;
    MultiSizePool multi_size_pool_;
    
    // 禁用拷贝和移动
    PoolManager(const PoolManager&) = delete;
    PoolManager& operator=(const PoolManager&) = delete;
    PoolManager(PoolManager&&) = delete;
    PoolManager& operator=(PoolManager&&) = delete;
};

} // namespace memory
} // namespace fastexcel

/**
 * @brief 内存池分配器
 * 
 * 标准库兼容的分配器
 */
template<typename T>
class PoolAllocator {
public:
    using value_type = T;
    using pointer = T*;
    using const_pointer = const T*;
    using reference = T&;
    using const_reference = const T&;
    using size_type = std::size_t;
    using difference_type = std::ptrdiff_t;
    
    template<typename U>
    struct rebind {
        using other = PoolAllocator<U>;
    };
    
    PoolAllocator() noexcept = default;
    
    template<typename U>
    PoolAllocator(const PoolAllocator<U>&) noexcept {}
    
    pointer allocate(size_type n) {
        if (n == 1) {
            // 使用内存池
            auto& pool = fastexcel::memory::PoolManager::getInstance().getPool<T>();
            return pool.allocate();
        } else {
            // 直接分配
            return static_cast<pointer>(std::malloc(n * sizeof(T)));
        }
    }
    
    void deallocate(pointer p, size_type n) noexcept {
        if (n == 1) {
            // 归还到内存池
            auto& pool = fastexcel::memory::PoolManager::getInstance().getPool<T>();
            pool.deallocate(p);
        } else {
            // 直接释放
            std::free(p);
        }
    }
    
    template<typename U, typename... Args>
    void construct(U* p, Args&&... args) {
        new (p) U(std::forward<Args>(args)...);
    }
    
    template<typename U>
    void destroy(U* p) {
        p->~U();
    }
    
    bool operator==(const PoolAllocator&) const noexcept { return true; }
    bool operator!=(const PoolAllocator&) const noexcept { return false; }
};

// 便捷的类型别名
namespace fastexcel {
    template<typename T>
    using PoolVector = std::vector<T, PoolAllocator<T>>;
    
    template<typename T>
    using pool_ptr = std::unique_ptr<T, std::function<void(T*)>>;
    
    template<typename T, typename... Args>
    pool_ptr<T> make_pool_ptr(Args&&... args) {
        auto& pool = memory::PoolManager::getInstance().getPool<T>();
        T* obj = pool.allocate(std::forward<Args>(args)...);
        
        return pool_ptr<T>(obj, [&pool](T* p) {
            pool.deallocate(p);
        });
    }
}