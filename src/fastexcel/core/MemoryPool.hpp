/**
 * @file MemoryPool.hpp
 * @brief 内存池实现，用于优化内存分配性能
 */

#pragma once


#include <memory>
#include <vector>
#include <mutex>
#include <cstddef>
#include <new>
#include "Expected.hpp"
#include "ErrorCode.hpp"

// 平台兼容的aligned_alloc
#ifdef _MSC_VER
    #include <malloc.h>
    inline void* aligned_alloc_compat(size_t alignment, size_t size) {
        return _aligned_malloc(size, alignment);
    }
    inline void aligned_free_compat(void* ptr) {
        _aligned_free(ptr);
    }
#else
    #include <cstdlib>
    inline void* aligned_alloc_compat(size_t alignment, size_t size) {
        return std::aligned_alloc(alignment, size);
    }
    inline void aligned_free_compat(void* ptr) {
        std::free(ptr);
    }
#endif

namespace fastexcel {
namespace core {

/**
 * @brief 简单的内存池实现
 * 
 * 用于减少频繁的内存分配和释放，提高性能
 */
class MemoryPool {
public:
    /**
     * @brief 构造函数
     * @param block_size 每个内存块的大小
     * @param initial_blocks 初始内存块数量
     */
    explicit MemoryPool(size_t block_size = 1024, size_t initial_blocks = 16);
    
    /**
     * @brief 析构函数
     */
    ~MemoryPool();
    
    /**
     * @brief 分配内存
     * @param size 需要分配的内存大小
     * @return Result<void*> 成功返回内存指针，失败返回错误
     */
    Result<void*> allocate(size_t size);
    
    /**
     * @brief 释放内存
     * @param ptr 要释放的内存指针
     * @return VoidResult 成功或失败信息
     */
    VoidResult deallocate(void* ptr);
    
    /**
     * @brief 清空内存池
     * @return VoidResult 成功或失败信息
     */
    VoidResult clear();
    
    /**
     * @brief 获取内存池统计信息
     */
    struct Statistics {
        size_t total_allocated;     // 总分配内存
        size_t total_deallocated;   // 总释放内存
        size_t current_usage;       // 当前使用量
        size_t peak_usage;          // 峰值使用量
        size_t allocation_count;    // 分配次数
        size_t deallocation_count;  // 释放次数
    };
    
    Statistics getStatistics() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStatistics();

private:
    struct Block {
        void* data;
        size_t size;
        bool in_use;
        
        static Result<std::unique_ptr<Block>> create(size_t s) {
            auto block = std::make_unique<Block>();
            block->size = s;
            block->in_use = false;
            block->data = aligned_alloc_compat(alignof(std::max_align_t), s);
            if (!block->data) {
                return makeError(ErrorCode::OutOfMemory, "Failed to allocate memory block");
            }
            return makeExpected(std::move(block));
        }
        
        Block() = default;
        
        ~Block() {
            if (data) {
                aligned_free_compat(data);
                data = nullptr;
            }
        }
    };
    
    mutable std::mutex mutex_;
    std::vector<std::unique_ptr<Block>> blocks_;
    size_t block_size_;
    
    // 统计信息
    mutable Statistics stats_;
    
    // 辅助方法
    Block* findAvailableBlock(size_t size);
    VoidResult addNewBlock(size_t size);
};

/**
 * @brief 内存池分配器
 * 
 * 可以与STL容器一起使用的分配器
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
    
    explicit PoolAllocator(MemoryPool* pool = nullptr) : pool_(pool) {}
    
    template<typename U>
    PoolAllocator(const PoolAllocator<U>& other) : pool_(other.pool_) {}
    
    pointer allocate(size_type n) {
        if (pool_) {
            auto result = pool_->allocate(n * sizeof(T));
            if (result.hasValue()) {
                return static_cast<pointer>(result.value());
            }
        }
        return static_cast<pointer>(std::malloc(n * sizeof(T)));
    }
    
    void deallocate(pointer p, size_type n) {
        if (pool_) {
            pool_->deallocate(p);
        } else {
            std::free(p);
        }
    }
    
    template<typename U>
    bool operator==(const PoolAllocator<U>& other) const {
        return pool_ == other.pool_;
    }
    
    template<typename U>
    bool operator!=(const PoolAllocator<U>& other) const {
        return !(*this == other);
    }
    
private:
    MemoryPool* pool_;
    
    template<typename U>
    friend class PoolAllocator;
};

/**
 * @brief 全局内存池管理器
 */
class MemoryManager {
public:
    /**
     * @brief 获取全局实例
     */
    static MemoryManager& getInstance();
    
    /**
     * @brief 获取默认内存池
     */
    MemoryPool& getDefaultPool();
    
    /**
     * @brief 获取指定大小的内存池
     * @return Result<MemoryPool&> 成功返回内存池引用，失败返回错误
     */
    Result<std::reference_wrapper<MemoryPool>> getPool(size_t block_size);
    
    /**
     * @brief 清理所有内存池
     * @return VoidResult 成功或失败信息
     */
    VoidResult cleanup();
    
    /**
     * @brief 获取总体统计信息
     */
    struct GlobalStatistics {
        size_t total_pools;
        size_t total_memory_allocated;
        size_t total_memory_in_use;
        MemoryPool::Statistics default_pool_stats;
    };
    
    GlobalStatistics getGlobalStatistics() const;

private:
    MemoryManager();
    ~MemoryManager();
    
    mutable std::mutex mutex_;
    std::unique_ptr<MemoryPool> default_pool_;
    std::vector<std::pair<size_t, std::unique_ptr<MemoryPool>>> pools_;
    
    // 禁用拷贝和赋值
    MemoryManager(const MemoryManager&) = delete;
    MemoryManager& operator=(const MemoryManager&) = delete;
};

} // namespace core
} // namespace fastexcel

