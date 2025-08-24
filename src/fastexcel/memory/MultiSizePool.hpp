/**
 * @file MultiSizePool.hpp
 * @brief 多大小内存池实现
 */

#pragma once

#include "AlignedAllocator.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <vector>
#include <mutex>
#include <memory>
#include <functional>

namespace fastexcel {
namespace memory {

/**
 * @brief 多大小内存池
 * 
 * 支持多种对象大小的内存分配，针对不同大小的对象进行优化
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
     * @brief 析构函数
     */
    ~MultiSizePool() = default;
    
    // 禁用拷贝和移动
    MultiSizePool(const MultiSizePool&) = delete;
    MultiSizePool& operator=(const MultiSizePool&) = delete;
    MultiSizePool(MultiSizePool&&) = delete;
    MultiSizePool& operator=(MultiSizePool&&) = delete;
    
    /**
     * @brief 分配指定大小的内存
     * @param size 要分配的内存大小
     * @param alignment 内存对齐要求
     * @return 分配的内存指针，失败时返回nullptr
     */
    void* allocate(size_t size, size_t alignment = AlignedAllocator::default_alignment()) {
        SizeClass* size_class = findSizeClass(size, alignment);
        
        if (size_class) {
            // 从对应的池中分配
            return allocateFromPool(size_class);
        } else {
            // 直接分配
            return AlignedAllocator::allocate(alignment, size);
        }
    }
    
    /**
     * @brief 释放内存
     * @param ptr 要释放的内存指针
     * @param size 内存大小
     * @param alignment 内存对齐要求
     */
    void deallocate(void* ptr, size_t size, size_t alignment = AlignedAllocator::default_alignment()) {
        if (!ptr) return;
        
        SizeClass* size_class = findSizeClass(size, alignment);
        
        if (size_class) {
            // 归还到对应的池
            deallocateToPool(size_class, ptr);
        } else {
            // 直接释放
            AlignedAllocator::deallocate(ptr);
        }
    }
    
    /**
     * @brief 获取支持的大小类数量
     * @return 大小类数量
     */
    size_t getSizeClassCount() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return size_classes_.size();
    }
    
    /**
     * @brief 获取指定大小的建议对齐值
     * @param size 内存大小
     * @return 建议的对齐值
     */
    size_t getRecommendedAlignment(size_t size) const {
        return std::max(size_t(8), size);
    }

private:
    std::vector<SizeClass> size_classes_;
    mutable std::mutex mutex_;
    
    void initializeSizeClasses() {
        // 常用大小：8, 16, 32, 64, 128, 256, 512, 1024 bytes
        std::vector<size_t> common_sizes = {8, 16, 32, 64, 128, 256, 512, 1024};
        
        for (size_t size : common_sizes) {
            size_classes_.emplace_back(size, getRecommendedAlignment(size));
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
        return AlignedAllocator::allocate(size_class->alignment, size_class->size);
    }
    
    void deallocateToPool(SizeClass* /*size_class*/, void* ptr) {
        // 简化实现：直接释放
        AlignedAllocator::deallocate(ptr);
    }
};

}} // namespace fastexcel::memory