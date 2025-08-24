/**
 * @file AlignedAllocator.hpp
 * @brief 平台兼容的对齐内存分配器
 */

#pragma once

#include <cstddef>

#ifdef _WIN32
#include <malloc.h>
#endif

namespace fastexcel {
namespace memory {

/**
 * @brief 平台兼容的对齐内存分配工具
 */
class AlignedAllocator {
public:
    /**
     * @brief 分配对齐的内存
     * @param alignment 对齐字节数
     * @param size 分配大小
     * @return 分配的内存指针，失败时返回nullptr
     */
    static void* allocate(size_t alignment, size_t size) {
#ifdef _WIN32
        return _aligned_malloc(size, alignment);
#else
        return std::aligned_alloc(alignment, size);
#endif
    }
    
    /**
     * @brief 释放对齐的内存
     * @param ptr 要释放的内存指针
     */
    static void deallocate(void* ptr) {
        if (!ptr) return;
        
#ifdef _WIN32
        _aligned_free(ptr);
#else
        std::free(ptr);
#endif
    }
    
    /**
     * @brief 获取默认对齐值
     * @return 默认对齐字节数
     */
    static constexpr size_t default_alignment() {
        return alignof(std::max_align_t);
    }
};

}} // namespace fastexcel::memory