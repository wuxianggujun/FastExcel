/**
 * @file PoolAllocator.hpp
 * @brief STL兼容的内存池分配器
 */

#pragma once

#include "PoolManager.hpp"
#include <memory>
#include <cstdlib>

namespace fastexcel {
namespace memory {

/**
 * @brief 内存池分配器
 * 
 * STL兼容的分配器，使用内存池进行内存分配，提供更好的性能
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
     * @brief 分配内存
     * @param n 要分配的对象数量
     * @return 分配的内存指针
     */
    pointer allocate(size_type n) {
        if (n == 1) {
            // 使用内存池分配单个对象
            auto& pool = PoolManager::getInstance().getPool<T>();
            return pool.allocate();
        } else {
            // 对于多个对象，使用标准分配
            return static_cast<pointer>(std::malloc(n * sizeof(T)));
        }
    }
    
    /**
     * @brief 释放内存
     * @param p 要释放的内存指针
     * @param n 对象数量
     */
    void deallocate(pointer p, size_type n) noexcept {
        if (n == 1) {
            // 归还到内存池
            auto& pool = PoolManager::getInstance().getPool<T>();
            pool.deallocate(p);
        } else {
            // 标准释放
            std::free(p);
        }
    }
    
    /**
     * @brief 构造对象
     * @param p 对象指针
     * @param args 构造函数参数
     */
    template<typename U, typename... Args>
    void construct(U* p, Args&&... args) {
        new (p) U(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 销毁对象
     * @param p 对象指针
     */
    template<typename U>
    void destroy(U* p) {
        p->~U();
    }
    
    /**
     * @brief 获取最大可分配数量
     * @return 最大可分配的对象数量
     */
    size_type max_size() const noexcept {
        return std::numeric_limits<size_type>::max() / sizeof(T);
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
};

}} // namespace fastexcel::memory

// 便捷的类型别名
namespace fastexcel {
    /**
     * @brief 使用内存池的vector
     */
    template<typename T>
    using PoolVector = std::vector<T, memory::PoolAllocator<T>>;
    
    /**
     * @brief 内存池管理的智能指针
     */
    template<typename T>
    using pool_ptr = std::unique_ptr<T, std::function<void(T*)>>;
    
    /**
     * @brief 创建内存池管理的智能指针
     * @param args 构造函数参数
     * @return 内存池管理的智能指针
     */
    template<typename T, typename... Args>
    pool_ptr<T> make_pool_ptr(Args&&... args) {
        auto& pool = memory::PoolManager::getInstance().getPool<T>();
        T* obj = pool.allocate(std::forward<Args>(args)...);
        
        return pool_ptr<T>(obj, [&pool](T* p) {
            pool.deallocate(p);
        });
    }
}