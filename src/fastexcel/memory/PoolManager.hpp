/**
 * @file PoolManager.hpp
 * @brief 全局内存池管理器
 */

#pragma once

#include "FixedSizePool.hpp"
#include "MultiSizePool.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <unordered_map>
#include <mutex>
#include <memory>
#include <typeinfo>
#include <functional>

namespace fastexcel {
namespace memory {

/**
 * @brief 内存池管理器
 * 
 * 全局管理所有类型的内存池，提供统一的内存池访问接口
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
     * @return 指定类型的FixedSizePool引用
     */
    template<typename T>
    FixedSizePool<T>& getPool() {
        // 首先尝试查找现有池
        {
            std::lock_guard<std::mutex> lock(mutex_);
            size_t type_hash = typeid(T).hash_code();
            auto it = pools_.find(type_hash);
            
            if (it != pools_.end()) {
                return *static_cast<FixedSizePool<T>*>(it->second.get());
            }
        }
        
        // 如果池不存在，创建新池（在锁外创建以避免死锁）
        auto new_pool = std::make_unique<FixedSizePool<T>>();
        FixedSizePool<T>* pool_ptr = new_pool.get();
        
        // 再次获取锁并插入新池
        {
            std::lock_guard<std::mutex> lock(mutex_);
            size_t type_hash = typeid(T).hash_code();
            auto it = pools_.find(type_hash);
            
            // 双重检查：可能其他线程已经创建了池
            if (it != pools_.end()) {
                return *static_cast<FixedSizePool<T>*>(it->second.get());
            }
            
            // 使用删除器函数正确删除池
            std::unique_ptr<void, void(*)(void*)> void_ptr(
                new_pool.release(),
                [](void* ptr) {
                    delete static_cast<FixedSizePool<T>*>(ptr);
                }
            );
            
            // 使用emplace避免默认构造
            pools_.emplace(type_hash, std::move(void_ptr));
        }
        
        return *pool_ptr;
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
    
    /**
     * @brief 获取池的统计信息
     * @return 池的数量
     */
    size_t getPoolCount() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return pools_.size();
    }
    
    /**
     * @brief 预热常用类型的内存池
     */
    template<typename... Types>
    void prewarmPools() {
        (getPool<Types>(), ...);
        FASTEXCEL_LOG_DEBUG("Prewarmed {} memory pools", sizeof...(Types));
    }

private:
    PoolManager() = default;
    ~PoolManager() {
        cleanup();
    }
    
    mutable std::mutex mutex_;
    std::unordered_map<size_t, std::unique_ptr<void, void(*)(void*)>> pools_;
    MultiSizePool multi_size_pool_;
    
    // 禁用拷贝和移动
    PoolManager(const PoolManager&) = delete;
    PoolManager& operator=(const PoolManager&) = delete;
    PoolManager(PoolManager&&) = delete;
    PoolManager& operator=(PoolManager&&) = delete;
};

}} // namespace fastexcel::memory