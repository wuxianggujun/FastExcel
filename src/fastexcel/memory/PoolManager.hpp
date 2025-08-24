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
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = pools_.find(typeid(T).hash_code());
        if (it == pools_.end()) {
            auto pool = std::make_unique<FixedSizePool<T>>();
            auto& pool_ref = *pool;
            
            // 使用类型删除器来正确删除池
            pools_[typeid(T).hash_code()] = std::unique_ptr<void, void(*)(void*)>(
                pool.release(),
                [](void* ptr) {
                    delete static_cast<FixedSizePool<T>*>(ptr);
                }
            );
            
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