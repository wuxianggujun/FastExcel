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
#include <atomic>
#include <cstdlib>
#include <thread>
#include <chrono>

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
        // 如果正在关闭，抛出异常
        if (is_shutting_down_.load(std::memory_order_acquire)) {
            throw std::runtime_error("PoolManager is shutting down");
        }
        
        // 首先尝试查找现有池
        {
            std::lock_guard<std::mutex> lock(mutex_);
            
            // 在锁内再次检查关闭状态
            if (is_shutting_down_.load(std::memory_order_acquire)) {
                throw std::runtime_error("PoolManager is shutting down");
            }
            
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
            
            // 在锁内再次检查关闭状态
            if (is_shutting_down_.load(std::memory_order_acquire)) {
                throw std::runtime_error("PoolManager is shutting down");
            }
            
            size_t type_hash = typeid(T).hash_code();
            auto it = pools_.find(type_hash);
            
            // 双重检查：可能其他线程已经创建了池
            if (it != pools_.end()) {
                return *static_cast<FixedSizePool<T>*>(it->second.get());
            }
            
            // 使用删除器函数正确删除池（添加异常保护）
            auto deleter = [](void* ptr) {
                try {
                    if (ptr) {
                        delete static_cast<FixedSizePool<T>*>(ptr);
                    }
                } catch (const std::exception& /*e*/) {
                    // 在deleter中不能抛出异常，只能记录
                    // 但此时Logger可能不可用
                }
            };
            
            std::unique_ptr<void, void(*)(void*)> void_ptr(
                new_pool.release(),
                deleter
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
    
    /**
     * @brief 尝试释放指针到正确的池
     * @param ptr 要释放的指针
     * @return 如果成功释放到某个池返回true，否则返回false
     */
    template<typename T>
    bool tryDeallocate(T* ptr) {
        if (!ptr) return false;
        
        // 如果正在关闭，不试图访问内存池
        if (is_shutting_down_.load(std::memory_order_acquire)) {
            return false; // 让调用者使用标准释放
        }
        
        try {
            std::lock_guard<std::mutex> lock(mutex_);
            
            // 再次检查是否正在关闭（双重检查）
            if (is_shutting_down_.load(std::memory_order_acquire)) {
                return false;
            }
            
            size_t type_hash = typeid(T).hash_code();
            auto it = pools_.find(type_hash);
            
            if (it != pools_.end()) {
                auto* pool = static_cast<FixedSizePool<T>*>(it->second.get());
                try {
                    pool->deallocate(ptr);
                    return true; // 成功释放
                } catch (const std::invalid_argument&) {
                    // 指针不属于这个池，这是正常情况
                    return false;
                } catch (const std::exception&) {
                    // 其他异常（可能是池已析构）
                    return false;
                }
            }
        } catch (const std::exception&) {
            // 锁获取失败或其他问题，直接返回
            return false;
        }
        
        return false; // 没有找到对应类型的池
    }

private:
    PoolManager() = default;
    ~PoolManager() {
        try {
            // 设置全局销毁标志，阻止新的操作
            is_shutting_down_.store(true, std::memory_order_release);
            
            // 等待所有正在进行的操作完成
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
            
            // 在析构时先停止所有正在进行的操作
            std::lock_guard<std::mutex> lock(mutex_);
            
            // 逐个安全地清理池，避免同时析构多个池
            for (auto it = pools_.begin(); it != pools_.end();) {
                try {
                    auto current = it++;
                    // 安全地重置单个池
                    current->second.reset();
                } catch (const std::exception&) {
                    // 忽略单个池清理失败，继续清理其他池
                }
                
                // 给其他线程一些时间
                std::this_thread::sleep_for(std::chrono::microseconds(1));
            }
            
            // 最后清空映射表
            pools_.clear();
            
        } catch (const std::exception& /*e*/) {
            // 在析构函数中不能抛出异常
            // 只能记录错误，但无法输出（Logger可能已析构）
            // FASTEXCEL_LOG_ERROR("清理PoolManager时出错: {}", e.what());
        }
    }
    
    mutable std::mutex mutex_;
    std::unordered_map<size_t, std::unique_ptr<void, void(*)(void*)>> pools_;
    MultiSizePool multi_size_pool_;
    std::atomic<bool> is_shutting_down_{false};  // 添加销毁标志
    
    // 禁用拷贝和移动
    PoolManager(const PoolManager&) = delete;
    PoolManager& operator=(const PoolManager&) = delete;
    PoolManager(PoolManager&&) = delete;
    PoolManager& operator=(PoolManager&&) = delete;
};

}} // namespace fastexcel::memory