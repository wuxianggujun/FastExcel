/**
 * @file LazyInitializer.hpp
 * @brief 内存管理专用的延迟初始化工具
 */

#pragma once

#include <memory>
#include <stdexcept>

namespace fastexcel {
namespace memory {

/**
 * @brief 延迟初始化容器 - 内存管理专用版本
 * 
 * 简化版本，专门用于内存池的延迟初始化，减少对其他模块的依赖
 */
template<typename T>
class LazyInitializer {
public:
    LazyInitializer() = default;
    
    // 禁止拷贝，允许移动
    LazyInitializer(const LazyInitializer&) = delete;
    LazyInitializer& operator=(const LazyInitializer&) = delete;
    
    LazyInitializer(LazyInitializer&& other) noexcept
        : instance_(std::move(other.instance_)) {
    }
    
    LazyInitializer& operator=(LazyInitializer&& other) noexcept {
        if (this != &other) {
            instance_ = std::move(other.instance_);
        }
        return *this;
    }
    
    /**
     * @brief 检查是否已初始化
     */
    bool isInitialized() const noexcept {
        return instance_ != nullptr;
    }
    
    /**
     * @brief 初始化对象
     * @param args 构造函数参数
     * @return 初始化的对象引用
     */
    template<typename... Args>
    T& initialize(Args&&... args) {
        if (isInitialized()) {
            throw std::logic_error("Object already initialized");
        }
        
        try {
            instance_ = std::make_unique<T>(std::forward<Args>(args)...);
            return *instance_;
        } catch (...) {
            instance_.reset();
            throw;
        }
    }
    
    /**
     * @brief 获取对象引用
     * @return 对象引用
     * @throws std::logic_error 如果对象未初始化
     */
    T& get() {
        if (!isInitialized()) {
            throw std::logic_error("Object not initialized");
        }
        return *instance_;
    }
    
    const T& get() const {
        if (!isInitialized()) {
            throw std::logic_error("Object not initialized");
        }
        return *instance_;
    }
    
    /**
     * @brief 重置对象
     */
    void reset() noexcept {
        instance_.reset();
    }
    
    /**
     * @brief 获取指针（可能为空）
     */
    T* operator->() {
        return instance_.get();
    }
    
    const T* operator->() const {
        return instance_.get();
    }
    
    /**
     * @brief 获取引用
     */
    T& operator*() {
        return get();
    }
    
    const T& operator*() const {
        return get();
    }
    
    /**
     * @brief 布尔转换
     */
    explicit operator bool() const noexcept {
        return isInitialized();
    }
    
    /**
     * @brief 获取原始指针
     */
    T* raw_ptr() noexcept {
        return instance_.get();
    }
    
    const T* raw_ptr() const noexcept {
        return instance_.get();
    }

private:
    std::unique_ptr<T> instance_;
};

}} // namespace fastexcel::memory