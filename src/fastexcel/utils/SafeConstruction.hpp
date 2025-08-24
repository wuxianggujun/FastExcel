/**
 * @file SafeConstruction.hpp
 * @brief 异常安全的构造模式和RAII辅助类
 */

#pragma once

#include <memory>
#include <vector>
#include <functional>
#include <exception>
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace utils {

/**
 * @brief RAII资源管理器
 * 
 * 用于管理多个资源的自动清理
 */
class ResourceManager {
public:
    using CleanupFunc = std::function<void()>;
    
    ResourceManager() = default;
    
    /**
     * @brief 析构函数，按照注册相反的顺序清理资源
     */
    ~ResourceManager() {
        cleanup();
    }
    
    // 禁用拷贝，允许移动
    ResourceManager(const ResourceManager&) = delete;
    ResourceManager& operator=(const ResourceManager&) = delete;
    
    ResourceManager(ResourceManager&& other) noexcept
        : cleanup_functions_(std::move(other.cleanup_functions_)) {
        other.cleanup_functions_.clear();
    }
    
    ResourceManager& operator=(ResourceManager&& other) noexcept {
        if (this != &other) {
            cleanup();
            cleanup_functions_ = std::move(other.cleanup_functions_);
            other.cleanup_functions_.clear();
        }
        return *this;
    }
    
    /**
     * @brief 注册清理函数
     * @param cleanup_func 清理函数
     */
    void addCleanup(CleanupFunc cleanup_func) {
        if (cleanup_func) {
            cleanup_functions_.push_back(std::move(cleanup_func));
        }
    }
    
    /**
     * @brief 注册需要清理的指针
     */
    template<typename T>
    void addResource(std::unique_ptr<T>& ptr) {
        if (ptr) {
            addCleanup([&ptr]() { ptr.reset(); });
        }
    }
    
    /**
     * @brief 手动执行清理
     */
    void cleanup() noexcept {
        // 按照注册相反的顺序清理
        for (auto it = cleanup_functions_.rbegin(); it != cleanup_functions_.rend(); ++it) {
            try {
                (*it)();
            } catch (const std::exception& e) {
                FASTEXCEL_LOG_ERROR("Error during cleanup: {}", e.what());
            } catch (...) {
                FASTEXCEL_LOG_ERROR("Unknown error during cleanup");
            }
        }
        cleanup_functions_.clear();
    }
    
    /**
     * @brief 取消所有清理操作（用于成功构造后）
     */
    void release() noexcept {
        cleanup_functions_.clear();
    }
    
    /**
     * @brief 获取注册的清理函数数量
     */
    size_t size() const noexcept {
        return cleanup_functions_.size();
    }

private:
    std::vector<CleanupFunc> cleanup_functions_;
};

/**
 * @brief 异常安全的构造辅助类
 * 
 * 提供构造过程中的异常安全保证
 */
template<typename T>
class SafeConstructor {
public:
    /**
     * @brief 构造函数
     */
    SafeConstructor() = default;
    
    /**
     * @brief 设置构造完成回调
     */
    SafeConstructor& onSuccess(std::function<void(T&)> callback) {
        success_callback_ = std::move(callback);
        return *this;
    }
    
    /**
     * @brief 设置构造失败回调
     */
    SafeConstructor& onFailure(std::function<void(const std::exception&)> callback) {
        failure_callback_ = std::move(callback);
        return *this;
    }
    
    /**
     * @brief 执行安全构造
     * @param constructor 构造函数
     * @return 构造的对象
     */
    template<typename Constructor>
    std::unique_ptr<T> construct(Constructor&& constructor) {
        ResourceManager resource_manager;
        
        try {
            // 执行构造
            auto result = constructor(resource_manager);
            
            // 构造成功，调用成功回调
            if (success_callback_) {
                success_callback_(*result);
            }
            
            // 取消清理（构造成功）
            resource_manager.release();
            
            return result;
            
        } catch (const std::exception& e) {
            // 构造失败，调用失败回调
            if (failure_callback_) {
                try {
                    failure_callback_(e);
                } catch (...) {
                    // 忽略回调中的异常
                }
            }
            
            // 资源管理器会自动清理
            throw;
        }
    }

private:
    std::function<void(T&)> success_callback_;
    std::function<void(const std::exception&)> failure_callback_;
};

/**
 * @brief 延迟初始化容器
 * 
 * 用于需要多步骤初始化的对象
 */
template<typename T>
class LazyInitializer {
public:
    LazyInitializer() = default;
    
    /**
     * @brief 检查是否已初始化
     */
    bool isInitialized() const noexcept {
        return instance_ != nullptr;
    }
    
    /**
     * @brief 初始化对象
     */
    template<typename... Args>
    T& initialize(Args&&... args) {
        if (isInitialized()) {
            throw core::OperationException(
                "Object already initialized",
                "initialize",
                core::ErrorCode::InvalidArgument,
                __FILE__, __LINE__
            );
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
     */
    T& get() {
        if (!isInitialized()) {
            throw core::OperationException(
                "Object not initialized",
                "get",
                core::ErrorCode::InvalidArgument,
                __FILE__, __LINE__
            );
        }
        return *instance_;
    }
    
    const T& get() const {
        if (!isInitialized()) {
            throw core::OperationException(
                "Object not initialized", 
                "get",
                core::ErrorCode::InvalidArgument,
                __FILE__, __LINE__
            );
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

private:
    std::unique_ptr<T> instance_;
};

/**
 * @brief 构造状态跟踪器
 * 
 * 用于跟踪复杂对象的构造进度
 */
class ConstructionTracker {
public:
    enum class State {
        NOT_STARTED,
        IN_PROGRESS,
        COMPLETED,
        FAILED
    };
    
    ConstructionTracker() : state_(State::NOT_STARTED) {}
    
    /**
     * @brief 开始构造
     */
    void start(const std::string& description = "") {
        if (state_ != State::NOT_STARTED) {
            throw core::OperationException(
                "Construction already started",
                "start",
                core::ErrorCode::InvalidArgument,
                __FILE__, __LINE__
            );
        }
        
        state_ = State::IN_PROGRESS;
        description_ = description;
        start_time_ = std::chrono::steady_clock::now();
        
        FASTEXCEL_LOG_DEBUG("Construction started: {}", description_);
    }
    
    /**
     * @brief 标记构造完成
     */
    void complete() {
        if (state_ != State::IN_PROGRESS) {
            throw core::OperationException(
                "Construction not in progress",
                "complete",
                core::ErrorCode::InvalidArgument,
                __FILE__, __LINE__
            );
        }
        
        state_ = State::COMPLETED;
        auto duration = std::chrono::steady_clock::now() - start_time_;
        auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
        
        FASTEXCEL_LOG_DEBUG("Construction completed: {} (took {}ms)", description_, ms);
    }
    
    /**
     * @brief 标记构造失败
     */
    void fail(const std::string& reason = "") {
        if (state_ != State::IN_PROGRESS) {
            return; // 可能已经失败了
        }
        
        state_ = State::FAILED;
        auto duration = std::chrono::steady_clock::now() - start_time_;
        auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
        
        FASTEXCEL_LOG_ERROR("Construction failed: {} - {} (after {}ms)", 
                           description_, reason, ms);
    }
    
    /**
     * @brief 获取当前状态
     */
    State getState() const noexcept {
        return state_;
    }
    
    /**
     * @brief 检查是否成功完成
     */
    bool isCompleted() const noexcept {
        return state_ == State::COMPLETED;
    }
    
    /**
     * @brief 检查是否失败
     */
    bool isFailed() const noexcept {
        return state_ == State::FAILED;
    }

private:
    State state_;
    std::string description_;
    std::chrono::steady_clock::time_point start_time_;
};

} // namespace utils
} // namespace fastexcel

/**
 * @brief 安全构造宏
 * 
 * 提供异常安全的构造模式
 */
#define FASTEXCEL_SAFE_CONSTRUCT(Class, ...) \
    fastexcel::utils::SafeConstructor<Class>() \
        .onFailure([](const std::exception& e) { \
            FASTEXCEL_LOG_ERROR("Failed to construct " #Class ": {}", e.what()); \
        }) \
        .construct([&](fastexcel::utils::ResourceManager& rm) -> std::unique_ptr<Class> { \
            auto instance = std::make_unique<Class>(__VA_ARGS__); \
            rm.addResource(instance); \
            return instance; \
        })

/**
 * @brief 延迟初始化宏
 */
#define FASTEXCEL_LAZY_INIT(var, Class, ...) \
    if (!var.isInitialized()) { \
        var.initialize(__VA_ARGS__); \
    }