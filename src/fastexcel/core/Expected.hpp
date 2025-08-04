#pragma once

#include "fastexcel/core/ErrorCode.hpp"
#include <type_traits>
#include <utility>
#include <new>

namespace fastexcel {
namespace core {

/**
 * @brief Expected<T, E> - 现代C++的错误处理类型
 * 
 * 类似于std::expected (C++23)，但针对FastExcel优化：
 * - 零成本抽象：无异常开销
 * - 移动语义优化：避免不必要的拷贝
 * - 内存布局优化：最小化内存占用
 * - 链式操作：支持函数式编程风格
 */
template<typename T, typename E = Error>
class Expected {
private:
    union {
        T value_;
        E error_;
    };
    bool has_value_;

public:
    using value_type = T;
    using error_type = E;

    // ========== 构造函数 ==========
    
    /**
     * @brief 默认构造函数（成功值）
     */
    Expected() : has_value_(true) {
        if constexpr (std::is_default_constructible_v<T>) {
            new(&value_) T{};
        }
    }
    
    /**
     * @brief 成功值构造函数
     */
    Expected(const T& value) : has_value_(true) {
        new(&value_) T(value);
    }
    
    Expected(T&& value) : has_value_(true) {
        new(&value_) T(std::move(value));
    }
    
    /**
     * @brief 错误构造函数
     */
    Expected(const E& error) : has_value_(false) {
        new(&error_) E(error);
    }
    
    Expected(E&& error) : has_value_(false) {
        new(&error_) E(std::move(error));
    }
    
    /**
     * @brief 拷贝构造函数
     */
    Expected(const Expected& other) : has_value_(other.has_value_) {
        if (has_value_) {
            new(&value_) T(other.value_);
        } else {
            new(&error_) E(other.error_);
        }
    }
    
    /**
     * @brief 移动构造函数
     */
    Expected(Expected&& other) noexcept : has_value_(other.has_value_) {
        if (has_value_) {
            new(&value_) T(std::move(other.value_));
        } else {
            new(&error_) E(std::move(other.error_));
        }
    }
    
    /**
     * @brief 析构函数
     */
    ~Expected() {
        if (has_value_) {
            value_.~T();
        } else {
            error_.~E();
        }
    }
    
    // ========== 赋值操作符 ==========
    
    Expected& operator=(const Expected& other) {
        if (this != &other) {
            this->~Expected();
            new(this) Expected(other);
        }
        return *this;
    }
    
    Expected& operator=(Expected&& other) noexcept {
        if (this != &other) {
            this->~Expected();
            new(this) Expected(std::move(other));
        }
        return *this;
    }
    
    Expected& operator=(const T& value) {
        this->~Expected();
        new(this) Expected(value);
        return *this;
    }
    
    Expected& operator=(T&& value) {
        this->~Expected();
        new(this) Expected(std::move(value));
        return *this;
    }
    
    Expected& operator=(const E& error) {
        this->~Expected();
        new(this) Expected(error);
        return *this;
    }
    
    Expected& operator=(E&& error) {
        this->~Expected();
        new(this) Expected(std::move(error));
        return *this;
    }
    
    // ========== 状态检查 ==========
    
    bool hasValue() const noexcept { return has_value_; }
    bool hasError() const noexcept { return !has_value_; }
    
    explicit operator bool() const noexcept { return has_value_; }
    
    // ========== 值访问 ==========
    
    /**
     * @brief 获取值（不检查）
     */
    T& value() & noexcept { return value_; }
    const T& value() const & noexcept { return value_; }
    T&& value() && noexcept { return std::move(value_); }
    const T&& value() const && noexcept { return std::move(value_); }
    
    /**
     * @brief 获取错误（不检查）
     */
    E& error() & noexcept { return error_; }
    const E& error() const & noexcept { return error_; }
    E&& error() && noexcept { return std::move(error_); }
    const E&& error() const && noexcept { return std::move(error_); }
    
    /**
     * @brief 安全获取值
     */
    T& valueOr(T& default_value) & noexcept {
        return has_value_ ? value_ : default_value;
    }
    
    const T& valueOr(const T& default_value) const & noexcept {
        return has_value_ ? value_ : default_value;
    }
    
    T valueOr(T&& default_value) && noexcept {
        return has_value_ ? std::move(value_) : std::move(default_value);
    }
    
    /**
     * @brief 操作符重载
     */
    T& operator*() & noexcept { return value_; }
    const T& operator*() const & noexcept { return value_; }
    T&& operator*() && noexcept { return std::move(value_); }
    const T&& operator*() const && noexcept { return std::move(value_); }
    
    T* operator->() noexcept { return &value_; }
    const T* operator->() const noexcept { return &value_; }
    
    // ========== 函数式操作 ==========
    
    /**
     * @brief 映射操作（成功时）
     */
    template<typename F>
    auto map(F&& func) -> Expected<decltype(func(value_)), E> {
        if (has_value_) {
            return Expected<decltype(func(value_)), E>(func(value_));
        } else {
            return Expected<decltype(func(value_)), E>(error_);
        }
    }
    
    /**
     * @brief 映射操作（失败时）
     */
    template<typename F>
    auto mapError(F&& func) -> Expected<T, decltype(func(error_))> {
        if (has_value_) {
            return Expected<T, decltype(func(error_))>(value_);
        } else {
            return Expected<T, decltype(func(error_))>(func(error_));
        }
    }
    
    /**
     * @brief 链式操作
     */
    template<typename F>
    auto andThen(F&& func) -> decltype(func(value_)) {
        if (has_value_) {
            return func(value_);
        } else {
            return decltype(func(value_))(error_);
        }
    }
    
    /**
     * @brief 错误恢复
     */
    template<typename F>
    auto orElse(F&& func) -> Expected<T, E> {
        if (has_value_) {
            return *this;
        } else {
            return func(error_);
        }
    }

    /**
     * @brief 抛出异常（如果是错误）
     */
    T& valueOrThrow() & {
        if (has_value_) {
            return value_;
        } else {
            if constexpr (std::is_same_v<E, Error>) {
                throwError(error_);
            } else {
                throw E(error_);
            }
        }
    }
    
    const T& valueOrThrow() const & {
        if (has_value_) {
            return value_;
        } else {
            if constexpr (std::is_same_v<E, Error>) {
                throwError(error_);
            } else {
                throw E(error_);
            }
        }
    }
    
    T valueOrThrow() && {
        if (has_value_) {
            return std::move(value_);
        } else {
            if constexpr (std::is_same_v<E, Error>) {
                throwError(error_);
            } else {
                throw E(std::move(error_));
            }
        }
    }
};

// ========== 便利函数 ==========

/**
 * @brief 创建成功结果
 */
template<typename T>
Expected<T> makeExpected(T&& value) {
    return Expected<T>(std::forward<T>(value));
}

/**
 * @brief 创建错误结果
 */
template<typename T, typename E>
Expected<T, E> makeUnexpected(E&& error) {
    return Expected<T, E>(std::forward<E>(error));
}

/**
 * @brief 特化：void类型的Expected
 */
template<typename E>
class Expected<void, E> {
private:
    E error_;
    bool has_value_;

public:
    using value_type = void;
    using error_type = E;

    Expected() : has_value_(true) {}
    
    Expected(const E& error) : error_(error), has_value_(false) {}
    Expected(E&& error) : error_(std::move(error)), has_value_(false) {}
    
    bool hasValue() const noexcept { return has_value_; }
    bool hasError() const noexcept { return !has_value_; }
    
    explicit operator bool() const noexcept { return has_value_; }
    
    E& error() & noexcept { return error_; }
    const E& error() const & noexcept { return error_; }
    E&& error() && noexcept { return std::move(error_); }
    const E&& error() const && noexcept { return std::move(error_); }

    void valueOrThrow() const {
        if (!has_value_) {
            if constexpr (std::is_same_v<E, Error>) {
                throwError(error_);
            } else {
                throw E(error_);
            }
        }
    }
};

// ========== 类型别名 ==========

template<typename T>
using Result = Expected<T, Error>;

using VoidResult = Expected<void, Error>;

}} // namespace fastexcel::core