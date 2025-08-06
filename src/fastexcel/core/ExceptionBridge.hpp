/**
 * @file ExceptionBridge.hpp
 * @brief 异常转换层：连接底层Result/Expected和用户层Exception
 */

#pragma once

#include "Expected.hpp"
#include "ErrorCode.hpp"
#include "Exception.hpp"
#include <type_traits>

namespace fastexcel {
namespace core {

/**
 * @brief 异常转换层
 * 
 * 提供底层高性能Result/Expected与上层传统异常之间的无缝转换
 * 
 * 设计原则：
 * 1. 底层使用Expected/Result：零异常开销，性能优先
 * 2. 用户层使用传统异常：开发友好，错误信息丰富
 * 3. 转换层：自动处理两者之间的转换
 */
class ExceptionBridge {
public:
    /**
     * @brief 将Result转换为异常抛出
     * @param result 底层Result对象
     * @throws FastExcelException 如果result包含错误
     */
    template<typename T>
    static T unwrap(Result<T>&& result) {
        if (result.hasError()) {
            throwFromError(result.error());
        }
        return std::move(result.value());
    }
    
    template<typename T>
    static T unwrap(const Result<T>& result) {
        if (result.hasError()) {
            throwFromError(result.error());
        }
        return result.value();
    }
    
    /**
     * @brief 将VoidResult转换为异常抛出
     * @param result 底层VoidResult对象
     * @throws FastExcelException 如果result包含错误
     */
    static void unwrap(VoidResult&& result) {
        if (result.hasError()) {
            throwFromError(result.error());
        }
    }
    
    static void unwrap(const VoidResult& result) {
        if (result.hasError()) {
            throwFromError(result.error());
        }
    }
    
    /**
     * @brief 捕获异常并转换为Result
     * @param func 可能抛出异常的函数
     * @return Result<T> 成功值或错误信息
     */
    template<typename F>
    static auto wrapCall(F&& func) -> Result<std::remove_reference_t<decltype(func())>> {
        try {
            using ReturnType = decltype(func());
            if constexpr (std::is_void_v<ReturnType>) {
                func();
                return success();
            } else {
                return makeExpected(func());
            }
        } catch (const FastExcelException& e) {
            return makeError(e.getErrorCode(), e.what());
        } catch (const std::bad_alloc&) {
            return makeError(ErrorCode::OutOfMemory, "Memory allocation failed");
        } catch (const std::exception& e) {
            return makeError(ErrorCode::InternalError, e.what());
        }
    }
    
    /**
     * @brief 捕获异常并转换为VoidResult
     * @param func 可能抛出异常的void函数
     * @return VoidResult 成功或错误信息
     */
    template<typename F>
    static VoidResult wrapVoidCall(F&& func) {
        try {
            func();
            return success();
        } catch (const FastExcelException& e) {
            return makeError(e.getErrorCode(), e.what());
        } catch (const std::bad_alloc&) {
            return makeError(ErrorCode::OutOfMemory, "Memory allocation failed");
        } catch (const std::exception& e) {
            return makeError(ErrorCode::InternalError, e.what());
        }
    }
    
    /**
     * @brief 从ErrorCode映射到异常类型并抛出
     * @param error 错误对象
     */
    [[noreturn]] static void throwFromError(const Error& error) {
        switch (error.code) {
            case ErrorCode::FileNotFound:
            case ErrorCode::FileAccessDenied:
            case ErrorCode::FileCorrupted:
            case ErrorCode::FileAlreadyExists:
            case ErrorCode::FileWriteError:
            case ErrorCode::FileReadError:
            case ErrorCode::FileTooLarge:
                throw FileException(error.fullMessage(), "", 
                    mapToLegacyErrorCode(error.code));
                
            case ErrorCode::OutOfMemory:
            case ErrorCode::MemoryPoolExhausted:
            case ErrorCode::BufferOverflow:
                throw MemoryException(error.fullMessage());
                
            case ErrorCode::InvalidArgument:
                throw ParameterException(error.fullMessage());
                
            case ErrorCode::XmlParseError:
            case ErrorCode::XmlInvalidFormat:
            case ErrorCode::XmlMissingElement:
            case ErrorCode::XmlInvalidAttribute:
            case ErrorCode::XmlEncodingError:
                throw XMLException(error.fullMessage());
                
            case ErrorCode::InvalidWorkbook:
            case ErrorCode::InvalidWorksheet:
                throw WorksheetException(error.fullMessage());
                
            case ErrorCode::InvalidCellReference:
                throw CellException(error.fullMessage());
                
            case ErrorCode::ZipCreateError:
            case ErrorCode::ZipWriteError:
            case ErrorCode::ZipReadError:
            case ErrorCode::ZipCorrupted:
            case ErrorCode::ZipCompressionError:
            case ErrorCode::ZipExtractionError:
                throw OperationException(error.fullMessage(), "ZIP操作");
                
            default:
                throw FastExcelException(error.fullMessage(), 
                    mapToLegacyErrorCode(error.code));
        }
    }
    
private:
    /**
     * @brief 将新ErrorCode映射到旧的ErrorCode枚举
     */
    static FastExcelException::ErrorCode mapToLegacyErrorCode(ErrorCode code) {
        switch (code) {
            case ErrorCode::FileNotFound: return FastExcelException::ErrorCode::FileNotFound;
            case ErrorCode::FileAccessDenied: return FastExcelException::ErrorCode::FileAccessDenied;
            case ErrorCode::OutOfMemory: return FastExcelException::ErrorCode::OutOfMemory;
            case ErrorCode::InvalidArgument: return FastExcelException::ErrorCode::InvalidParameter;
            case ErrorCode::XmlParseError: return FastExcelException::ErrorCode::XMLParseError;
            case ErrorCode::ZipCompressionError: return FastExcelException::ErrorCode::CompressionError;
            default: return FastExcelException::ErrorCode::Unknown;
        }
    }
};

/**
 * @brief 便利宏定义
 */

// 在用户层API中使用，自动转换Result为异常
#define FASTEXCEL_UNWRAP(result) \
    fastexcel::core::ExceptionBridge::unwrap(result)

// 在底层调用用户层代码时使用，自动捕获异常转换为Result
#define FASTEXCEL_WRAP_CALL(func) \
    fastexcel::core::ExceptionBridge::wrapCall([&]() { return func; })

#define FASTEXCEL_WRAP_VOID_CALL(func) \
    fastexcel::core::ExceptionBridge::wrapVoidCall([&]() { func; })

/**
 * @brief 用户层API助手类
 * 
 * 为用户层API提供便利的结果处理
 */
template<typename T>
class UserAPIWrapper {
private:
    Result<T> result_;

public:
    explicit UserAPIWrapper(Result<T>&& result) : result_(std::move(result)) {}
    
    /**
     * @brief 获取结果，失败时抛出异常
     */
    T get() && {
        return ExceptionBridge::unwrap(std::move(result_));
    }
    
    T get() const & {
        return ExceptionBridge::unwrap(result_);
    }
    
    /**
     * @brief 检查是否成功
     */
    bool isSuccess() const { return result_.hasValue(); }
    
    /**
     * @brief 获取错误信息（不抛出异常）
     */
    std::string getErrorMessage() const {
        return result_.hasError() ? result_.error().fullMessage() : "";
    }
};

/**
 * @brief void类型的用户API包装器
 */
class VoidUserAPIWrapper {
private:
    VoidResult result_;

public:
    explicit VoidUserAPIWrapper(VoidResult&& result) : result_(std::move(result)) {}
    
    /**
     * @brief 检查结果，失败时抛出异常
     */
    void check() {
        ExceptionBridge::unwrap(result_);
    }
    
    /**
     * @brief 检查是否成功
     */
    bool isSuccess() const { return result_.hasValue(); }
    
    /**
     * @brief 获取错误信息（不抛出异常）
     */
    std::string getErrorMessage() const {
        return result_.hasError() ? result_.error().fullMessage() : "";
    }
};

/**
 * @brief 创建用户API包装器的便利函数
 */
template<typename T>
UserAPIWrapper<T> wrapForUser(Result<T>&& result) {
    return UserAPIWrapper<T>(std::move(result));
}

inline VoidUserAPIWrapper wrapForUser(VoidResult&& result) {
    return VoidUserAPIWrapper(std::move(result));
}

} // namespace core
} // namespace fastexcel
