#pragma once

#include <cstdint>
#include <string>
#include <system_error>

namespace fastexcel {
namespace core {

/**
 * @brief FastExcel统一错误码
 * 
 * 设计原则：
 * - 零成本抽象：底层使用错误码，无异常开销
 * - 双通道模型：可选择抛异常或返回错误码
 * - 性能优先：热路径完全无异常
 */
enum class ErrorCode : uint16_t {
    // 成功
    Ok = 0,
    
    // 通用错误 (1-99)
    InvalidArgument = 1,
    OutOfMemory = 2,
    NotImplemented = 3,
    InternalError = 4,
    Timeout = 5,
    
    // 文件操作错误 (100-199)
    FileNotFound = 100,
    FileAccessDenied = 101,
    FileCorrupted = 102,
    FileAlreadyExists = 103,
    FileWriteError = 104,
    FileReadError = 105,
    FileTooLarge = 106,
    
    // ZIP/压缩错误 (200-299)
    ZipCreateError = 200,
    ZipWriteError = 201,
    ZipReadError = 202,
    ZipCorrupted = 203,
    ZipCompressionError = 204,
    ZipExtractionError = 205,
    
    // XML解析错误 (300-399)
    XmlParseError = 300,
    XmlInvalidFormat = 301,
    XmlMissingElement = 302,
    XmlInvalidAttribute = 303,
    XmlEncodingError = 304,
    
    // Excel格式错误 (400-499)
    InvalidWorkbook = 400,
    InvalidWorksheet = 401,
    InvalidCellReference = 402,
    InvalidFormat = 403,
    InvalidFormula = 404,
    UnsupportedFeature = 405,
    CorruptedStyles = 406,
    CorruptedSharedStrings = 407,
    
    // 内存/性能错误 (500-599)
    MemoryPoolExhausted = 500,
    CacheOverflow = 501,
    BufferOverflow = 502,
    ResourceExhausted = 503,
    
    // 并发错误 (600-699)
    ThreadPoolError = 600,
    LockTimeout = 601,
    ConcurrencyError = 602,
    
    // 网络/IO错误 (700-799)
    NetworkError = 700,
    IoError = 701,
    StreamError = 702
};

/**
 * @brief 错误信息结构
 */
struct Error {
    ErrorCode code;
    std::string message;
    std::string context;  // 额外上下文信息
    
    Error() : code(ErrorCode::Ok) {}
    
    explicit Error(ErrorCode c);
    
    Error(ErrorCode c, const std::string& msg) : code(c), message(msg) {}
    
    Error(ErrorCode c, const std::string& msg, const std::string& ctx) 
        : code(c), message(msg), context(ctx) {}
    
    bool isOk() const noexcept { return code == ErrorCode::Ok; }
    bool isError() const noexcept { return code != ErrorCode::Ok; }
    
    operator bool() const noexcept { return isError(); }
    
    std::string fullMessage() const {
        if (context.empty()) {
            return message;
        }
        return message + " (Context: " + context + ")";
    }
};

/**
 * @brief 错误码转字符串
 */
const char* toString(ErrorCode code) noexcept;

/**
 * @brief 创建错误对象的便利函数
 */
inline Error makeError(ErrorCode code) {
    return Error(code);
}

inline Error makeError(ErrorCode code, const std::string& message) {
    return Error(code, message);
}

inline Error makeError(ErrorCode code, const std::string& message, const std::string& context) {
    return Error(code, message, context);
}

/**
 * @brief 成功结果
 */
inline Error success() {
    return Error(ErrorCode::Ok);
}

#include <stdexcept>

/**
 * @brief FastExcel异常类
 */
class FastExcelException : public std::runtime_error {
private:
    ErrorCode code_;
    std::string context_;

public:
    explicit FastExcelException(const Error& error)
        : std::runtime_error(error.fullMessage())
        , code_(error.code)
        , context_(error.context) {}
    
    explicit FastExcelException(ErrorCode code)
        : std::runtime_error(toString(code))
        , code_(code) {}
    
    FastExcelException(ErrorCode code, const std::string& message)
        : std::runtime_error(message)
        , code_(code) {}
    
    ErrorCode code() const noexcept { return code_; }
    const std::string& context() const noexcept { return context_; }
};

/**
 * @brief 抛出异常的便利函数
 */
[[noreturn]] inline void throwError(const Error& error) {
    throw FastExcelException(error);
}

[[noreturn]] inline void throwError(ErrorCode code) {
    throw FastExcelException(code);
}

[[noreturn]] inline void throwError(ErrorCode code, const std::string& message) {
    throw FastExcelException(code, message);
}

}} // namespace fastexcel::core