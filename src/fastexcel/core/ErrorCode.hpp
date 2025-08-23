#pragma once

#include <cstdint>
#include <string>
#include <system_error>
#include <fmt/format.h>

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
/**
 * @brief 简化的FastExcel错误码 - 专注于Excel处理核心错误
 * 
 * 架构优化原则：
 * - 只包含Excel处理相关的错误
 * - 移除网络、并发等不相关错误
 * - 使用更紧凑的编码
 */
enum class ErrorCode : uint8_t {
    // 成功
    Ok = 0,
    
    // 通用错误 (1-19)
    InvalidArgument = 1,
    OutOfMemory = 2,
    InternalError = 3,
    
    // 文件操作错误 (20-39)
    FileNotFound = 20,
    FileAccessDenied = 21,
    FileCorrupted = 22,
    FileWriteError = 23,
    FileReadError = 24,
    
    // Excel格式错误 (40-59)
    InvalidWorkbook = 40,
    InvalidWorksheet = 41,
    InvalidCellReference = 42,
    InvalidFormat = 43,
    InvalidFormula = 44,
    CorruptedStyles = 45,
    CorruptedSharedStrings = 46,
    
    // ZIP/XML处理错误 (60-79)
    ZipError = 60,
    XmlParseError = 61,
    XmlInvalidFormat = 62,
    XmlMissingElement = 63,
    
    // 功能实现状态 (80-89)
    NotImplemented = 80
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
        return fmt::format("{} (Context: {})", message, context);
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

}} // namespace fastexcel::core
