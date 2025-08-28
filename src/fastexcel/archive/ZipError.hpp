#pragma once

namespace fastexcel {
namespace archive {

// 错误码枚举
enum class ZipError {
    Ok,                    // 操作成功
    NotOpen,              // ZIP 文件未打开
    IoFail,               // I/O 操作失败
    BadFormat,            // ZIP 格式错误
    TooLarge,             // 文件太大
    FileNotFound,         // 文件未找到
    InvalidParameter,     // 无效参数
    CompressionFail,      // 压缩失败
    InternalError         // 内部错误
};

// 为 ZipError 提供 bool 转换操作符，使其可以在布尔上下文中使用
// 只有 ZipError::Ok 被视为 true，其他所有错误都被视为 false
constexpr bool operator!(ZipError error) noexcept {
    return error != ZipError::Ok;
}

// 显式 bool 转换函数，用于更清晰的语义
constexpr bool isSuccess(ZipError error) noexcept {
    return error == ZipError::Ok;
}

constexpr bool isError(ZipError error) noexcept {
    return error != ZipError::Ok;
}

}} // namespace fastexcel::archive
