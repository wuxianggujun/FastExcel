#include "fastexcel/core/ErrorCode.hpp"

namespace fastexcel {
namespace core {

// Error类构造函数实现
Error::Error(ErrorCode c) : code(c), message(toString(c)) {}

const char* toString(ErrorCode code) noexcept {
    switch (code) {
        // 成功
        case ErrorCode::Ok:
            return "Success";
        
        // 通用错误 (1-19)
        case ErrorCode::InvalidArgument:
            return "Invalid argument";
        case ErrorCode::OutOfMemory:
            return "Out of memory";
        case ErrorCode::InternalError:
            return "Internal error";
        
        // 文件操作错误 (20-39)
        case ErrorCode::FileNotFound:
            return "File not found";
        case ErrorCode::FileAccessDenied:
            return "File access denied";
        case ErrorCode::FileCorrupted:
            return "File corrupted";
        case ErrorCode::FileWriteError:
            return "File write error";
        case ErrorCode::FileReadError:
            return "File read error";
        
        // Excel格式错误 (40-59)
        case ErrorCode::InvalidWorkbook:
            return "Invalid workbook";
        case ErrorCode::InvalidWorksheet:
            return "Invalid worksheet";
        case ErrorCode::InvalidCellReference:
            return "Invalid cell reference";
        case ErrorCode::InvalidFormat:
            return "Invalid format";
        case ErrorCode::InvalidFormula:
            return "Invalid formula";
        case ErrorCode::CorruptedStyles:
            return "Corrupted styles";
        case ErrorCode::CorruptedSharedStrings:
            return "Corrupted shared strings";
        
        // ZIP/XML处理错误 (60-79)
        case ErrorCode::ZipError:
            return "ZIP error";
        case ErrorCode::XmlParseError:
            return "XML parse error";
        case ErrorCode::XmlInvalidFormat:
            return "Invalid XML format";
        case ErrorCode::XmlMissingElement:
            return "Missing XML element";
        
        // 功能实现状态 (80-89)
        case ErrorCode::NotImplemented:
            return "Feature not implemented";
        
        default:
            return "Unknown error";
    }
}

}} // namespace fastexcel::core