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
        
        // 通用错误
        case ErrorCode::InvalidArgument:
            return "Invalid argument";
        case ErrorCode::OutOfMemory:
            return "Out of memory";
        case ErrorCode::NotImplemented:
            return "Feature not implemented";
        case ErrorCode::InternalError:
            return "Internal error";
        case ErrorCode::Timeout:
            return "Operation timeout";
        
        // 文件操作错误
        case ErrorCode::FileNotFound:
            return "File not found";
        case ErrorCode::FileAccessDenied:
            return "File access denied";
        case ErrorCode::FileCorrupted:
            return "File corrupted";
        case ErrorCode::FileAlreadyExists:
            return "File already exists";
        case ErrorCode::FileWriteError:
            return "File write error";
        case ErrorCode::FileReadError:
            return "File read error";
        case ErrorCode::FileTooLarge:
            return "File too large";
        
        // ZIP/压缩错误
        case ErrorCode::ZipCreateError:
            return "ZIP creation error";
        case ErrorCode::ZipWriteError:
            return "ZIP write error";
        case ErrorCode::ZipReadError:
            return "ZIP read error";
        case ErrorCode::ZipCorrupted:
            return "ZIP file corrupted";
        case ErrorCode::ZipCompressionError:
            return "ZIP compression error";
        case ErrorCode::ZipExtractionError:
            return "ZIP extraction error";
        
        // XML解析错误
        case ErrorCode::XmlParseError:
            return "XML parse error";
        case ErrorCode::XmlInvalidFormat:
            return "Invalid XML format";
        case ErrorCode::XmlMissingElement:
            return "Missing XML element";
        case ErrorCode::XmlInvalidAttribute:
            return "Invalid XML attribute";
        case ErrorCode::XmlEncodingError:
            return "XML encoding error";
        
        // Excel格式错误
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
        case ErrorCode::UnsupportedFeature:
            return "Unsupported feature";
        case ErrorCode::CorruptedStyles:
            return "Corrupted styles";
        case ErrorCode::CorruptedSharedStrings:
            return "Corrupted shared strings";
        
        // 内存/性能错误
        case ErrorCode::MemoryPoolExhausted:
            return "Memory pool exhausted";
        case ErrorCode::CacheOverflow:
            return "Cache overflow";
        case ErrorCode::BufferOverflow:
            return "Buffer overflow";
        case ErrorCode::ResourceExhausted:
            return "Resource exhausted";
        
        // 并发错误
        case ErrorCode::ThreadPoolError:
            return "Thread pool error";
        case ErrorCode::LockTimeout:
            return "Lock timeout";
        case ErrorCode::ConcurrencyError:
            return "Concurrency error";
        
        // 网络/IO错误
        case ErrorCode::NetworkError:
            return "Network error";
        case ErrorCode::IoError:
            return "I/O error";
        case ErrorCode::StreamError:
            return "Stream error";
        
        default:
            return "Unknown error";
    }
}

}} // namespace fastexcel::core