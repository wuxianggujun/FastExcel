/**
 * @file Exception.cpp
 * @brief FastExcel异常类实现
 */

#include "Exception.hpp"
#include <sstream>
#include <iostream>
#include <mutex>

namespace fastexcel {
namespace core {

// FastExcelException 实现
FastExcelException::FastExcelException(const std::string& message, 
                                     ErrorCode code,
                                     const char* file,
                                     int line)
    : std::runtime_error(message)
    , error_code_(code)
    , file_(file)
    , line_(line) {
}

std::string FastExcelException::getErrorCodeString() const {
    switch (error_code_) {
        case ErrorCode::Ok: return "Ok";
        case ErrorCode::InvalidArgument: return "InvalidArgument";
        case ErrorCode::OutOfMemory: return "OutOfMemory";
        case ErrorCode::NotImplemented: return "NotImplemented";
        case ErrorCode::InternalError: return "InternalError";
        case ErrorCode::Timeout: return "Timeout";
        case ErrorCode::FileNotFound: return "FileNotFound";
        case ErrorCode::FileAccessDenied: return "FileAccessDenied";
        case ErrorCode::FileCorrupted: return "FileCorrupted";
        case ErrorCode::FileAlreadyExists: return "FileAlreadyExists";
        case ErrorCode::FileWriteError: return "FileWriteError";
        case ErrorCode::FileReadError: return "FileReadError";
        case ErrorCode::FileTooLarge: return "FileTooLarge";
        case ErrorCode::ZipCreateError: return "ZipCreateError";
        case ErrorCode::ZipWriteError: return "ZipWriteError";
        case ErrorCode::ZipReadError: return "ZipReadError";
        case ErrorCode::ZipCorrupted: return "ZipCorrupted";
        case ErrorCode::ZipCompressionError: return "ZipCompressionError";
        case ErrorCode::ZipExtractionError: return "ZipExtractionError";
        case ErrorCode::XmlParseError: return "XmlParseError";
        case ErrorCode::XmlInvalidFormat: return "XmlInvalidFormat";
        case ErrorCode::XmlMissingElement: return "XmlMissingElement";
        case ErrorCode::XmlInvalidAttribute: return "XmlInvalidAttribute";
        case ErrorCode::XmlEncodingError: return "XmlEncodingError";
        case ErrorCode::InvalidWorkbook: return "InvalidWorkbook";
        case ErrorCode::InvalidWorksheet: return "InvalidWorksheet";
        case ErrorCode::InvalidCellReference: return "InvalidCellReference";
        case ErrorCode::InvalidFormat: return "InvalidFormat";
        case ErrorCode::InvalidFormula: return "InvalidFormula";
        case ErrorCode::UnsupportedFeature: return "UnsupportedFeature";
        case ErrorCode::CorruptedStyles: return "CorruptedStyles";
        case ErrorCode::CorruptedSharedStrings: return "CorruptedSharedStrings";
        default: return "Unknown";
    }
}

std::string FastExcelException::getDetailedMessage() const {
    std::ostringstream oss;
    oss << "[" << getErrorCodeString() << "] " << what();
    
    if (file_ && line_ > 0) {
        oss << " (at " << file_ << ":" << line_ << ")";
    }
    
    if (!context_.empty()) {
        oss << "\nContext:";
        for (const auto& ctx : context_) {
            oss << "\n  - " << ctx;
        }
    }
    
    return oss.str();
}

void FastExcelException::addContext(const std::string& context) {
    context_.push_back(context);
}

// FileException 实现
FileException::FileException(const std::string& message, const std::string& filename,
                           ErrorCode code, const char* file, int line)
    : FastExcelException(message + " (file: " + filename + ")", code, file, line)
    , filename_(filename) {
}

// FormatException 实现
FormatException::FormatException(const std::string& message,
                               ErrorCode code, const char* file, int line)
    : FastExcelException(message, code, file, line) {
}

// MemoryException 实现
MemoryException::MemoryException(const std::string& message,
                               size_t requested_size, const char* file, int line)
    : FastExcelException(message, ErrorCode::OutOfMemory, file, line)
    , requested_size_(requested_size) {
}

// ParameterException 实现
ParameterException::ParameterException(const std::string& message,
                                     const std::string& parameter_name,
                                     const char* file, int line)
    : FastExcelException(message + " (parameter: " + parameter_name + ")", 
                        ErrorCode::InvalidArgument, file, line)
    , parameter_name_(parameter_name) {
}

// OperationException 实现
OperationException::OperationException(const std::string& message,
                                     const std::string& operation,
                                     ErrorCode code, const char* file, int line)
    : FastExcelException(message + " (operation: " + operation + ")", code, file, line)
    , operation_(operation) {
}

// WorksheetException 实现
WorksheetException::WorksheetException(const std::string& message,
                                     const std::string& worksheet_name,
                                     ErrorCode code, const char* file, int line)
    : FastExcelException(message + " (worksheet: " + worksheet_name + ")", code, file, line)
    , worksheet_name_(worksheet_name) {
}

// CellException 实现
CellException::CellException(const std::string& message,
                           int row, int col, ErrorCode code,
                           const char* file, int line)
    : FastExcelException(message, code, file, line)
    , row_(row)
    , col_(col) {
}

std::string CellException::getCellReference() const {
    if (row_ < 0 || col_ < 0) {
        return "Unknown";
    }
    
    std::string col_str;
    int col_num = col_;
    while (col_num >= 0) {
        col_str = char('A' + (col_num % 26)) + col_str;
        col_num = col_num / 26 - 1;
    }
    
    return col_str + std::to_string(row_ + 1);
}

// XMLException 实现
XMLException::XMLException(const std::string& message,
                         const std::string& xml_path,
                         int xml_line, const char* file, int line)
    : FastExcelException(message, ErrorCode::XmlParseError, file, line)
    , xml_path_(xml_path)
    , xml_line_(xml_line) {
}

// DefaultErrorHandler 实现
DefaultErrorHandler::DefaultErrorHandler(bool throw_on_error, bool log_warnings)
    : throw_on_error_(throw_on_error)
    , log_warnings_(log_warnings) {
}

bool DefaultErrorHandler::handleError(const FastExcelException& exception) {
    std::cerr << "FastExcel Error: " << exception.getDetailedMessage() << std::endl;
    
    if (throw_on_error_) {
        throw exception;
    }
    
    return false; // 不继续执行
}

void DefaultErrorHandler::handleWarning(const std::string& message, 
                                       const std::string& context) {
    if (log_warnings_) {
        std::cerr << "FastExcel Warning: " << message;
        if (!context.empty()) {
            std::cerr << " (context: " << context << ")";
        }
        std::cerr << std::endl;
    }
}

// ErrorManager 实现
ErrorManager::ErrorManager()
    : error_handler_(std::make_unique<DefaultErrorHandler>()) {
}

ErrorManager& ErrorManager::getInstance() {
    static ErrorManager instance;
    return instance;
}

void ErrorManager::setErrorHandler(std::unique_ptr<ErrorHandler> handler) {
    std::lock_guard<std::mutex> lock(mutex_);
    error_handler_ = std::move(handler);
}

bool ErrorManager::handleError(const FastExcelException& exception) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    stats_.total_errors++;
    
    if (error_handler_) {
        try {
            bool result = error_handler_->handleError(exception);
            stats_.handled_errors++;
            return result;
        } catch (...) {
            stats_.unhandled_errors++;
            throw;
        }
    } else {
        stats_.unhandled_errors++;
        throw exception;
    }
}

void ErrorManager::handleWarning(const std::string& message, const std::string& context) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    stats_.total_warnings++;
    
    if (error_handler_) {
        error_handler_->handleWarning(message, context);
    }
}

ErrorManager::ErrorStatistics ErrorManager::getStatistics() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return stats_;
}

void ErrorManager::resetStatistics() {
    std::lock_guard<std::mutex> lock(mutex_);
    stats_ = {};
}

} // namespace core
} // namespace fastexcel