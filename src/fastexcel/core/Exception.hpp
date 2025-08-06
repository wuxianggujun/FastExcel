/**
 * @file Exception.hpp
 * @brief FastExcel异常类定义
 */

#ifndef FASTEXCEL_EXCEPTION_HPP
#define FASTEXCEL_EXCEPTION_HPP

#include <stdexcept>
#include <string>
#include <memory>
#include <vector>
#include <mutex>
#include "ErrorCode.hpp"

namespace fastexcel {
namespace core {

/**
 * @brief FastExcel基础异常类
 */
class FastExcelException : public std::runtime_error {
public:
    /**
     * @brief 构造函数
     * @param message 错误消息
     * @param code 错误代码
     * @param file 发生错误的文件名
     * @param line 发生错误的行号
     */
    FastExcelException(const std::string& message, 
                      ErrorCode code = ErrorCode::InternalError,
                      const char* file = nullptr,
                      int line = 0);
    
    /**
     * @brief 获取错误代码
     */
    ErrorCode getErrorCode() const noexcept { return error_code_; }
    
    /**
     * @brief 获取错误代码字符串
     */
    std::string getErrorCodeString() const;
    
    /**
     * @brief 获取详细错误信息
     */
    std::string getDetailedMessage() const;
    
    /**
     * @brief 获取发生错误的文件名
     */
    const char* getFile() const noexcept { return file_; }
    
    /**
     * @brief 获取发生错误的行号
     */
    int getLine() const noexcept { return line_; }
    
    /**
     * @brief 添加上下文信息
     */
    void addContext(const std::string& context);
    
    /**
     * @brief 获取上下文信息
     */
    const std::vector<std::string>& getContext() const { return context_; }

private:
    ErrorCode error_code_;
    const char* file_;
    int line_;
    std::vector<std::string> context_;
};

/**
 * @brief 文件相关异常
 */
class FileException : public FastExcelException {
public:
    FileException(const std::string& message, const std::string& filename,
                 ErrorCode code = ErrorCode::FileNotFound,
                 const char* file = nullptr, int line = 0);
    
    const std::string& getFilename() const { return filename_; }

private:
    std::string filename_;
};

/**
 * @brief 格式相关异常
 */
class FormatException : public FastExcelException {
public:
    FormatException(const std::string& message,
                   ErrorCode code = ErrorCode::InvalidFormat,
                   const char* file = nullptr, int line = 0);
};

/**
 * @brief 内存相关异常
 */
class MemoryException : public FastExcelException {
public:
    MemoryException(const std::string& message,
                   size_t requested_size = 0,
                   const char* file = nullptr, int line = 0);
    
    size_t getRequestedSize() const { return requested_size_; }

private:
    size_t requested_size_;
};

/**
 * @brief 参数相关异常
 */
class ParameterException : public FastExcelException {
public:
    ParameterException(const std::string& message,
                      const std::string& parameter_name = "",
                      const char* file = nullptr, int line = 0);
    
    const std::string& getParameterName() const { return parameter_name_; }

private:
    std::string parameter_name_;
};

/**
 * @brief 操作相关异常
 */
class OperationException : public FastExcelException {
public:
    OperationException(const std::string& message,
                      const std::string& operation = "",
                      ErrorCode code = ErrorCode::InvalidArgument,
                      const char* file = nullptr, int line = 0);
    
    const std::string& getOperation() const { return operation_; }

private:
    std::string operation_;
};

/**
 * @brief 工作表相关异常
 */
class WorksheetException : public FastExcelException {
public:
    WorksheetException(const std::string& message,
                      const std::string& worksheet_name = "",
                      ErrorCode code = ErrorCode::InvalidWorksheet,
                      const char* file = nullptr, int line = 0);
    
    const std::string& getWorksheetName() const { return worksheet_name_; }

private:
    std::string worksheet_name_;
};

/**
 * @brief 单元格相关异常
 */
class CellException : public FastExcelException {
public:
    CellException(const std::string& message,
                 int row = -1, int col = -1,
                 ErrorCode code = ErrorCode::InvalidCellReference,
                 const char* file = nullptr, int line = 0);
    
    int getRow() const { return row_; }
    int getCol() const { return col_; }
    std::string getCellReference() const;

private:
    int row_;
    int col_;
};

/**
 * @brief XML解析异常
 */
class XMLException : public FastExcelException {
public:
    XMLException(const std::string& message,
                const std::string& xml_path = "",
                int xml_line = -1,
                const char* file = nullptr, int line = 0);
    
    const std::string& getXMLPath() const { return xml_path_; }
    int getXMLLine() const { return xml_line_; }

private:
    std::string xml_path_;
    int xml_line_;
};

/**
 * @brief 错误处理器接口
 */
class ErrorHandler {
public:
    virtual ~ErrorHandler() = default;
    
    /**
     * @brief 处理错误
     * @param exception 异常对象
     * @return 是否继续执行
     */
    virtual bool handleError(const FastExcelException& exception) = 0;
    
    /**
     * @brief 处理警告
     * @param message 警告消息
     * @param context 上下文信息
     */
    virtual void handleWarning(const std::string& message, 
                              const std::string& context = "") = 0;
};

/**
 * @brief 默认错误处理器
 */
class DefaultErrorHandler : public ErrorHandler {
public:
    DefaultErrorHandler(bool throw_on_error = true, bool log_warnings = true);
    
    bool handleError(const FastExcelException& exception) override;
    void handleWarning(const std::string& message, 
                      const std::string& context = "") override;
    
    void setThrowOnError(bool throw_on_error) { throw_on_error_ = throw_on_error; }
    void setLogWarnings(bool log_warnings) { log_warnings_ = log_warnings; }

private:
    bool throw_on_error_;
    bool log_warnings_;
};

/**
 * @brief 错误管理器
 */
class ErrorManager {
public:
    static ErrorManager& getInstance();
    
    /**
     * @brief 设置错误处理器
     */
    void setErrorHandler(std::unique_ptr<ErrorHandler> handler);
    
    /**
     * @brief 获取错误处理器
     */
    ErrorHandler* getErrorHandler() const { return error_handler_.get(); }
    
    /**
     * @brief 处理错误
     */
    bool handleError(const FastExcelException& exception);
    
    /**
     * @brief 处理警告
     */
    void handleWarning(const std::string& message, const std::string& context = "");
    
    /**
     * @brief 获取错误统计
     */
    struct ErrorStatistics {
        size_t total_errors = 0;
        size_t total_warnings = 0;
        size_t handled_errors = 0;
        size_t unhandled_errors = 0;
    };
    
    ErrorStatistics getStatistics() const;
    void resetStatistics();

private:
    ErrorManager();
    ~ErrorManager() = default;
    
    std::unique_ptr<ErrorHandler> error_handler_;
    mutable std::mutex mutex_;
    ErrorStatistics stats_;
    
    // 禁用拷贝和赋值
    ErrorManager(const ErrorManager&) = delete;
    ErrorManager& operator=(const ErrorManager&) = delete;
};

} // namespace core
} // namespace fastexcel

// 便捷宏定义
#define FASTEXCEL_THROW(ExceptionType, message) \
    throw ExceptionType(message, __FILE__, __LINE__)

#define FASTEXCEL_THROW_IF(condition, ExceptionType, message) \
    do { if (condition) { FASTEXCEL_THROW(ExceptionType, message); } } while(0)

#define FASTEXCEL_HANDLE_ERROR(exception) \
    fastexcel::core::ErrorManager::getInstance().handleError(exception)

#define FASTEXCEL_HANDLE_WARNING(message, context) \
    fastexcel::core::ErrorManager::getInstance().handleWarning(message, context)

#endif // FASTEXCEL_EXCEPTION_HPP