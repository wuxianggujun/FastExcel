#pragma once

// FastExcel库 - 高性能Excel文件处理库
// 版本 2.0 - 重构后的现代C++架构

// === 核心公共接口 ===

// 标准库依赖
#include <string>
#include <memory>
#include <stdexcept>

// 公共类型定义
#include "fastexcel/core/WorkbookTypes.hpp"
#include "fastexcel/core/WorksheetTypes.hpp"
#include "fastexcel/core/FormatTypes.hpp"

// === 前向声明 ===

namespace fastexcel {
namespace core {
    // 核心类型前向声明
    class Workbook;
    class Worksheet;
    class Cell;
    class FormatDescriptor;
    class StyleBuilder;
    class Color;
    
    // 稳定接口
    namespace interfaces {
        class IWorkbook;
        class IWorksheet;
    }
}

namespace utils {
    class Logger;
}
}

// 版本信息
#define FASTEXCEL_VERSION_MAJOR 2
#define FASTEXCEL_VERSION_MINOR 0
#define FASTEXCEL_VERSION_PATCH 0
#define FASTEXCEL_VERSION_STRING "2.0.0"

namespace fastexcel {

// 版本信息
inline std::string getVersion() {
    return FASTEXCEL_VERSION_STRING;
}

} // namespace fastexcel

// 平台检测
#ifdef _WIN32
    #define FASTEXCEL_WINDOWS
    #ifndef NOMINMAX
        #define NOMINMAX
    #endif
#elif defined(__linux__)
    #define FASTEXCEL_LINUX
#elif defined(__APPLE__)
    #define FASTEXCEL_MACOS
#endif

// 编译器检测
#ifdef _MSC_VER
    #define FASTEXCEL_COMPILER_MSVC
#elif defined(__GNUC__)
    #define FASTEXCEL_COMPILER_GCC
#elif defined(__clang__)
    #define FASTEXCEL_COMPILER_CLANG
#endif

// C++标准检查
#if __cplusplus >= 201703L
    #define FASTEXCEL_CPP17
#endif

// 导出宏定义
#ifdef FASTEXCEL_WINDOWS
    #ifdef FASTEXCEL_SHARED
        #ifdef FASTEXCEL_EXPORTS
            #define FASTEXCEL_API __declspec(dllexport)
        #else
            #define FASTEXCEL_API __declspec(dllimport)
        #endif
    #else
        #define FASTEXCEL_API
    #endif
#else
    #define FASTEXCEL_API
#endif

// 库初始化和清理
namespace fastexcel {

/**
 * @brief 初始化FastExcel库
 * @param log_file_path 日志文件路径
 * @param enable_console 是否启用控制台日志
 * @return 初始化是否成功
 */
FASTEXCEL_API bool initialize(const std::string& log_file_path = "logs/fastexcel.log",
                             bool enable_console = true);

/**
 * @brief 清理FastExcel库资源
 */
FASTEXCEL_API void cleanup();

// === 公共接口 ===

/**
 * @brief 创建新的Excel工作簿
 */
FASTEXCEL_API std::unique_ptr<core::Workbook> createWorkbook();

/**
 * @brief 打开现有的Excel文件
 */
FASTEXCEL_API std::unique_ptr<core::Workbook> openWorkbook(const std::string& filename);

/**
 * @brief 创建样式构建器
 */
FASTEXCEL_API core::StyleBuilder createStyle();

// 类型别名
using StyleBuilder = core::StyleBuilder;
using FormatDescriptor = core::FormatDescriptor;
using BorderStyle = core::BorderStyle;
using PatternType = core::PatternType;
using HorizontalAlign = core::HorizontalAlign;
using VerticalAlign = core::VerticalAlign;

// 错误处理
class FASTEXCEL_API FastExcelException : public std::exception {
private:
    std::string message_;

public:
    explicit FastExcelException(const std::string& message) : message_(message) {}
    
    const char* what() const noexcept override {
        return message_.c_str();
    }
};

class FASTEXCEL_API StyleException : public FastExcelException {
public:
    explicit StyleException(const std::string& message)
        : FastExcelException("Style error: " + message) {}
};

} // namespace fastexcel
