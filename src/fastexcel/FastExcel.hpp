#pragma once

// FastExcel库 - 高性能Excel文件处理库
// 基于minizip-ng、zlib-ng、libexpat、spdlog等依赖库实现
// 版本 2.0 - 重构后的现代C++架构，支持新旧API共存

// 核心组件
#include "fastexcel/core/FormatDescriptor.hpp"    // 不可变格式描述符
#include "fastexcel/core/FormatRepository.hpp"    // 线程安全格式仓储  
#include "fastexcel/core/StyleBuilder.hpp"        // 流畅样式构建器
#include "fastexcel/core/StyleTransferContext.hpp" // 跨工作簿样式传输

// XML序列化层
#include "fastexcel/xml/StyleSerializer.hpp"      // XLSX样式序列化器

// === 旧架构组件 (向后兼容) ===

// 核心组件
#include "fastexcel/core/Color.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/ThreadPool.hpp"

// XML处理层
#include "fastexcel/xml/ContentTypes.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"

// 压缩归档层
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/archive/ZipArchive.hpp"

// 工具层
#include "fastexcel/utils/Logger.hpp"  // 直接包含日志头文件，用户无需手动包含

// 版本信息
#define FASTEXCEL_VERSION_MAJOR 2
#define FASTEXCEL_VERSION_MINOR 0
#define FASTEXCEL_VERSION_PATCH 0

#define FASTEXCEL_VERSION_STRING "2.0.0 - Modern C++ Architecture"

// 便捷宏定义
#define FASTEXCEL_STRINGIFY(x) #x
#define FASTEXCEL_TOSTRING(x) FASTEXCEL_STRINGIFY(x)

#define FASTEXCEL_VERSION \
    FASTEXCEL_TOSTRING(FASTEXCEL_VERSION_MAJOR) "." \
    FASTEXCEL_TOSTRING(FASTEXCEL_VERSION_MINOR) "." \
    FASTEXCEL_TOSTRING(FASTEXCEL_VERSION_PATCH)

// 命名空间
namespace fastexcel {

// 版本信息函数
inline std::string getVersion() {
    return FASTEXCEL_VERSION_STRING;
}

// 初始化和清理函数已移至下方统一声明，此处不再重复声明

// 日志宏现在统一在 utils/Logger.hpp 中定义
// FastExcel.hpp 已自动包含 Logger.hpp，用户可直接使用日志宏

} // namespace fastexcel

// 平台特定的宏定义
#ifdef _WIN32
    #define FASTEXCEL_WINDOWS
    #ifndef NOMINMAX
        #define NOMINMAX
    #endif
    // 避免在公共头直接引入 <windows.h> 以减少命名污染和编译开销
#endif

#ifdef __linux__
    #define FASTEXCEL_LINUX
#endif

#ifdef __APPLE__
    #define FASTEXCEL_MACOS
#endif

// 编译器特定的宏定义
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

// 导出宏定义（用于Windows DLL）
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
 * @brief 初始化FastExcel库（统一接口）
 * @param log_file_path 日志文件路径（默认："logs/fastexcel.log"）
 * @param enable_console 是否启用控制台日志（默认：true）
 * @return 初始化是否成功
 *
 * @note 此函数是幂等的，多次调用是安全的
 */
FASTEXCEL_API bool initialize(const std::string& log_file_path = "logs/fastexcel.log",
                             bool enable_console = true);

/**
 * @brief 初始化FastExcel库（简化版本）
 * 使用默认参数初始化
 */
FASTEXCEL_API inline void initialize() {
    initialize("logs/fastexcel.log", true);
}

/**
 * @brief 清理FastExcel库资源
 *
 * 清理全局资源、日志系统等
 * 程序结束前调用，幂等操作
 */
FASTEXCEL_API void cleanup();

// API 2.0 （推荐使用）

/**
 * @namespace fastexcel::core
 * @brief 核心命名空间
 * 
 * 包含现代C++设计的核心类：
 * - 不可变值对象 (FormatDescriptor)
 * - Repository模式 (FormatRepository)  
 * - Builder模式 (StyleBuilder)
 * - 线程安全设计
 * - 自动样式去重
 */
namespace core {
    // 前向声明核心类型
    class FormatDescriptor;
    class FormatRepository;  
    class StyleBuilder;
    class StyleTransferContext;
}

/**
 * @brief 版本信息
 */
struct NewArchitectureVersion {
    static constexpr int MAJOR = 2;
    static constexpr int MINOR = 0;
    static constexpr int PATCH = 0;
    static constexpr const char* STRING = "2.0.0";
    static constexpr const char* DESCRIPTION = "Modern C++ Architecture with immutable design patterns";
};

// 样式构建便捷函数

/**
 * @brief 创建样式构建器
 * @return 样式构建器
 */
FASTEXCEL_API core::StyleBuilder createStyle();

// 类型别名 (便于使用)

// 核心类型别名
using StyleBuilder = core::StyleBuilder;
using FormatDescriptor = core::FormatDescriptor;
using FormatRepository = core::FormatRepository;
using StyleTransferContext = core::StyleTransferContext;

// 保持向后兼容的别名
using LegacyWorkbook = fastexcel::core::Workbook;   // 当前工作簿类型
using LegacyWorksheet = fastexcel::core::Worksheet; // 当前工作表类型

// 枚举类型别名
using BorderStyle = core::BorderStyle;
using PatternType = core::PatternType;  
using HorizontalAlign = core::HorizontalAlign;
using VerticalAlign = core::VerticalAlign;
using UnderlineType = core::UnderlineType;
using FontScript = core::FontScript;

// 错误处理

/**
 * @brief FastExcel异常基类
 */
class FASTEXCEL_API FastExcelException : public std::exception {
private:
    std::string message_;

public:
    explicit FastExcelException(const std::string& message) : message_(message) {}
    
    const char* what() const noexcept override {
        return message_.c_str();
    }
};

/**
 * @brief 专用异常
 */
class FASTEXCEL_API StyleException : public FastExcelException {
public:
    explicit StyleException(const std::string& message)
        : FastExcelException("Style error: " + message) {}
};

} // namespace fastexcel
