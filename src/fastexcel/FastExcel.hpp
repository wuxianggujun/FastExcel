#pragma once

// FastExcel库 - 高性能Excel文件处理库
// 基于minizip-ng、zlib-ng、libexpat、spdlog等依赖库实现
// 版本 2.0 - 重构后的现代C++架构，支持新旧API共存

// === 新架构组件 (推荐使用) ===
// 核心组件 - 新架构
#include "fastexcel/core/FormatDescriptor.hpp"    // 不可变格式描述符
#include "fastexcel/core/FormatRepository.hpp"    // 线程安全格式仓储  
#include "fastexcel/core/StyleBuilder.hpp"        // 流畅样式构建器
#include "fastexcel/core/StyleTransferContext.hpp" // 跨工作簿样式传输

// XML序列化层 - 新架构
#include "fastexcel/xml/StyleSerializer.hpp"      // XLSX样式序列化器

// === 旧架构组件 (向后兼容) ===

// 核心组件
#include "core/Color.hpp"
#include "core/Cell.hpp"
#include "core/Workbook.hpp"
#include "core/Worksheet.hpp"
#include "core/ThreadPool.hpp"

// XML处理层
#include "xml/ContentTypes.hpp"
#include "xml/Relationships.hpp"
#include "xml/SharedStrings.hpp"
#include "xml/XMLStreamWriter.hpp"

// 压缩归档层
#include "archive/FileManager.hpp"
#include "archive/ZipArchive.hpp"

// 工具层
#include "utils/Logger.hpp"

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

/**
 * @brief 初始化FastExcel库
 * 
 * 初始化日志系统、内存池、全局资源等
 * 在使用FastExcel任何功能前调用
 */
void initialize();

/**
 * @brief 清理FastExcel库
 * 
 * 清理全局资源、日志系统等
 * 程序结束前调用
 */
void cleanup();

// 日志便捷宏定义
#define LOG_TRACE(...)    fastexcel::Logger::getInstance().trace(__VA_ARGS__)
#define LOG_DEBUG(...)    fastexcel::Logger::getInstance().debug(__VA_ARGS__)
#define LOG_INFO(...)     fastexcel::Logger::getInstance().info(__VA_ARGS__)
#define LOG_WARN(...)     fastexcel::Logger::getInstance().warn(__VA_ARGS__)
#define LOG_ERROR(...)    fastexcel::Logger::getInstance().error(__VA_ARGS__)
#define LOG_CRITICAL(...) fastexcel::Logger::getInstance().critical(__VA_ARGS__)

} // namespace fastexcel

// 平台特定的宏定义
#ifdef _WIN32
    #define FASTEXCEL_WINDOWS
    #ifndef NOMINMAX
        #define NOMINMAX
    #endif
    #include <windows.h>
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

// ========== 新架构 2.0 API (推荐使用) ==========

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
 * @brief 新架构版本信息
 */
struct NewArchitectureVersion {
    static constexpr int MAJOR = 2;
    static constexpr int MINOR = 0;
    static constexpr int PATCH = 0;
    static constexpr const char* STRING = "2.0.0";
    static constexpr const char* DESCRIPTION = "Modern C++ Architecture with immutable design patterns";
};

// ========== 样式构建便捷函数 ==========

/**
 * @brief 创建样式构建器
 * @return 样式构建器
 */
FASTEXCEL_API core::StyleBuilder createStyle();

// ========== 类型别名 (便于使用) ==========

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

// ========== 错误处理 ==========

/**
 * @brief FastExcel异常基类(新架构)
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
 * @brief 新架构专用异常
 */
class FASTEXCEL_API StyleException : public FastExcelException {
public:
    explicit StyleException(const std::string& message)
        : FastExcelException("Style error: " + message) {}
};

} // namespace fastexcel
