#pragma once

// FastExcel库 - 高性能Excel文件处理库
// 基于minizip-ng、zlib-ng、libexpat、spdlog等依赖库实现
// 版本 2.0 - 重构后的现代C++架构，支持新旧API共存

// === 新架构组件 (推荐使用) ===
// 核心组件 - 新架构
#include "core/FormatDescriptor.hpp"    // 不可变格式描述符
#include "core/FormatRepository.hpp"    // 线程安全格式仓储  
#include "core/StyleBuilder.hpp"        // 流畅样式构建器
#include "core/StyleTransferContext.hpp" // 跨工作簿样式传输
#include "core/WorkbookNew.hpp"         // 新工作簿API
#include "core/WorksheetNew.hpp"        // 新工作表API

// XML序列化层 - 新架构
#include "xml/StyleSerializer.hpp"      // XLSX样式序列化器

// === 旧架构组件 (向后兼容) ===

// 核心组件
#include "core/Color.hpp"
#include "core/Cell.hpp"
#include "core/Format.hpp"
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
 * @brief 新架构的核心命名空间
 * 
 * 采用现代C++设计模式：
 * - 不可变值对象 (FormatDescriptor)
 * - Repository模式 (FormatRepository)
 * - Builder模式 (StyleBuilder)
 * - 线程安全设计
 * - 自动样式去重
 */
namespace core {
    // 前向声明新架构类型
    class FormatDescriptor;
    class FormatRepository;  
    class StyleBuilder;
    class StyleTransferContext;
    class NewWorkbook;
    class NewWorksheet;
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

// ========== 新架构便捷工厂函数 ==========

/**
 * @brief 创建新架构的Excel工作簿
 * @param filename 文件名
 * @return 工作簿智能指针
 */
FASTEXCEL_API std::unique_ptr<core::NewWorkbook> createWorkbook(const std::string& filename);

/**
 * @brief 打开现有Excel文件(新架构)
 * @param filename 文件名  
 * @return 工作簿智能指针
 */
FASTEXCEL_API std::unique_ptr<core::NewWorkbook> openWorkbook(const std::string& filename);

/**
 * @brief 创建样式构建器
 * @return 样式构建器
 */
FASTEXCEL_API core::StyleBuilder createStyle();

/**
 * @brief 创建命名样式
 * @param name 样式名称
 * @param builder 样式构建器
 * @return 命名样式
 */
FASTEXCEL_API core::NamedStyle createNamedStyle(const std::string& name, const core::StyleBuilder& builder);

// ========== 预定义样式工厂 ==========

/**
 * @brief 预定义样式命名空间
 */
namespace styles {
    FASTEXCEL_API core::StyleBuilder title();      // 标题样式
    FASTEXCEL_API core::StyleBuilder header();     // 表头样式  
    FASTEXCEL_API core::StyleBuilder money();      // 货币样式
    FASTEXCEL_API core::StyleBuilder percent();    // 百分比样式
    FASTEXCEL_API core::StyleBuilder date();       // 日期样式
    
    // 便捷样式创建函数
    FASTEXCEL_API core::StyleBuilder border(core::BorderStyle style, core::Color color = core::Color::BLACK);
    FASTEXCEL_API core::StyleBuilder fill(core::Color color);
    FASTEXCEL_API core::StyleBuilder font(const std::string& name, double size, bool bold = false);
}

// ========== 类型别名 (便于使用) ==========

// 新架构类型别名
using StyleBuilder = core::StyleBuilder;
using NamedStyle = core::NamedStyle; 
using FormatDescriptor = core::FormatDescriptor;
using FormatRepository = core::FormatRepository;
using StyleTransferContext = core::StyleTransferContext;
using NewWorkbook = core::NewWorkbook;      // 新工作簿类型
using NewWorksheet = core::NewWorksheet;    // 新工作表类型

// 保持旧架构别名(向后兼容)
using LegacyWorkbook = fastexcel::core::Workbook;   // 旧工作簿类型
using LegacyWorksheet = fastexcel::core::Worksheet; // 旧工作表类型

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

// ========== 功能特性检测 ==========

/**
 * @brief 检查新架构功能是否可用
 * @return 是否支持新架构功能
 */
FASTEXCEL_API bool hasNewArchitectureSupport();

/**
 * @brief 获取新架构功能描述
 * @return 功能描述字符串
 */
FASTEXCEL_API std::string getNewArchitectureFeatures();

/**
 * @brief 迁移指南信息
 */  
struct FASTEXCEL_API MigrationGuide {
    static const char* OLD_API_EXAMPLE;
    static const char* NEW_API_EXAMPLE;
    static const char* MIGRATION_STEPS;
};

} // namespace fastexcel
