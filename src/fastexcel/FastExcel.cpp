#include "FastExcel.hpp"

// 实际实现的头文件包含（仅在.cpp中）
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/ReadOnlyWorkbook.hpp"
#include "fastexcel/core/StyleBuilder.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>

namespace fastexcel {

// 初始化和清理函数实现

FASTEXCEL_API bool initialize(const std::string& log_file_path, bool enable_console) {
    try {
        // 初始化日志系统
        Logger::getInstance().initialize(log_file_path, Logger::Level::INFO, enable_console);
        FASTEXCEL_LOG_INFO("FastExcel library initialized successfully");
        FASTEXCEL_LOG_INFO("Version: {}", getVersion());
        
        return true;
    } catch (const std::exception& e) {
        // 如果日志系统初始化失败，输出到标准错误
        if (enable_console) {
            std::cerr << "Failed to initialize FastExcel: " << e.what() << std::endl;
        }
        return false;
    }
}

FASTEXCEL_API void cleanup() {
    try {
        FASTEXCEL_LOG_INFO("FastExcel library cleanup completed");
        // 日志系统会在程序结束时自动清理
    } catch (const std::exception& e) {
        // 清理时记录异常但不抛出，避免析构函数中抛出异常
        try {
            FASTEXCEL_LOG_ERROR("Exception during cleanup: {}", e.what());
        } catch (...) {
            // 如果连日志都失败了，则静默忽略
        }
    }
}

// === 新的类型安全工厂方法 ===

FASTEXCEL_API std::unique_ptr<core::Workbook> createWorkbook(const std::string& filename) {
    return core::Workbook::create(filename);
}

FASTEXCEL_API std::unique_ptr<core::ReadOnlyWorkbook> openReadOnly(const std::string& filename) {
    return core::ReadOnlyWorkbook::fromFile(filename);
}

FASTEXCEL_API std::unique_ptr<core::ReadOnlyWorkbook> openReadOnly(const std::string& filename, 
                                                                   const core::WorkbookOptions& options) {
    return core::ReadOnlyWorkbook::fromFile(filename, options);
}

FASTEXCEL_API std::unique_ptr<core::Workbook> openEditable(const std::string& filename) {
    return core::Workbook::openEditable(filename);
}

// === 已废弃的旧方法（为了向后兼容，但建议使用新方法）===
// 注意：这个方法不再推荐使用，请改用 openEditable() 或 openReadOnly()

FASTEXCEL_API std::unique_ptr<core::Workbook> openWorkbook(const std::string& filename) {
    try {
        auto workbook = std::make_unique<core::Workbook>(core::Path(filename));
        
        if (!workbook->open(filename)) {
            FASTEXCEL_LOG_ERROR("Failed to open workbook: {}", filename);
            return nullptr;
        }
        
        FASTEXCEL_LOG_DEBUG("Opened workbook: {}", filename);
        return workbook;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to open workbook {}: {}", filename, e.what());
        return nullptr;
    }
}

FASTEXCEL_API core::StyleBuilder createStyle() {
        return core::StyleBuilder();
}

} // namespace fastexcel
