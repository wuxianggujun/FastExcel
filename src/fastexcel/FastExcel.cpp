#include "FastExcel.hpp"

// 实际实现的头文件包含（仅在.cpp中）
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
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
        // 初始化日志系统 - 使用静态初始化
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
    } catch (...) {
        // 清理时忽略异常
    }
}

// API 2.0 工厂函数实现

FASTEXCEL_API std::unique_ptr<core::Workbook> createWorkbook() {
    try {
        // 使用默认的临时文件名
        auto workbook = std::make_unique<core::Workbook>(core::Path("temp.xlsx"));
        FASTEXCEL_LOG_DEBUG("Created new workbook");
        return workbook;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to create workbook: {}", e.what());
        return nullptr;
    }
}

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
    try {
        return core::StyleBuilder();
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to create style builder: {}", e.what());
        throw StyleException("Failed to create style builder: " + std::string(e.what()));
    }
}

} // namespace fastexcel