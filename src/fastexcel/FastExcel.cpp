#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/MemoryPool.hpp"
#include "fastexcel/core/StyleBuilder.hpp"

namespace fastexcel {

bool initialize(const std::string& log_file_path, bool enable_console) {
    // 初始化日志系统
    try {
        fastexcel::Logger::getInstance().initialize(log_file_path,
                                                   fastexcel::Logger::Level::DEBUG,
                                                   enable_console);
        return true;
    } catch (...) {
        return false;
    }
}

void initialize() {
    // 初始化日志系统（使用默认设置）
    try {
        fastexcel::Logger::getInstance().initialize("fastexcel.log",
                                                   fastexcel::Logger::Level::INFO,
                                                   true);
    } catch (...) {
        // 忽略日志初始化错误
    }
}

void cleanup() {
    // 清理内存池
    try {
        fastexcel::core::MemoryManager::getInstance().cleanup();
    } catch (...) {
        // 忽略清理错误
    }
    
    // 清理日志系统
    fastexcel::Logger::getInstance().shutdown();
}

core::StyleBuilder createStyle() {
    return core::StyleBuilder();
}

} // namespace fastexcel
