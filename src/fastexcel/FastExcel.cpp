#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/MemoryPool.hpp"

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
    
    // 初始化全局内存池等其他资源
    // 这里可以添加其他全局初始化代码
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

} // namespace fastexcel