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

// 注意：简化版本的 initialize() 现在是 inline 函数，在头文件中定义

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
