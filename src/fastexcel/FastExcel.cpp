#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/Logger.hpp"

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

void cleanup() {
    // 清理日志系统
    fastexcel::Logger::getInstance().shutdown();
}

} // namespace fastexcel