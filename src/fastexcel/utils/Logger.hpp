#pragma once

// 防止 Windows 宏冲突
#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#endif

#include <spdlog/spdlog.h>
#include <spdlog/sinks/stdout_color_sinks.h>
#include <spdlog/sinks/daily_file_sink.h>
#include <spdlog/sinks/rotating_file_sink.h>
#include <memory>
#include <string>

// 取消可能冲突的 Windows 宏定义
#ifdef ERROR
#undef ERROR
#endif

namespace fastexcel {

class Logger {
public:
    enum class Level {
        TRACE = SPDLOG_LEVEL_TRACE,
        DEBUG = SPDLOG_LEVEL_DEBUG,
        INFO = SPDLOG_LEVEL_INFO,
        WARN = SPDLOG_LEVEL_WARN,
        ERROR = SPDLOG_LEVEL_ERROR,
        CRITICAL = SPDLOG_LEVEL_CRITICAL,
        OFF = SPDLOG_LEVEL_OFF
    };

    static Logger& getInstance();
    
    void initialize(const std::string& log_file_path = "logs/fastexcel.log", 
                   Level level = Level::INFO, 
                   bool enable_console = true,
                   size_t max_file_size = 10 * 1024 * 1024,  // 10MB
                   size_t max_files = 5);
    
    void setLevel(Level level);
    Level getLevel() const;
    
    template<typename... Args>
    void trace(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->trace(fmt, std::forward<Args>(args)...);
        }
    }
    
    template<typename... Args>
    void debug(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->debug(fmt, std::forward<Args>(args)...);
        }
    }
    
    template<typename... Args>
    void info(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->info(fmt, std::forward<Args>(args)...);
        }
    }
    
    template<typename... Args>
    void warn(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->warn(fmt, std::forward<Args>(args)...);
        }
    }
    
    template<typename... Args>
    void error(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->error(fmt, std::forward<Args>(args)...);
        }
    }
    
    template<typename... Args>
    void critical(const std::string& fmt, Args&&... args) {
        if (logger_) {
            logger_->critical(fmt, std::forward<Args>(args)...);
        }
    }
    
    void flush();
    void shutdown();

private:
    Logger() = default;
    ~Logger();
    Logger(const Logger&) = delete;
    Logger& operator=(const Logger&) = delete;
    
    std::shared_ptr<spdlog::logger> logger_;
    Level current_level_ = Level::INFO;
};

// 便捷宏定义
#define LOG_TRACE(...)    fastexcel::Logger::getInstance().trace(__VA_ARGS__)
#define LOG_DEBUG(...)    fastexcel::Logger::getInstance().debug(__VA_ARGS__)
#define LOG_INFO(...)     fastexcel::Logger::getInstance().info(__VA_ARGS__)
#define LOG_WARN(...)     fastexcel::Logger::getInstance().warn(__VA_ARGS__)
#define LOG_ERROR(...)    fastexcel::Logger::getInstance().error(__VA_ARGS__)
#define LOG_CRITICAL(...) fastexcel::Logger::getInstance().critical(__VA_ARGS__)

} // namespace fastexcel