#pragma once

#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#endif

#include <memory>
#include <string>
#include <fstream>
#include <iostream>
#include <mutex>
#include <chrono>
#include <iomanip>
#include <sstream>
#include <vector>
#include <atomic>
#include <fmt/format.h>
#include <fmt/chrono.h>

#ifdef ERROR
#undef ERROR
#endif

namespace fastexcel {

class Logger {
public:
    enum class Level {
        TRACE = 0,
        DEBUG = 1,
        INFO = 2,
        WARN = 3,
        ERROR = 4,
        CRITICAL = 5,
        OFF = 6
    };

    static Logger& getInstance();
    
    void initialize(const std::string& log_file_path = "logs/fastexcel.log", 
                   Level level = Level::INFO, 
                   bool enable_console = true,
                   size_t max_file_size = 10 * 1024 * 1024,
                   size_t max_files = 5);
    
    void setLevel(Level level);
    Level getLevel() const;
    
    void trace(const std::string& message);
    void debug(const std::string& message);
    void info(const std::string& message);
    void warn(const std::string& message);
    void error(const std::string& message);
    void critical(const std::string& message);
    
    template<typename... Args>
    inline void trace(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::TRACE)) {
            try {
                trace(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                trace(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void debug(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::DEBUG)) {
            try {
                debug(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                debug(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void info(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::INFO)) {
            try {
                info(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                info(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void warn(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::WARN)) {
            try {
                warn(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                warn(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void error(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::ERROR)) {
            try {
                error(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                error(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void critical(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::CRITICAL)) {
            try {
                critical(fmt::format(fmt_str, std::forward<Args>(args)...));
            } catch (...) {
                critical(fmt_str);
            }
        }
    }
    
    void flush();
    void shutdown();

private:
    Logger() = default;
    ~Logger();
    Logger(const Logger&) = delete;
    Logger& operator=(const Logger&) = delete;
    
    bool should_log(Level level) const;
    void log_to_console(Level level, const std::string& message);
    void log_to_file(const std::string& message);
    std::string format_message(Level level, const std::string& message) const;
    std::string level_to_string(Level level) const;
    std::string get_timestamp() const;
    void rotate_file_if_needed();
    std::string get_rotated_filename(size_t index) const;
    
    mutable std::mutex mutex_;
    std::atomic<Level> current_level_{Level::INFO};
    std::atomic<bool> initialized_{false};
    std::atomic<bool> enable_console_{true};
    std::atomic<bool> shutting_down_{false};
    
    std::string log_file_path_;
    std::ofstream file_stream_;
    std::atomic<size_t> current_file_size_{0};
    size_t max_file_size_ = 10 * 1024 * 1024;
    size_t max_files_ = 5;
};

#define LOG_TRACE(...)    fastexcel::Logger::getInstance().trace(__VA_ARGS__)
#define LOG_DEBUG(...)    fastexcel::Logger::getInstance().debug(__VA_ARGS__)
#define LOG_INFO(...)     fastexcel::Logger::getInstance().info(__VA_ARGS__)
#define LOG_WARN(...)     fastexcel::Logger::getInstance().warn(__VA_ARGS__)
#define LOG_ERROR(...)    fastexcel::Logger::getInstance().error(__VA_ARGS__)
#define LOG_CRITICAL(...) fastexcel::Logger::getInstance().critical(__VA_ARGS__)

}