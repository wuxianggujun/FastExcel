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
#include <cstring>
#include <algorithm>

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
    
    enum class WriteMode {
        TRUNCATE = 0,  // 覆盖模式（默认）
        APPEND = 1     // 追加模式
    };

    static Logger& getInstance();
    
    void initialize(const std::string& log_file_path = "logs/fastexcel.log", 
                   Level level = Level::INFO, 
                   bool enable_console = true,
                   size_t max_file_size = 10 * 1024 * 1024,
                   size_t max_files = 5,
                   WriteMode write_mode = WriteMode::TRUNCATE);
    
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
                trace(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                trace(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void debug(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::DEBUG)) {
            try {
                debug(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                debug(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void info(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::INFO)) {
            try {
                info(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                info(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void warn(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::WARN)) {
            try {
                warn(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                warn(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void error(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::ERROR)) {
            try {
                error(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                error(fmt_str);
            }
        }
    }
    
    template<typename... Args>
    inline void critical(const std::string& fmt_str, Args&&... args) {
        if (should_log(Level::CRITICAL)) {
            try {
                critical(fmt::vformat(fmt_str, fmt::make_format_args(args...)));
            } catch (...) {
                critical(fmt_str);
            }
        }
    }
    
    void flush();
    void shutdown();

    // 带源码位置信息的便捷接口（在宏中使用）
    template<typename... Args>
    inline void traceCtx(const char* file, int line, const char* func,
                         const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::TRACE)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        trace(fmt_with_ctx, std::forward<Args>(args)...);
    }

    template<typename... Args>
    inline void debugCtx(const char* file, int line, const char* func,
                         const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::DEBUG)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        debug(fmt_with_ctx, std::forward<Args>(args)...);
    }

    template<typename... Args>
    inline void infoCtx(const char* file, int line, const char* func,
                        const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::INFO)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        info(fmt_with_ctx, std::forward<Args>(args)...);
    }

    template<typename... Args>
    inline void warnCtx(const char* file, int line, const char* func,
                        const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::WARN)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        warn(fmt_with_ctx, std::forward<Args>(args)...);
    }

    template<typename... Args>
    inline void errorCtx(const char* file, int line, const char* func,
                         const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::ERROR)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        error(fmt_with_ctx, std::forward<Args>(args)...);
    }

    template<typename... Args>
    inline void criticalCtx(const char* file, int line, const char* func,
                            const std::string& fmt_str, Args&&... args) {
        if (!should_log(Level::CRITICAL)) return;
        const std::string fmt_with_ctx = fmt::format("[{}:{}:{}] {}", baseFilename(file), line, extractFunctionName(func), fmt_str);
        critical(fmt_with_ctx, std::forward<Args>(args)...);
    }

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
    
    // 提取文件名（去除路径）
    static inline const char* baseFilename(const char* path) {
        if (!path) return "";
        const char* slash1 = std::strrchr(path, '/');
        const char* slash2 = std::strrchr(path, '\\');
        const char* p = (slash1 && slash2) ? (std::max(slash1, slash2)) : (slash1 ? slash1 : slash2);
        return p ? (p + 1) : path;
    }
    
    // 提取函数名（去除命名空间和参数）
    static inline std::string extractFunctionName(const char* func_sig) {
        if (!func_sig) return "";
        
        std::string sig(func_sig);
        
        // 找到最后一个双冒号的位置（处理命名空间）
        size_t lastColon = sig.rfind("::");
        if (lastColon != std::string::npos) {
            sig = sig.substr(lastColon + 2);
        }
        
        // 找到第一个括号的位置（去除参数）
        size_t paren = sig.find('(');
        if (paren != std::string::npos) {
            sig = sig.substr(0, paren);
        }
        
        return sig;
    }
    
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
    WriteMode write_mode_ = WriteMode::TRUNCATE;
};

// 跨编译器的函数签名宏，使用简洁的函数名而不是完整签名
#if defined(_MSC_VER)
#  define FASTEXCEL_FUNC __FUNCTION__
#elif defined(__GNUC__) || defined(__clang__)
#  define FASTEXCEL_FUNC __FUNCTION__
#else
#  define FASTEXCEL_FUNC __func__
#endif

// 统一日志宏（带源码位置信息，不包含模块前缀）
#define FASTEXCEL_LOG_TRACE(fmt, ...)    fastexcel::Logger::getInstance().traceCtx   (__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)
#define FASTEXCEL_LOG_DEBUG(fmt, ...)    fastexcel::Logger::getInstance().debugCtx   (__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)
#define FASTEXCEL_LOG_INFO(fmt, ...)     fastexcel::Logger::getInstance().infoCtx    (__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)
#define FASTEXCEL_LOG_WARN(fmt, ...)     fastexcel::Logger::getInstance().warnCtx    (__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)
#define FASTEXCEL_LOG_ERROR(fmt, ...)    fastexcel::Logger::getInstance().errorCtx   (__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)
#define FASTEXCEL_LOG_CRITICAL(fmt, ...) fastexcel::Logger::getInstance().criticalCtx(__FILE__, __LINE__, FASTEXCEL_FUNC, fmt, ##__VA_ARGS__)

// 统一的错误/告警处理宏（仅记录日志，不做额外处理）
#ifndef FASTEXCEL_HANDLE_WARNING
#  define FASTEXCEL_HANDLE_WARNING(message, context) \
    do { FASTEXCEL_LOG_WARN("[ctx:{}] {}", (context), (message)); } while (0)
#endif

#ifndef FASTEXCEL_HANDLE_ERROR
#  define FASTEXCEL_HANDLE_ERROR(ex) \
    do { FASTEXCEL_LOG_ERROR("{}", (ex).what()); } while (0)
#endif

} 
