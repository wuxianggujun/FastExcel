#include "Logger.hpp"
#include <filesystem>
#include <iostream>
#include <thread>
#include <sstream>
#include <fmt/format.h>
#include <fmt/chrono.h>

#ifdef _WIN32
// 防止 Windows 宏冲突
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <io.h>
#include <fcntl.h>
// 取消可能冲突的 Windows 宏定义
#ifdef ERROR
#undef ERROR
#endif
#endif

namespace fastexcel {

Logger& Logger::getInstance() {
    static Logger instance;
    return instance;
}

void Logger::initialize(const std::string& log_file_path, 
                       Level level, 
                       bool enable_console,
                       size_t max_file_size,
                       size_t max_files,
                       WriteMode write_mode) {
    
    // 如果已经初始化或正在关闭，直接返回
    if (initialized_.load() || shutting_down_.load()) {
        return;
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 双重检查
    if (initialized_.load() || shutting_down_.load()) {
        return;
    }
    
    try {
        // 设置参数
        current_level_.store(level);
        enable_console_.store(enable_console);
        log_file_path_ = log_file_path;
        max_file_size_ = max_file_size;
        max_files_ = max_files;
        write_mode_ = write_mode;
        
        // 创建日志目录
        std::filesystem::path log_path(log_file_path_);
        std::filesystem::path log_dir = log_path.parent_path();
        if (!log_dir.empty() && !std::filesystem::exists(log_dir)) {
            std::filesystem::create_directories(log_dir);
        }
        
        // 打开日志文件
        std::ios::openmode open_mode = (write_mode_ == WriteMode::APPEND) ? 
                                      (std::ios::out | std::ios::app) : 
                                      (std::ios::out | std::ios::trunc);
        file_stream_.open(log_file_path_, open_mode);
        if (file_stream_.is_open() && write_mode_ == WriteMode::APPEND) {
            file_stream_.seekp(0, std::ios::end);
            current_file_size_.store(static_cast<size_t>(file_stream_.tellp()));
        } else {
            current_file_size_.store(0);
        }
        
#ifdef _WIN32
        // 设置UTF-8编码支持
        if (enable_console) {
            SetConsoleOutputCP(CP_UTF8);
            SetConsoleCP(CP_UTF8);
            
            // 启用ANSI转义序列支持（Windows 10+）
            HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
            if (hOut != INVALID_HANDLE_VALUE) {
                DWORD dwMode = 0;
                if (GetConsoleMode(hOut, &dwMode)) {
                    dwMode |= ENABLE_VIRTUAL_TERMINAL_PROCESSING;
                    SetConsoleMode(hOut, dwMode);
                }
            }
        }
#endif
        
        initialized_.store(true);
        
        // 记录初始化成功（不使用递归调用）
        if (should_log(Level::INFO)) {
            std::string msg = format_message(Level::INFO, 
                fmt::format("Logger initialized successfully. Log file: {}, Mode: {}", 
                    log_file_path, (write_mode_ == WriteMode::APPEND ? "APPEND" : "TRUNCATE")));
            
            if (enable_console_.load()) {
                log_to_console(Level::INFO, msg);
            }
            log_to_file(msg);
        }
        
    } catch (const std::exception& ex) {
        std::cerr << "Logger initialization failed: " << ex.what() << std::endl;
    }
}

void Logger::setLevel(Level level) {
    current_level_.store(level);
}

Logger::Level Logger::getLevel() const {
    return current_level_.load();
}

bool Logger::should_log(Level level) const {
    return static_cast<int>(level) >= static_cast<int>(current_level_.load()) && 
           !shutting_down_.load();
}

void Logger::trace(const std::string& message) {
    if (!should_log(Level::TRACE)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::TRACE, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::TRACE, formatted_message);
    }
    log_to_file(formatted_message);
}

void Logger::debug(const std::string& message) {
    if (!should_log(Level::DEBUG)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::DEBUG, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::DEBUG, formatted_message);
    }
    log_to_file(formatted_message);
}

void Logger::info(const std::string& message) {
    if (!should_log(Level::INFO)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::INFO, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::INFO, formatted_message);
    }
    log_to_file(formatted_message);
}

void Logger::warn(const std::string& message) {
    if (!should_log(Level::WARN)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::WARN, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::WARN, formatted_message);
    }
    log_to_file(formatted_message);
    
    // 立即刷新
    flush();
}

void Logger::error(const std::string& message) {
    if (!should_log(Level::ERROR)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::ERROR, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::ERROR, formatted_message);
    }
    log_to_file(formatted_message);
    
    // 立即刷新
    flush();
}

void Logger::critical(const std::string& message) {
    if (!should_log(Level::CRITICAL)) return;
    
    if (!initialized_.load()) {
        initialize();
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    if (shutting_down_.load()) return;
    
    std::string formatted_message = format_message(Level::CRITICAL, message);
    
    if (enable_console_.load()) {
        log_to_console(Level::CRITICAL, formatted_message);
    }
    log_to_file(formatted_message);
    
    // 立即刷新
    flush();
}

void Logger::log_to_console(Level level, const std::string& message) {
    // 根据日志级别设置颜色
    std::string color_code;
    switch (level) {
        case Level::TRACE:    color_code = "\033[37m";   break; // 白色
        case Level::DEBUG:    color_code = "\033[36m";   break; // 青色
        case Level::INFO:     color_code = "\033[32m";   break; // 绿色
        case Level::WARN:     color_code = "\033[33m";   break; // 黄色
        case Level::ERROR:    color_code = "\033[31m";   break; // 红色
        case Level::CRITICAL: color_code = "\033[35m";   break; // 紫色
        default:              color_code = "\033[0m";    break; // 默认
    }
    
    std::cout << color_code << message << "\033[0m" << std::endl;
}

void Logger::log_to_file(const std::string& message) {
    if (!file_stream_.is_open()) {
        return;
    }
    
    rotate_file_if_needed();
    
    file_stream_ << message << std::endl;
    current_file_size_.fetch_add(message.length() + 1); // +1 for newline
}

void Logger::rotate_file_if_needed() {
    if (current_file_size_.load() < max_file_size_) {
        return;
    }
    
    file_stream_.close();
    
    // 旋转日志文件
    for (size_t i = max_files_ - 1; i > 0; --i) {
        std::string old_file = get_rotated_filename(i - 1);
        std::string new_file = get_rotated_filename(i);
        
        if (std::filesystem::exists(old_file)) {
            std::filesystem::rename(old_file, new_file);
        }
    }
    
    // 将当前文件重命名为 .1
    if (std::filesystem::exists(log_file_path_)) {
        std::filesystem::rename(log_file_path_, get_rotated_filename(1));
    }
    
    // 重新打开新文件
    file_stream_.open(log_file_path_, std::ios::out | std::ios::trunc);
    current_file_size_.store(0);
}

std::string Logger::get_rotated_filename(size_t index) const {
    if (index == 0) {
        return log_file_path_;
    }
    return fmt::format("{}.{}", log_file_path_, index);
}

std::string Logger::format_message(Level level, const std::string& message) const {
    auto thread_id = std::this_thread::get_id();
    std::ostringstream oss;
    oss << thread_id;
    return fmt::format("[{}] [{}] [{}] {}", 
                      get_timestamp(), 
                      level_to_string(level), 
                      oss.str(),
                      message);
}

std::string Logger::level_to_string(Level level) const {
    switch (level) {
        case Level::TRACE:    return "TRACE";
        case Level::DEBUG:    return "DEBUG";
        case Level::INFO:     return "INFO ";
        case Level::WARN:     return "WARN ";
        case Level::ERROR:    return "ERROR";
        case Level::CRITICAL: return "CRIT ";
        default:              return "UNKN ";
    }
}

std::string Logger::get_timestamp() const {
    auto now = std::chrono::system_clock::now();
    return fmt::format("{:%Y-%m-%d %H:%M:%S}", now);
}

void Logger::flush() {
    // 不加锁的刷新，避免在已持有锁的情况下再次加锁
    if (file_stream_.is_open()) {
        file_stream_.flush();
    }
    std::cout.flush();
}

void Logger::shutdown() {
    // 设置关闭标志，防止新的日志调用
    shutting_down_.store(true);
    
    // 尝试获取锁，但不要无限等待
    std::unique_lock<std::mutex> lock(mutex_, std::try_to_lock);
    if (!lock.owns_lock()) {
        // 如果无法获取锁，直接返回，避免死锁
        return;
    }
    
    if (initialized_.load()) {
        try {
            if (file_stream_.is_open()) {
                file_stream_.flush();
                file_stream_.close();
            }
        } catch (...) {
            // 忽略关闭时的异常
        }
        initialized_.store(false);
    }
}

Logger::~Logger() {
    shutdown();
}

} // namespace fastexcel
