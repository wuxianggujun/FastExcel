#include "Logger.hpp"
#include <filesystem>
#include <iostream>
#include <spdlog/async.h>

#ifdef _WIN32
// 防止 Windows 宏冲突
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
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
                       size_t max_files) {
    try {
        // 创建日志目录
        std::filesystem::path log_path(log_file_path);
        std::filesystem::path log_dir = log_path.parent_path();
        if (!log_dir.empty() && !std::filesystem::exists(log_dir)) {
            std::filesystem::create_directories(log_dir);
        }

        // 创建sinks
        std::vector<spdlog::sink_ptr> sinks;

        // 控制台sink (支持中文)
        if (enable_console) {
            auto console_sink = std::make_shared<spdlog::sinks::stdout_color_sink_mt>();
            console_sink->set_pattern("[%Y-%m-%d %H:%M:%S.%e] [%^%l%$] [%t] %v");
            sinks.push_back(console_sink);
        }

        // 文件sink (旋转日志)
        auto file_sink = std::make_shared<spdlog::sinks::rotating_file_sink_mt>(
            log_file_path, max_file_size, max_files);
        file_sink->set_pattern("[%Y-%m-%d %H:%M:%S.%e] [%l] [%t] %v");
        sinks.push_back(file_sink);

        // 创建logger
        logger_ = std::make_shared<spdlog::logger>("fastexcel", sinks.begin(), sinks.end());
        
        // 设置日志级别
        setLevel(level);
        
        // 设置刷新策略
        logger_->flush_on(spdlog::level::warn);
        
        // 注册logger
        spdlog::register_logger(logger_);
        
        // 设置为默认logger
        spdlog::set_default_logger(logger_);

        // 设置UTF-8编码支持 (Windows)
#ifdef _WIN32
        SetConsoleOutputCP(CP_UTF8);
        SetConsoleCP(CP_UTF8);
#endif

        LOG_INFO("Logger initialized successfully. Log file: {}", log_file_path);
        
    } catch (const std::exception& ex) {
        std::cerr << "Logger initialization failed: " << ex.what() << std::endl;
    }
}

void Logger::setLevel(Level level) {
    current_level_ = level;
    if (logger_) {
        logger_->set_level(static_cast<spdlog::level::level_enum>(level));
    }
}

Logger::Level Logger::getLevel() const {
    return current_level_;
}

void Logger::flush() {
    if (logger_) {
        logger_->flush();
    }
}

void Logger::shutdown() {
    if (logger_) {
        logger_->flush();
        spdlog::shutdown();
        logger_.reset();
    }
}

Logger::~Logger() {
    shutdown();
}

} // namespace fastexcel
