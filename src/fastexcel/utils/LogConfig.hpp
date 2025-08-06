#pragma once

// 日志控制宏
// 设置为 0 禁用特定类型的日志，设置为 1 启用

#define ENABLE_ZIP_DEBUG_LOGS 0      // 禁用ZIP相关的调试日志
#define ENABLE_BATCH_DEBUG_LOGS 0    // 禁用批处理相关的调试日志
#define ENABLE_WORKSHEET_DEBUG_LOGS 0 // 禁用工作表调试日志

// 条件日志宏
#if ENABLE_ZIP_DEBUG_LOGS
    #define LOG_ZIP_DEBUG(...) LOG_DEBUG(__VA_ARGS__)
#else
    #define LOG_ZIP_DEBUG(...) do {} while(0)
#endif

#if ENABLE_BATCH_DEBUG_LOGS
    #define LOG_BATCH_DEBUG(...) LOG_DEBUG(__VA_ARGS__)
#else
    #define LOG_BATCH_DEBUG(...) do {} while(0)
#endif

#if ENABLE_WORKSHEET_DEBUG_LOGS
    #define LOG_WORKSHEET_DEBUG(...) LOG_DEBUG(__VA_ARGS__)
#else
    #define LOG_WORKSHEET_DEBUG(...) do {} while(0)
#endif