#pragma once
#include "Logger.hpp"
#include "LogConfig.hpp"

/**
 * @file ModuleLoggers.hpp
 * @brief 模块化日志宏定义
 * 
 * 每个模块都有自己的日志宏，格式: [等级][模块] 消息
 * 便于调试时快速识别日志来源和等级
 */

// 核心模块 (core)
#define CORE_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][CORE] " __VA_ARGS__)
#define CORE_INFO(...)     FASTEXCEL_LOG_INFO("[INF][CORE] " __VA_ARGS__)
#define CORE_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][CORE] " __VA_ARGS__)
#define CORE_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][CORE] " __VA_ARGS__)
#define CORE_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][CORE] " __VA_ARGS__)

// 读取模块 (reader)
#define READER_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][READ] " __VA_ARGS__)
#define READER_INFO(...)     FASTEXCEL_LOG_INFO("[INF][read] " __VA_ARGS__)
#define READER_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][read] " __VA_ARGS__)
#define READER_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][read] " __VA_ARGS__)
#define READER_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][read] " __VA_ARGS__)

// XML模块 (xml)
#define XML_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][xml ] " __VA_ARGS__)
#define XML_INFO(...)     FASTEXCEL_LOG_INFO("[INF][xml ] " __VA_ARGS__)
#define XML_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][xml ] " __VA_ARGS__)
#define XML_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][xml ] " __VA_ARGS__)
#define XML_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][xml ] " __VA_ARGS__)

// 归档模块 (archive)
#define ARCHIVE_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][arch] " __VA_ARGS__)
#define ARCHIVE_INFO(...)     FASTEXCEL_LOG_INFO("[INF][arch] " __VA_ARGS__)
#define ARCHIVE_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][arch] " __VA_ARGS__)
#define ARCHIVE_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][arch] " __VA_ARGS__)
#define ARCHIVE_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][arch] " __VA_ARGS__)

// OPC模块 (opc)
#define OPC_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][opc ] " __VA_ARGS__)
#define OPC_INFO(...)     FASTEXCEL_LOG_INFO("[INF][opc ] " __VA_ARGS__)
#define OPC_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][opc ] " __VA_ARGS__)
#define OPC_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][opc ] " __VA_ARGS__)
#define OPC_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][opc ] " __VA_ARGS__)

// 变更跟踪模块 (tracking)
#define TRACKING_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][trak] " __VA_ARGS__)
#define TRACKING_INFO(...)     FASTEXCEL_LOG_INFO("[INF][trak] " __VA_ARGS__)
#define TRACKING_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][trak] " __VA_ARGS__)
#define TRACKING_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][trak] " __VA_ARGS__)
#define TRACKING_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][trak] " __VA_ARGS__)

// 工具模块 (utils)
#define UTILS_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][util] " __VA_ARGS__)
#define UTILS_INFO(...)     FASTEXCEL_LOG_INFO("[INF][util] " __VA_ARGS__)
#define UTILS_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][util] " __VA_ARGS__)
#define UTILS_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][util] " __VA_ARGS__)
#define UTILS_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][util] " __VA_ARGS__)

// 示例和测试模块 (demo)
#define DEMO_DEBUG(...)    FASTEXCEL_LOG_DEBUG("[DBG][demo] " __VA_ARGS__)
#define DEMO_INFO(...)     FASTEXCEL_LOG_INFO("[INF][demo] " __VA_ARGS__)
#define DEMO_WARN(...)     FASTEXCEL_LOG_WARN("[WRN][demo] " __VA_ARGS__)
#define DEMO_ERROR(...)    FASTEXCEL_LOG_ERROR("[ERR][demo] " __VA_ARGS__)
#define DEMO_CRITICAL(...) FASTEXCEL_LOG_CRITICAL("[CRT][demo] " __VA_ARGS__)

// 示例模块别名 (examples)
#define EXAMPLE_DEBUG(...)    DEMO_DEBUG(__VA_ARGS__)
#define EXAMPLE_INFO(...)     DEMO_INFO(__VA_ARGS__)
#define EXAMPLE_WARN(...)     DEMO_WARN(__VA_ARGS__)
#define EXAMPLE_ERROR(...)    DEMO_ERROR(__VA_ARGS__)
#define EXAMPLE_CRITICAL(...) DEMO_CRITICAL(__VA_ARGS__)

// 条件日志宏 (使用模块宏实现)
#if ENABLE_ZIP_DEBUG_LOGS
    #define FASTEXCEL_LOG_ZIP_DEBUG(...) ARCHIVE_DEBUG(__VA_ARGS__)
#else
    #define FASTEXCEL_LOG_ZIP_DEBUG(...) do {} while(0)
#endif

#if ENABLE_BATCH_DEBUG_LOGS
    #define FASTEXCEL_LOG_BATCH_DEBUG(...) CORE_DEBUG(__VA_ARGS__)
#else
    #define FASTEXCEL_LOG_BATCH_DEBUG(...) do {} while(0)
#endif

#if ENABLE_WORKSHEET_DEBUG_LOGS
    #define FASTEXCEL_LOG_WORKSHEET_DEBUG(...) CORE_DEBUG(__VA_ARGS__)
#else
    #define FASTEXCEL_LOG_WORKSHEET_DEBUG(...) do {} while(0)
#endif
