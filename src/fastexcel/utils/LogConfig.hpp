#pragma once

// 日志控制宏
// 设置为 0 禁用特定类型的日志，设置为 1 启用

#define ENABLE_ZIP_DEBUG_LOGS 0      // 禁用ZIP相关的调试日志
#define ENABLE_BATCH_DEBUG_LOGS 0    // 禁用批处理相关的调试日志
#define ENABLE_WORKSHEET_DEBUG_LOGS 0 // 禁用工作表调试日志

// 前向声明 - 这些宏在ModuleLoggers.hpp中实际定义
// 这里只是为了避免包含循环依赖，真正的定义在ModuleLoggers.hpp

// 条件日志宏将在ModuleLoggers.hpp中使用模块宏来定义
// 例如：
// FASTEXCEL_LOG_ZIP_DEBUG -> ARCHIVE_DEBUG 
// FASTEXCEL_LOG_BATCH_DEBUG -> CORE_DEBUG
// FASTEXCEL_LOG_WORKSHEET_DEBUG -> CORE_DEBUG