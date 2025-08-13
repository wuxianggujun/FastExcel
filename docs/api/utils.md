# 工具层（utils）API 与类关系

提供日志、路径、地址解析、时间与通用工具、XML 工具等支撑能力；被各层广泛依赖。

## 关系图（简述）
- `Logger` 为全局日志服务（线程安全），宏 `FASTEXCEL_LOG_*` 提供便捷调用
- `Path` 被 `archive/opc/reader` 等作为统一路径类型
- `AddressParser` 为 `Workbook/Worksheet` 提供 `"Sheet!A1"`/`"A1"` 地址解析
- 其他工具：`CommonUtils/TimeUtils/XMLUtils/LogConfig/ModuleLoggers` 等

---

## class fastexcel::Logger
- 职责：文件/控制台日志，支持轮转、级别/模式配置、fmt 格式化与线程安全。
- 主要 API：
  - 初始化：`initialize(log_path, level, enable_console, max_size, max_files, write_mode)`、`shutdown()`、`flush()`
  - 级别：`setLevel(level)`、`getLevel()`
  - 记录：`trace/debug/info/warn/error/critical(message|fmt,args...)`
  - 宏：`FASTEXCEL_LOG_TRACE/DEBUG/INFO/WARN/ERROR/CRITICAL`

## class fastexcel::core::Path
- 职责：跨平台路径封装与字符串化（细节见头文件）。

## namespace fastexcel::utils
- `AddressParser`：`parseAddress("Sheet!A1"|"A1") -> (sheet,row,col)`；被 `Workbook/Worksheet` 使用。
- `CommonUtils/TimeUtils/XMLUtils/LogConfig/ModuleLoggers`：通用与模块化日志工具集合。

