#pragma once

#include "WorkbookModeSelector.hpp"
#include <string>
#include <cstddef>
#include <vector>
#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file WorkbookTypes.hpp
 * @brief 工作簿相关的类型定义
 * 
 * 包含工作簿配置和状态管理所需的枚举和结构体：
 * - 工作簿状态管理
 * - 文件来源类型
 * - 工作簿配置选项
 */

/**
 * @brief 工作簿状态枚举 - 统一的状态管理
 */
enum class WorkbookState {
    CLOSED,      // 未打开状态
    CREATING,    // 正在创建新文件
    READING,     // 只读模式打开
    EDITING      // 编辑模式打开
};

/**
 * @brief 文件来源类型
 */
enum class FileSource {
    NEW_FILE,        // 全新创建的文件
    EXISTING_FILE    // 从现有文件加载
};

/**
 * @brief 工作簿选项配置结构体
 */
struct WorkbookOptions {
    bool constant_memory = false;      // 常量内存模式
    bool use_zip64 = false;           // 使用ZIP64格式
    std::string tmpdir;               // 临时目录
    bool optimize_for_speed = false;  // 速度优化
    bool read_only_recommended = false; // 建议只读
    
    // 计算选项
    bool calc_on_load = true;         // 加载时计算
    bool full_calc_on_load = false;   // 加载时完全计算
    
    // 安全选项
    std::string password;             // 工作簿密码
    bool encrypt_metadata = false;    // 加密元数据
    
    // 性能优化选项
    bool use_shared_strings = true;   // 使用共享字符串（默认启用以匹配Excel格式）
    WorkbookMode mode = WorkbookMode::AUTO;  // 工作簿模式（默认自动选择）
    size_t row_buffer_size = 5000;    // 行缓冲大小（默认较大缓冲）
    int compression_level = 6;        // ZIP压缩级别（默认平衡压缩）
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小（4MB）
    
    // 列式存储选项（只读模式下的优化配置）
    std::vector<uint32_t> projected_columns;  // 投影列（空表示全部）
    uint32_t max_rows = 0;           // 最大读取行数（0表示全部）
    
    // 自动模式阈值
    size_t auto_mode_cell_threshold = 1000000;     // 100万单元格
    size_t auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
};

}} // namespace fastexcel::core