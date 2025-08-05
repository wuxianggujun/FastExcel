#pragma once

namespace fastexcel {
namespace core {

/**
 * @brief 工作簿模式选择器
 * 
 * 提供三种模式：
 * 1. AUTO - 根据数据量自动选择
 * 2. BATCH - 强制批量模式
 * 3. STREAMING - 强制流式模式
 */
enum class WorkbookMode {
    AUTO = 0,      // 自动选择（默认）
    BATCH = 1,     // 强制批量模式
    STREAMING = 2  // 强制流式模式
};

}} // namespace fastexcel::core