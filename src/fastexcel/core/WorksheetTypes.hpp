#pragma once

namespace fastexcel {
namespace core {

/**
 * @file WorksheetTypes.hpp
 * @brief 工作表相关的类型定义
 * 
 * 包含工作表配置和设置所需的结构体：
 * - 页面视图设置
 */

/**
 * @brief 页面视图设置结构体
 */
struct SheetView {
    bool show_gridlines = true;    // 显示网格线
    bool show_row_col_headers = true; // 显示行列标题
    bool show_zeros = true;        // 显示零值
    bool right_to_left = false;    // 从右到左
    bool tab_selected = false;     // 选项卡选中
    bool show_ruler = true;        // 显示标尺
    bool show_outline_symbols = true; // 显示大纲符号
    bool show_white_space = true;  // 显示空白
    int zoom_scale = 100;          // 缩放比例
    int zoom_scale_normal = 100;   // 正常缩放比例
};

}} // namespace fastexcel::core