#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file WorksheetTypes.hpp
 * @brief 工作表相关的类型定义
 * 
 * 包含工作表配置和设置所需的结构体：
 * - 页面视图设置
 * - 列信息和行信息
 * - 合并单元格、自动筛选、冻结窗格等
 */

// 列信息结构
struct ColumnInfo {
    double width = -1.0;
    int format_id = -1;
    bool hidden = false;
    bool collapsed = false;
    uint8_t outline_level = 0;
    bool precise_width = false;
    
    // 相等比较操作符
    bool operator==(const ColumnInfo& other) const {
        return width == other.width &&
               format_id == other.format_id &&
               hidden == other.hidden &&
               collapsed == other.collapsed &&
               outline_level == other.outline_level &&
               precise_width == other.precise_width;
    }
    
    // 小于比较操作符（用于排序）
    bool operator<(const ColumnInfo& other) const {
        if (width != other.width) return width < other.width;
        if (format_id != other.format_id) return format_id < other.format_id;
        if (hidden != other.hidden) return hidden < other.hidden;
        if (collapsed != other.collapsed) return collapsed < other.collapsed;
        if (outline_level != other.outline_level) return outline_level < other.outline_level;
        return precise_width < other.precise_width;
    }
};

// 行信息结构
struct RowInfo {
    double height = -1.0;
    int format_id = -1;
    bool hidden = false;
    bool collapsed = false;
    uint8_t outline_level = 0;
};

// 合并单元格范围
struct MergeRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
    
    MergeRange(int fr, int fc, int lr, int lc) 
        : first_row(fr), first_col(fc), last_row(lr), last_col(lc) {}
};

// 自动筛选范围
struct AutoFilterRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
    
    AutoFilterRange(int fr = 0, int fc = 0, int lr = 0, int lc = 0)
        : first_row(fr), first_col(fc), last_row(lr), last_col(lc) {}
};

// 冻结窗格信息
struct FreezePanes {
    int row = 0;
    int col = 0;
    int top_left_row = 0;
    int top_left_col = 0;
    
    FreezePanes(int r = 0, int c = 0, int tlr = 0, int tlc = 0)
        : row(r), col(c), top_left_row(tlr), top_left_col(tlc) {}
};

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