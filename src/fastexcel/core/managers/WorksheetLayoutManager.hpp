#pragma once

#include "fastexcel/core/ColumnWidthManager.hpp"
#include <map>
#include <vector>
#include <memory>

namespace fastexcel {
namespace core {

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

class WorksheetLayoutManager {
public:
    explicit WorksheetLayoutManager();
    
    // 列宽管理
    double setColumnWidth(int col, double width);
    std::pair<double, int> setColumnWidthAdvanced(int col, double target_width,
                                                  const std::string& font_name = "Calibri",
                                                  double font_size = 11.0,
                                                  ColumnWidthManager::WidthStrategy strategy = ColumnWidthManager::WidthStrategy::EXACT,
                                                  const std::vector<std::string>& cell_contents = {});
    std::unordered_map<int, std::pair<double, int>> setColumnWidthsBatch(
        const std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig>& configs);
    double calculateOptimalWidth(double target_width, const std::string& font_name = "Calibri", double font_size = 11.0) const;
    
    double getColumnWidth(int col) const;
    void setColumnFormatId(int col, int format_id);
    void setColumnFormatId(int first_col, int last_col, int format_id);
    int getColumnFormatId(int col) const;
    
    // 列操作
    void hideColumn(int col);
    void hideColumn(int first_col, int last_col);
    bool isColumnHidden(int col) const;
    
    // 行高管理
    void setRowHeight(int row, double height);
    double getRowHeight(int row) const;
    
    // 行操作
    void hideRow(int row);
    void hideRow(int first_row, int last_row);
    bool isRowHidden(int row) const;
    
    // 合并单元格
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    const std::vector<MergeRange>& getMergeRanges() const { return merge_ranges_; }
    
    // 自动筛选
    void setAutoFilter(int first_row, int first_col, int last_row, int last_col);
    void removeAutoFilter();
    AutoFilterRange getAutoFilterRange() const;
    bool hasAutoFilter() const { return autofilter_ != nullptr; }
    
    // 冻结窗格
    void freezePanes(int row, int col);
    void freezePanes(int row, int col, int top_left_row, int top_left_col);
    void splitPanes(int row, int col);
    FreezePanes getFreezeInfo() const;
    bool hasFreezePane() const { return freeze_panes_ != nullptr; }
    
    // 清理
    void clear();
    
    // 设置默认值
    void setDefaultColumnWidth(double width) { default_col_width_ = width; }
    void setDefaultRowHeight(double height) { default_row_height_ = height; }
    double getDefaultColumnWidth() const { return default_col_width_; }
    double getDefaultRowHeight() const { return default_row_height_; }
    
    // 设置FormatRepository（用于高级列宽管理）
    void setFormatRepository(std::shared_ptr<FormatRepository> format_repo);

private:
    std::unordered_map<int, ColumnInfo> column_info_;
    std::unordered_map<int, RowInfo> row_info_;
    std::vector<MergeRange> merge_ranges_;
    std::unique_ptr<AutoFilterRange> autofilter_;
    std::unique_ptr<FreezePanes> freeze_panes_;
    
    double default_col_width_ = 8.43;  // Excel默认列宽
    double default_row_height_ = 15.0; // Excel默认行高
    
    std::unique_ptr<ColumnWidthManager> column_width_manager_;
    std::shared_ptr<FormatRepository> format_repo_;
    
    void validateCellPosition(int row, int col) const;
    void validateRange(int first_row, int first_col, int last_row, int last_col) const;
};

} // namespace core
} // namespace fastexcel