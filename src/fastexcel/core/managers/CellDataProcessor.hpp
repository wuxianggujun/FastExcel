#pragma once

#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/CellRangeManager.hpp"
#include <map>
#include <memory>
#include <functional>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;
class DirtyManager;

class CellDataProcessor {
public:
    explicit CellDataProcessor(std::map<std::pair<int, int>, Cell>& cells, 
                              CellRangeManager& range_manager,
                              std::shared_ptr<Workbook> parent_workbook = nullptr,
                              int sheet_id = -1);
    
    // 基础单元格操作
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
    
    // 单元格值设置
    template<typename T>
    void setValue(int row, int col, T&& value);
    
    void setFormula(int row, int col, const std::string& formula, double result = 0.0);
    void setHyperlink(int row, int col, const std::string& url, const std::string& display_text = "");
    
    // 单元格操作
    void clearCell(int row, int col);
    bool hasCellAt(int row, int col) const;
    
    // 单元格复制和移动
    void copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format = true);
    void moveCell(int src_row, int src_col, int dst_row, int dst_col);
    
    // 范围操作
    void copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col, bool copy_format = true);
    void moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col);
    void clearRange(int first_row, int first_col, int last_row, int last_col);
    
    // 查找和替换
    std::vector<std::pair<int, int>> findCells(const std::string& search_text,
                                               bool match_case = false,
                                               bool match_entire_cell = false) const;
    int findAndReplace(const std::string& find_text, const std::string& replace_text,
                      bool match_case = false, bool match_entire_cell = false);
    
    // 排序功能
    void sortRange(int first_row, int first_col, int last_row, int last_col,
                   int sort_column, bool ascending = true, bool has_header = false);
    
    // 统计信息
    int getCellCount() const;
    int getCellCountInRow(int row) const;
    int getCellCountInColumn(int col) const;
    
    // 范围信息
    std::pair<int, int> getUsedRange() const;
    std::tuple<int, int, int, int> getUsedRangeFull() const;
    
    // 清理操作
    void clearRow(int row);
    void clearColumn(int col);
    void clearAll();

private:
    std::map<std::pair<int, int>, Cell>& cells_;
    CellRangeManager& range_manager_;
    std::shared_ptr<Workbook> parent_workbook_;
    int sheet_id_;
    
    void validateCellPosition(int row, int col) const;
    void validateRange(int first_row, int first_col, int last_row, int last_col) const;
    void updateUsedRange(int row, int col);
    void markDirty() const;  // 标记工作表为脏数据
};

// 模板实现
template<typename T>
void CellDataProcessor::setValue(int row, int col, T&& value) {
    validateCellPosition(row, col);
    updateUsedRange(row, col);
    markDirty();
    getCell(row, col).setValue(std::forward<T>(value));
}

} // namespace core
} // namespace fastexcel