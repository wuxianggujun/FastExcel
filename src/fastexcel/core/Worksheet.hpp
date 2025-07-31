#pragma once

#include "Cell.h"
#include "Format.h"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <utility>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;

class Worksheet {
private:
    std::string name_;
    std::map<std::pair<int, int>, Cell> cells_; // (row, col) -> Cell
    std::shared_ptr<Workbook> parent_workbook_;
    int sheet_id_;
    
public:
    explicit Worksheet(const std::string& name, std::shared_ptr<Workbook> workbook, int sheet_id = 1);
    ~Worksheet() = default;
    
    // 单元格操作
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
    
    // 写入数据
    void writeString(int row, int col, const std::string& value, std::shared_ptr<Format> format = nullptr);
    void writeNumber(int row, int col, double value, std::shared_ptr<Format> format = nullptr);
    void writeBoolean(int row, int col, bool value, std::shared_ptr<Format> format = nullptr);
    void writeFormula(int row, int col, const std::string& formula, std::shared_ptr<Format> format = nullptr);
    
    // 批量写入
    void writeRange(int start_row, int start_col, const std::vector<std::vector<std::string>>& data);
    void writeRange(int start_row, int start_col, const std::vector<std::vector<double>>& data);
    
    // 获取工作表信息
    std::string getName() const { return name_; }
    int getSheetId() const { return sheet_id_; }
    std::pair<int, int> getUsedRange() const; // 返回 (max_row, max_col)
    
    // 生成XML
    std::string generateXML() const;
    
    // 清空工作表
    void clear();
    
private:
    // 辅助函数
    std::string columnToLetter(int col) const;
    std::string cellReference(int row, int col) const;
    void validateCellPosition(int row, int col) const;
};

}} // namespace fastexcel::core