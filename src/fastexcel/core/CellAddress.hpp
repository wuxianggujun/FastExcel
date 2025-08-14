#pragma once

#include "fastexcel/utils/AddressParser.hpp"
#include <string>
#include <stdexcept>

namespace fastexcel {
namespace core {

/**
 * @brief Excel单元格地址工具类 - 支持隐式转换
 * 
 * 通过隐式转换构造函数，统一处理字符串地址和数字坐标：
 * - Address("A1") - 字符串地址
 * - Address(0, 0) - 行列坐标  
 * - Address{0, 0} - 列表初始化
 * 
 * 设计原则：
 * - KISS: 简单统一的接口
 * - DRY: 避免重复的方法重载
 * - 类型安全: 编译时地址格式验证
 * - 兼容性: 与现有Range结构体协同工作
 */
class Address {
private:
    int row_;
    int col_;
    std::string sheet_name_;

public:
    /**
     * @brief 从行列坐标构造地址（隐式转换）
     * @param row 行号（0基索引）
     * @param col 列号（0基索引）
     */
    Address(int row, int col) 
        : row_(row), col_(col), sheet_name_("") {
        if (row < 0 || col < 0) {
            throw std::invalid_argument("行列索引不能为负数");
        }
    }
    
    /**
     * @brief 从字符串地址构造（隐式转换）
     * @param address Excel地址字符串（如 "A1", "Sheet1!B2"）
     */
    Address(const std::string& address) {
        auto [sheet, row, col] = utils::AddressParser::parseAddress(address);
        sheet_name_ = sheet;
        row_ = row;
        col_ = col;
    }
    
    /**
     * @brief 从C字符串构造（隐式转换）
     * @param address C字符串地址
     */
    Address(const char* address) : Address(std::string(address)) {}
    
    // 获取器方法
    int getRow() const { return row_; }
    int getCol() const { return col_; }
    const std::string& getSheetName() const { return sheet_name_; }
    
    /**
     * @brief 转换为Excel地址字符串
     * @param include_sheet 是否包含工作表名
     * @return Excel地址字符串
     */
    std::string toString(bool include_sheet = false) const {
        if (include_sheet && !sheet_name_.empty()) {
            return utils::AddressParser::indexToAddress(row_, col_, sheet_name_);
        }
        return utils::AddressParser::indexToAddress(row_, col_);
    }
    
    /**
     * @brief 验证地址是否有效
     * @return 是否有效
     */
    bool isValid() const {
        return row_ >= 0 && col_ >= 0;
    }
    
    // 比较操作符
    bool operator==(const Address& other) const {
        return row_ == other.row_ && col_ == other.col_ && sheet_name_ == other.sheet_name_;
    }
    
    bool operator!=(const Address& other) const {
        return !(*this == other);
    }
    
    bool operator<(const Address& other) const {
        if (sheet_name_ != other.sheet_name_) return sheet_name_ < other.sheet_name_;
        if (row_ != other.row_) return row_ < other.row_;
        return col_ < other.col_;
    }
};

/**
 * @brief 简单范围工具类 - 用于隐式转换
 * 
 * 注意：这个类主要用于方法参数的隐式转换。
 * 项目中的实际范围结构体如MergeRange、AutoFilterRange等保持不变。
 * 
 * 这个类作为统一的接口，可以转换为具体的范围结构：
 * - CellRange range("A1:C3") -> 可以转换为MergeRange等
 */
class CellRange {
private:
    int start_row_;
    int start_col_;
    int end_row_;
    int end_col_;
    std::string sheet_name_;

public:
    /**
     * @brief 从坐标构造范围（隐式转换）
     */
    CellRange(int start_row, int start_col, int end_row, int end_col)
        : start_row_(start_row), start_col_(start_col), 
          end_row_(end_row), end_col_(end_col), sheet_name_("") {
        if (start_row < 0 || start_col < 0 || end_row < 0 || end_col < 0) {
            throw std::invalid_argument("行列索引不能为负数");
        }
        
        // 确保起始位置 <= 结束位置
        if (start_row > end_row) std::swap(start_row_, end_row_);
        if (start_col > end_col) std::swap(start_col_, end_col_);
    }
    
    /**
     * @brief 从字符串范围构造（隐式转换）
     */
    CellRange(const std::string& range) {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        sheet_name_ = sheet;
        start_row_ = start_row;
        start_col_ = start_col;
        end_row_ = end_row;
        end_col_ = end_col;
    }
    
    /**
     * @brief 从C字符串构造（隐式转换）
     */
    CellRange(const char* range) : CellRange(std::string(range)) {}
    
    /**
     * @brief 从单个地址构造1x1范围（隐式转换）
     */
    CellRange(const Address& address) 
        : start_row_(address.getRow()), start_col_(address.getCol()),
          end_row_(address.getRow()), end_col_(address.getCol()),
          sheet_name_(address.getSheetName()) {}
    
    // 获取器方法
    int getStartRow() const { return start_row_; }
    int getStartCol() const { return start_col_; }
    int getEndRow() const { return end_row_; }
    int getEndCol() const { return end_col_; }
    const std::string& getSheetName() const { return sheet_name_; }
    
    /**
     * @brief 转换为Excel范围字符串
     */
    std::string toString(bool include_sheet = false) const {
        if (include_sheet && !sheet_name_.empty()) {
            return utils::AddressParser::indexToRange(start_row_, start_col_, end_row_, end_col_, sheet_name_);
        }
        return utils::AddressParser::indexToRange(start_row_, start_col_, end_row_, end_col_);
    }
    
    /**
     * @brief 验证范围是否有效
     */
    bool isValid() const {
        return start_row_ >= 0 && start_col_ >= 0 && 
               end_row_ >= start_row_ && end_col_ >= start_col_;
    }
    
    /**
     * @brief 检查是否为单个单元格
     */
    bool isSingleCell() const {
        return start_row_ == end_row_ && start_col_ == end_col_;
    }
    
    /**
     * @brief 获取范围的行数
     */
    int getRowCount() const { return end_row_ - start_row_ + 1; }
    
    /**
     * @brief 获取范围的列数
     */
    int getColCount() const { return end_col_ - start_col_ + 1; }
    
    /**
     * @brief 检查指定地址是否在范围内
     */
    bool contains(const Address& address) const {
        return address.getRow() >= start_row_ && address.getRow() <= end_row_ &&
               address.getCol() >= start_col_ && address.getCol() <= end_col_;
    }
};

}} // namespace fastexcel::core