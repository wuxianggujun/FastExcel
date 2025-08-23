#pragma once

#include <string>
#include <fmt/format.h>
#include <utility>
#include <limits>

namespace fastexcel {
namespace core {

/**
 * @brief 单元格范围管理器
 * 
 * 封装单元格范围的计算和管理逻辑，简化Worksheet中的范围处理
 */
class CellRangeManager {
private:
    int min_row_ = std::numeric_limits<int>::max();
    int max_row_ = std::numeric_limits<int>::min();
    int min_col_ = std::numeric_limits<int>::max();
    int max_col_ = std::numeric_limits<int>::min();
    bool has_data_ = false;

    /**
     * @brief 将列索引转换为Excel列名（A, B, C, ..., Z, AA, AB, ...）
     * @param col 列索引（0基）
     * @return Excel列名
     */
    std::string columnToExcelName(int col) const {
        std::string result;
        col++; // 转换为1基
        
        while (col > 0) {
            col--; // 调整为0基进行计算
            result = char('A' + (col % 26)) + result;
            col /= 26;
        }
        
        return result;
    }
    
    /**
     * @brief 将行列索引转换为Excel单元格引用
     * @param row 行索引（0基）
     * @param col 列索引（0基）
     * @return Excel单元格引用（如"A1"）
     */
    std::string cellReference(int row, int col) const {
        return fmt::format("{}{}", columnToExcelName(col), row + 1);
    }

public:
    CellRangeManager() = default;
    ~CellRangeManager() = default;
    
    // 允许拷贝和移动
    CellRangeManager(const CellRangeManager&) = default;
    CellRangeManager& operator=(const CellRangeManager&) = default;
    CellRangeManager(CellRangeManager&&) = default;
    CellRangeManager& operator=(CellRangeManager&&) = default;
    
    /**
     * @brief 更新使用范围
     * @param row 行索引（0基）
     * @param col 列索引（0基）
     */
    void updateRange(int row, int col) {
        if (row < 0 || col < 0) {
            return; // 忽略无效坐标
        }
        
        if (!has_data_) {
            // 第一个有效单元格
            min_row_ = max_row_ = row;
            min_col_ = max_col_ = col;
            has_data_ = true;
        } else {
            // 扩展范围
            min_row_ = std::min(min_row_, row);
            max_row_ = std::max(max_row_, row);
            min_col_ = std::min(min_col_, col);
            max_col_ = std::max(max_col_, col);
        }
    }
    
    /**
     * @brief 更新使用范围（批量）
     * @param min_row 最小行
     * @param max_row 最大行
     * @param min_col 最小列
     * @param max_col 最大列
     */
    void updateRange(int min_row, int max_row, int min_col, int max_col) {
        if (min_row < 0 || max_row < 0 || min_col < 0 || max_col < 0 ||
            min_row > max_row || min_col > max_col) {
            return; // 忽略无效范围
        }
        
        if (!has_data_) {
            min_row_ = min_row;
            max_row_ = max_row;
            min_col_ = min_col;
            max_col_ = max_col;
            has_data_ = true;
        } else {
            min_row_ = std::min(min_row_, min_row);
            max_row_ = std::max(max_row_, max_row);
            min_col_ = std::min(min_col_, min_col);
            max_col_ = std::max(max_col_, max_col);
        }
    }
    
    /**
     * @brief 重置范围
     */
    void resetRange() {
        min_row_ = std::numeric_limits<int>::max();
        max_row_ = std::numeric_limits<int>::min();
        min_col_ = std::numeric_limits<int>::max();
        max_col_ = std::numeric_limits<int>::min();
        has_data_ = false;
    }
    
    /**
     * @brief 获取使用的行范围
     * @return {最小行, 最大行}，如果没有数据则返回{-1, -1}
     */
    std::pair<int, int> getUsedRowRange() const {
        if (!has_data_) {
            return {-1, -1};
        }
        return {min_row_, max_row_};
    }
    
    /**
     * @brief 获取使用的列范围
     * @return {最小列, 最大列}，如果没有数据则返回{-1, -1}
     */
    std::pair<int, int> getUsedColRange() const {
        if (!has_data_) {
            return {-1, -1};
        }
        return {min_col_, max_col_};
    }
    
    /**
     * @brief 获取完整的使用范围
     * @return {最小行, 最大行, 最小列, 最大列}，如果没有数据则全部返回-1
     */
    std::tuple<int, int, int, int> getUsedRange() const {
        if (!has_data_) {
            return {-1, -1, -1, -1};
        }
        return {min_row_, max_row_, min_col_, max_col_};
    }
    
    /**
     * @brief 获取范围引用字符串（Excel格式）
     * @return 范围引用字符串，如"A1:C10"，如果没有数据则返回"A1"
     */
    std::string getRangeReference() const {
        if (!has_data_) {
            return "A1";
        }
        
        if (min_row_ == max_row_ && min_col_ == max_col_) {
            // 单个单元格
            return cellReference(min_row_, min_col_);
        } else {
            // 范围
            return fmt::format("{}:{}", cellReference(min_row_, min_col_), cellReference(max_row_, max_col_));
        }
    }
    
    /**
     * @brief 获取左上角单元格引用
     * @return 左上角单元格引用，如"A1"
     */
    std::string getTopLeftReference() const {
        if (!has_data_) {
            return "A1";
        }
        return cellReference(min_row_, min_col_);
    }
    
    /**
     * @brief 获取右下角单元格引用
     * @return 右下角单元格引用，如"C10"
     */
    std::string getBottomRightReference() const {
        if (!has_data_) {
            return "A1";
        }
        return cellReference(max_row_, max_col_);
    }
    
    /**
     * @brief 检查范围是否为空
     * @return 是否没有任何使用的单元格
     */
    bool isEmpty() const {
        return !has_data_;
    }
    
    /**
     * @brief 检查指定位置是否在使用范围内
     * @param row 行索引
     * @param col 列索引
     * @return 是否在范围内
     */
    bool contains(int row, int col) const {
        if (!has_data_) {
            return false;
        }
        return row >= min_row_ && row <= max_row_ && col >= min_col_ && col <= max_col_;
    }
    
    /**
     * @brief 获取范围内的行数
     * @return 行数，如果没有数据则返回0
     */
    int getRowCount() const {
        if (!has_data_) {
            return 0;
        }
        return max_row_ - min_row_ + 1;
    }
    
    /**
     * @brief 获取范围内的列数
     * @return 列数，如果没有数据则返回0
     */
    int getColCount() const {
        if (!has_data_) {
            return 0;
        }
        return max_col_ - min_col_ + 1;
    }
    
    /**
     * @brief 获取范围内的总单元格数
     * @return 单元格数量
     */
    size_t getTotalCellCount() const {
        if (!has_data_) {
            return 0;
        }
        return static_cast<size_t>(getRowCount()) * static_cast<size_t>(getColCount());
    }
    
    /**
     * @brief 扩展范围到指定位置
     * @param row 目标行
     * @param col 目标列
     */
    void expandTo(int row, int col) {
        updateRange(row, col);
    }
    
    /**
     * @brief 收缩范围（移除指定位置）
     * 注意：这是一个简化版本，只有当指定位置是边界时才会收缩
     * @param row 要移除的行
     * @param col 要移除的列
     * @return 是否成功收缩
     */
    bool shrinkFrom(int row, int col) {
        if (!has_data_ || !contains(row, col)) {
            return false;
        }
        
        // 只在是边界位置时才收缩
        bool shrunk = false;
        
        if (row == min_row_ && min_row_ < max_row_) {
            min_row_++;
            shrunk = true;
        } else if (row == max_row_ && min_row_ < max_row_) {
            max_row_--;
            shrunk = true;
        }
        
        if (col == min_col_ && min_col_ < max_col_) {
            min_col_++;
            shrunk = true;
        } else if (col == max_col_ && min_col_ < max_col_) {
            max_col_--;
            shrunk = true;
        }
        
        // 检查是否收缩到无效状态
        if (min_row_ > max_row_ || min_col_ > max_col_) {
            resetRange();
            shrunk = true;
        }
        
        return shrunk;
    }
    
    /**
     * @brief 克隆当前范围管理器
     * @return 新的范围管理器实例
     */
    CellRangeManager clone() const {
        CellRangeManager copy;
        copy.min_row_ = min_row_;
        copy.max_row_ = max_row_;
        copy.min_col_ = min_col_;
        copy.max_col_ = max_col_;
        copy.has_data_ = has_data_;
        return copy;
    }
    
    /**
     * @brief 比较两个范围是否相等
     * @param other 另一个范围管理器
     * @return 是否相等
     */
    bool operator==(const CellRangeManager& other) const {
        return min_row_ == other.min_row_ && max_row_ == other.max_row_ &&
               min_col_ == other.min_col_ && max_col_ == other.max_col_ &&
               has_data_ == other.has_data_;
    }
    
    bool operator!=(const CellRangeManager& other) const {
        return !(*this == other);
    }
};

}} // namespace fastexcel::core
