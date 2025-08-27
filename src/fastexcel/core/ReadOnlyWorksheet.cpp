/**
 * @file ReadOnlyWorksheet.cpp
 * @brief 只读工作表实现 - 使用现有的ColumnarStorageManager
 */

#include "fastexcel/core/ReadOnlyWorksheet.hpp"
#include "fastexcel/core/ColumnarStorageManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <set>

namespace fastexcel {
namespace core {

// 私有构造函数
ReadOnlyWorksheet::ReadOnlyWorksheet(const std::string& name, 
                                   std::shared_ptr<ColumnarStorageManager> storage_manager,
                                   int first_row, int first_col, int last_row, int last_col)
    : name_(name)
    , storage_manager_(std::move(storage_manager))
    , first_row_(first_row)
    , first_col_(first_col)
    , last_row_(last_row)
    , last_col_(last_col) {
}

// 基本信息接口
std::string ReadOnlyWorksheet::getName() const {
    return name_;
}

std::pair<int, int> ReadOnlyWorksheet::getUsedRange() const {
    return {last_row_ + 1, last_col_ + 1};  // 返回行列数（而不是最大索引）
}

std::tuple<int, int, int, int> ReadOnlyWorksheet::getUsedRangeFull() const {
    return {first_row_, first_col_, last_row_, last_col_};
}

// 列式数据访问接口
std::unordered_map<uint32_t, double> ReadOnlyWorksheet::getNumberColumn(uint32_t col) const {
    if (!storage_manager_) {
        return {};
    }
    
    return storage_manager_->getNumberColumn(col);
}

std::unordered_map<uint32_t, uint32_t> ReadOnlyWorksheet::getStringColumn(uint32_t col) const {
    if (!storage_manager_) {
        return {};
    }
    
    return storage_manager_->getStringColumn(col);
}

std::unordered_map<uint32_t, bool> ReadOnlyWorksheet::getBooleanColumn(uint32_t col) const {
    if (!storage_manager_) {
        return {};
    }
    
    return storage_manager_->getBooleanColumn(col);
}

std::unordered_map<uint32_t, std::string> ReadOnlyWorksheet::getErrorColumn(uint32_t col) const {
    if (!storage_manager_) {
        return {};
    }
    
    return storage_manager_->getErrorColumn(col);
}

// 列遍历接口
void ReadOnlyWorksheet::forEachInColumn(uint32_t col, const ColumnCallback& callback) const {
    if (!storage_manager_ || !callback) {
        return;
    }
    
    // 使用ColumnarStorageManager的forEachInColumn，但需要适配CellValue类型
    storage_manager_->forEachInColumn(col, [&](uint32_t row, const ColumnarStorageManager::ColumnarValueVariant& value) {
        // 将ColumnarValueVariant转换为CellValue
        std::visit([&](const auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, std::monostate>) {
                // 跳过空值
                return;
            } else if constexpr (std::is_same_v<T, double>) {
                callback(row, CellValue(v));
            } else if constexpr (std::is_same_v<T, uint32_t>) {
                callback(row, CellValue(v));
            } else if constexpr (std::is_same_v<T, bool>) {
                callback(row, CellValue(v));
            } else if constexpr (std::is_same_v<T, std::string>) {
                callback(row, CellValue(v));
            } else if constexpr (std::is_same_v<T, ColumnarStorageManager::FormulaValue>) {
                // 对于公式，返回计算结果
                callback(row, CellValue(v.result));
            }
        }, value);
    });
}

void ReadOnlyWorksheet::forEachInColumnRange(uint32_t col, uint32_t start_row, uint32_t end_row,
                                           const ColumnCallback& callback) const {
    if (!storage_manager_ || !callback) {
        return;
    }
    
    // 使用forEachInColumn然后过滤行范围
    forEachInColumn(col, [&](uint32_t row, const CellValue& value) {
        if (row >= start_row && row <= end_row) {
            callback(row, value);
        }
    });
}

// 批量数据接口
std::vector<ReadOnlyWorksheet::ColumnData> ReadOnlyWorksheet::getBatchColumns(const std::vector<uint32_t>& columns) const {
    std::vector<ColumnData> result;
    result.reserve(columns.size());
    
    for (uint32_t col : columns) {
        ColumnData column_data;
        
        // 收集数字数据
        auto numbers = getNumberColumn(col);
        for (const auto& [row, value] : numbers) {
            column_data[row] = CellValue(value);
        }
        
        // 收集字符串数据
        auto strings = getStringColumn(col);
        for (const auto& [row, sst_index] : strings) {
            column_data[row] = CellValue(sst_index);
        }
        
        // 收集布尔数据
        auto booleans = getBooleanColumn(col);
        for (const auto& [row, value] : booleans) {
            column_data[row] = CellValue(value);
        }
        
        // 收集错误/内联字符串数据
        auto errors = getErrorColumn(col);
        for (const auto& [row, value] : errors) {
            column_data[row] = CellValue(value);
        }
        
        // 收集日期时间数据（作为数字）
        auto datetimes = storage_manager_->getDateTimeColumn(col);
        for (const auto& [row, value] : datetimes) {
            column_data[row] = CellValue(value);
        }
        
        // 收集公式数据（返回计算结果）
        auto formulas = storage_manager_->getFormulaColumn(col);
        for (const auto& [row, formula_val] : formulas) {
            column_data[row] = CellValue(formula_val.result);
        }
        
        result.push_back(std::move(column_data));
    }
    
    return result;
}

std::unordered_map<uint32_t, std::unordered_map<uint32_t, ReadOnlyWorksheet::CellValue>> 
ReadOnlyWorksheet::getRowRangeData(uint32_t start_row, uint32_t end_row) const {
    std::unordered_map<uint32_t, std::unordered_map<uint32_t, CellValue>> result;
    
    if (!storage_manager_) {
        return result;
    }
    
    // 获取所有有数据的列
    auto data_columns = getDataColumns();
    
    // 遍历每一列收集指定行范围的数据
    for (uint32_t col : data_columns) {
        forEachInColumnRange(col, start_row, end_row, [&](uint32_t row, const CellValue& value) {
            result[row][col] = value;
        });
    }
    
    return result;
}

// 统计信息接口
size_t ReadOnlyWorksheet::getColumnarDataCount() const {
    if (!storage_manager_) {
        return 0;
    }
    
    return storage_manager_->getDataCount();
}

size_t ReadOnlyWorksheet::getColumnarMemoryUsage() const {
    if (!storage_manager_) {
        return 0;
    }
    
    return storage_manager_->getMemoryUsage();
}

bool ReadOnlyWorksheet::isColumnarMode() const {
    return storage_manager_ != nullptr && storage_manager_->isColumnarEnabled();
}

ReadOnlyWorksheet::Stats ReadOnlyWorksheet::getStats() const {
    Stats stats{};
    
    stats.total_data_points = getColumnarDataCount();
    stats.memory_usage = getColumnarMemoryUsage();
    stats.used_range = getUsedRange();
    
    if (storage_manager_) {
        // 统计各类型列数（通过检查是否有数据）
        stats.number_columns = 0;
        stats.string_columns = 0;
        stats.boolean_columns = 0;
        stats.error_columns = 0;
        
        // 获取所有有数据的列，然后检查每列包含的数据类型
        auto data_columns = getDataColumns();
        for (uint32_t col : data_columns) {
            if (!storage_manager_->getNumberColumn(col).empty() || 
                !storage_manager_->getDateTimeColumn(col).empty() ||
                !storage_manager_->getFormulaColumn(col).empty()) {
                stats.number_columns++;
            }
            if (!storage_manager_->getStringColumn(col).empty()) {
                stats.string_columns++;
            }
            if (!storage_manager_->getBooleanColumn(col).empty()) {
                stats.boolean_columns++;
            }
            if (!storage_manager_->getErrorColumn(col).empty()) {
                stats.error_columns++;
            }
        }
    }
    
    return stats;
}

// 数据查询接口
bool ReadOnlyWorksheet::hasDataAt(uint32_t row, uint32_t col) const {
    if (!storage_manager_) {
        return false;
    }
    
    return storage_manager_->hasValue(row, col);
}

size_t ReadOnlyWorksheet::getColumnDataCount(uint32_t col) const {
    if (!storage_manager_) {
        return 0;
    }
    
    size_t count = 0;
    
    // 统计该列所有类型的数据数量
    count += storage_manager_->getNumberColumn(col).size();
    count += storage_manager_->getStringColumn(col).size();
    count += storage_manager_->getBooleanColumn(col).size();
    count += storage_manager_->getDateTimeColumn(col).size();
    count += storage_manager_->getFormulaColumn(col).size();
    count += storage_manager_->getErrorColumn(col).size();
    
    return count;
}

std::vector<uint32_t> ReadOnlyWorksheet::getDataColumns() const {
    std::vector<uint32_t> columns;
    
    if (!storage_manager_) {
        return columns;
    }
    
    // 收集所有有数据的列索引
    std::set<uint32_t> column_set;
    
    // 简单实现：遍历一定范围的列来找到有数据的列
    // 这里可以优化为从storage_manager直接获取所有列索引
    for (uint32_t col = 0; col <= static_cast<uint32_t>(last_col_); ++col) {
        if (getColumnDataCount(col) > 0) {
            column_set.insert(col);
        }
    }
    
    // 转换为vector
    columns.assign(column_set.begin(), column_set.end());
    return columns;
}

ReadOnlyWorksheet::~ReadOnlyWorksheet() = default;

}} // namespace fastexcel::core