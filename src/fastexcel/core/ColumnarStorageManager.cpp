/**
 * @file ColumnarStorageManager.cpp
 * @brief 列式存储管理器实现
 */

#include "fastexcel/core/ColumnarStorageManager.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <map>

namespace fastexcel {
namespace core {

void ColumnarStorageManager::enableColumnarStorage(const WorkbookOptions* options) {
    if (!data_) {
        data_ = std::make_unique<ColumnarData>();
        options_ = options;
        FASTEXCEL_LOG_INFO("启用列式存储管理器");
    }
}

void ColumnarStorageManager::setValue(uint32_t row, uint32_t col, double value) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    data_->number_columns[col][row] = value;
}

void ColumnarStorageManager::setValue(uint32_t row, uint32_t col, uint32_t sst_index) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    data_->string_columns[col][row] = sst_index;
}

void ColumnarStorageManager::setValue(uint32_t row, uint32_t col, bool value) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    data_->boolean_columns[col][row] = value;
}

void ColumnarStorageManager::setValue(uint32_t row, uint32_t col, const std::tm& datetime) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    // 将日期时间转换为Excel序列号存储
    double excel_serial = utils::TimeUtils::toExcelSerialNumber(datetime);
    data_->datetime_columns[col][row] = excel_serial;
}

void ColumnarStorageManager::setFormula(uint32_t row, uint32_t col, uint32_t formula_index, double result) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    data_->formula_columns[col][row] = {formula_index, result};
}

void ColumnarStorageManager::setError(uint32_t row, uint32_t col, const std::string& error_code) {
    if (!data_) {
        enableColumnarStorage();
    }
    
    if (shouldSkipColumn(col)) {
        return;  // 跳过过滤的列
    }
    
    data_->error_columns[col][row] = error_code;
}

bool ColumnarStorageManager::hasValue(uint32_t row, uint32_t col) const {
    if (!data_) {
        return false;
    }
    
    // 检查所有列类型是否包含该位置的数据
    auto num_col_it = data_->number_columns.find(col);
    if (num_col_it != data_->number_columns.end()) {
        if (num_col_it->second.find(row) != num_col_it->second.end()) {
            return true;
        }
    }
    
    auto str_col_it = data_->string_columns.find(col);
    if (str_col_it != data_->string_columns.end()) {
        if (str_col_it->second.find(row) != str_col_it->second.end()) {
            return true;
        }
    }
    
    auto bool_col_it = data_->boolean_columns.find(col);
    if (bool_col_it != data_->boolean_columns.end()) {
        if (bool_col_it->second.find(row) != bool_col_it->second.end()) {
            return true;
        }
    }
    
    auto dt_col_it = data_->datetime_columns.find(col);
    if (dt_col_it != data_->datetime_columns.end()) {
        if (dt_col_it->second.find(row) != dt_col_it->second.end()) {
            return true;
        }
    }
    
    auto formula_col_it = data_->formula_columns.find(col);
    if (formula_col_it != data_->formula_columns.end()) {
        if (formula_col_it->second.find(row) != formula_col_it->second.end()) {
            return true;
        }
    }
    
    auto err_col_it = data_->error_columns.find(col);
    if (err_col_it != data_->error_columns.end()) {
        if (err_col_it->second.find(row) != err_col_it->second.end()) {
            return true;
        }
    }
    
    return false;
}

ColumnarStorageManager::ColumnarValueVariant ColumnarStorageManager::getValue(uint32_t row, uint32_t col) const {
    if (!data_) {
        return std::monostate{};
    }
    
    // 按优先级检查不同类型的列
    auto num_col_it = data_->number_columns.find(col);
    if (num_col_it != data_->number_columns.end()) {
        auto row_it = num_col_it->second.find(row);
        if (row_it != num_col_it->second.end()) {
            return row_it->second;
        }
    }
    
    auto str_col_it = data_->string_columns.find(col);
    if (str_col_it != data_->string_columns.end()) {
        auto row_it = str_col_it->second.find(row);
        if (row_it != str_col_it->second.end()) {
            return row_it->second;  // 返回SST索引
        }
    }
    
    auto bool_col_it = data_->boolean_columns.find(col);
    if (bool_col_it != data_->boolean_columns.end()) {
        auto row_it = bool_col_it->second.find(row);
        if (row_it != bool_col_it->second.end()) {
            return row_it->second;
        }
    }
    
    auto dt_col_it = data_->datetime_columns.find(col);
    if (dt_col_it != data_->datetime_columns.end()) {
        auto row_it = dt_col_it->second.find(row);
        if (row_it != dt_col_it->second.end()) {
            return row_it->second;  // Excel序列号
        }
    }
    
    auto formula_col_it = data_->formula_columns.find(col);
    if (formula_col_it != data_->formula_columns.end()) {
        auto row_it = formula_col_it->second.find(row);
        if (row_it != formula_col_it->second.end()) {
            return row_it->second;  // FormulaValue结构
        }
    }
    
    auto err_col_it = data_->error_columns.find(col);
    if (err_col_it != data_->error_columns.end()) {
        auto row_it = err_col_it->second.find(row);
        if (row_it != err_col_it->second.end()) {
            return row_it->second;
        }
    }
    
    return std::monostate{};  // 无数据
}

void ColumnarStorageManager::forEachInColumn(uint32_t col, std::function<void(uint32_t row, const ColumnarValueVariant& value)> callback) const {
    if (!data_ || !callback) {
        return;
    }
    
    // 遍历该列的所有行数据，按行号排序
    std::map<uint32_t, ColumnarValueVariant> sorted_rows;
    
    // 收集数字列数据
    auto num_col_it = data_->number_columns.find(col);
    if (num_col_it != data_->number_columns.end()) {
        for (const auto& [row, value] : num_col_it->second) {
            sorted_rows[row] = value;
        }
    }
    
    // 收集字符串列数据
    auto str_col_it = data_->string_columns.find(col);
    if (str_col_it != data_->string_columns.end()) {
        for (const auto& [row, sst_index] : str_col_it->second) {
            sorted_rows[row] = sst_index;
        }
    }
    
    // 收集布尔列数据
    auto bool_col_it = data_->boolean_columns.find(col);
    if (bool_col_it != data_->boolean_columns.end()) {
        for (const auto& [row, value] : bool_col_it->second) {
            sorted_rows[row] = value;
        }
    }
    
    // 收集日期时间列数据
    auto dt_col_it = data_->datetime_columns.find(col);
    if (dt_col_it != data_->datetime_columns.end()) {
        for (const auto& [row, value] : dt_col_it->second) {
            sorted_rows[row] = value;
        }
    }
    
    // 收集公式列数据
    auto formula_col_it = data_->formula_columns.find(col);
    if (formula_col_it != data_->formula_columns.end()) {
        for (const auto& [row, formula_val] : formula_col_it->second) {
            sorted_rows[row] = formula_val;
        }
    }
    
    // 收集错误列数据
    auto err_col_it = data_->error_columns.find(col);
    if (err_col_it != data_->error_columns.end()) {
        for (const auto& [row, error] : err_col_it->second) {
            sorted_rows[row] = error;
        }
    }
    
    // 按行号顺序调用回调
    for (const auto& [row, value] : sorted_rows) {
        callback(row, value);
    }
}

std::unordered_map<uint32_t, double> ColumnarStorageManager::getNumberColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->number_columns.find(col);
    if (it != data_->number_columns.end()) {
        return it->second;
    }
    return {};
}

std::unordered_map<uint32_t, uint32_t> ColumnarStorageManager::getStringColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->string_columns.find(col);
    if (it != data_->string_columns.end()) {
        return it->second;
    }
    return {};
}

std::unordered_map<uint32_t, bool> ColumnarStorageManager::getBooleanColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->boolean_columns.find(col);
    if (it != data_->boolean_columns.end()) {
        return it->second;
    }
    return {};
}

std::unordered_map<uint32_t, double> ColumnarStorageManager::getDateTimeColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->datetime_columns.find(col);
    if (it != data_->datetime_columns.end()) {
        return it->second;
    }
    return {};
}

std::unordered_map<uint32_t, ColumnarStorageManager::FormulaValue> ColumnarStorageManager::getFormulaColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->formula_columns.find(col);
    if (it != data_->formula_columns.end()) {
        return it->second;
    }
    return {};
}

std::unordered_map<uint32_t, std::string> ColumnarStorageManager::getErrorColumn(uint32_t col) const {
    if (!data_) {
        return {};
    }
    
    auto it = data_->error_columns.find(col);
    if (it != data_->error_columns.end()) {
        return it->second;
    }
    return {};
}

size_t ColumnarStorageManager::getDataCount() const {
    if (!data_) {
        return 0;
    }
    
    size_t total = 0;
    
    // 统计所有列的数据数量
    for (const auto& [col, row_data] : data_->number_columns) {
        total += row_data.size();
    }
    for (const auto& [col, row_data] : data_->string_columns) {
        total += row_data.size();
    }
    for (const auto& [col, row_data] : data_->boolean_columns) {
        total += row_data.size();
    }
    for (const auto& [col, row_data] : data_->datetime_columns) {
        total += row_data.size();
    }
    for (const auto& [col, row_data] : data_->formula_columns) {
        total += row_data.size();
    }
    for (const auto& [col, row_data] : data_->error_columns) {
        total += row_data.size();
    }
    
    return total;
}

size_t ColumnarStorageManager::getMemoryUsage() const {
    if (!data_) {
        return 0;
    }
    
    size_t memory = sizeof(ColumnarData);
    
    // 估算各类型列的内存使用
    memory += data_->number_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, double>>);
    for (const auto& [col, row_data] : data_->number_columns) {
        memory += row_data.size() * sizeof(std::pair<uint32_t, double>);
    }
    
    memory += data_->string_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, uint32_t>>);
    for (const auto& [col, row_data] : data_->string_columns) {
        memory += row_data.size() * sizeof(std::pair<uint32_t, uint32_t>);
    }
    
    memory += data_->boolean_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, bool>>);
    for (const auto& [col, row_data] : data_->boolean_columns) {
        memory += row_data.size() * sizeof(std::pair<uint32_t, bool>);
    }
    
    memory += data_->datetime_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, double>>);
    for (const auto& [col, row_data] : data_->datetime_columns) {
        memory += row_data.size() * sizeof(std::pair<uint32_t, double>);
    }
    
    memory += data_->formula_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, FormulaValue>>);
    for (const auto& [col, row_data] : data_->formula_columns) {
        memory += row_data.size() * sizeof(std::pair<uint32_t, FormulaValue>);
    }
    
    memory += data_->error_columns.size() * sizeof(std::pair<uint32_t, std::unordered_map<uint32_t, std::string>>);
    for (const auto& [col, row_data] : data_->error_columns) {
        for (const auto& [row, error_str] : row_data) {
            memory += sizeof(std::pair<uint32_t, std::string>) + error_str.capacity();
        }
    }
    
    return memory;
}

void ColumnarStorageManager::clearData() {
    if (data_) {
        data_.reset();
        options_ = nullptr;
        FASTEXCEL_LOG_INFO("清除列式存储管理器数据");
    }
}

bool ColumnarStorageManager::shouldSkipColumn(uint32_t col) const {
    if (!options_ || !options_->enable_columnar_storage) {
        return false;
    }
    
    // 检查列投影过滤
    if (!options_->projected_columns.empty()) {
        return std::find(options_->projected_columns.begin(), 
                        options_->projected_columns.end(), 
                        col) == options_->projected_columns.end();
    }
    
    return false;
}

} // namespace core
} // namespace fastexcel