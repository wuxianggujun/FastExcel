#include "ReadOnlyWorksheet.hpp"
#include <algorithm>
#include <cctype>

namespace fastexcel {
namespace core {
namespace columnar {

ReadOnlyWorksheet::ReadOnlyWorksheet(const std::string& name,
                                   const SharedStringTable* sst,
                                   const FormatRepository* format_repo)
    : name_(name), sst_(sst), format_repo_(format_repo) {
}

void ReadOnlyWorksheet::setValue(uint32_t row, uint32_t col, double value) {
    getOrCreateColumn(col)->setValue(row, value);
    invalidateUsedRangeCache();
}

void ReadOnlyWorksheet::setValue(uint32_t row, uint32_t col, uint32_t sst_index) {
    getOrCreateColumn(col)->setValue(row, sst_index);
    invalidateUsedRangeCache();
}

void ReadOnlyWorksheet::setValue(uint32_t row, uint32_t col, bool value) {
    getOrCreateColumn(col)->setValue(row, value);
    invalidateUsedRangeCache();
}

void ReadOnlyWorksheet::setValue(uint32_t row, uint32_t col, std::string_view value) {
    getOrCreateColumn(col)->setValue(row, value);
    invalidateUsedRangeCache();
}

ReadOnlyValue ReadOnlyWorksheet::getValue(uint32_t row, uint32_t col) const {
    auto it = columns_.find(col);
    if (it == columns_.end() || !it->second->hasValue(row)) {
        return ReadOnlyValue();  // Empty
    }
    
    const auto* column = it->second.get();
    switch (column->getColumnType()) {
        case ColumnType::Number:
            return ReadOnlyValue(column->getValue<double>(row));
        case ColumnType::SharedStringIndex:
            return ReadOnlyValue(column->getValue<uint32_t>(row));
        case ColumnType::Boolean:
            return ReadOnlyValue(column->getValue<bool>(row));
        case ColumnType::InlineString:
            // 对于内联字符串，我们需要特殊处理
            // 这里简化处理，实际可能需要更复杂的逻辑
            return ReadOnlyValue();  // Empty for now
        default:
            return ReadOnlyValue();  // Empty
    }
}

bool ReadOnlyWorksheet::hasValue(uint32_t row, uint32_t col) const {
    auto it = columns_.find(col);
    return it != columns_.end() && it->second->hasValue(row);
}

double ReadOnlyWorksheet::getNumberValue(uint32_t row, uint32_t col) const {
    auto value = getValue(row, col);
    return value.asNumber();
}

std::string ReadOnlyWorksheet::getStringValue(uint32_t row, uint32_t col) const {
    auto it = columns_.find(col);
    if (it == columns_.end() || !it->second->hasValue(row)) {
        return "";
    }
    
    const auto* column = it->second.get();
    switch (column->getColumnType()) {
        case ColumnType::SharedStringIndex: {
            uint32_t index = column->getValue<uint32_t>(row);
            return sst_ ? sst_->getString(index) : "";
        }
        case ColumnType::InlineString: {
            auto sv = column->getValue<std::string_view>(row);
            return std::string(sv);
        }
        case ColumnType::Number: {
            double num = column->getValue<double>(row);
            return std::to_string(num);
        }
        case ColumnType::Boolean: {
            bool b = column->getValue<bool>(row);
            return b ? "TRUE" : "FALSE";
        }
        default:
            return "";
    }
}

bool ReadOnlyWorksheet::getBooleanValue(uint32_t row, uint32_t col) const {
    auto value = getValue(row, col);
    return value.asBoolean();
}

std::pair<uint32_t, uint32_t> ReadOnlyWorksheet::getUsedRange() const {
    if (used_range_dirty_) {
        updateUsedRangeCache();
    }
    
    if (cached_used_range_) {
        return *cached_used_range_;
    }
    
    return {0, 0};  // 空工作表
}

std::tuple<uint32_t, uint32_t, uint32_t, uint32_t> ReadOnlyWorksheet::getUsedRangeFull() const {
    if (columns_.empty()) {
        return {0, 0, 0, 0};
    }
    
    uint32_t min_row = UINT32_MAX, max_row = 0;
    uint32_t min_col = UINT32_MAX, max_col = 0;
    
    for (const auto& [col_index, column] : columns_) {
        if (!column->isEmpty()) {
            min_col = std::min(min_col, col_index);
            max_col = std::max(max_col, col_index);
            
            uint32_t col_rows = static_cast<uint32_t>(column->getRowCount());
            if (col_rows > 0) {
                min_row = std::min(min_row, 0U);  // 第一行从0开始
                max_row = std::max(max_row, col_rows - 1);
            }
        }
    }
    
    if (min_row == UINT32_MAX) {
        return {0, 0, 0, 0};  // 空工作表
    }
    
    return {min_row, min_col, max_row, max_col};
}

size_t ReadOnlyWorksheet::getCellCount() const {
    size_t total = 0;
    for (const auto& [col_index, column] : columns_) {
        // 这里需要计算实际有数据的单元格数，而不是容量
        // 简化实现，实际应该遍历validity bitmap
        if (column && !column->isEmpty()) {
            total += column->getRowCount();  // 这是一个近似值
        }
    }
    return total;
}

size_t ReadOnlyWorksheet::getMemoryUsage() const {
    size_t total = sizeof(*this);
    
    for (const auto& [col_index, column] : columns_) {
        if (column) {
            total += column->getMemoryUsage();
        }
    }
    
    return total;
}

std::vector<std::pair<uint32_t, uint32_t>> ReadOnlyWorksheet::findCells(
    const std::string& search_text, bool match_case, bool match_entire_cell) const {
    
    std::vector<std::pair<uint32_t, uint32_t>> results;
    
    std::string search_lower;
    if (!match_case) {
        search_lower = search_text;
        std::transform(search_lower.begin(), search_lower.end(), 
                      search_lower.begin(), ::tolower);
    }
    
    for (const auto& [col_index, column] : columns_) {
        if (!column || column->isEmpty()) {
            continue;
        }
        
        uint32_t max_row = static_cast<uint32_t>(column->getRowCount());
        for (uint32_t row = 0; row < max_row; ++row) {
            if (!column->hasValue(row)) {
                continue;
            }
            
            std::string cell_value = getStringValue(row, col_index);
            if (cell_value.empty()) {
                continue;
            }
            
            std::string compare_value = cell_value;
            if (!match_case) {
                std::transform(compare_value.begin(), compare_value.end(),
                              compare_value.begin(), ::tolower);
            }
            
            bool found = false;
            if (match_entire_cell) {
                found = (match_case ? cell_value == search_text 
                                   : compare_value == search_lower);
            } else {
                found = (match_case ? cell_value.find(search_text) != std::string::npos
                                   : compare_value.find(search_lower) != std::string::npos);
            }
            
            if (found) {
                results.emplace_back(row, col_index);
            }
        }
    }
    
    return results;
}

void ReadOnlyWorksheet::clear() {
    columns_.clear();
    cached_used_range_.reset();
    used_range_dirty_ = true;
}

ColumnStorage* ReadOnlyWorksheet::getOrCreateColumn(uint32_t col) {
    auto it = columns_.find(col);
    if (it != columns_.end()) {
        return it->second.get();
    }
    
    auto column = std::make_unique<ColumnStorage>(col);
    ColumnStorage* result = column.get();
    columns_[col] = std::move(column);
    
    return result;
}

void ReadOnlyWorksheet::updateUsedRangeCache() const {
    if (columns_.empty()) {
        cached_used_range_.reset();
        used_range_dirty_ = false;
        return;
    }
    
    uint32_t max_row = 0;
    uint32_t max_col = 0;
    bool has_data = false;
    
    for (const auto& [col_index, column] : columns_) {
        if (column && !column->isEmpty()) {
            has_data = true;
            max_col = std::max(max_col, col_index);
            max_row = std::max(max_row, static_cast<uint32_t>(column->getRowCount()));
        }
    }
    
    if (has_data) {
        cached_used_range_ = {max_row, max_col};
    } else {
        cached_used_range_.reset();
    }
    
    used_range_dirty_ = false;
}

}}} // namespace fastexcel::core::columnar