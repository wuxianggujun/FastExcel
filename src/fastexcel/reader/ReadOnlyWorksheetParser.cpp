/**
 * @file ReadOnlyWorksheetParser.cpp
 * @brief 只读模式专用工作表XML流式解析器实现
 */

#include "ReadOnlyWorksheetParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <stdexcept>

namespace fastexcel {
namespace reader {

uint32_t ReadOnlyWorksheetParser::parseColumnReference(const std::string& cell_ref) {
    uint32_t col = 0;
    for (char c : cell_ref) {
        if (c >= 'A' && c <= 'Z') {
            col = col * 26 + (c - 'A' + 1);
        } else {
            // 遇到数字就停止（行号部分）
            break;
        }
    }
    return col;
}

bool ReadOnlyWorksheetParser::shouldSkipCell(uint32_t row, uint32_t col) const {
    // 检查行限制
    if (options_->max_rows > 0 && row > options_->max_rows) {
        return true;
    }
    
    // 检查列投影
    if (!options_->projected_columns.empty()) {
        if (std::find(options_->projected_columns.begin(), options_->projected_columns.end(), col) 
            == options_->projected_columns.end()) {
            return true;
        }
    }
    
    return false;
}

void ReadOnlyWorksheetParser::processCellValue() {
    if (current_cell_value_.empty() || !storage_) {
        return;
    }
    
    // 检查是否应该跳过这个单元格
    if (shouldSkipCell(static_cast<uint32_t>(current_row_), current_col_)) {
        return;
    }
    
    try {
        if (current_cell_type_ == "s") {
            // 共享字符串
            int sst_index = std::stoi(current_cell_value_);
            storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, 
                             static_cast<uint32_t>(sst_index));
            cells_processed_++;
        } else if (current_cell_type_ == "b") {
            // 布尔值
            bool bool_value = (current_cell_value_ == "1");
            storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, bool_value);
            cells_processed_++;
        } else if (current_cell_type_ == "str" || current_cell_type_ == "inlineStr") {
            // 内联字符串
            storage_->setError(static_cast<uint32_t>(current_row_), current_col_, current_cell_value_);
            cells_processed_++;
        } else {
            // 数值（默认）或者空类型
            double numeric_value = std::stod(current_cell_value_);
            storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, numeric_value);
            cells_processed_++;
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("解析单元格值失败: {} (type: {}, value: {})", 
                           e.what(), current_cell_type_, current_cell_value_);
        // 解析失败不影响整体进程，跳过这个单元格
    }
}

void ReadOnlyWorksheetParser::onStartElement(std::string_view name, 
                                            span<const xml::XMLAttribute> attributes, int depth) {
    std::string element_name(name);
    
    if (element_name == "sheetData") {
        in_sheet_data_ = true;
        FASTEXCEL_LOG_DEBUG("开始解析工作表数据");
    } else if (element_name == "row" && in_sheet_data_) {
        in_row_ = true;
        
        // 提取行号
        for (const auto& attr : attributes) {
            if (attr.name == "r") {
                try {
                    current_row_ = std::stoi(std::string(attr.value));
                } catch (const std::exception&) {
                    FASTEXCEL_LOG_DEBUG("无效的行号: {}", std::string(attr.value));
                    current_row_ = 0;
                }
                break;
            }
        }
    } else if (element_name == "c" && in_row_) {
        in_cell_ = true;
        
        // 提取单元格引用和类型
        for (const auto& attr : attributes) {
            if (attr.name == "r") {
                current_col_ = parseColumnReference(std::string(attr.value));
            } else if (attr.name == "t") {
                current_cell_type_ = std::string(attr.value);
            }
        }
        
        current_cell_value_.clear();
    } else if (element_name == "v" && in_cell_) {
        in_value_ = true;
        current_cell_value_.clear();
    }
}

void ReadOnlyWorksheetParser::onEndElement(std::string_view name, int depth) {
    std::string element_name(name);
    
    if (element_name == "sheetData") {
        in_sheet_data_ = false;
        FASTEXCEL_LOG_INFO("完成工作表数据解析，处理了 {} 个单元格", cells_processed_);
    } else if (element_name == "row") {
        in_row_ = false;
        current_row_ = 0;
    } else if (element_name == "c") {
        in_cell_ = false;
        // 单元格结束时处理数据
        if (!current_cell_value_.empty()) {
            processCellValue();
        }
        current_col_ = 0;
        current_cell_type_.clear();
        current_cell_value_.clear();
    } else if (element_name == "v") {
        in_value_ = false;
        // 值已经在onText中收集
    }
}

void ReadOnlyWorksheetParser::onText(std::string_view data, int depth) {
    if (in_value_) {
        current_cell_value_ += std::string(data);
    }
}

}} // namespace fastexcel::reader