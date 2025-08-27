/**
 * @file ReadOnlyWorksheetParser.cpp
 * @brief 只读模式专用工作表XML流式解析器实现
 */

#include "ReadOnlyWorksheetParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/ColumnReferenceUtils.hpp"
#include <fast_float/fast_float.h>
#include <algorithm>
#include <stdexcept>
#include <cstring>

namespace fastexcel {
namespace reader {

uint32_t ReadOnlyWorksheetParser::parseColumnReference(std::string_view cell_ref) {
    // 使用高性能预计算查找表
    return utils::ColumnReferenceUtils::parseColumnFast(cell_ref);
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
    
    // 直接处理单元格 - 针对性SIMD优化
    try {
        auto cell_value_view = current_cell_value_.view();
        
        if (current_cell_type_ == "s") {
            // 共享字符串 - 使用 fast_float 解析整数索引
            uint32_t sst_index;
            auto result = fast_float::from_chars(cell_value_view.data(),
                                               cell_value_view.data() + cell_value_view.size(),
                                               sst_index);
            if (result.ec == std::errc{}) {
                storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, sst_index);
                cells_processed_++;
            }
        } else if (current_cell_type_ == "b") {
            // 布尔值 - 简单高效的字符比较
            bool bool_value = (cell_value_view == "1");
            storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, bool_value);
            cells_processed_++;
        } else if (current_cell_type_ == "str" || current_cell_type_ == "inlineStr") {
            // 内联字符串
            storage_->setError(static_cast<uint32_t>(current_row_), current_col_, std::string(cell_value_view));
            cells_processed_++;
        } else {
            // 数值（默认）- fast_float已经高度优化
            double numeric_value;
            auto result = fast_float::from_chars(cell_value_view.data(),
                                               cell_value_view.data() + cell_value_view.size(),
                                               numeric_value);
            if (result.ec == std::errc{}) {
                storage_->setValue(static_cast<uint32_t>(current_row_), current_col_, numeric_value);
                cells_processed_++;
            }
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("解析单元格值失败: {} (type: {}, value: {})", 
                           e.what(), current_cell_type_, current_cell_value_.view());
        // 解析失败不影响整体进程，跳过这个单元格
    }
}

void ReadOnlyWorksheetParser::onStartElement(std::string_view name, 
                                            span<const xml::XMLAttribute> attributes, int depth) {
    
    if (name == "sheetData") {
        in_sheet_data_ = true;
        FASTEXCEL_LOG_DEBUG("开始解析工作表数据");
    } else if (name == "row" && in_sheet_data_) {
        in_row_ = true;
        
        // 提取行号 - 使用 fast_float 高性能解析
        for (const auto& attr : attributes) {
            if (attr.name == "r") {
                try {
                    uint32_t row_number;
                    auto result = fast_float::from_chars(attr.value.data(),
                                                       attr.value.data() + attr.value.size(),
                                                       row_number);
                    if (result.ec == std::errc{}) {
                        current_row_ = static_cast<int>(row_number);
                    }
                } catch (const std::exception&) {
                    FASTEXCEL_LOG_DEBUG("无效的行号: {}", attr.value);
                    current_row_ = 0;
                }
                break;
            }
        }
    } else if (name == "c" && in_row_) {
        in_cell_ = true;
        
        // 提取单元格引用和类型 - 避免临时string创建
        for (const auto& attr : attributes) {
            if (attr.name == "r") {
                current_col_ = parseColumnReference(attr.value);
            } else if (attr.name == "t") {
                current_cell_type_ = attr.value;
            }
        }
        
        current_cell_value_.clear();
    } else if (name == "v" && in_cell_) {
        in_value_ = true;
        current_cell_value_.clear();
    }
}

void ReadOnlyWorksheetParser::onEndElement(std::string_view name, int depth) {
    
    if (name == "sheetData") {
        in_sheet_data_ = false;
        FASTEXCEL_LOG_INFO("完成工作表数据解析，处理了 {} 个单元格", cells_processed_);
    } else if (name == "row") {
        in_row_ = false;
        current_row_ = 0;
    } else if (name == "c") {
        in_cell_ = false;
        // 单元格结束时处理数据
        if (!current_cell_value_.empty()) {
            processCellValue();
        }
        current_col_ = 0;
        current_cell_type_ = std::string_view{};
        current_cell_value_.clear();
    } else if (name == "v") {
        in_value_ = false;
        // 值已经在onText中收集
    }
}

void ReadOnlyWorksheetParser::onText(std::string_view data, int depth) {
    if (in_value_) {
        current_cell_value_.append(data);
    }
}

}} // namespace fastexcel::reader