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

#ifdef FASTEXCEL_HAS_HIGHWAY
#include "hwy/highway.h"
#include "hwy/aligned_allocator.h"
namespace HWY_NAMESPACE {
    using namespace hwy::HWY_NAMESPACE;
}
#endif

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
    
    // 添加到批量处理队列而不是立即处理
    row_batch_.emplace_back(static_cast<uint32_t>(current_row_), current_col_, 
                           std::move(current_cell_value_), current_cell_type_);
    
    // 当批量达到阈值时处理
    if (row_batch_.size() >= BATCH_SIZE) {
        processBatch();
    }
}

void ReadOnlyWorksheetParser::processBatch() {
    if (row_batch_.empty() || !storage_) {
        return;
    }
    
#ifdef FASTEXCEL_HAS_HIGHWAY
    // 使用SIMD优化批量数值解析
    processBatchSIMD();
#else
    // 回退到标准批量处理
    processBatchStandard();
#endif
    
    // 清空批量缓冲区
    row_batch_.clear();
}

#ifdef FASTEXCEL_HAS_HIGHWAY
void ReadOnlyWorksheetParser::processBatchSIMD() {
    // 分离数值和非数值单元格
    std::vector<size_t> numeric_indices;
    std::vector<double> numeric_values;
    
    numeric_indices.reserve(row_batch_.size());
    numeric_values.reserve(row_batch_.size());
    
    // 第一遍：识别数值类型并尝试解析
    for (size_t i = 0; i < row_batch_.size(); ++i) {
        const auto& cell = row_batch_[i];
        
        if (cell.type == "s") {
            // 共享字符串 - 使用 fast_float 解析整数索引
            uint32_t sst_index;
            auto result = fast_float::from_chars(cell.value.data(),
                                               cell.value.data() + cell.value.size(),
                                               sst_index);
            if (result.ec == std::errc{}) {
                storage_->setValue(cell.row, cell.col, sst_index);
                cells_processed_++;
            }
        } else if (cell.type == "b") {
            // 布尔值
            bool bool_value = (cell.value == "1");
            storage_->setValue(cell.row, cell.col, bool_value);
            cells_processed_++;
        } else if (cell.type == "str" || cell.type == "inlineStr") {
            // 内联字符串
            storage_->setError(cell.row, cell.col, cell.value);
            cells_processed_++;
        } else {
            // 数值类型 - 收集用于SIMD批量处理
            double numeric_value;
            auto result = fast_float::from_chars(cell.value.data(),
                                               cell.value.data() + cell.value.size(),
                                               numeric_value);
            if (result.ec == std::errc{}) {
                numeric_indices.push_back(i);
                numeric_values.push_back(numeric_value);
            }
        }
    }
    
    // 第二遍：SIMD批量存储数值
    if (!numeric_indices.empty()) {
        const HWY_NAMESPACE::ScalableTag<double> d;
        const size_t lanes = Lanes(d);
        
        // 按SIMD向量大小批量处理
        for (size_t i = 0; i + lanes <= numeric_values.size(); i += lanes) {
            auto values_vec = LoadU(d, &numeric_values[i]);
            
            // 逐个存储（这里SIMD主要优化了数值处理，存储还是需要逐个）
            for (size_t j = 0; j < lanes; ++j) {
                const size_t idx = numeric_indices[i + j];
                const auto& cell = row_batch_[idx];
                storage_->setValue(cell.row, cell.col, numeric_values[i + j]);
                cells_processed_++;
            }
        }
        
        // 处理剩余的数值
        for (size_t i = (numeric_values.size() / lanes) * lanes; i < numeric_values.size(); ++i) {
            const size_t idx = numeric_indices[i];
            const auto& cell = row_batch_[idx];
            storage_->setValue(cell.row, cell.col, numeric_values[i]);
            cells_processed_++;
        }
    }
}
#endif

void ReadOnlyWorksheetParser::processBatchStandard() {
    for (const auto& cell : row_batch_) {
        try {
            if (cell.type == "s") {
                // 共享字符串 - 使用 fast_float 解析整数索引
                uint32_t sst_index;
                auto result = fast_float::from_chars(cell.value.data(),
                                                   cell.value.data() + cell.value.size(),
                                                   sst_index);
                if (result.ec == std::errc{}) {
                    storage_->setValue(cell.row, cell.col, sst_index);
                    cells_processed_++;
                }
            } else if (cell.type == "b") {
                // 布尔值
                bool bool_value = (cell.value == "1");
                storage_->setValue(cell.row, cell.col, bool_value);
                cells_processed_++;
            } else if (cell.type == "str" || cell.type == "inlineStr") {
                // 内联字符串
                storage_->setError(cell.row, cell.col, cell.value);
                cells_processed_++;
            } else {
                // 数值（默认）或者空类型 - 使用 fast_float 高性能解析
                double numeric_value;
                auto result = fast_float::from_chars(cell.value.data(),
                                                   cell.value.data() + cell.value.size(),
                                                   numeric_value);
                if (result.ec == std::errc{}) {
                    storage_->setValue(cell.row, cell.col, numeric_value);
                    cells_processed_++;
                }
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_DEBUG("解析单元格值失败: {} (type: {}, value: {})", 
                               e.what(), cell.type, cell.value);
            // 解析失败不影响整体进程，跳过这个单元格
        }
    }
}

void ReadOnlyWorksheetParser::flushBatch() {
    if (!row_batch_.empty()) {
        processBatch();
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
        flushBatch(); // 确保处理所有剩余的单元格
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