#include "fastexcel/edit/RowWriter.hpp"
#include "fastexcel/core/Logger.hpp"
#include <sstream>
#include <iomanip>

namespace fastexcel::edit {

// 构造函数
RowWriter::RowWriter(std::shared_ptr<core::Worksheet> worksheet, 
                     core::DirtyManager* dirty_manager)
    : worksheet_(worksheet)
    , dirty_manager_(dirty_manager)
    , current_row_(0)
    , current_col_(0)
    , rows_written_(0)
    , cells_written_(0)
    , buffer_size_(1000)
    , auto_flush_enabled_(true)
    , streaming_mode_(false) {
    
    if (!worksheet_) {
        throw std::invalid_argument("Worksheet cannot be null");
    }
    
    // 预分配缓冲区
    row_buffer_.reserve(buffer_size_);
    
    LOG_DEBUG("RowWriter created for worksheet: {}", worksheet_->getName());
}

// 析构函数
RowWriter::~RowWriter() {
    // 确保所有数据都被写入
    if (!row_buffer_.empty()) {
        flush();
    }
    
    LOG_DEBUG("RowWriter destroyed - Rows: {}, Cells: {}", 
             rows_written_, cells_written_);
}

// 写入字符串值
RowWriter& RowWriter::writeString(const std::string& value) {
    addToBuffer(CellData{current_row_, current_col_, value});
    current_col_++;
    cells_written_++;
    
    checkAutoFlush();
    return *this;
}

// 写入数值
RowWriter& RowWriter::writeNumber(double value) {
    addToBuffer(CellData{current_row_, current_col_, value});
    current_col_++;
    cells_written_++;
    
    checkAutoFlush();
    return *this;
}

// 写入布尔值
RowWriter& RowWriter::writeBool(bool value) {
    addToBuffer(CellData{current_row_, current_col_, value});
    current_col_++;
    cells_written_++;
    
    checkAutoFlush();
    return *this;
}

// 写入公式
RowWriter& RowWriter::writeFormula(const std::string& formula) {
    // 确保公式以=开头
    std::string full_formula = formula;
    if (!formula.empty() && formula[0] != '=') {
        full_formula = "=" + formula;
    }
    
    addToBuffer(CellData{current_row_, current_col_, full_formula, true});
    current_col_++;
    cells_written_++;
    
    checkAutoFlush();
    return *this;
}

// 写入日期时间
RowWriter& RowWriter::writeDateTime(const std::chrono::system_clock::time_point& datetime) {
    // 转换为Excel日期格式（从1900-01-01开始的天数）
    static const auto excel_epoch = std::chrono::system_clock::from_time_t(-2209161600);
    auto duration = datetime - excel_epoch;
    auto days = std::chrono::duration_cast<std::chrono::duration<double, std::ratio<86400>>>(duration);
    
    writeNumber(days.count());
    
    // 标记单元格需要日期格式
    if (dirty_manager_) {
        dirty_manager_->markCellFormatted(worksheet_->getName(), current_row_, current_col_ - 1);
    }
    
    return *this;
}

// 写入空单元格
RowWriter& RowWriter::writeEmpty() {
    // 跳过当前列
    current_col_++;
    return *this;
}

// 移动到下一行
RowWriter& RowWriter::nextRow() {
    current_row_++;
    current_col_ = 0;
    rows_written_++;
    
    // 在流式模式下，每行结束时自动刷新
    if (streaming_mode_) {
        flush();
    }
    
    return *this;
}

// 跳过指定数量的行
RowWriter& RowWriter::skipRows(size_t count) {
    current_row_ += count;
    current_col_ = 0;
    return *this;
}

// 跳过指定数量的列
RowWriter& RowWriter::skipColumns(size_t count) {
    current_col_ += count;
    return *this;
}

// 移动到指定位置
RowWriter& RowWriter::moveTo(size_t row, size_t col) {
    current_row_ = row;
    current_col_ = col;
    return *this;
}

// 写入整行数据
RowWriter& RowWriter::writeRow(const std::vector<std::variant<std::string, double, bool>>& values) {
    for (const auto& value : values) {
        std::visit([this](const auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, std::string>) {
                writeString(v);
            } else if constexpr (std::is_same_v<T, double>) {
                writeNumber(v);
            } else if constexpr (std::is_same_v<T, bool>) {
                writeBool(v);
            }
        }, value);
    }
    
    nextRow();
    return *this;
}

// 批量写入多行
RowWriter& RowWriter::writeRows(const std::vector<std::vector<std::variant<std::string, double, bool>>>& rows) {
    for (const auto& row : rows) {
        writeRow(row);
    }
    return *this;
}

// 写入CSV格式的行
RowWriter& RowWriter::writeCSVRow(const std::string& csv_line, char delimiter) {
    std::vector<std::string> values;
    std::stringstream ss(csv_line);
    std::string cell;
    
    while (std::getline(ss, cell, delimiter)) {
        // 去除引号（如果有）
        if (!cell.empty() && cell.front() == '"' && cell.back() == '"') {
            cell = cell.substr(1, cell.length() - 2);
        }
        
        // 尝试解析为数字
        try {
            size_t pos;
            double num = std::stod(cell, &pos);
            if (pos == cell.length()) {
                writeNumber(num);
            } else {
                writeString(cell);
            }
        } catch (...) {
            // 不是数字，作为字符串写入
            writeString(cell);
        }
    }
    
    nextRow();
    return *this;
}

// 设置列宽
RowWriter& RowWriter::setColumnWidth(size_t col, double width) {
    worksheet_->setColumnWidth(col, width);
    
    if (dirty_manager_) {
        dirty_manager_->markWorksheetModified(worksheet_->getName());
    }
    
    return *this;
}

// 设置行高
RowWriter& RowWriter::setRowHeight(size_t row, double height) {
    worksheet_->setRowHeight(row, height);
    
    if (dirty_manager_) {
        dirty_manager_->markWorksheetModified(worksheet_->getName());
    }
    
    return *this;
}

// 应用单元格样式
RowWriter& RowWriter::applyStyle(size_t row, size_t col, const CellStyle& style) {
    // 应用样式到指定单元格
    auto cell = worksheet_->getCell(row, col);
    if (cell) {
        if (style.bold || style.italic) {
            cell->setFontStyle(style.bold, style.italic);
        }
        if (style.font_size > 0) {
            cell->setFontSize(style.font_size);
        }
        if (!style.font_color.empty()) {
            cell->setFontColor(style.font_color);
        }
        if (!style.bg_color.empty()) {
            cell->setBackgroundColor(style.bg_color);
        }
        if (style.has_border) {
            cell->setBorder(true);
        }
        if (!style.number_format.empty()) {
            cell->setNumberFormat(style.number_format);
        }
        
        if (dirty_manager_) {
            dirty_manager_->markCellModified(worksheet_->getName(), row, col);
        }
    }
    
    return *this;
}

// 合并单元格
RowWriter& RowWriter::mergeCells(size_t start_row, size_t start_col, 
                                 size_t end_row, size_t end_col) {
    worksheet_->mergeCells(start_row, start_col, end_row, end_col);
    
    if (dirty_manager_) {
        dirty_manager_->markWorksheetModified(worksheet_->getName());
    }
    
    return *this;
}

// 刷新缓冲区
void RowWriter::flush() {
    if (row_buffer_.empty()) {
        return;
    }
    
    LOG_TRACE("Flushing {} cells to worksheet", row_buffer_.size());
    
    // 批量写入所有缓冲的单元格
    for (const auto& cell_data : row_buffer_) {
        std::visit([this, &cell_data](const auto& value) {
            using T = std::decay_t<decltype(value)>;
            
            if constexpr (std::is_same_v<T, std::string>) {
                if (cell_data.is_formula) {
                    worksheet_->setCellFormula(cell_data.row, cell_data.col, value);
                } else {
                    worksheet_->setCellString(cell_data.row, cell_data.col, value);
                }
            } else if constexpr (std::is_same_v<T, double>) {
                worksheet_->setCellNumber(cell_data.row, cell_data.col, value);
            } else if constexpr (std::is_same_v<T, bool>) {
                worksheet_->setCellBool(cell_data.row, cell_data.col, value);
            }
            
            // 标记单元格已修改
            if (dirty_manager_) {
                dirty_manager_->markCellModified(worksheet_->getName(), 
                                                cell_data.row, cell_data.col);
            }
        }, cell_data.value);
    }
    
    // 清空缓冲区
    row_buffer_.clear();
}

// 启用自动刷新
void RowWriter::enableAutoFlush(size_t threshold) {
    auto_flush_enabled_ = true;
    buffer_size_ = threshold;
    LOG_DEBUG("Auto-flush enabled with threshold: {}", threshold);
}

// 禁用自动刷新
void RowWriter::disableAutoFlush() {
    auto_flush_enabled_ = false;
    LOG_DEBUG("Auto-flush disabled");
}

// 启用流式模式
void RowWriter::enableStreamingMode() {
    streaming_mode_ = true;
    auto_flush_enabled_ = true;
    buffer_size_ = 100;  // 流式模式下使用较小的缓冲区
    LOG_INFO("Streaming mode enabled");
}

// 禁用流式模式
void RowWriter::disableStreamingMode() {
    streaming_mode_ = false;
    buffer_size_ = 1000;  // 恢复默认缓冲区大小
    LOG_INFO("Streaming mode disabled");
}

// 获取当前位置
RowWriter::Position RowWriter::getCurrentPosition() const {
    return {current_row_, current_col_};
}

// 获取统计信息
RowWriter::Stats RowWriter::getStats() const {
    Stats stats;
    stats.rows_written = rows_written_;
    stats.cells_written = cells_written_;
    stats.buffer_size = row_buffer_.size();
    stats.is_streaming = streaming_mode_;
    return stats;
}

// 重置写入器
void RowWriter::reset() {
    // 刷新所有待写入的数据
    flush();
    
    // 重置状态
    current_row_ = 0;
    current_col_ = 0;
    rows_written_ = 0;
    cells_written_ = 0;
    
    LOG_DEBUG("RowWriter reset");
}

// 添加到缓冲区
void RowWriter::addToBuffer(const CellData& data) {
    row_buffer_.push_back(data);
}

// 检查是否需要自动刷新
void RowWriter::checkAutoFlush() {
    if (auto_flush_enabled_ && row_buffer_.size() >= buffer_size_) {
        flush();
    }
}

// 创建表格头部的辅助方法
RowWriter& RowWriter::writeHeader(const std::vector<std::string>& headers, 
                                  const CellStyle& style) {
    size_t col = 0;
    for (const auto& header : headers) {
        writeString(header);
        if (style.bold || style.font_size > 0 || !style.bg_color.empty()) {
            applyStyle(current_row_, col, style);
        }
        col++;
    }
    nextRow();
    return *this;
}

// 写入数据表格的辅助方法
RowWriter& RowWriter::writeTable(
    const std::vector<std::string>& headers,
    const std::vector<std::vector<std::variant<std::string, double, bool>>>& data,
    bool with_borders) {
    
    // 写入表头
    CellStyle header_style;
    header_style.bold = true;
    header_style.bg_color = "#E0E0E0";
    header_style.has_border = with_borders;
    writeHeader(headers, header_style);
    
    // 写入数据
    size_t start_row = current_row_;
    for (const auto& row : data) {
        writeRow(row);
        
        // 应用边框
        if (with_borders) {
            for (size_t col = 0; col < row.size(); ++col) {
                CellStyle cell_style;
                cell_style.has_border = true;
                applyStyle(current_row_ - 1, col, cell_style);
            }
        }
    }
    
    // 自动调整列宽
    for (size_t col = 0; col < headers.size(); ++col) {
        worksheet_->autoFitColumn(col);
    }
    
    LOG_INFO("Table written: {} headers, {} rows", headers.size(), data.size());
    
    return *this;
}

// 写入汇总行的辅助方法
RowWriter& RowWriter::writeSummaryRow(const std::string& label, 
                                      const std::vector<std::string>& formulas) {
    // 写入标签
    writeString(label);
    
    // 写入公式
    for (const auto& formula : formulas) {
        writeFormula(formula);
    }
    
    // 应用样式（加粗）
    CellStyle summary_style;
    summary_style.bold = true;
    for (size_t col = 0; col <= formulas.size(); ++col) {
        applyStyle(current_row_, col, summary_style);
    }
    
    nextRow();
    return *this;
}

} // namespace fastexcel::edit