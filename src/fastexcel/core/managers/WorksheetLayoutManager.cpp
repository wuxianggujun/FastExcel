#include "WorksheetLayoutManager.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace core {

WorksheetLayoutManager::WorksheetLayoutManager() {
}

double WorksheetLayoutManager::setColumnWidth(int col, double width) {
    validateCellPosition(0, col);
    
    // 使用基础列宽计算器
    auto calculator = utils::ColumnWidthCalculator(utils::ColumnWidthCalculator::FontType::CALIBRI_11);
    double actual_width = calculator.quantize(width);
    
    column_info_[col].width = actual_width;
    column_info_[col].precise_width = true;
    
    FASTEXCEL_LOG_DEBUG("设置列{}宽度: {} -> {}", col, width, actual_width);
    
    return actual_width;
}

std::pair<double, int> WorksheetLayoutManager::setColumnWidthAdvanced(int col, double target_width,
                                                                     const std::string& font_name,
                                                                     double font_size,
                                                                     ColumnWidthManager::WidthStrategy strategy,
                                                                     const std::vector<std::string>& cell_contents) {
    validateCellPosition(0, col);
    
    // 确保列宽管理器已初始化
    if (!column_width_manager_) {
        column_width_manager_ = std::make_unique<ColumnWidthManager>(format_repo_.get());
    }
    
    // 构建配置
    ColumnWidthManager::ColumnWidthConfig config(target_width, font_name, font_size, strategy);
    
    // 使用内容感知模式时，传递单元格内容
    std::pair<double, int> result;
    if (strategy == ColumnWidthManager::WidthStrategy::CONTENT_AWARE && !cell_contents.empty()) {
        result = column_width_manager_->setSmartColumnWidth(col, target_width, cell_contents);
    } else {
        result = column_width_manager_->setColumnWidth(col, config);
    }
    
    // 更新列信息
    column_info_[col].width = result.first;
    column_info_[col].precise_width = true;
    if (result.second >= 0) {
        column_info_[col].format_id = result.second;
    }
    
    return result;
}

std::unordered_map<int, std::pair<double, int>> WorksheetLayoutManager::setColumnWidthsBatch(
    const std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig>& configs) {
    
    // 确保列宽管理器已初始化
    if (!column_width_manager_) {
        column_width_manager_ = std::make_unique<ColumnWidthManager>(format_repo_.get());
    }
    
    // 验证所有列位置
    for (const auto& [col, config] : configs) {
        validateCellPosition(0, col);
    }
    
    // 批量处理
    auto results = column_width_manager_->setColumnWidths(configs);
    
    // 更新列信息
    for (const auto& [col, result] : results) {
        column_info_[col].width = result.first;
        column_info_[col].precise_width = true;
        if (result.second >= 0) {
            column_info_[col].format_id = result.second;
        }
    }
    
    return results;
}

double WorksheetLayoutManager::calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const {
    if (!column_width_manager_) {
        auto temp_manager = std::make_unique<ColumnWidthManager>(format_repo_.get());
        return temp_manager->calculateOptimalWidth(target_width, font_name, font_size);
    }
    
    return column_width_manager_->calculateOptimalWidth(target_width, font_name, font_size);
}

double WorksheetLayoutManager::getColumnWidth(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end() && it->second.width > 0) {
        return it->second.width;
    }
    return default_col_width_;
}

void WorksheetLayoutManager::setColumnFormatId(int col, int format_id) {
    validateCellPosition(0, col);
    column_info_[col].format_id = format_id;
}

void WorksheetLayoutManager::setColumnFormatId(int first_col, int last_col, int format_id) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].format_id = format_id;
    }
}

int WorksheetLayoutManager::getColumnFormatId(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end()) {
        return it->second.format_id;
    }
    return -1;
}

void WorksheetLayoutManager::hideColumn(int col) {
    validateCellPosition(0, col);
    column_info_[col].hidden = true;
}

void WorksheetLayoutManager::hideColumn(int first_col, int last_col) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].hidden = true;
    }
}

bool WorksheetLayoutManager::isColumnHidden(int col) const {
    auto it = column_info_.find(col);
    return it != column_info_.end() && it->second.hidden;
}

void WorksheetLayoutManager::setRowHeight(int row, double height) {
    validateCellPosition(row, 0);
    row_info_[row].height = height;
}

double WorksheetLayoutManager::getRowHeight(int row) const {
    auto it = row_info_.find(row);
    if (it != row_info_.end() && it->second.height > 0) {
        return it->second.height;
    }
    return default_row_height_;
}

void WorksheetLayoutManager::hideRow(int row) {
    validateCellPosition(row, 0);
    row_info_[row].hidden = true;
}

void WorksheetLayoutManager::hideRow(int first_row, int last_row) {
    validateRange(first_row, 0, last_row, 0);
    for (int row = first_row; row <= last_row; ++row) {
        row_info_[row].hidden = true;
    }
}

bool WorksheetLayoutManager::isRowHidden(int row) const {
    auto it = row_info_.find(row);
    return it != row_info_.end() && it->second.hidden;
}

void WorksheetLayoutManager::mergeCells(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    merge_ranges_.emplace_back(first_row, first_col, last_row, last_col);
}

void WorksheetLayoutManager::setAutoFilter(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    autofilter_ = std::make_unique<AutoFilterRange>(first_row, first_col, last_row, last_col);
}

void WorksheetLayoutManager::removeAutoFilter() {
    autofilter_.reset();
}

AutoFilterRange WorksheetLayoutManager::getAutoFilterRange() const {
    if (autofilter_) {
        return *autofilter_;
    }
    return AutoFilterRange(0, 0, 0, 0);
}

void WorksheetLayoutManager::freezePanes(int row, int col) {
    validateCellPosition(row, col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

void WorksheetLayoutManager::freezePanes(int row, int col, int top_left_row, int top_left_col) {
    validateCellPosition(row, col);
    validateCellPosition(top_left_row, top_left_col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col, top_left_row, top_left_col);
}

void WorksheetLayoutManager::splitPanes(int row, int col) {
    validateCellPosition(row, col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

FreezePanes WorksheetLayoutManager::getFreezeInfo() const {
    if (freeze_panes_) {
        return *freeze_panes_;
    }
    return FreezePanes();
}

void WorksheetLayoutManager::clear() {
    column_info_.clear();
    row_info_.clear();
    merge_ranges_.clear();
    autofilter_.reset();
    freeze_panes_.reset();
}

void WorksheetLayoutManager::setFormatRepository(std::shared_ptr<FormatRepository> format_repo) {
    format_repo_ = format_repo;
    if (column_width_manager_) {
        column_width_manager_ = std::make_unique<ColumnWidthManager>(format_repo_.get());
    }
}

void WorksheetLayoutManager::validateCellPosition(int row, int col) const {
    FASTEXCEL_VALIDATE_CELL_POSITION(row, col);
}

void WorksheetLayoutManager::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col);
}

} // namespace core
} // namespace fastexcel
