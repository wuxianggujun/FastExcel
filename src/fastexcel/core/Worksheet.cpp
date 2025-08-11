#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/WorksheetChain.hpp"  // 🚀 新增：链式调用支持
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/CellRangeManager.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/WorksheetXMLGenerator.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/LogConfig.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/core/Exception.hpp"
#include <sstream>
#include <stdexcept>
#include <algorithm>
#include <iomanip>
#include <cmath>

namespace fastexcel {
namespace core {

Worksheet::Worksheet(const std::string& name, std::shared_ptr<Workbook> workbook, int sheet_id)
    : name_(name), parent_workbook_(workbook), sheet_id_(sheet_id) {
    // 初始化共享公式管理器
    shared_formula_manager_ = std::make_unique<SharedFormulaManager>();
}

// ========== 基本单元格操作 ==========

Cell& Worksheet::getCell(int row, int col) {
    validateCellPosition(row, col);
    updateUsedRange(row, col);
    return cells_[std::make_pair(row, col)];
}

const Cell& Worksheet::getCell(int row, int col) const {
    validateCellPosition(row, col);
    auto it = cells_.find(std::make_pair(row, col));
    if (it == cells_.end()) {
        static Cell empty_cell;
        return empty_cell;
    }
    return it->second;
}

// ========== 基本写入方法 ==========

void Worksheet::writeString(int row, int col, const std::string& value) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    Cell cell;
    if (sst_) {
        // 使用共享字符串表
        sst_->addString(value);
        cell.setValue(value);
    } else {
        cell.setValue(value);
    }
    
    if (optimize_mode_) {
        this->writeOptimizedCell(row, col, std::move(cell));
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        this->updateUsedRange(row, col);
    }
}

void Worksheet::writeNumber(int row, int col, double value) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    Cell cell;
    cell.setValue(value);
    
    if (optimize_mode_) {
        this->writeOptimizedCell(row, col, std::move(cell));
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        this->updateUsedRange(row, col);
    }
}

void Worksheet::writeBoolean(int row, int col, bool value) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    Cell cell;
    cell.setValue(value);
    
    if (optimize_mode_) {
        this->writeOptimizedCell(row, col, std::move(cell));
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        this->updateUsedRange(row, col);
    }
}

void Worksheet::writeFormula(int row, int col, const std::string& formula) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    Cell cell;
    cell.setFormula(formula);
    
    if (optimize_mode_) {
        this->writeOptimizedCell(row, col, std::move(cell));
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        this->updateUsedRange(row, col);
    }
}

int Worksheet::createSharedFormula(int first_row, int first_col, int last_row, int last_col, const std::string& formula) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    
    // 验证范围
    validateRange(first_row, first_col, last_row, last_col);
    
    if (!shared_formula_manager_) {
        shared_formula_manager_ = std::make_unique<SharedFormulaManager>();
    }
    
    // 创建共享公式范围字符串
    std::string range = utils::CommonUtils::cellReference(first_row, first_col) + ":" +
                       utils::CommonUtils::cellReference(last_row, last_col);
    
    // 注册共享公式
    int shared_index = shared_formula_manager_->registerSharedFormula(formula, range);
    if (shared_index < 0) {
        CORE_ERROR("Failed to register shared formula in range {}", range);
        return -1;
    }
    
    // 获取注册的共享公式对象并更新受影响的单元格列表
    const SharedFormula* shared_formula = shared_formula_manager_->getSharedFormula(shared_index);
    if (shared_formula) {
        // 手动添加受影响的单元格到统计中
        for (int row = first_row; row <= last_row; ++row) {
            for (int col = first_col; col <= last_col; ++col) {
                // 这里需要调用非const版本来更新affected_cells_
                auto* mutable_formula = const_cast<SharedFormula*>(shared_formula);
                mutable_formula->addAffectedCell(row, col);
            }
        }
    }
    
    // 为范围内的每个单元格设置共享公式引用
    for (int row = first_row; row <= last_row; ++row) {
        for (int col = first_col; col <= last_col; ++col) {
            Cell cell;
            if (row == first_row && col == first_col) {
                // 主单元格存储完整的基础公式和共享公式索引
                cell.setFormula(formula);  // 先设置常规公式
                cell.setSharedFormula(shared_index);  // 然后转换为共享公式
            } else {
                // 其他单元格只存储共享公式引用
                cell.setSharedFormulaReference(shared_index);
            }
            
            if (optimize_mode_) {
                this->writeOptimizedCell(row, col, std::move(cell));
            } else {
                auto& target_cell = cells_[std::make_pair(row, col)];
                target_cell = std::move(cell);
                this->updateUsedRange(row, col);
            }
        }
    }
    
    CORE_DEBUG("Created shared formula: index={}, range={}, formula='{}'", 
             shared_index, range, formula);
    
    return shared_index;
}

void Worksheet::writeDateTime(int row, int col, const std::tm& datetime) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    // 使用 TimeUtils 将日期时间转换为Excel序列号
    double excel_serial = utils::TimeUtils::toExcelSerialNumber(datetime);
    this->writeNumber(row, col, excel_serial);
}

void Worksheet::writeUrl(int row, int col, const std::string& url, const std::string& string) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        std::string rels_path = "xl/worksheets/_rels/sheet" + std::to_string(sheet_id_) + ".xml.rels";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
        parent_workbook_->getDirtyManager()->markDirty(rels_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    
    std::string display_text = string.empty() ? url : string;
    cell.setValue(display_text);
    cell.setHyperlink(url);
    
    this->updateUsedRange(row, col);
}

// ========== 批量数据操作 ==========

void Worksheet::writeRange(int start_row, int start_col, const std::vector<std::vector<std::string>>& data) {
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            writeString(static_cast<int>(start_row + row), static_cast<int>(start_col + col), data[row][col]);
        }
    }
}

void Worksheet::writeRange(int start_row, int start_col, const std::vector<std::vector<double>>& data) {
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            writeNumber(static_cast<int>(start_row + row), static_cast<int>(start_col + col), data[row][col]);
        }
    }
}

// ========== 行列操作 ==========

void Worksheet::setColumnWidth(int col, double width) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(0, col);
    column_info_[col].width = width;
}

void Worksheet::setColumnWidth(int first_col, int last_col, double width) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].width = width;
    }
}

void Worksheet::setColumnFormatId(int col, int format_id) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(0, col);
    column_info_[col].format_id = format_id;
}

void Worksheet::setColumnFormatId(int first_col, int last_col, int format_id) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].format_id = format_id;
    }
}

// setColumnFormat方法已移除，请使用FormatDescriptor架构

// setColumnFormat范围方法已移除，请使用FormatDescriptor架构

void Worksheet::hideColumn(int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(0, col);
    column_info_[col].hidden = true;
}

void Worksheet::hideColumn(int first_col, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].hidden = true;
    }
}

void Worksheet::setRowHeight(int row, double height) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(row, 0);
    row_info_[row].height = height;
}

// setRowFormat方法已移除，请使用FormatDescriptor架构

void Worksheet::hideRow(int row) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(row, 0);
    row_info_[row].hidden = true;
}

void Worksheet::hideRow(int first_row, int last_row) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(first_row, 0, last_row, 0);
    for (int row = first_row; row <= last_row; ++row) {
        row_info_[row].hidden = true;
    }
}

// ========== 合并单元格 ==========

void Worksheet::mergeCells(int first_row, int first_col, int last_row, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(first_row, first_col, last_row, last_col);
    merge_ranges_.emplace_back(first_row, first_col, last_row, last_col);
}

// ========== 自动筛选 ==========

void Worksheet::setAutoFilter(int first_row, int first_col, int last_row, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(first_row, first_col, last_row, last_col);
    autofilter_ = std::make_unique<AutoFilterRange>(first_row, first_col, last_row, last_col);
}

void Worksheet::removeAutoFilter() {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    autofilter_.reset();
}

// ========== 冻结窗格 ==========

void Worksheet::freezePanes(int row, int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(row, col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

void Worksheet::freezePanes(int row, int col, int top_left_row, int top_left_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(row, col);
    validateCellPosition(top_left_row, top_left_col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col, top_left_row, top_left_col);
}

void Worksheet::splitPanes(int row, int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateCellPosition(row, col);
    // 分割窗格的实现与冻结窗格类似，但使用不同的XML属性
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

// ========== 打印设置 ==========

void Worksheet::setPrintArea(int first_row, int first_col, int last_row, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(first_row, first_col, last_row, last_col);
    print_settings_.print_area_first_row = first_row;
    print_settings_.print_area_first_col = first_col;
    print_settings_.print_area_last_row = last_row;
    print_settings_.print_area_last_col = last_col;
}

void Worksheet::setRepeatRows(int first_row, int last_row) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(first_row, 0, last_row, 0);
    print_settings_.repeat_rows_first = first_row;
    print_settings_.repeat_rows_last = last_row;
}

void Worksheet::setRepeatColumns(int first_col, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    validateRange(0, first_col, 0, last_col);
    print_settings_.repeat_cols_first = first_col;
    print_settings_.repeat_cols_last = last_col;
}

void Worksheet::setLandscape(bool landscape) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    print_settings_.landscape = landscape;
}

void Worksheet::setPaperSize(int paper_size) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    // 纸张大小代码的实现
    // 这里可以根据需要添加具体的纸张大小映射
    (void)paper_size; // 避免未使用参数警告
}

void Worksheet::setMargins(double left, double right, double top, double bottom) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    print_settings_.left_margin = left;
    print_settings_.right_margin = right;
    print_settings_.top_margin = top;
    print_settings_.bottom_margin = bottom;
}

void Worksheet::setHeaderFooterMargins(double header, double footer) {
    print_settings_.header_margin = header;
    print_settings_.footer_margin = footer;
}

void Worksheet::setPrintScale(int scale) {
    print_settings_.scale = std::max(10, std::min(400, scale));
    print_settings_.fit_to_pages_wide = 0;
    print_settings_.fit_to_pages_tall = 0;
}

void Worksheet::setFitToPages(int width, int height) {
    print_settings_.fit_to_pages_wide = width;
    print_settings_.fit_to_pages_tall = height;
    print_settings_.scale = 100;
}

void Worksheet::setPrintGridlines(bool print) {
    print_settings_.print_gridlines = print;
}

void Worksheet::setPrintHeadings(bool print) {
    print_settings_.print_headings = print;
}

void Worksheet::setCenterOnPage(bool horizontal, bool vertical) {
    print_settings_.center_horizontally = horizontal;
    print_settings_.center_vertically = vertical;
}

// ========== 工作表保护 ==========

void Worksheet::protect(const std::string& password) {
    protected_ = true;
    protection_password_ = password;
}

void Worksheet::unprotect() {
    protected_ = false;
    protection_password_.clear();
}

// ========== 视图设置 ==========

void Worksheet::setZoom(int scale) {
    sheet_view_.zoom_scale = std::max(10, std::min(400, scale));
}

void Worksheet::showGridlines(bool show) {
    sheet_view_.show_gridlines = show;
}

void Worksheet::showRowColHeaders(bool show) {
    sheet_view_.show_row_col_headers = show;
}

void Worksheet::setRightToLeft(bool rtl) {
    sheet_view_.right_to_left = rtl;
}

void Worksheet::setTabSelected(bool selected) {
    sheet_view_.tab_selected = selected;
}

void Worksheet::setActiveCell(int row, int col) {
    validateCellPosition(row, col);
    active_cell_ = utils::CommonUtils::cellReference(row, col);
}

void Worksheet::setSelection(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    if (first_row == last_row && first_col == last_col) {
        selection_ = utils::CommonUtils::cellReference(first_row, first_col);
    } else {
        selection_ = utils::CommonUtils::rangeReference(first_row, first_col, last_row, last_col);
    }
}

// ========== 获取信息 ==========

std::pair<int, int> Worksheet::getUsedRange() const {
    return range_manager_.getUsedRowRange().first != -1 ? 
           std::make_pair(range_manager_.getUsedRowRange().second, range_manager_.getUsedColRange().second) :
           std::make_pair(-1, -1);
}

bool Worksheet::hasCellAt(int row, int col) const {
    auto it = cells_.find(std::make_pair(row, col));
    // 🔧 关键修复：有格式的空单元格也应该被认为是存在的，以便保持格式信息
    return it != cells_.end() && (!it->second.isEmpty() || it->second.hasFormat());
}

// ========== 获取方法实现 ==========

double Worksheet::getColumnWidth(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end() && it->second.width > 0) {
        return it->second.width;
    }
    return default_col_width_;
}

int Worksheet::getColumnFormatId(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end()) {
        return it->second.format_id;
    }
    return -1;
}

double Worksheet::getRowHeight(int row) const {
    auto it = row_info_.find(row);
    if (it != row_info_.end() && it->second.height > 0) {
        return it->second.height;
    }
    return default_row_height_;
}

// getColumnFormat方法已移除，请使用FormatDescriptor架构

// getRowFormat方法已移除，请使用FormatDescriptor架构

bool Worksheet::isColumnHidden(int col) const {
    auto it = column_info_.find(col);
    return it != column_info_.end() && it->second.hidden;
}

bool Worksheet::isRowHidden(int row) const {
    auto it = row_info_.find(row);
    return it != row_info_.end() && it->second.hidden;
}

AutoFilterRange Worksheet::getAutoFilterRange() const {
    if (autofilter_) {
        return *autofilter_;
    }
    return AutoFilterRange(0, 0, 0, 0);
}

FreezePanes Worksheet::getFreezeInfo() const {
    if (freeze_panes_) {
        return *freeze_panes_;
    }
    return FreezePanes();
}

AutoFilterRange Worksheet::getPrintArea() const {
    return AutoFilterRange(
        print_settings_.print_area_first_row,
        print_settings_.print_area_first_col,
        print_settings_.print_area_last_row,
        print_settings_.print_area_last_col
    );
}

std::pair<int, int> Worksheet::getRepeatRows() const {
    return {print_settings_.repeat_rows_first, print_settings_.repeat_rows_last};
}

std::pair<int, int> Worksheet::getRepeatColumns() const {
    return {print_settings_.repeat_cols_first, print_settings_.repeat_cols_last};
}

Worksheet::Margins Worksheet::getMargins() const {
    return {
        print_settings_.left_margin,
        print_settings_.right_margin,
        print_settings_.top_margin,
        print_settings_.bottom_margin
    };
}

std::pair<int, int> Worksheet::getFitToPages() const {
    return {print_settings_.fit_to_pages_wide, print_settings_.fit_to_pages_tall};
}

// ========== XML生成 ==========

void Worksheet::generateXML(const std::function<void(const char*, size_t)>& callback) const {
    // 使用独立的WorksheetXMLGenerator生成XML
    auto generator = xml::WorksheetXMLGeneratorFactory::create(this);
    generator->generate(callback);
}

void Worksheet::generateXMLBatch(const std::function<void(const char*, size_t)>& callback) const {
    // 委托给WorksheetXMLGenerator
    auto generator = xml::WorksheetXMLGeneratorFactory::createBatch(this);
    generator->generate(callback);
}

void Worksheet::generateXMLStreaming(const std::function<void(const char*, size_t)>& callback) const {
    // 委托给WorksheetXMLGenerator
    auto generator = xml::WorksheetXMLGeneratorFactory::createStreaming(this);
    generator->generate(callback);
}

void Worksheet::generateRelsXML(const std::function<void(const char*, size_t)>& callback) const {
    // 关键修复：只有在有超链接时才生成关系XML
    bool has_hyperlinks = false;
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            has_hyperlinks = true;
            break;
        }
    }
    
    // 如果没有超链接，不生成任何内容
    if (!has_hyperlinks) {
        return;
    }
    
    // 生成工作表关系XML（如果有超链接等）
    xml::XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    int rel_id = 1;
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            writer.startElement("Relationship");
            writer.writeAttribute("Id", ("rId" + std::to_string(rel_id)).c_str());
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
            writer.writeAttribute("Target", cell.getHyperlink().c_str());
            writer.writeAttribute("TargetMode", "External");
            writer.endElement(); // Relationship
            rel_id++;
        }
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
}

void Worksheet::generateRelsXMLToFile(const std::string& filename) const {
    // 关键修复：只有在有超链接时才生成关系XML文件
    bool has_hyperlinks = false;
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            has_hyperlinks = true;
            break;
        }
    }
    
    // 如果没有超链接，不生成文件
    if (!has_hyperlinks) {
        return;
    }
    
    // 生成工作表关系XML（如果有超链接等）
    xml::XMLStreamWriter writer(filename);
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    int rel_id = 1;
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            writer.startElement("Relationship");
            writer.writeAttribute("Id", ("rId" + std::to_string(rel_id)).c_str());
            writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
            writer.writeAttribute("Target", cell.getHyperlink().c_str());
            writer.writeAttribute("TargetMode", "External");
            writer.endElement(); // Relationship
            rel_id++;
        }
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
}

// ========== 工具方法 ==========

void Worksheet::clear() {
    cells_.clear();
    column_info_.clear();
    row_info_.clear();
    merge_ranges_.clear();
    autofilter_.reset();
    freeze_panes_.reset();
    print_settings_ = PrintSettings{};
    sheet_view_ = SheetView{};
    protected_ = false;
    protection_password_.clear();
    selection_ = "A1";
    active_cell_ = "A1";
}

void Worksheet::clearRange(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    
    for (int row = first_row; row <= last_row; ++row) {
        for (int col = first_col; col <= last_col; ++col) {
            auto it = cells_.find(std::make_pair(row, col));
            if (it != cells_.end()) {
                cells_.erase(it);
            }
        }
    }
}

void Worksheet::insertRows(int row, int count) {
    validateCellPosition(row, 0);
    shiftCellsForRowInsertion(row, count);
}

void Worksheet::insertColumns(int col, int count) {
    validateCellPosition(0, col);
    shiftCellsForColumnInsertion(col, count);
}

void Worksheet::deleteRows(int row, int count) {
    validateCellPosition(row, 0);
    shiftCellsForRowDeletion(row, count);
}

void Worksheet::deleteColumns(int col, int count) {
    validateCellPosition(0, col);
    shiftCellsForColumnDeletion(col, count);
}

// ========== 内部辅助方法 ==========


void Worksheet::validateCellPosition(int row, int col) const {
    FASTEXCEL_VALIDATE_CELL_POSITION(row, col);
}

void Worksheet::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col);
}
// ========== 内部状态管理 ==========

void Worksheet::updateUsedRange(int row, int col) {
    range_manager_.updateRange(row, col);
}

void Worksheet::shiftCellsForRowInsertion(int row, int count) {
    std::map<std::pair<int, int>, Cell> new_cells;
    
    for (auto& [pos, cell] : cells_) {
        if (pos.first >= row) {
            // 向下移动
            new_cells[{pos.first + count, pos.second}] = std::move(cell);
        } else {
            new_cells[pos] = std::move(cell);
        }
    }
    
    cells_ = std::move(new_cells);
    
    // 更新合并单元格
    for (auto& range : merge_ranges_) {
        if (range.first_row >= row) {
            range.first_row += count;
        }
        if (range.last_row >= row) {
            range.last_row += count;
        }
    }
}

void Worksheet::shiftCellsForColumnInsertion(int col, int count) {
    std::map<std::pair<int, int>, Cell> new_cells;
    
    for (auto& [pos, cell] : cells_) {
        if (pos.second >= col) {
            // 向右移动
            new_cells[{pos.first, pos.second + count}] = std::move(cell);
        } else {
            new_cells[pos] = std::move(cell);
        }
    }
    
    cells_ = std::move(new_cells);
    
    // 更新合并单元格
    for (auto& range : merge_ranges_) {
        if (range.first_col >= col) {
            range.first_col += count;
        }
        if (range.last_col >= col) {
            range.last_col += count;
        }
    }
}

void Worksheet::shiftCellsForRowDeletion(int row, int count) {
    std::map<std::pair<int, int>, Cell> new_cells;
    
    for (auto& [pos, cell] : cells_) {
        if (pos.first >= row + count) {
            // 向上移动
            new_cells[{pos.first - count, pos.second}] = std::move(cell);
        } else if (pos.first < row) {
            new_cells[pos] = std::move(cell);
        }
        // 删除范围内的单元格被忽略
    }
    
    cells_ = std::move(new_cells);
    
    // 更新合并单元格
    auto it = merge_ranges_.begin();
    while (it != merge_ranges_.end()) {
        if (it->last_row < row) {
            // 在删除范围之前，不变
            ++it;
        } else if (it->first_row >= row + count) {
            // 在删除范围之后，向上移动
            it->first_row -= count;
            it->last_row -= count;
            ++it;
        } else {
            // 与删除范围重叠，删除合并单元格
            it = merge_ranges_.erase(it);
        }
    }
}

void Worksheet::shiftCellsForColumnDeletion(int col, int count) {
    std::map<std::pair<int, int>, Cell> new_cells;
    
    for (auto& [pos, cell] : cells_) {
        if (pos.second >= col + count) {
            // 向左移动
            new_cells[{pos.first, pos.second - count}] = std::move(cell);
        } else if (pos.second < col) {
            new_cells[pos] = std::move(cell);
        }
        // 删除范围内的单元格被忽略
    }
    
    cells_ = std::move(new_cells);
    
    // 更新合并单元格
    auto it = merge_ranges_.begin();
    while (it != merge_ranges_.end()) {
        if (it->last_col < col) {
            // 在删除范围之前，不变
            ++it;
        } else if (it->first_col >= col + count) {
            // 在删除范围之后，向左移动
            it->first_col -= count;
            it->last_col -= count;
            ++it;
        } else {
            // 与删除范围重叠，删除合并单元格
            it = merge_ranges_.erase(it);
        }
    }
}

// ========== 优化功能实现 ==========

void Worksheet::setOptimizeMode(bool enable) {
    if (optimize_mode_ == enable) {
        return;  // 状态未改变
    }
    
    if (optimize_mode_ && !enable) {
        // 从优化模式切换到标准模式
        flushCurrentRow();
        current_row_.reset();
        row_buffer_.clear();
    } else if (!optimize_mode_ && enable) {
        // 从标准模式切换到优化模式
        row_buffer_.reserve(16384);  // Excel最大列数
    }
    
    optimize_mode_ = enable;
}

void Worksheet::flushCurrentRow() {
    if (!optimize_mode_ || !current_row_ || !current_row_->data_changed) {
        return;
    }
    
    // 将当前行数据移动到主存储中
    int row_num = current_row_->row_num;
    for (auto& [col, cell] : current_row_->cells) {
        cells_[std::make_pair(row_num, col)] = std::move(cell);
    }
    
    // 更新行信息
    if (current_row_->height > 0 || current_row_->hidden) {
        RowInfo& row_info = row_info_[row_num];
        if (current_row_->height > 0) {
            row_info.height = current_row_->height;
        }
        // format字段已移除，请使用FormatDescriptor架构
        if (current_row_->hidden) {
            row_info.hidden = current_row_->hidden;
        }
    }
    
    // 重置当前行
    current_row_.reset();
}

size_t Worksheet::getMemoryUsage() const {
    size_t usage = sizeof(Worksheet);
    
    // 单元格内存
    for (const auto& [pos, cell] : cells_) {
        usage += sizeof(std::pair<std::pair<int, int>, Cell>);
        usage += cell.getMemoryUsage();
    }
    
    // 当前行内存（优化模式）
    if (current_row_) {
        usage += sizeof(WorksheetRow);
        usage += current_row_->cells.size() * sizeof(std::pair<int, Cell>);
        for (const auto& [col, cell] : current_row_->cells) {
            usage += cell.getMemoryUsage();
        }
    }
    
    // 行缓冲区内存
    usage += row_buffer_.capacity() * sizeof(Cell);
    
    // 行列信息内存
    usage += column_info_.size() * sizeof(std::pair<int, ColumnInfo>);
    usage += row_info_.size() * sizeof(std::pair<int, RowInfo>);
    
    // 合并单元格内存
    usage += merge_ranges_.size() * sizeof(MergeRange);
    
    return usage;
}

Worksheet::PerformanceStats Worksheet::getPerformanceStats() const {
    PerformanceStats stats;
    stats.total_cells = getCellCount();
    stats.memory_usage = getMemoryUsage();
    
    if (sst_) {
        stats.sst_strings = sst_->getStringCount();
        auto compression_stats = sst_->getCompressionStats();
        stats.sst_compression_ratio = compression_stats.compression_ratio;
    } else {
        stats.sst_strings = 0;
        stats.sst_compression_ratio = 0.0;
    }
    
    if (format_repo_) {
        stats.unique_formats = format_repo_->getFormatCount();
        auto dedup_stats = format_repo_->getDeduplicationStats();
        stats.format_deduplication_ratio = dedup_stats.deduplication_ratio;
    } else {
        stats.unique_formats = 0;
        stats.format_deduplication_ratio = 0.0;
    }
    
    return stats;
}

void Worksheet::ensureCurrentRow(int row_num) {
    if (!current_row_ || current_row_->row_num != row_num) {
        switchToNewRow(row_num);
    }
}

void Worksheet::switchToNewRow(int row_num) {
    // 刷新当前行
    flushCurrentRow();
    
    // 创建新的当前行
    current_row_ = std::make_unique<WorksheetRow>(row_num);
}

void Worksheet::writeOptimizedCell(int row, int col, Cell&& cell) {
    this->updateUsedRangeOptimized(row, col);
    
    this->ensureCurrentRow(row);
    current_row_->cells[col] = std::move(cell);
    current_row_->data_changed = true;
}

void Worksheet::updateUsedRangeOptimized(int row, int col) {
    range_manager_.updateRange(row, col);
}

// ========== 单元格编辑功能实现 ==========

// 私有辅助方法：通用的单元格编辑逻辑
template<typename T>
void Worksheet::editCellValueImpl(int row, int col, T&& value, bool preserve_format) {
    validateCellPosition(row, col);
    
    auto& cell = getCell(row, col);
    auto old_format = preserve_format ? cell.getFormatDescriptor() : nullptr;
    
    cell.setValue(std::forward<T>(value));
    
    if (preserve_format && old_format) {
        cell.setFormat(old_format);
    }
    
    updateUsedRange(row, col);
}

void Worksheet::editCellValue(int row, int col, const std::string& value, bool preserve_format) {
    editCellValueImpl(row, col, value, preserve_format);
}

void Worksheet::editCellValue(int row, int col, double value, bool preserve_format) {
    editCellValueImpl(row, col, value, preserve_format);
}

void Worksheet::editCellValue(int row, int col, bool value, bool preserve_format) {
    editCellValueImpl(row, col, value, preserve_format);
}

// editCellFormat方法已移除，请使用FormatDescriptor架构

void Worksheet::copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format, bool copy_row_height) {
    validateCellPosition(src_row, src_col);
    validateCellPosition(dst_row, dst_col);
    
    const auto& src_cell = getCell(src_row, src_col);
    if (src_cell.isEmpty()) {
        return; // 源单元格为空，无需复制
    }
    
    auto& dst_cell = getCell(dst_row, dst_col);
    
    // 复制值
    if (src_cell.isString()) {
        dst_cell.setValue(src_cell.getStringValue());
    } else if (src_cell.isNumber()) {
        dst_cell.setValue(src_cell.getNumberValue());
    } else if (src_cell.isBoolean()) {
        dst_cell.setValue(src_cell.getBooleanValue());
    } else if (src_cell.isFormula()) {
        dst_cell.setFormula(src_cell.getFormula(), src_cell.getFormulaResult());
    }
    
    // 复制格式
    if (copy_format && src_cell.hasFormat()) {
        dst_cell.setFormat(src_cell.getFormatDescriptor());
    }
    
    // 复制超链接
    if (src_cell.hasHyperlink()) {
        dst_cell.setHyperlink(src_cell.getHyperlink());
    }
    
    // 🔧 新增功能：复制行高
    if (copy_row_height && src_row != dst_row) {
        double src_row_height = getRowHeight(src_row);
        if (src_row_height != getRowHeight(dst_row)) {
            setRowHeight(dst_row, src_row_height);
        }
    }
    
    updateUsedRange(dst_row, dst_col);
}

void Worksheet::moveCell(int src_row, int src_col, int dst_row, int dst_col) {
    validateCellPosition(src_row, src_col);
    validateCellPosition(dst_row, dst_col);
    
    if (src_row == dst_row && src_col == dst_col) {
        return; // 源和目标相同，无需移动
    }
    
    // 复制单元格（包括行高）
    copyCell(src_row, src_col, dst_row, dst_col, true, true);
    
    // 清空源单元格
    auto it = cells_.find(std::make_pair(src_row, src_col));
    if (it != cells_.end()) {
        cells_.erase(it);
    }
}

void Worksheet::copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                         int dst_row, int dst_col, bool copy_format) {
    validateRange(src_first_row, src_first_col, src_last_row, src_last_col);
    
    int rows = src_last_row - src_first_row + 1;
    int cols = src_last_col - src_first_col + 1;
    
    // 检查目标范围是否有效
    validateCellPosition(dst_row + rows - 1, dst_col + cols - 1);
    
    for (int r = 0; r < rows; ++r) {
        for (int c = 0; c < cols; ++c) {
            // 🔧 优化：对于范围复制，默认复制行高（智能判断）
            bool copy_row_height = (c == 0); // 只在每行的第一列复制行高
            copyCell(src_first_row + r, src_first_col + c,
                    dst_row + r, dst_col + c, copy_format, copy_row_height);
        }
    }
}

void Worksheet::moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                         int dst_row, int dst_col) {
    validateRange(src_first_row, src_first_col, src_last_row, src_last_col);
    
    int rows = src_last_row - src_first_row + 1;
    int cols = src_last_col - src_first_col + 1;
    
    // 检查目标范围是否有效
    validateCellPosition(dst_row + rows - 1, dst_col + cols - 1);
    
    // 检查源和目标范围是否重叠
    bool overlaps = !(dst_row + rows <= src_first_row || dst_row >= src_last_row + 1 ||
                     dst_col + cols <= src_first_col || dst_col >= src_last_col + 1);
    
    if (overlaps) {
        // 如果重叠，需要使用临时存储
        std::map<std::pair<int, int>, Cell> temp_cells;
        
        // 先复制到临时存储
        for (int r = 0; r < rows; ++r) {
            for (int c = 0; c < cols; ++c) {
                int src_r = src_first_row + r;
                int src_c = src_first_col + c;
                auto it = cells_.find(std::make_pair(src_r, src_c));
                if (it != cells_.end()) {
                    temp_cells[std::make_pair(r, c)] = std::move(it->second);
                    cells_.erase(it);
                }
            }
        }
        
        // 从临时存储移动到目标位置
        for (const auto& [temp_pos, cell] : temp_cells) {
            int dst_r = dst_row + temp_pos.first;
            int dst_c = dst_col + temp_pos.second;
            cells_[std::make_pair(dst_r, dst_c)] = std::move(const_cast<Cell&>(cell));
            updateUsedRange(dst_r, dst_c);
        }
    } else {
        // 不重叠，直接移动
        copyRange(src_first_row, src_first_col, src_last_row, src_last_col, dst_row, dst_col, true);
        clearRange(src_first_row, src_first_col, src_last_row, src_last_col);
    }
}

int Worksheet::findAndReplace(const std::string& find_text, const std::string& replace_text,
                             bool match_case, bool match_entire_cell) {
    int replace_count = 0;
    
    for (auto& [pos, cell] : cells_) {
        if (!cell.isString()) {
            continue; // 只处理字符串单元格
        }
        
        std::string cell_text = cell.getStringValue();
        std::string search_text = find_text;
        std::string target_text = cell_text;
        
        // 处理大小写敏感性
        if (!match_case) {
            std::transform(search_text.begin(), search_text.end(), search_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
            std::transform(target_text.begin(), target_text.end(), target_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
        }
        
        if (match_entire_cell) {
            // 匹配整个单元格
            if (target_text == search_text) {
                cell.setValue(replace_text);
                replace_count++;
            }
        } else {
            // 部分匹配
            size_t pos_found = target_text.find(search_text);
            if (pos_found != std::string::npos) {
                // 在原始文本中进行替换
                std::string new_text = cell_text;
                size_t actual_pos = pos_found;
                
                // 如果不区分大小写，需要找到原始文本中的实际位置
                if (!match_case) {
                    actual_pos = cell_text.find(find_text);
                    if (actual_pos == std::string::npos) {
                        // 尝试不区分大小写的查找
                        for (size_t i = 0; i <= cell_text.length() - find_text.length(); ++i) {
                            bool match = true;
                            for (size_t j = 0; j < find_text.length(); ++j) {
                                if (std::tolower(static_cast<unsigned char>(cell_text[i + j])) !=
                                    std::tolower(static_cast<unsigned char>(find_text[j]))) {
                                    match = false;
                                    break;
                                }
                            }
                            if (match) {
                                actual_pos = i;
                                break;
                            }
                        }
                    }
                }
                
                if (actual_pos != std::string::npos) {
                    new_text.replace(actual_pos, find_text.length(), replace_text);
                    cell.setValue(new_text);
                    replace_count++;
                }
            }
        }
    }
    
    return replace_count;
}

std::vector<std::pair<int, int>> Worksheet::findCells(const std::string& search_text,
                                                      bool match_case,
                                                      bool match_entire_cell) const {
    std::vector<std::pair<int, int>> results;
    
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isString()) {
            continue; // 只搜索字符串单元格
        }
        
        std::string cell_text = cell.getStringValue();
        std::string target_text = cell_text;
        std::string find_text = search_text;
        
        // 处理大小写敏感性
        if (!match_case) {
            std::transform(find_text.begin(), find_text.end(), find_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
            std::transform(target_text.begin(), target_text.end(), target_text.begin(),
                         [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
        }
        
        bool found = false;
        if (match_entire_cell) {
            found = (target_text == find_text);
        } else {
            found = (target_text.find(find_text) != std::string::npos);
        }
        
        if (found) {
            results.push_back(pos);
        }
    }
    
    return results;
}

void Worksheet::sortRange(int first_row, int first_col, int last_row, int last_col,
                         int sort_column, bool ascending, bool has_header) {
    validateRange(first_row, first_col, last_row, last_col);
    
    int data_start_row = has_header ? first_row + 1 : first_row;
    if (data_start_row > last_row) {
        return; // 没有数据行需要排序
    }
    
    int sort_col = first_col + sort_column;
    if (sort_col > last_col) {
        FASTEXCEL_THROW_PARAM("Sort column is outside the range");
    }
    
    // 收集需要排序的行数据
    std::vector<std::pair<int, std::map<int, Cell>>> rows_data;
    
    for (int row = data_start_row; row <= last_row; ++row) {
        std::map<int, Cell> row_cells;
        for (int col = first_col; col <= last_col; ++col) {
            auto it = cells_.find(std::make_pair(row, col));
            if (it != cells_.end()) {
                row_cells[col] = std::move(it->second);
                cells_.erase(it);
            }
        }
        rows_data.emplace_back(row, std::move(row_cells));
    }
    
    // 排序
    std::sort(rows_data.begin(), rows_data.end(),
        [sort_col, ascending](const auto& a, const auto& b) {
            const auto& a_cells = a.second;
            const auto& b_cells = b.second;
            
            auto a_it = a_cells.find(sort_col);
            auto b_it = b_cells.find(sort_col);
            
            // 处理空单元格
            if (a_it == a_cells.end() && b_it == b_cells.end()) {
                return false; // 两个都为空，认为相等
            }
            if (a_it == a_cells.end()) {
                return ascending; // 空单元格排在后面（升序）或前面（降序）
            }
            if (b_it == b_cells.end()) {
                return !ascending;
            }
            
            const Cell& a_cell = a_it->second;
            const Cell& b_cell = b_it->second;
            
            // 比较单元格值
            if (a_cell.isNumber() && b_cell.isNumber()) {
                double a_val = a_cell.getNumberValue();
                double b_val = b_cell.getNumberValue();
                return ascending ? (a_val < b_val) : (a_val > b_val);
            } else if (a_cell.isString() && b_cell.isString()) {
                const std::string& a_str = a_cell.getStringValue();
                const std::string& b_str = b_cell.getStringValue();
                return ascending ? (a_str < b_str) : (a_str > b_str);
            } else {
                // 混合类型：数字 < 字符串
                if (a_cell.isNumber() && b_cell.isString()) {
                    return ascending;
                } else if (a_cell.isString() && b_cell.isNumber()) {
                    return !ascending;
                }
                return false; // 其他情况认为相等
            }
        });
    
    // 将排序后的数据放回工作表
    for (size_t i = 0; i < rows_data.size(); ++i) {
        int target_row = data_start_row + static_cast<int>(i);
        const auto& row_cells = rows_data[i].second;
        
        for (const auto& [col, cell] : row_cells) {
            cells_[std::make_pair(target_row, col)] = std::move(const_cast<Cell&>(cell));
            updateUsedRange(target_row, col);
        }
    }
}

// ========== 公式优化方法 ==========

int Worksheet::optimizeFormulas(int min_similar_count) {
    if (!shared_formula_manager_) {
        shared_formula_manager_ = std::make_unique<SharedFormulaManager>();
    }
    
    // 收集所有非共享公式
    std::map<std::pair<int, int>, std::string> formulas;
    auto [max_row, max_col] = getUsedRange();
    
    for (int row = 0; row <= max_row; ++row) {
        for (int col = 0; col <= max_col; ++col) {
            if (hasCellAt(row, col)) {
                const auto& cell = getCell(row, col);
                if (cell.isFormula() && !cell.isSharedFormula()) {
                    formulas[{row, col}] = cell.getFormula();
                }
            }
        }
    }
    
    if (formulas.empty()) {
        CORE_DEBUG("工作表中没有可优化的公式");
        return 0;
    }
    
    // 使用共享公式管理器进行优化
    int optimized_count = shared_formula_manager_->optimizeFormulas(formulas, min_similar_count);
    
    if (optimized_count > 0) {
        CORE_DEBUG("成功优化 {} 个公式为共享公式", optimized_count);
        
        // 🔧 关键修复：将对应的Cell对象标记为共享公式
        // 获取所有共享公式索引，并更新对应的单元格
        auto shared_indices = shared_formula_manager_->getAllSharedIndices();
        for (int shared_index : shared_indices) {
            const SharedFormula* shared_formula = shared_formula_manager_->getSharedFormula(shared_index);
            if (shared_formula) {
                const auto& affected_cells = shared_formula->getAffectedCells();
                for (const auto& [row, col] : affected_cells) {
                    if (hasCellAt(row, col)) {
                        Cell& cell = getCell(row, col);
                        if (cell.isFormula() && !cell.isSharedFormula()) {
                            // 将普通公式转换为共享公式引用
                            cell.setSharedFormula(shared_index);
                        }
                    }
                }
            }
        }
        
        // 标记为修改状态
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
        }
    }
    
    return optimized_count;
}

Worksheet::FormulaOptimizationReport Worksheet::analyzeFormulaOptimization() const {
    FormulaOptimizationReport report;
    
    // 收集所有公式
    std::map<std::pair<int, int>, std::string> formulas;
    auto [max_row, max_col] = getUsedRange();
    
    for (int row = 0; row <= max_row; ++row) {
        for (int col = 0; col <= max_col; ++col) {
            if (hasCellAt(row, col)) {
                const auto& cell = getCell(row, col);
                if (cell.isFormula()) {
                    formulas[{row, col}] = cell.getFormula();
                }
            }
        }
    }
    
    report.total_formulas = formulas.size();
    
    if (formulas.empty()) {
        return report;
    }
    
    // 使用临时的共享公式管理器进行分析
    SharedFormulaManager temp_manager;
    auto patterns = temp_manager.detectSharedFormulaPatterns(formulas);
    
    // 分析优化潜力
    size_t optimizable_count = 0;
    size_t estimated_savings = 0;
    
    for (const auto& pattern : patterns) {
        if (pattern.matching_cells.size() >= 3) { // 至少3个相似公式
            optimizable_count += pattern.matching_cells.size();
            estimated_savings += pattern.estimated_savings;
            
            // 添加模式示例（限制最多5个）
            if (report.pattern_examples.size() < 5) {
                std::string example = "模式: " + std::to_string(pattern.matching_cells.size()) + 
                    " 个相似公式，预估节省 " + std::to_string(pattern.estimated_savings) + " 字节";
                
                // 添加具体公式示例
                if (!pattern.matching_cells.empty()) {
                    auto first_pos = pattern.matching_cells[0];
                    auto formula_it = formulas.find(first_pos);
                    if (formula_it != formulas.end()) {
                        std::string cell_ref = utils::CommonUtils::cellReference(first_pos.first, first_pos.second);
                        example += " (示例: " + cell_ref + " = " + formula_it->second + ")";
                    }
                }
                report.pattern_examples.push_back(example);
            }
        }
    }
    
    report.optimizable_formulas = optimizable_count;
    report.estimated_memory_savings = estimated_savings;
    
    if (report.total_formulas > 0) {
        report.optimization_ratio = static_cast<double>(optimizable_count) / report.total_formulas * 100.0;
    }
    
    return report;
}

// 🚀 新API：链式调用方法实现
WorksheetChain Worksheet::chain() {
    return WorksheetChain(*this);
}

} // namespace core
} // namespace fastexcel
