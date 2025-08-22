#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/WorksheetChain.hpp"
#include "fastexcel/core/RangeFormatter.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/CellRangeManager.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include <iomanip>
#include <sstream>
#include <cmath>
#include "fastexcel/core/Image.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/WorksheetXMLGenerator.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/AddressParser.hpp"
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
    
    // 初始化单元格数据处理器
    cell_processor_ = std::make_unique<CellDataProcessor>(cells_, range_manager_, parent_workbook_, sheet_id_);
    
    // 初始化布局管理器
    layout_manager_ = std::make_unique<WorksheetLayoutManager>();
    
    // 初始化图片管理器
    image_manager_ = std::make_unique<WorksheetImageManager>();
    
    // 初始化CSV处理器
    csv_handler_ = std::make_unique<WorksheetCSVHandler>(*this);
    
    // 列宽管理器延迟初始化，待 setFormatRepository 设置
}

// 基本单元格操作

Cell& Worksheet::getCell(int row, int col) {
    return cell_processor_->getCell(row, col);
}

const Cell& Worksheet::getCell(int row, int col) const {
    return cell_processor_->getCell(row, col);
}

// 基本写入方法

void Worksheet::writeDateTime(int row, int col, const std::tm& datetime) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
    }
    this->validateCellPosition(row, col);
    
    // 使用 TimeUtils 将日期时间转换为Excel序列号
    double excel_serial = utils::TimeUtils::toExcelSerialNumber(datetime);
    this->setValue(row, col, excel_serial);
}

void Worksheet::writeUrl(int row, int col, const std::string& url, const std::string& string) {
    cell_processor_->setHyperlink(row, col, url, string);
}


// 智能列宽管理方法

std::pair<double, int> Worksheet::setColumnWidthAdvanced(int col, double target_width,
                                                        const std::string& font_name,
                                                        double font_size,
                                                        ColumnWidthManager::WidthStrategy strategy,
                                                        const std::vector<std::string>& cell_contents) {
    validateCellPosition(0, col);
    
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    auto result = layout_manager_->setColumnWidthAdvanced(col, target_width, font_name, font_size, strategy, cell_contents);
    
    // 同步更新本地column_info_（保持兼容性）
    column_info_[col].width = result.first;
    column_info_[col].precise_width = true;
    if (result.second >= 0) {
        column_info_[col].format_id = result.second;
    }
    
    return result;
}

std::unordered_map<int, std::pair<double, int>> Worksheet::setColumnWidthsBatch(
    const std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig>& configs) {
    
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    // 验证所有列位置
    for (const auto& [col, config] : configs) {
        validateCellPosition(0, col);
    }
    
    auto results = layout_manager_->setColumnWidthsBatch(configs);
    
    // 同步更新本地column_info_（保持兼容性）
    for (const auto& [col, result] : results) {
        column_info_[col].width = result.first;
        column_info_[col].precise_width = true;
        if (result.second >= 0) {
            column_info_[col].format_id = result.second;
        }
    }
    
    return results;
}

double Worksheet::calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const {
    return layout_manager_->calculateOptimalWidth(target_width, font_name, font_size);
}

// 简化版列宽设置方法

double Worksheet::setColumnWidth(int col, double width) {
    validateCellPosition(0, col);
    
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    double actual_width = layout_manager_->setColumnWidth(col, width);
    
    // 同步更新本地column_info_（保持兼容性）
    column_info_[col].width = actual_width;
    column_info_[col].precise_width = true;
    
    FASTEXCEL_LOG_DEBUG("设置列{}宽度: {} -> {}", col, width, actual_width);
    
    return actual_width;
}

std::pair<double, int> Worksheet::setColumnWidthWithFont(int col, double width, 
                                                         const std::string& font_name, 
                                                         double font_size) {
    // 调用 setColumnWidthAdvanced，使用EXACT策略
    return setColumnWidthAdvanced(col, width, font_name, font_size, 
                                 ColumnWidthManager::WidthStrategy::EXACT, {});
}

// 行列操作

// 旧的列宽方法已移除，请使用 ColumnWidthManager 架构：
// - setColumnWidth() 单列智能设置
// - setColumnWidthsBatch() 批量高性能设置
// - calculateOptimalWidth() 预计算

// 获取工作簿默认字体信息的辅助方法
std::string Worksheet::getWorkbookDefaultFont() const {
    if (parent_workbook_) {
        // 尝试从工作簿的格式仓储获取默认字体
        if (format_repo_) {
            const auto& default_format = core::FormatDescriptor::getDefault();
            return default_format.getFontName();
        }
    }
    return "Calibri";  // 默认字体
}

double Worksheet::getWorkbookDefaultFontSize() const {
    if (parent_workbook_) {
        // 尝试从工作簿的格式仓储获取默认字体大小
        if (format_repo_) {
            const auto& default_format = core::FormatDescriptor::getDefault();
            return default_format.getFontSize();
        }
    }
    return 11.0;  // 默认字体大小
}

void Worksheet::setColumnFormatId(int col, int format_id) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->setColumnFormatId(col, format_id);
    
    // 同步更新本地column_info_（保持兼容性）
    validateCellPosition(0, col);
    column_info_[col].format_id = format_id;
}

void Worksheet::setColumnFormatId(int first_col, int last_col, int format_id) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->setColumnFormatId(first_col, last_col, format_id);
    
    // 同步更新本地column_info_（保持兼容性）
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].format_id = format_id;
    }
}

// setColumnFormat方法已移除（请使用FormatDescriptor架构）

// setColumnFormat范围方法已移除（请使用FormatDescriptor架构）

void Worksheet::hideColumn(int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->hideColumn(col);
    
    // 同步更新本地column_info_（保持兼容性）
    validateCellPosition(0, col);
    column_info_[col].hidden = true;
}

void Worksheet::hideColumn(int first_col, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->hideColumn(first_col, last_col);
    
    // 同步更新本地column_info_（保持兼容性）
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
    
    layout_manager_->setRowHeight(row, height);
    
    // 同步更新本地row_info_（保持兼容性）
    validateCellPosition(row, 0);
    row_info_[row].height = height;
}

// setRowFormat方法已移除（请使用FormatDescriptor架构）

void Worksheet::hideRow(int row) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->hideRow(row);
    
    // 同步更新本地row_info_（保持兼容性）
    validateCellPosition(row, 0);
    row_info_[row].hidden = true;
}

void Worksheet::hideRow(int first_row, int last_row) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->hideRow(first_row, last_row);
    
    // 同步更新本地row_info_（保持兼容性）
    validateRange(first_row, 0, last_row, 0);
    for (int row = first_row; row <= last_row; ++row) {
        row_info_[row].hidden = true;
    }
}

// 合并单元格

void Worksheet::mergeCells(int first_row, int first_col, int last_row, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->mergeCells(first_row, first_col, last_row, last_col);
    
    // 同步更新本地merge_ranges_（保持兼容性）
    validateRange(first_row, first_col, last_row, last_col);
    merge_ranges_.emplace_back(first_row, first_col, last_row, last_col);
}

// 自动筛选

void Worksheet::setAutoFilter(int first_row, int first_col, int last_row, int last_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->setAutoFilter(first_row, first_col, last_row, last_col);
    
    // 同步更新本地autofilter_（保持兼容性）
    validateRange(first_row, first_col, last_row, last_col);
    autofilter_ = std::make_unique<AutoFilterRange>(first_row, first_col, last_row, last_col);
}

void Worksheet::removeAutoFilter() {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->removeAutoFilter();
    
    // 同步更新本地autofilter_（保持兼容性）
    autofilter_.reset();
}

// 冻结窗格

void Worksheet::freezePanes(int row, int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->freezePanes(row, col);
    
    // 同步更新本地freeze_panes_（保持兼容性）
    validateCellPosition(row, col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

void Worksheet::freezePanes(int row, int col, int top_left_row, int top_left_col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->freezePanes(row, col, top_left_row, top_left_col);
    
    // 同步更新本地freeze_panes_（保持兼容性）
    validateCellPosition(row, col);
    validateCellPosition(top_left_row, top_left_col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col, top_left_row, top_left_col);
}

void Worksheet::splitPanes(int row, int col) {
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
    
    layout_manager_->splitPanes(row, col);
    
    // 同步更新本地freeze_panes_（保持兼容性）
    validateCellPosition(row, col);
    // 分割窗格的实现与冻结窗格类似，但使用不同的XML属性
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

// 工作表保护

void Worksheet::protect(const std::string& password) {
    protected_ = true;
    protection_password_ = password;
}

void Worksheet::unprotect() {
    protected_ = false;
    protection_password_.clear();
}

// 视图设置

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

// 获取信息

std::pair<int, int> Worksheet::getUsedRange() const {
    return cell_processor_->getUsedRange();
}

std::tuple<int, int, int, int> Worksheet::getUsedRangeFull() const {
    return cell_processor_->getUsedRangeFull();
}

bool Worksheet::hasCellAt(int row, int col) const {
    return cell_processor_->hasCellAt(row, col);
}

// 获取方法实现

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

// 统一样式API实现

void Worksheet::setColumnFormat(int col, const FormatDescriptor& format) {
    validateCellPosition(0, col);
    
    if (!parent_workbook_) {
        throw std::runtime_error("工作簿未初始化，无法进行智能格式优化");
    }
    
    // 🎯 核心优化：自动添加到FormatRepository（去重）
    int styleId = parent_workbook_->addStyle(format);
    
    // 设置列格式ID
    column_info_[col].format_id = styleId;
    
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
}

void Worksheet::setRowFormat(int row, const FormatDescriptor& format) {
    validateCellPosition(row, 0);
    
    if (!parent_workbook_) {
        throw std::runtime_error("工作簿未初始化，无法进行智能格式优化");
    }
    
    // 🎯 核心优化：自动添加到FormatRepository（去重）
    int styleId = parent_workbook_->addStyle(format);
    
    // 设置行格式ID
    row_info_[row].format_id = styleId;
    
    if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
        std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
        parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::METADATA);
    }
}

std::shared_ptr<const FormatDescriptor> Worksheet::getColumnFormat(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end() && it->second.format_id >= 0) {
        if (parent_workbook_) {
            return parent_workbook_->getStyle(it->second.format_id);
        }
    }
    return nullptr;
}

std::shared_ptr<const FormatDescriptor> Worksheet::getRowFormat(int row) const {
    auto it = row_info_.find(row);
    if (it != row_info_.end() && it->second.format_id >= 0) {
        if (parent_workbook_) {
            return parent_workbook_->getStyle(it->second.format_id);
        }
    }
    return nullptr;
}

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

// XML生成

void Worksheet::generateXML(const std::function<void(const char*, size_t)>& callback) const {
    // 使用独立的WorksheetXMLGenerator生成XML
    auto generator = xml::WorksheetXMLGeneratorFactory::create(this);
    generator->generate(callback);
}

void Worksheet::generateXMLBatch(const std::function<void(const char*, size_t)>& callback) const {
    auto generator = xml::WorksheetXMLGeneratorFactory::createBatch(this);
    generator->generate(callback);
}

void Worksheet::generateXMLStreaming(const std::function<void(const char*, size_t)>& callback) const {
    auto generator = xml::WorksheetXMLGeneratorFactory::createStreaming(this);
    generator->generate(callback);
}

void Worksheet::generateRelsXML(const std::function<void(const char*, size_t)>& callback) const {
    // 使用专用的Relationships类生成XML
    xml::Relationships relationships;
    
    // 添加超链接关系
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            relationships.addAutoRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                cell.getHyperlink(),
                "External"
            );
        }
    }
    
    if (!images_.empty()) {
        std::string drawing_target = "../drawings/drawing" + std::to_string(sheet_id_) + ".xml";
        relationships.addAutoRelationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            drawing_target
        );
    }
    
    // 如果没有关系，不生成任何内容
    if (relationships.size() == 0) {
        return;
    }
    
    // 使用Relationships类生成XML到回调
    relationships.generate(callback);
}

void Worksheet::generateRelsXMLToFile(const std::string& filename) const {
    // 使用专用的Relationships类生成XML
    xml::Relationships relationships;
    
    // 添加超链接关系
    for (const auto& [pos, cell] : cells_) {
        if (cell.hasHyperlink()) {
            relationships.addAutoRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                cell.getHyperlink(),
                "External"
            );
        }
    }
    
    // 如果没有关系，不生成文件
    if (relationships.size() == 0) {
        return;
    }
    
    // 使用Relationships类生成XML文件
    relationships.generateToFile(filename);
}

// 工具方法

void Worksheet::clear() {
    cells_.clear();
    column_info_.clear();
    row_info_.clear();
    merge_ranges_.clear();
    autofilter_.reset();
    freeze_panes_.reset();
    sheet_view_ = SheetView{};
    protected_ = false;
    protection_password_.clear();
    selection_ = "A1";
    active_cell_ = "A1";
    
    // 清空图片
    images_.clear();
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

// 内部辅助方法


void Worksheet::validateCellPosition(int row, int col) const {
    FASTEXCEL_VALIDATE_CELL_POSITION(row, col);
}

void Worksheet::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col);
}
// 内部状态管理

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

// 优化功能实现

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
        // 行格式由 FormatDescriptor 管理
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

// 单元格编辑功能实现

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

// editCellFormat方法已移除（请使用FormatDescriptor架构）

// 智能单元格格式设置方法

void Worksheet::setCellFormat(int row, int col, const core::FormatDescriptor& format) {
    validateCellPosition(row, col);
    
    if (!parent_workbook_) {
        throw std::runtime_error("工作簿未初始化，无法进行智能格式优化");
    }
    
    // 🎯 核心优化：自动添加到FormatRepository（去重）
    int styleId = parent_workbook_->addStyle(format);
    
    // 获取FormatRepository中优化后的格式引用
    auto optimizedFormat = parent_workbook_->getStyle(styleId);
    
    // 直接应用到指定单元格
    getCell(row, col).setFormat(optimizedFormat);
}

void Worksheet::setCellFormat(int row, int col, std::shared_ptr<const core::FormatDescriptor> format) {
    if (!format) {
        // 清除单元格格式
        getCell(row, col).setFormat(nullptr);
        return;
    }
    
    // 调用值版本
    setCellFormat(row, col, *format);
}

void Worksheet::setCellFormat(int row, int col, const core::StyleBuilder& builder) {
    validateCellPosition(row, col);
    
    if (!parent_workbook_) {
        throw std::runtime_error("工作簿未初始化，无法进行智能格式优化");
    }
    
    // 🎯 一步到位：构建、优化、应用到指定单元格
    int styleId = parent_workbook_->addStyle(builder);
    auto optimizedFormat = parent_workbook_->getStyle(styleId);
    getCell(row, col).setFormat(optimizedFormat);
}

// 范围格式化API方法

RangeFormatter Worksheet::rangeFormatter(const std::string& range) {
    return std::move(RangeFormatter(this).setRange(range));
}

RangeFormatter Worksheet::rangeFormatter(int start_row, int start_col, int end_row, int end_col) {
    return std::move(RangeFormatter(this).setRange(start_row, start_col, end_row, end_col));
}

void Worksheet::copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format, bool copy_row_height) {
    cell_processor_->copyCell(src_row, src_col, dst_row, dst_col, copy_format);
    
    if (copy_row_height && src_row != dst_row) {
        double src_row_height = getRowHeight(src_row);
        if (src_row_height != getRowHeight(dst_row)) {
            setRowHeight(dst_row, src_row_height);
        }
    }
}

void Worksheet::moveCell(int src_row, int src_col, int dst_row, int dst_col) {
    cell_processor_->moveCell(src_row, src_col, dst_row, dst_col);
}

void Worksheet::copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                         int dst_row, int dst_col, bool copy_format) {
    cell_processor_->copyRange(src_first_row, src_first_col, src_last_row, src_last_col, dst_row, dst_col, copy_format);
}

void Worksheet::moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                         int dst_row, int dst_col) {
    cell_processor_->moveRange(src_first_row, src_first_col, src_last_row, src_last_col, dst_row, dst_col);
}

int Worksheet::findAndReplace(const std::string& find_text, const std::string& replace_text,
                             bool match_case, bool match_entire_cell) {
    return cell_processor_->findAndReplace(find_text, replace_text, match_case, match_entire_cell);
}

std::vector<std::pair<int, int>> Worksheet::findCells(const std::string& search_text,
                                                      bool match_case,
                                                      bool match_entire_cell) const {
    return cell_processor_->findCells(search_text, match_case, match_entire_cell);
}

void Worksheet::sortRange(int first_row, int first_col, int last_row, int last_col,
                         int sort_column, bool ascending, bool has_header) {
    cell_processor_->sortRange(first_row, first_col, last_row, last_col, sort_column, ascending, has_header);
}

// 共享公式管理

int Worksheet::createSharedFormula(int first_row, int first_col, int last_row, int last_col, const std::string& formula) {
    if (!shared_formula_manager_) {
        shared_formula_manager_ = std::make_unique<SharedFormulaManager>();
    }
    
    // 构建范围引用字符串
    std::string range_ref = utils::CommonUtils::rangeReference(first_row, first_col, last_row, last_col);
    
    // 使用SharedFormulaManager注册共享公式
    int shared_index = shared_formula_manager_->registerSharedFormula(formula, range_ref);
    
    if (shared_index >= 0) {
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
                Cell& cell = getCell(row, col);
                if (row == first_row && col == first_col) {
                    // 主单元格存储完整的基础公式和共享公式索引
                    cell.setFormula(formula);  // 先设置常规公式
                    cell.setSharedFormula(shared_index);  // 然后转换为共享公式
                } else {
                    // 其他单元格只存储共享公式引用
                    cell.setSharedFormulaReference(shared_index);
                }
            }
        }
    }
    
    return shared_index;
}

// 便捷的公式设置方法实现
void Worksheet::setFormula(const Address& address, const std::string& formula, double result) {
    setFormula(address.getRow(), address.getCol(), formula, result);
}

void Worksheet::setFormula(int row, int col, const std::string& formula, double result) {
    Cell& cell = getCell(row, col);
    cell.setFormula(formula, result);
}

int Worksheet::createSharedFormula(const CellRange& range, const std::string& formula) {
    return createSharedFormula(range.getStartRow(), range.getStartCol(), 
                              range.getEndRow(), range.getEndCol(), formula);
}

// 公式优化方法已移除（请使用新的架构）

// 便捷的工作表状态检查方法实现
int Worksheet::getRowCount() const {
    if (cells_.empty()) {
        return 0;
    }
    
    int max_row = -1;
    for (const auto& [pos, cell] : cells_) {
        max_row = std::max(max_row, pos.first);
    }
    
    return max_row + 1; // 返回实际行数（从0开始，所以+1）
}

int Worksheet::getColumnCount() const {
    if (cells_.empty()) {
        return 0;
    }
    
    int max_col = -1;
    for (const auto& [pos, cell] : cells_) {
        max_col = std::max(max_col, pos.second);
    }
    
    return max_col + 1; // 返回实际列数（从0开始，所以+1）
}

int Worksheet::getCellCountInRow(int row) const {
    return cell_processor_->getCellCountInRow(row);
}

int Worksheet::getCellCountInColumn(int col) const {
    return cell_processor_->getCellCountInColumn(col);
}

// 注意：这些方法与头文件中的声明重复，已在头文件的实现中定义

void Worksheet::clearRow(int row) {
    cell_processor_->clearRow(row);
}

void Worksheet::clearColumn(int col) {
    cell_processor_->clearColumn(col);
}

void Worksheet::clearAll() {
    cell_processor_->clearAll();
    FASTEXCEL_LOG_DEBUG("Cleared all cells in worksheet '{}'", name_);
}

// 链式调用方法实现
WorksheetChain Worksheet::chain() {
    return WorksheetChain(*this);
}

// 图片插入功能实现

std::string Worksheet::insertImage(int row, int col, const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} at cell ({}, {})", image_path, row, col);
    
    std::string image_id = image_manager_->insertImage(row, col, image_path);
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            // 使用显式clone深拷贝，避免调用已删除的拷贝构造
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} at cell ({}, {})", image_id, row, col);
    }
    
    return image_id;
}

std::string Worksheet::insertImage(int row, int col, std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    std::string image_id = image_manager_->insertImage(row, col, std::move(image));
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} at cell position ({}, {})", image_id, row, col);
    }
    
    return image_id;
}

std::string Worksheet::insertImage(int from_row, int from_col, int to_row, int to_col,
                                  const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} in range ({},{}) to ({},{})",
                       image_path, from_row, from_col, to_row, to_col);
    
    std::string image_id = image_manager_->insertImage(from_row, from_col, to_row, to_col, image_path);
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} in range ({},{}) to ({},{})",
                          image_id, from_row, from_col, to_row, to_col);
    }
    
    return image_id;
}

std::string Worksheet::insertImage(int from_row, int from_col, int to_row, int to_col,
                                  std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    std::string image_id = image_manager_->insertImage(from_row, from_col, to_row, to_col, std::move(image));
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} in range ({},{}) to ({},{})",
                          image_id, from_row, from_col, to_row, to_col);
    }
    
    return image_id;
}

std::string Worksheet::insertImageAt(double x, double y, double width, double height,
                                    const std::string& image_path) {
    FASTEXCEL_LOG_DEBUG("Inserting image from file: {} at absolute position ({}, {}) with size {}x{}",
                       image_path, x, y, width, height);
    
    std::string image_id = image_manager_->insertImageAt(x, y, width, height, image_path);
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} at absolute position ({}, {}) with size {}x{}",
                          image_id, x, y, width, height);
    }
    
    return image_id;
}

std::string Worksheet::insertImageAt(double x, double y, double width, double height,
                                    std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    std::string image_id = image_manager_->insertImageAt(x, y, width, height, std::move(image));
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
        
        FASTEXCEL_LOG_INFO("Successfully inserted image: {} at absolute position ({}, {}) with size {}x{}",
                          image_id, x, y, width, height);
    }
    
    return image_id;
}

std::string Worksheet::insertImage(const std::string& address, const std::string& image_path) {
    std::string image_id = image_manager_->insertImage(address, image_path);
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
    }
    
    return image_id;
}

std::string Worksheet::insertImage(const std::string& address, std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    std::string image_id = image_manager_->insertImage(address, std::move(image));
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
    }
    
    return image_id;
}

std::string Worksheet::insertImageRange(const std::string& range, const std::string& image_path) {
    std::string image_id = image_manager_->insertImageRange(range, image_path);
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
    }
    
    return image_id;
}

std::string Worksheet::insertImageRange(const std::string& range, std::unique_ptr<Image> image) {
    if (!image) {
        FASTEXCEL_LOG_ERROR("Cannot insert null image");
        return "";
    }
    
    std::string image_id = image_manager_->insertImageRange(range, std::move(image));
    
    if (!image_id.empty()) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        next_image_id_ = static_cast<int>(image_manager_->getImageCount()) + 1;
    }
    
    return image_id;
}

// 图片管理功能实现

const Image* Worksheet::findImage(const std::string& image_id) const {
    return image_manager_->findImage(image_id);
}

Image* Worksheet::findImage(const std::string& image_id) {
    return image_manager_->findImage(image_id);
}

bool Worksheet::removeImage(const std::string& image_id) {
    bool result = image_manager_->removeImage(image_id);
    
    if (result) {
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        const auto& manager_images = image_manager_->getImages();
        images_.clear();
        for (const auto& img : manager_images) {
            images_.push_back(img->clone());
        }
        
        FASTEXCEL_LOG_INFO("Removed image: {}", image_id);
    } else {
        FASTEXCEL_LOG_WARN("Image not found for removal: {}", image_id);
    }
    
    return result;
}

void Worksheet::clearImages() {
    size_t count = image_manager_->getImageCount();
    
    if (count > 0) {
        image_manager_->clearImages();
        
        // 标记工作表为脏数据
        if (parent_workbook_ && parent_workbook_->getDirtyManager()) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(sheet_id_) + ".xml";
            std::string drawing_path = "xl/drawings/drawing" + std::to_string(sheet_id_) + ".xml";
            parent_workbook_->getDirtyManager()->markDirty(sheet_path, DirtyManager::DirtyLevel::CONTENT);
            parent_workbook_->getDirtyManager()->markDirty(drawing_path, DirtyManager::DirtyLevel::CONTENT);
        }
        
        // 同步更新本地images_（保持兼容性）
        images_.clear();
        next_image_id_ = 1;
        
        FASTEXCEL_LOG_INFO("Cleared {} images from worksheet", count);
    }
}

size_t Worksheet::getImagesMemoryUsage() const {
    return image_manager_->getImagesMemoryUsage();
}

// 架构优化完成状态标记

// ✅ 已完成的管理器委托架构优化：
// - CellDataProcessor: 单元格数据操作、复制移动、查找替换、排序等 (已完成)
// - WorksheetLayoutManager: 列宽行高、隐藏、合并、筛选、冻结窗格等 (已完成)
// - WorksheetImageManager: 图片插入、查找、移除等 (已完成)
// - WorksheetCSVHandler: CSV导入导出处理 (已完成)

// 架构设计说明：
// 1. 保持向后兼容性：所有原有API保持不变，内部委托给管理器
// 2. 数据同步策略：委托操作后同步更新本地数据结构（如column_info_、row_info_等）
// 3. 责任分离清晰：Worksheet专注协调，具体功能由专门管理器处理
// 4. 性能优化考虑：避免重复数据复制，使用引用和移动语义

// CSV功能实现

CSVParseInfo Worksheet::loadFromCSV(const std::string& filepath, const CSVOptions& options) {
    FASTEXCEL_LOG_INFO("Loading CSV from file: {} into worksheet: {}", filepath, name_);
    return csv_handler_->loadFromCSV(filepath, options);
}

CSVParseInfo Worksheet::loadFromCSVString(const std::string& csv_content, const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Loading CSV from string into worksheet: {}, content length: {}", name_, csv_content.length());
    return csv_handler_->loadFromCSVString(csv_content, options);
}

bool Worksheet::saveAsCSV(const std::string& filepath, const CSVOptions& options) const {
    FASTEXCEL_LOG_INFO("Saving worksheet: {} as CSV to file: {}", name_, filepath);
    return csv_handler_->saveAsCSV(filepath, options);
}

std::string Worksheet::toCSVString(const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting worksheet: {} to CSV string", name_);
    return csv_handler_->toCSVString(options);
}

std::string Worksheet::rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                                       const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting range ({},{}) to ({},{}) of worksheet: {} to CSV string", 
                       start_row, start_col, end_row, end_col, name_);
    return csv_handler_->rangeToCSVString(start_row, start_col, end_row, end_col, options);
}

CSVParseInfo Worksheet::previewCSV(const std::string& filepath, const CSVOptions& options) {
    return WorksheetCSVHandler::previewCSV(filepath, options);
}

CSVOptions Worksheet::detectCSVOptions(const std::string& filepath) {
    return WorksheetCSVHandler::detectCSVOptions(filepath);
}

bool Worksheet::isCSVFile(const std::string& filepath) {
    return WorksheetCSVHandler::isCSVFile(filepath);
}

std::string Worksheet::getCellDisplayValue(int row, int col) const {
    return csv_handler_->getCellDisplayValue(row, col);
}

} // namespace core
} // namespace fastexcel
