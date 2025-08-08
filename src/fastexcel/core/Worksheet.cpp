#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/CellRangeManager.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/LogConfig.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
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

void Worksheet::writeDateTime(int row, int col, const std::tm& datetime) {
    this->validateCellPosition(row, col);
    
    // 使用 TimeUtils 将日期时间转换为Excel序列号
    double excel_serial = utils::TimeUtils::toExcelSerialNumber(datetime);
    this->writeNumber(row, col, excel_serial);
}

void Worksheet::writeUrl(int row, int col, const std::string& url, const std::string& string) {
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
    validateCellPosition(0, col);
    column_info_[col].width = width;
}

void Worksheet::setColumnWidth(int first_col, int last_col, double width) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].width = width;
    }
}

void Worksheet::setColumnFormatId(int col, int format_id) {
    validateCellPosition(0, col);
    column_info_[col].format_id = format_id;
}

void Worksheet::setColumnFormatId(int first_col, int last_col, int format_id) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].format_id = format_id;
    }
}

// setColumnFormat方法已移除，请使用FormatDescriptor架构

// setColumnFormat范围方法已移除，请使用FormatDescriptor架构

void Worksheet::hideColumn(int col) {
    validateCellPosition(0, col);
    column_info_[col].hidden = true;
}

void Worksheet::hideColumn(int first_col, int last_col) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].hidden = true;
    }
}

void Worksheet::setRowHeight(int row, double height) {
    validateCellPosition(row, 0);
    row_info_[row].height = height;
}

// setRowFormat方法已移除，请使用FormatDescriptor架构

void Worksheet::hideRow(int row) {
    validateCellPosition(row, 0);
    row_info_[row].hidden = true;
}

void Worksheet::hideRow(int first_row, int last_row) {
    validateRange(first_row, 0, last_row, 0);
    for (int row = first_row; row <= last_row; ++row) {
        row_info_[row].hidden = true;
    }
}

// ========== 合并单元格 ==========

void Worksheet::mergeCells(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    merge_ranges_.emplace_back(first_row, first_col, last_row, last_col);
}

// ========== 自动筛选 ==========

void Worksheet::setAutoFilter(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    autofilter_ = std::make_unique<AutoFilterRange>(first_row, first_col, last_row, last_col);
}

void Worksheet::removeAutoFilter() {
    autofilter_.reset();
}

// ========== 冻结窗格 ==========

void Worksheet::freezePanes(int row, int col) {
    validateCellPosition(row, col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

void Worksheet::freezePanes(int row, int col, int top_left_row, int top_left_col) {
    validateCellPosition(row, col);
    validateCellPosition(top_left_row, top_left_col);
    freeze_panes_ = std::make_unique<FreezePanes>(row, col, top_left_row, top_left_col);
}

void Worksheet::splitPanes(int row, int col) {
    validateCellPosition(row, col);
    // 分割窗格的实现与冻结窗格类似，但使用不同的XML属性
    freeze_panes_ = std::make_unique<FreezePanes>(row, col);
}

// ========== 打印设置 ==========

void Worksheet::setPrintArea(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    print_settings_.print_area_first_row = first_row;
    print_settings_.print_area_first_col = first_col;
    print_settings_.print_area_last_row = last_row;
    print_settings_.print_area_last_col = last_col;
}

void Worksheet::setRepeatRows(int first_row, int last_row) {
    validateRange(first_row, 0, last_row, 0);
    print_settings_.repeat_rows_first = first_row;
    print_settings_.repeat_rows_last = last_row;
}

void Worksheet::setRepeatColumns(int first_col, int last_col) {
    validateRange(0, first_col, 0, last_col);
    print_settings_.repeat_cols_first = first_col;
    print_settings_.repeat_cols_last = last_col;
}

void Worksheet::setLandscape(bool landscape) {
    print_settings_.landscape = landscape;
}

void Worksheet::setPaperSize(int paper_size) {
    // 纸张大小代码的实现
    // 这里可以根据需要添加具体的纸张大小映射
    (void)paper_size; // 避免未使用参数警告
}

void Worksheet::setMargins(double left, double right, double top, double bottom) {
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
    return it != cells_.end() && !it->second.isEmpty();
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
    // 检查是否为流式模式：通过检查parent_workbook的模式
    bool is_streaming_mode = false;
    if (parent_workbook_) {
        auto mode = parent_workbook_->getOptions().mode;
        is_streaming_mode = (mode == WorkbookMode::STREAMING);
    }
    
    if (is_streaming_mode) {
        // 流式模式：直接字符串拼接，真正的流式处理
        generateXMLStreaming(callback);
    } else {
        // 批量模式：使用XMLStreamWriter生成标准XML
        generateXMLBatch(callback);
    }
}

void Worksheet::generateXMLBatch(const std::function<void(const char*, size_t)>& callback) const {
    // 🔧 根本性修复：使用统一的 XMLStreamWriter 实例，确保正确的 flush 机制
    xml::XMLStreamWriter writer(callback);
    
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 尺寸信息
    writer.startElement("dimension");
    writer.writeAttribute("ref", range_manager_.getRangeReference().c_str());
    writer.endElement(); // dimension
    
    // 工作表视图
    writer.startElement("sheetViews");
    writer.startElement("sheetView");
    if (sheet_view_.tab_selected) {
        writer.writeAttribute("tabSelected", "1");
    }
    writer.writeAttribute("workbookViewId", "0");
    writer.endElement(); // sheetView
    writer.endElement(); // sheetViews
    
    // 工作表格式信息
    writer.startElement("sheetFormatPr");
    writer.writeAttribute("defaultRowHeight", "15");
    writer.endElement(); // sheetFormatPr
    
    // 🔧 关键修复：列信息生成 - 确保每一步都使用同一个 writer（支持范围合并）
    if (!column_info_.empty()) {
        LOG_INFO("🔧 UNIFIED WRITER: 生成<cols>XML，column_info_大小: {}", column_info_.size());
        
        writer.startElement("cols");
        
        // 🔧 关键修复：合并相邻的相同属性列
        std::vector<std::pair<int, ColumnInfo>> sorted_columns(column_info_.begin(), column_info_.end());
        std::sort(sorted_columns.begin(), sorted_columns.end());
        
        for (size_t i = 0; i < sorted_columns.size(); ) {
            int min_col = sorted_columns[i].first;
            const auto& col_info = sorted_columns[i].second;
            int max_col = min_col;
            
            // 查找具有相同属性的连续列
            while (i + 1 < sorted_columns.size() && 
                   sorted_columns[i + 1].first == max_col + 1 &&
                   sorted_columns[i + 1].second.width == col_info.width &&
                   sorted_columns[i + 1].second.format_id == col_info.format_id &&
                   sorted_columns[i + 1].second.hidden == col_info.hidden) {
                max_col = sorted_columns[i + 1].first;
                ++i;
            }
            
            // 生成合并后的<col>标签
            writer.startElement("col");
            writer.writeAttribute("min", std::to_string(min_col + 1).c_str());
            writer.writeAttribute("max", std::to_string(max_col + 1).c_str());
            
            if (col_info.width > 0) {
                writer.writeAttribute("width", std::to_string(col_info.width).c_str());
                writer.writeAttribute("customWidth", "1");
            }
            
            if (col_info.format_id >= 0) {
                writer.writeAttribute("style", std::to_string(col_info.format_id).c_str());
            }
            
            if (col_info.hidden) {
                writer.writeAttribute("hidden", "1");
            }
            
            writer.endElement(); // col
            
            LOG_DEBUG("🔧 生成列范围: {}-{} (width={}, style={}, hidden={})", 
                     min_col, max_col, col_info.width, col_info.format_id, col_info.hidden);
            
            ++i;
        }
        
        writer.endElement(); // cols
        LOG_INFO("🔧 UNIFIED WRITER: <cols>XML生成完成");
    }
    
    // 🔧 继续使用同一个 XMLStreamWriter 生成单元格数据
    writer.startElement("sheetData");
    
    // 按行排序输出单元格数据
    std::map<int, std::map<int, const Cell*>> sorted_cells;
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isEmpty() || cell.hasFormat()) {
            sorted_cells[pos.first][pos.second] = &cell;
        }
    }
    
    for (const auto& [row_num, row_cells] : sorted_cells) {
        writer.startElement("row");
        writer.writeAttribute("r", std::to_string(row_num + 1).c_str());
        
        // 检查行信息
        auto row_it = row_info_.find(row_num);
        if (row_it != row_info_.end()) {
            if (row_it->second.height > 0) {
                writer.writeAttribute("ht", std::to_string(row_it->second.height).c_str());
                writer.writeAttribute("customHeight", "1");
            }
            if (row_it->second.hidden) {
                writer.writeAttribute("hidden", "1");
            }
        }
        
        for (const auto& [col_num, cell] : row_cells) {
            writer.startElement("c");
            writer.writeAttribute("r", utils::CommonUtils::cellReference(row_num, col_num).c_str());
            
            // 应用单元格格式
            if (cell->hasFormat()) {
                int xf_index = -1;
                
                // 获取FormatDescriptor并查找其在FormatRepository中的ID
                auto format_descriptor = cell->getFormatDescriptor();
                if (format_descriptor && parent_workbook_) {
                    auto& format_repo = parent_workbook_->getStyleRepository();
                    
                    // 通过比较找到匹配的格式ID
                    for (size_t i = 0; i < format_repo.getFormatCount(); ++i) {
                        auto stored_format = format_repo.getFormat(static_cast<int>(i));
                        if (stored_format && *stored_format == *format_descriptor) {
                            xf_index = static_cast<int>(i);
                            break;
                        }
                    }
                }
                
                // 如果找不到格式，使用默认样式ID 0
                if (xf_index < 0) {
                    xf_index = 0;
                }
                
                writer.writeAttribute("s", std::to_string(xf_index).c_str());
            }
            
            // 只有在单元格不为空时才写入值
            if (!cell->isEmpty()) {
                if (cell->isFormula()) {
                    writer.writeAttribute("t", "str");
                    writer.startElement("f");
                    writer.writeText(cell->getFormula().c_str());
                    writer.endElement(); // f
                } else if (cell->isString()) {
                    // 根据工作簿设置决定使用共享字符串还是内联字符串
                    if (parent_workbook_ && parent_workbook_->getOptions().use_shared_strings) {
                        // 使用共享字符串表
                        writer.writeAttribute("t", "s");
                        writer.startElement("v");
                        int sst_index = parent_workbook_->getSharedStringIndex(cell->getStringValue());
                        if (sst_index >= 0) {
                            writer.writeText(std::to_string(sst_index).c_str());
                        } else {
                            // 如果字符串不在SST中，添加它
                            sst_index = parent_workbook_->addSharedString(cell->getStringValue());
                            writer.writeText(std::to_string(sst_index).c_str());
                        }
                        writer.endElement(); // v
                    } else {
                        writer.writeAttribute("t", "inlineStr");
                        writer.startElement("is");
                        writer.startElement("t");
                        writer.writeText(cell->getStringValue().c_str());
                        writer.endElement(); // t
                        writer.endElement(); // is
                    }
                } else if (cell->isNumber()) {
                    writer.startElement("v");
                    writer.writeText(std::to_string(cell->getNumberValue()).c_str());
                    writer.endElement(); // v
                } else if (cell->isBoolean()) {
                    writer.writeAttribute("t", "b");
                    writer.startElement("v");
                    writer.writeText(cell->getBooleanValue() ? "1" : "0");
                    writer.endElement(); // v
                }
            }
            
            writer.endElement(); // c
        }
        
        writer.endElement(); // row
    }
    
    writer.endElement(); // sheetData
    
    // 🔧 继续使用同一个 writer 生成其他部分
    generateOtherXMLWithWriter(writer);
    
    writer.endElement(); // worksheet
    writer.endDocument();
    
    LOG_INFO("🔧 UNIFIED WRITER: XML生成完成");
}

void Worksheet::generateXMLStreaming(const std::function<void(const char*, size_t)>& callback) const {
    // 真正的流式XML生成：直接写入，不缓存完整XML
    
    // XML头部 - 关键修复：添加换行符以匹配libxlsxwriter格式
    const char* xml_header = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    callback(xml_header, strlen(xml_header));
    
    // 工作表开始标签
    const char* worksheet_start = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
    callback(worksheet_start, strlen(worksheet_start));
    
    // 尺寸信息
    std::string dimension_str = "<dimension ref=\"" + range_manager_.getRangeReference() + "\"/>";
    callback(dimension_str.c_str(), dimension_str.length());
    
    // 关键修复：流式模式也要正确处理tabSelected属性
    std::string sheet_views_start = "<sheetViews><sheetView";
    if (sheet_view_.tab_selected) {
        sheet_views_start += " tabSelected=\"1\"";
    }
    sheet_views_start += " workbookViewId=\"0\"";
    callback(sheet_views_start.c_str(), sheet_views_start.length());
    
    // 冻结窗格
    if (freeze_panes_) {
        callback("><pane", 7);
        if (freeze_panes_->col > 0) {
            std::string xsplit = " xSplit=\"" + std::to_string(freeze_panes_->col) + "\"";
            callback(xsplit.c_str(), xsplit.length());
        }
        if (freeze_panes_->row > 0) {
            std::string ysplit = " ySplit=\"" + std::to_string(freeze_panes_->row) + "\"";
            callback(ysplit.c_str(), ysplit.length());
        }
        if (freeze_panes_->top_left_row >= 0 && freeze_panes_->top_left_col >= 0) {
            std::string top_left = " topLeftCell=\"" + utils::CommonUtils::cellReference(freeze_panes_->top_left_row, freeze_panes_->top_left_col) + "\"";
            callback(top_left.c_str(), top_left.length());
        }
        const char* sheet_views_end_with_pane = " state=\"frozen\"/></sheetView></sheetViews>";
        callback(sheet_views_end_with_pane, strlen(sheet_views_end_with_pane));
    } else {
        // 关键修复：修正XML标签结构，确保正确闭合
        const char* sheet_views_end = "/></sheetViews>";
        callback(sheet_views_end, strlen(sheet_views_end));
    }
    
    // 工作表格式信息
    const char* sheet_format = "<sheetFormatPr defaultRowHeight=\"15\"/>";
    callback(sheet_format, strlen(sheet_format));
    
    // 列信息（如果有）
    if (!column_info_.empty()) {
        callback("<cols>", 6);
        for (const auto& [col_num, col_info] : column_info_) {
            std::string col_xml = "<col min=\"" + std::to_string(col_num + 1) + "\" max=\"" + std::to_string(col_num + 1) + "\"";
            if (col_info.width > 0) {
                col_xml += " width=\"" + std::to_string(col_info.width) + "\" customWidth=\"1\"";
            }
            if (col_info.hidden) {
                col_xml += " hidden=\"1\"";
            }
            col_xml += "/>";
            callback(col_xml.c_str(), col_xml.length());
        }
        callback("</cols>", 7);
    }
    
    // 关键修复：确保始终生成sheetData元素（即使为空）
    callback("<sheetData", 10);
    
    // 检查是否有数据需要生成
    auto [max_row, max_col] = getUsedRange();
    if (max_row >= 0 && max_col >= 0) {
        callback(">", 1);
        // 真正的流式处理：按行处理，不在内存中缓存所有数据
        generateSheetDataStreaming(callback);
        callback("</sheetData>", 12);
    } else {
        // 空的sheetData元素
        callback("/>", 2);
    }
    
    // 工作表保护
    if (protected_) {
        std::string protection = "<sheetProtection sheet=\"1\"";
        if (!protection_password_.empty()) {
            protection += " password=\"" + protection_password_ + "\"";
        }
        protection += "/>";
        callback(protection.c_str(), protection.length());
    }
    
    // 自动筛选
    if (autofilter_) {
        std::string autofilter = "<autoFilter ref=\"" +
            utils::CommonUtils::rangeReference(autofilter_->first_row, autofilter_->first_col,
                                             autofilter_->last_row, autofilter_->last_col) + "\"/>";
        callback(autofilter.c_str(), autofilter.length());
    }
    
    // 合并单元格
    if (!merge_ranges_.empty()) {
        std::string merge_start = "<mergeCells count=\"" + std::to_string(merge_ranges_.size()) + "\">";
        callback(merge_start.c_str(), merge_start.length());
        
        for (const auto& range : merge_ranges_) {
            std::string merge_cell = "<mergeCell ref=\"" +
                utils::CommonUtils::rangeReference(range.first_row, range.first_col, range.last_row, range.last_col) + "\"/>";
            callback(merge_cell.c_str(), merge_cell.length());
        }
        
        callback("</mergeCells>", 13);
    }
    
    // 打印选项
    if (print_settings_.print_gridlines || print_settings_.print_headings ||
        print_settings_.center_horizontally || print_settings_.center_vertically) {
        std::string print_opts = "<printOptions";
        if (print_settings_.print_gridlines) print_opts += " gridLines=\"1\"";
        if (print_settings_.print_headings) print_opts += " headings=\"1\"";
        if (print_settings_.center_horizontally) print_opts += " horizontalCentered=\"1\"";
        if (print_settings_.center_vertically) print_opts += " verticalCentered=\"1\"";
        print_opts += "/>";
        callback(print_opts.c_str(), print_opts.length());
    }
    
    // 关键修复：确保始终生成pageMargins元素（按照libxlsxwriter模版）
    const char* margins = "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";
    callback(margins, strlen(margins));
    
    // 工作表结束标签
    const char* worksheet_end = "</worksheet>";
    callback(worksheet_end, strlen(worksheet_end));
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

// ========== XML生成辅助方法 ==========

void Worksheet::generateSheetDataXML(const std::function<void(const char*, size_t)>& callback) const {
    LOG_WORKSHEET_DEBUG("=== WORKSHEET XML GENERATION DEBUG START ===");
    LOG_WORKSHEET_DEBUG("Worksheet name: {}", name_);
    LOG_WORKSHEET_DEBUG("Total cells in worksheet: {}", cells_.size());
    
    // 直接使用字符串拼接，避免XMLStreamWriter的问题
    std::string xml_content = "<sheetData>";
    
    // 按行排序输出单元格数据
    std::map<int, std::map<int, const Cell*>> sorted_cells;
    for (const auto& [pos, cell] : cells_) {
        LOG_WORKSHEET_DEBUG("Processing cell at ({}, {}): isEmpty={}, hasFormat={}",
                           pos.first, pos.second, cell.isEmpty(), cell.hasFormat());
        if (!cell.isEmpty() || cell.hasFormat()) {
            sorted_cells[pos.first][pos.second] = &cell;
            LOG_WORKSHEET_DEBUG("Added cell ({}, {}) to sorted_cells", pos.first, pos.second);
        }
    }
    
    LOG_WORKSHEET_DEBUG("Grouped cells into {} rows", sorted_cells.size());
    
    for (const auto& [row_num, row_cells] : sorted_cells) {
        LOG_WORKSHEET_DEBUG("Generating row {}: {} cells", row_num, row_cells.size());
        xml_content += "<row r=\"" + std::to_string(row_num + 1) + "\"";
        
        // 检查行信息
        auto row_it = row_info_.find(row_num);
        if (row_it != row_info_.end()) {
            if (row_it->second.height > 0) {
                xml_content += " ht=\"" + std::to_string(row_it->second.height) + "\" customHeight=\"1\"";
            }
            if (row_it->second.hidden) {
                xml_content += " hidden=\"1\"";
            }
        }
        xml_content += ">";
        
        for (const auto& [col_num, cell] : row_cells) {
            LOG_WORKSHEET_DEBUG("Generating cell ({}, {}): isEmpty={}, isString={}, isNumber={}",
                               row_num, col_num, cell->isEmpty(), cell->isString(), cell->isNumber());
            
            xml_content += "<c r=\"" + utils::CommonUtils::cellReference(row_num, col_num) + "\"";
            
            // 应用单元格格式 - 已移除，现在使用FormatDescriptor架构
            // 这个旧方法已经不再使用
            
            // 只有在单元格不为空时才写入值
            if (!cell->isEmpty()) {
                if (cell->isFormula()) {
                    xml_content += " t=\"str\"><f>" + cell->getFormula() + "</f></c>";
                    LOG_WORKSHEET_DEBUG("Added formula cell: ({}, {})", row_num, col_num);
                } else if (cell->isString()) {
                    // 根据工作簿设置决定使用共享字符串还是内联字符串
                    if (parent_workbook_ && parent_workbook_->getOptions().use_shared_strings) {
                        // 使用共享字符串表
                        xml_content += " t=\"s\"><v>";
                        int sst_index = parent_workbook_->getSharedStringIndex(cell->getStringValue());
                        if (sst_index >= 0) {
                            xml_content += std::to_string(sst_index);
                        } else {
                            // 如果字符串不在SST中，添加它
                            sst_index = parent_workbook_->addSharedString(cell->getStringValue());
                            xml_content += std::to_string(sst_index);
                        }
                        xml_content += "</v></c>";
                        LOG_WORKSHEET_DEBUG("Added shared string cell: ({}, {}) with SST index {}", row_num, col_num, sst_index);
                    } else {
                        xml_content += " t=\"inlineStr\"><is><t>" + cell->getStringValue() + "</t></is></c>";
                        LOG_WORKSHEET_DEBUG("Added inline string cell: ({}, {}) with value '{}'", row_num, col_num, cell->getStringValue());
                    }
                } else if (cell->isNumber()) {
                    xml_content += "><v>" + std::to_string(cell->getNumberValue()) + "</v></c>";
                    LOG_WORKSHEET_DEBUG("Added number cell: ({}, {}) with value {}", row_num, col_num, cell->getNumberValue());
                } else if (cell->isBoolean()) {
                    xml_content += " t=\"b\"><v>" + std::string(cell->getBooleanValue() ? "1" : "0") + "</v></c>";
                    LOG_WORKSHEET_DEBUG("Added boolean cell: ({}, {}) with value {}", row_num, col_num, cell->getBooleanValue());
                } else {
                    xml_content += "/>";
                    LOG_WORKSHEET_DEBUG("Added empty cell with format: ({}, {})", row_num, col_num);
                }
            } else {
                xml_content += "/>";
                LOG_WORKSHEET_DEBUG("Added empty cell: ({}, {})", row_num, col_num);
            }
        }
        
        xml_content += "</row>";
        LOG_WORKSHEET_DEBUG("Completed row {}", row_num);
    }
    
    xml_content += "</sheetData>";
    
    LOG_WORKSHEET_DEBUG("Generated XML length: {} characters", xml_content.length());
    LOG_WORKSHEET_DEBUG("XML preview (first 500 chars): {}",
                       xml_content.length() > 500 ? xml_content.substr(0, 500) + "..." : xml_content);
    LOG_WORKSHEET_DEBUG("=== WORKSHEET XML GENERATION DEBUG END ===");
    
    callback(xml_content.c_str(), xml_content.length());
}

void Worksheet::generateColumnsXML(const std::function<void(const char*, size_t)>& callback) const {
    LOG_INFO("🔧 CRITICAL DEBUG: generateColumnsXML被调用，column_info_大小: {}", column_info_.size());
    
    if (column_info_.empty()) {
        LOG_INFO("🔧 CRITICAL DEBUG: column_info_为空，不生成<cols>标签");
        return;
    }
    
    LOG_INFO("🔧 CRITICAL DEBUG: 开始生成<cols>XML");
    
    xml::XMLStreamWriter writer(callback);
    LOG_INFO("🔧 CRITICAL DEBUG: XMLStreamWriter创建完成");
    writer.startElement("cols");
    LOG_INFO("🔧 CRITICAL DEBUG: <cols>元素开始");
    
    int processed_cols = 0;
    for (const auto& [col_num, col_info] : column_info_) {
        LOG_INFO("🔧 CRITICAL DEBUG: 处理列 {} width={} format_id={}", col_num, col_info.width, col_info.format_id);
        
        writer.startElement("col");
        writer.writeAttribute("min", std::to_string(col_num + 1).c_str());
        writer.writeAttribute("max", std::to_string(col_num + 1).c_str());
        
        if (col_info.width > 0) {
            writer.writeAttribute("width", std::to_string(col_info.width).c_str());
            writer.writeAttribute("customWidth", "1");
        }
        
        // 🔧 关键修复：添加样式属性
        if (col_info.format_id >= 0) {
            writer.writeAttribute("style", std::to_string(col_info.format_id).c_str());
        }
        
        if (col_info.hidden) {
            writer.writeAttribute("hidden", "1");
        }
        
        writer.endElement(); // col
        processed_cols++;
    }
    
    LOG_INFO("🔧 CRITICAL DEBUG: 处理了 {} 列，结束<cols>元素", processed_cols);
    writer.endElement(); // cols
    LOG_INFO("🔧 CRITICAL DEBUG: generateColumnsXML完成");
}

void Worksheet::generateMergeCellsXML(const std::function<void(const char*, size_t)>& callback) const {
    if (merge_ranges_.empty()) {
        return;
    }
    
    xml::XMLStreamWriter writer(callback);
    writer.startElement("mergeCells");
    writer.writeAttribute("count", std::to_string(merge_ranges_.size()).c_str());
    
    for (const auto& range : merge_ranges_) {
        writer.startElement("mergeCell");
        std::string ref = utils::CommonUtils::rangeReference(range.first_row, range.first_col, range.last_row, range.last_col);
        writer.writeAttribute("ref", ref.c_str());
        writer.endElement(); // mergeCell
    }
    
    writer.endElement(); // mergeCells
}

void Worksheet::generateAutoFilterXML(const std::function<void(const char*, size_t)>& callback) const {
    if (!autofilter_) {
        return;
    }
    
    xml::XMLStreamWriter writer(callback);
    writer.startElement("autoFilter");
    std::string ref = utils::CommonUtils::rangeReference(autofilter_->first_row, autofilter_->first_col,
                                   autofilter_->last_row, autofilter_->last_col);
    writer.writeAttribute("ref", ref.c_str());
    writer.endElement(); // autoFilter
}

void Worksheet::generateSheetViewsXML(const std::function<void(const char*, size_t)>& callback) const {
    xml::XMLStreamWriter writer(callback);
    writer.startElement("sheetViews");
    writer.startElement("sheetView");
    
    if (sheet_view_.tab_selected) {
        writer.writeAttribute("tabSelected", "1");
    }
    
    writer.writeAttribute("workbookViewId", "0");
    
    if (sheet_view_.zoom_scale != 100) {
        writer.writeAttribute("zoomScale", std::to_string(sheet_view_.zoom_scale).c_str());
    }
    
    if (!sheet_view_.show_gridlines) {
        writer.writeAttribute("showGridLines", "0");
    }
    
    if (!sheet_view_.show_row_col_headers) {
        writer.writeAttribute("showRowColHeaders", "0");
    }
    
    if (sheet_view_.right_to_left) {
        writer.writeAttribute("rightToLeft", "1");
    }
    
    // 选择区域 - 关键修复：不生成selection元素
    // 根据修复后的文件，所有工作表都不应该有selection元素
    
    // 冻结窗格
    if (freeze_panes_) {
        writer.startElement("pane");
        if (freeze_panes_->col > 0) {
            writer.writeAttribute("xSplit", std::to_string(freeze_panes_->col).c_str());
        }
        if (freeze_panes_->row > 0) {
            writer.writeAttribute("ySplit", std::to_string(freeze_panes_->row).c_str());
        }
        if (freeze_panes_->top_left_row >= 0 && freeze_panes_->top_left_col >= 0) {
            std::string top_left = utils::CommonUtils::cellReference(freeze_panes_->top_left_row, freeze_panes_->top_left_col);
            writer.writeAttribute("topLeftCell", top_left.c_str());
        }
        writer.writeAttribute("state", "frozen");
        writer.endElement(); // pane
    }
    
    writer.endElement(); // sheetView
    writer.endElement(); // sheetViews
}

void Worksheet::generatePageSetupXML(const std::function<void(const char*, size_t)>& callback) const {
    xml::XMLStreamWriter writer(callback);
    writer.startElement("pageSetup");
    
    if (print_settings_.landscape) {
        writer.writeAttribute("orientation", "landscape");
    }
    
    if (print_settings_.scale != 100) {
        writer.writeAttribute("scale", std::to_string(print_settings_.scale).c_str());
    }
    
    if (print_settings_.fit_to_pages_wide > 0 || print_settings_.fit_to_pages_tall > 0) {
        writer.writeAttribute("fitToWidth", std::to_string(print_settings_.fit_to_pages_wide).c_str());
        writer.writeAttribute("fitToHeight", std::to_string(print_settings_.fit_to_pages_tall).c_str());
    }
    
    writer.endElement(); // pageSetup
}

void Worksheet::generatePrintOptionsXML(const std::function<void(const char*, size_t)>& callback) const {
    if (!print_settings_.print_gridlines && !print_settings_.print_headings) {
        return;
    }
    
    xml::XMLStreamWriter writer(callback);
    writer.startElement("printOptions");
    
    if (print_settings_.print_gridlines) {
        writer.writeAttribute("gridLines", "1");
    }
    
    if (print_settings_.print_headings) {
        writer.writeAttribute("headings", "1");
    }
    
    if (print_settings_.center_horizontally) {
        writer.writeAttribute("horizontalCentered", "1");
    }
    
    if (print_settings_.center_vertically) {
        writer.writeAttribute("verticalCentered", "1");
    }
    
    writer.endElement(); // printOptions
}

void Worksheet::generatePageMarginsXML(const std::function<void(const char*, size_t)>& callback) const {
    xml::XMLStreamWriter writer(callback);
    writer.startElement("pageMargins");
    
    writer.writeAttribute("left", std::to_string(print_settings_.left_margin).c_str());
    writer.writeAttribute("right", std::to_string(print_settings_.right_margin).c_str());
    writer.writeAttribute("top", std::to_string(print_settings_.top_margin).c_str());
    writer.writeAttribute("bottom", std::to_string(print_settings_.bottom_margin).c_str());
    writer.writeAttribute("header", std::to_string(print_settings_.header_margin).c_str());
    writer.writeAttribute("footer", std::to_string(print_settings_.footer_margin).c_str());
    
    writer.endElement(); // pageMargins
}

void Worksheet::generateSheetProtectionXML(const std::function<void(const char*, size_t)>& callback) const {
    if (!protected_) {
        return;
    }
    
    xml::XMLStreamWriter writer(callback);
    writer.startElement("sheetProtection");
    writer.writeAttribute("sheet", "1");
    
    if (!protection_password_.empty()) {
        // 这里应该实现密码哈希，简化处理
        writer.writeAttribute("password", protection_password_.c_str());
    }
    
    writer.endElement(); // sheetProtection
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

void Worksheet::copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format) {
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
    
    updateUsedRange(dst_row, dst_col);
}

void Worksheet::moveCell(int src_row, int src_col, int dst_row, int dst_col) {
    validateCellPosition(src_row, src_col);
    validateCellPosition(dst_row, dst_col);
    
    if (src_row == dst_row && src_col == dst_col) {
        return; // 源和目标相同，无需移动
    }
    
    // 复制单元格
    copyCell(src_row, src_col, dst_row, dst_col, true);
    
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
            copyCell(src_first_row + r, src_first_col + c,
                    dst_row + r, dst_col + c, copy_format);
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

void Worksheet::generateSheetDataStreaming(const std::function<void(const char*, size_t)>& callback) const {
    // 真正的流式处理：分块处理数据，常量内存使用
    
    // 获取数据范围
    auto [max_row, max_col] = getUsedRange();
    if (max_row < 0 || max_col < 0) {
        return; // 没有数据
    }
    
    // 分块处理：每次处理一定数量的行，避免内存占用过大
    const int CHUNK_SIZE = 1000; // 每次处理1000行
    
    for (int chunk_start = 0; chunk_start <= max_row; chunk_start += CHUNK_SIZE) {
        int chunk_end = std::min(chunk_start + CHUNK_SIZE - 1, max_row);
        
        // 处理当前块的行
        for (int row_num = chunk_start; row_num <= chunk_end; ++row_num) {
            // 检查这一行是否有数据
            bool has_data = false;
            int min_col_in_row = INT_MAX;
            int max_col_in_row = INT_MIN;
            
            // 快速扫描这一行的列范围
            for (int col = 0; col <= max_col; ++col) {
                auto it = cells_.find(std::make_pair(row_num, col));
                if (it != cells_.end() && (!it->second.isEmpty() || it->second.hasFormat())) {
                    has_data = true;
                    min_col_in_row = std::min(min_col_in_row, col);
                    max_col_in_row = std::max(max_col_in_row, col);
                }
            }
            
            if (!has_data) {
                continue; // 跳过空行
            }
            
            // 生成行开始标签
            std::string row_start = "<row r=\"" + std::to_string(row_num + 1) + "\"";
            
            // 添加spans属性
            std::string spans = " spans=\"" + std::to_string(min_col_in_row + 1) + ":" + std::to_string(max_col_in_row + 1) + "\"";
            row_start += spans;
            
            // 检查行信息
            auto row_it = row_info_.find(row_num);
            if (row_it != row_info_.end()) {
                if (row_it->second.height > 0) {
                    row_start += " ht=\"" + std::to_string(row_it->second.height) + "\" customHeight=\"1\"";
                }
                if (row_it->second.hidden) {
                    row_start += " hidden=\"1\"";
                }
            }
            
            row_start += ">";
            callback(row_start.c_str(), row_start.length());
            
            // 生成单元格数据
            for (int col = min_col_in_row; col <= max_col_in_row; ++col) {
                auto it = cells_.find(std::make_pair(row_num, col));
                if (it == cells_.end() || (it->second.isEmpty() && !it->second.hasFormat())) {
                    continue; // 跳过空单元格
                }
                
                const Cell& cell = it->second;
                
                // 生成单元格XML
                std::string cell_xml = "<c r=\"" + utils::CommonUtils::cellReference(row_num, col) + "\"";
                
                // 应用单元格格式
                if (cell.hasFormat()) {
                    int xf_index = -1;
                    
                    // 获取FormatDescriptor并查找其在FormatRepository中的ID
                    auto format_descriptor = cell.getFormatDescriptor();
                    if (format_descriptor && parent_workbook_) {
                        auto& format_repo = parent_workbook_->getStyleRepository();
                        
                        // 通过比较找到匹配的格式ID
                        for (size_t i = 0; i < format_repo.getFormatCount(); ++i) {
                            auto stored_format = format_repo.getFormat(static_cast<int>(i));
                            if (stored_format && *stored_format == *format_descriptor) {
                                xf_index = static_cast<int>(i);
                                break;
                            }
                        }
                    }
                    
                    // 如果找不到格式，使用默认样式ID 0
                    if (xf_index < 0) {
                        xf_index = 0;
                    }
                    
                    cell_xml += " s=\"" + std::to_string(xf_index) + "\"";
                }
                
                // 只有在单元格不为空时才写入值
                if (!cell.isEmpty()) {
                    if (cell.isFormula()) {
                        cell_xml += " t=\"str\"><f>" + cell.getFormula() + "</f></c>";
                    } else if (cell.isString()) {
                        // 根据工作簿设置决定使用共享字符串还是内联字符串
                        if (parent_workbook_ && parent_workbook_->getOptions().use_shared_strings) {
                            // 使用共享字符串表
                            cell_xml += " t=\"s\"><v>";
                            int sst_index = parent_workbook_->getSharedStringIndex(cell.getStringValue());
                            if (sst_index >= 0) {
                                cell_xml += std::to_string(sst_index);
                            } else {
                                // 如果字符串不在SST中，添加它
                            sst_index = parent_workbook_->addSharedString(cell.getStringValue());
                                cell_xml += std::to_string(sst_index);
                            }
                            cell_xml += "</v></c>";
                        } else {
                            cell_xml += " t=\"inlineStr\"><is><t>" + cell.getStringValue() + "</t></is></c>";
                        }
                    } else if (cell.isNumber()) {
                        cell_xml += "><v>" + std::to_string(cell.getNumberValue()) + "</v></c>";
                    } else if (cell.isBoolean()) {
                        cell_xml += " t=\"b\"><v>" + std::string(cell.getBooleanValue() ? "1" : "0") + "</v></c>";
                    } else {
                        cell_xml += "/>";
                    }
                } else {
                    cell_xml += "/>";
                }
                
                callback(cell_xml.c_str(), cell_xml.length());
            }
            
            // 行结束标签
            const char* row_end = "</row>";
            callback(row_end, strlen(row_end));
        }
        
        // 可选：在处理完每个块后，可以进行垃圾回收或内存清理
        // 这里保持简单，让系统自动管理内存
    }
}

// 🔧 新增的统一XML生成辅助方法
std::string Worksheet::escapeXmlText(const std::string& text) const {
    std::string result;
    result.reserve(text.size() * 1.2); // 预估大小
    
    for (char c : text) {
        switch (c) {
            case '&': result += "&amp;"; break;
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default: result += c; break;
        }
    }
    
    return result;
}

void Worksheet::generateOtherXMLWithWriter(xml::XMLStreamWriter& writer) const {
    // 工作表保护
    if (!protection_password_.empty()) {
        writer.startElement("sheetProtection");
        writer.writeAttribute("sheet", "1");
        writer.writeAttribute("objects", "1");
        writer.writeAttribute("scenarios", "1");
        writer.endElement(); // sheetProtection
    }
    
    // 自动筛选
    if (autofilter_) {
        writer.startElement("autoFilter");
        // 需要实际的 autofilter 范围信息，这里先使用简单版本
        writer.writeAttribute("ref", "A1:Z1000");
        writer.endElement(); // autoFilter
    }
    
    // 合并单元格
    if (!merge_ranges_.empty()) {
        writer.startElement("mergeCells");
        writer.writeAttribute("count", std::to_string(merge_ranges_.size()).c_str());
        for (const auto& merge_range : merge_ranges_) {
            writer.startElement("mergeCell");
            // 需要将 MergeRange 转换为字符串，这里先使用简单版本
            std::string ref_str = utils::CommonUtils::cellReference(merge_range.first_row, merge_range.first_col) + 
                                 ":" + utils::CommonUtils::cellReference(merge_range.last_row, merge_range.last_col);
            writer.writeAttribute("ref", ref_str.c_str());
            writer.endElement(); // mergeCell
        }
        writer.endElement(); // mergeCells
    }
    
    // 打印选项
    writer.startElement("pageMargins");
    writer.writeAttribute("left", "0.7");
    writer.writeAttribute("right", "0.7");
    writer.writeAttribute("top", "0.75");
    writer.writeAttribute("bottom", "0.75");
    writer.writeAttribute("header", "0.3");
    writer.writeAttribute("footer", "0.3");
    writer.endElement(); // pageMargins
}

void Worksheet::generateOtherXMLSections(std::ostringstream& xml_stream) const {
    // 工作表保护
    if (!protection_password_.empty()) {
        xml_stream << "<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"/>";
    }
    
    // 自动筛选
    if (autofilter_) {
        xml_stream << "<autoFilter ref=\"A1:Z1000\"/>";
    }
    
    // 合并单元格
    if (!merge_ranges_.empty()) {
        xml_stream << "<mergeCells count=\"" << merge_ranges_.size() << "\">";
        for (const auto& merge_range : merge_ranges_) {
            std::string ref_str = utils::CommonUtils::cellReference(merge_range.first_row, merge_range.first_col) + 
                                 ":" + utils::CommonUtils::cellReference(merge_range.last_row, merge_range.last_col);
            xml_stream << "<mergeCell ref=\"" << ref_str << "\"/>";
        }
        xml_stream << "</mergeCells>";
    }
    
    // 打印选项
    xml_stream << "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" "
               << "bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";
}

}} // namespace fastexcel::core
