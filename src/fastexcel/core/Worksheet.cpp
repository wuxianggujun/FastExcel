#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatPool.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
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

void Worksheet::writeString(int row, int col, const std::string& value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    
    Cell cell;
    if (sst_) {
        // 使用共享字符串表
        sst_->addString(value);
        // 注意：这里需要修改Cell类来支持SST ID，暂时直接设置值
        cell.setValue(value);
    } else {
        cell.setValue(value);
    }
    
    if (optimize_mode_) {
        writeOptimizedCell(row, col, std::move(cell), format);
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        if (format) {
            target_cell.setFormat(format);
        }
        updateUsedRange(row, col);
    }
}

void Worksheet::writeNumber(int row, int col, double value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    
    Cell cell;
    cell.setValue(value);
    
    if (optimize_mode_) {
        writeOptimizedCell(row, col, std::move(cell), format);
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        if (format) {
            target_cell.setFormat(format);
        }
        updateUsedRange(row, col);
    }
}

void Worksheet::writeBoolean(int row, int col, bool value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    
    Cell cell;
    cell.setValue(value);
    
    if (optimize_mode_) {
        writeOptimizedCell(row, col, std::move(cell), format);
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        if (format) {
            target_cell.setFormat(format);
        }
        updateUsedRange(row, col);
    }
}

void Worksheet::writeFormula(int row, int col, const std::string& formula, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    
    Cell cell;
    cell.setFormula(formula);
    
    if (optimize_mode_) {
        writeOptimizedCell(row, col, std::move(cell), format);
    } else {
        auto& target_cell = cells_[std::make_pair(row, col)];
        target_cell = std::move(cell);
        if (format) {
            target_cell.setFormat(format);
        }
        updateUsedRange(row, col);
    }
}

void Worksheet::writeDateTime(int row, int col, const std::tm& datetime, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    
    // 将tm结构转换为Excel日期序列号
    // Excel使用1900年1月1日作为起始日期（序列号1）
    std::tm epoch = {};
    epoch.tm_year = 100; // 1900年
    epoch.tm_mon = 0;    // 1月
    epoch.tm_mday = 1;   // 1日
    
    std::time_t epoch_time = std::mktime(&epoch);
    std::time_t input_time = std::mktime(const_cast<std::tm*>(&datetime));
    
    double days = std::difftime(input_time, epoch_time) / (24 * 60 * 60);
    days += 1; // Excel从1开始计数
    
    // Excel错误地认为1900年是闰年，所以需要调整
    if (days >= 60) {
        days += 1;
    }
    
    writeNumber(row, col, days, format);
}

void Worksheet::writeUrl(int row, int col, const std::string& url, const std::string& string, 
                        std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    
    std::string display_text = string.empty() ? url : string;
    cell.setValue(display_text);
    cell.setHyperlink(url);
    
    if (format) {
        cell.setFormat(format);
    }
    updateUsedRange(row, col);
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

void Worksheet::setColumnFormat(int col, std::shared_ptr<Format> format) {
    validateCellPosition(0, col);
    column_info_[col].format = format;
}

void Worksheet::setColumnFormat(int first_col, int last_col, std::shared_ptr<Format> format) {
    validateRange(0, first_col, 0, last_col);
    for (int col = first_col; col <= last_col; ++col) {
        column_info_[col].format = format;
    }
}

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

void Worksheet::setRowFormat(int row, std::shared_ptr<Format> format) {
    validateCellPosition(row, 0);
    row_info_[row].format = format;
}

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

void Worksheet::mergeRange(int first_row, int first_col, int last_row, int last_col, 
                          const std::string& value, std::shared_ptr<Format> format) {
    mergeCells(first_row, first_col, last_row, last_col);
    writeString(first_row, first_col, value, format);
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
    active_cell_ = cellReference(row, col);
}

void Worksheet::setSelection(int first_row, int first_col, int last_row, int last_col) {
    validateRange(first_row, first_col, last_row, last_col);
    if (first_row == last_row && first_col == last_col) {
        selection_ = cellReference(first_row, first_col);
    } else {
        selection_ = rangeReference(first_row, first_col, last_row, last_col);
    }
}

// ========== 获取信息 ==========

std::pair<int, int> Worksheet::getUsedRange() const {
    int max_row = -1;
    int max_col = -1;
    
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isEmpty()) {
            max_row = std::max(max_row, pos.first);
            max_col = std::max(max_col, pos.second);
        }
    }
    
    return {max_row, max_col};
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

double Worksheet::getRowHeight(int row) const {
    auto it = row_info_.find(row);
    if (it != row_info_.end() && it->second.height > 0) {
        return it->second.height;
    }
    return default_row_height_;
}

std::shared_ptr<Format> Worksheet::getColumnFormat(int col) const {
    auto it = column_info_.find(col);
    if (it != column_info_.end()) {
        return it->second.format;
    }
    return nullptr;
}

std::shared_ptr<Format> Worksheet::getRowFormat(int row) const {
    auto it = row_info_.find(row);
    if (it != row_info_.end()) {
        return it->second.format;
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
    xml::XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 工作表属性
    if (sheet_view_.right_to_left) {
        writer.startElement("sheetPr");
        writer.writeAttribute("rightToLeft", "1");
        writer.endElement(); // sheetPr
    }
    
    // 尺寸信息
    auto [max_row, max_col] = getUsedRange();
    if (max_row >= 0 && max_col >= 0) {
        writer.startElement("dimension");
        std::string ref = "A1:" + cellReference(max_row, max_col);
        writer.writeAttribute("ref", ref.c_str());
        writer.endElement(); // dimension
    }
    
    // 工作表视图
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
    
    // 选择区域
    if (!selection_.empty()) {
        writer.startElement("selection");
        writer.writeAttribute("sqref", selection_.c_str());
        if (!active_cell_.empty()) {
            writer.writeAttribute("activeCell", active_cell_.c_str());
        }
        writer.endElement(); // selection
    }
    
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
            std::string top_left = cellReference(freeze_panes_->top_left_row, freeze_panes_->top_left_col);
            writer.writeAttribute("topLeftCell", top_left.c_str());
        }
        writer.writeAttribute("state", "frozen");
        writer.endElement(); // pane
    }
    
    writer.endElement(); // sheetView
    writer.endElement(); // sheetViews
    
    // 工作表格式信息
    writer.writeRaw("<sheetFormatPr defaultRowHeight=\"");
    writer.writeRaw(std::to_string(default_row_height_).c_str());
    writer.writeRaw("\" defaultColWidth=\"");
    writer.writeRaw(std::to_string(default_col_width_).c_str());
    writer.writeRaw("\"/>");
    
    // 列信息
    if (!column_info_.empty()) {
        writer.startElement("cols");
        
        for (const auto& [col_num, col_info] : column_info_) {
            writer.startElement("col");
            writer.writeAttribute("min", std::to_string(col_num + 1).c_str());
            writer.writeAttribute("max", std::to_string(col_num + 1).c_str());
            
            if (col_info.width > 0) {
                writer.writeAttribute("width", std::to_string(col_info.width).c_str());
                writer.writeAttribute("customWidth", "1");
            }
            
            if (col_info.hidden) {
                writer.writeAttribute("hidden", "1");
            }
            
            writer.endElement(); // col
        }
        
        writer.endElement(); // cols
    }
    
    // 工作表数据
    writer.startElement("sheetData");
    
    // 按行排序输出单元格数据
    std::map<int, std::map<int, const Cell*>> sorted_cells;
    // 关键修复 #1: 只要单元格有值或有格式，就必须处理
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
            writer.writeAttribute("r", cellReference(row_num, col_num).c_str());
            
            // 关键修复 #2: 应用单元格格式
            if (cell->hasFormat()) {
                writer.writeAttribute("s", std::to_string(cell->getFormat()->getXfIndex()).c_str());
            }
            
            // 只有在单元格不为空时才写入值
            if (!cell->isEmpty()) {
                if (cell->isFormula()) {
                    writer.writeAttribute("t", "str");
                    writer.startElement("f");
                    writer.writeText(cell->getFormula().c_str());
                    writer.endElement(); // f
                } else if (cell->isString()) {
                    // 关键修复 #3: 根据是否使用SST，决定写入共享字符串还是内联字符串
                    if (sst_) {
                        writer.writeAttribute("t", "s");
                        writer.startElement("v");
                        writer.writeText(std::to_string(sst_->getStringId(cell->getStringValue())).c_str());
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
    
    // 工作表保护
    if (protected_) {
        writer.startElement("sheetProtection");
        writer.writeAttribute("sheet", "1");
        
        if (!protection_password_.empty()) {
            // 这里应该实现密码哈希，简化处理
            writer.writeAttribute("password", protection_password_.c_str());
        }
        
        writer.endElement(); // sheetProtection
    }
    
    // 自动筛选
    if (autofilter_) {
        writer.startElement("autoFilter");
        std::string ref = rangeReference(autofilter_->first_row, autofilter_->first_col,
                                       autofilter_->last_row, autofilter_->last_col);
        writer.writeAttribute("ref", ref.c_str());
        writer.endElement(); // autoFilter
    }
    
    // 合并单元格
    if (!merge_ranges_.empty()) {
        writer.startElement("mergeCells");
        writer.writeAttribute("count", std::to_string(merge_ranges_.size()).c_str());
        
        for (const auto& range : merge_ranges_) {
            writer.startElement("mergeCell");
            std::string ref = rangeReference(range.first_row, range.first_col, range.last_row, range.last_col);
            writer.writeAttribute("ref", ref.c_str());
            writer.endElement(); // mergeCell
        }
        
        writer.endElement(); // mergeCells
    }
    
    // 打印选项
    if (print_settings_.print_gridlines || print_settings_.print_headings ||
        print_settings_.center_horizontally || print_settings_.center_vertically) {
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
    
    // 页边距
    writer.startElement("pageMargins");
    writer.writeAttribute("left", std::to_string(print_settings_.left_margin).c_str());
    writer.writeAttribute("right", std::to_string(print_settings_.right_margin).c_str());
    writer.writeAttribute("top", std::to_string(print_settings_.top_margin).c_str());
    writer.writeAttribute("bottom", std::to_string(print_settings_.bottom_margin).c_str());
    writer.writeAttribute("header", std::to_string(print_settings_.header_margin).c_str());
    writer.writeAttribute("footer", std::to_string(print_settings_.footer_margin).c_str());
    writer.endElement(); // pageMargins
    
    // 页面设置
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
    
    writer.endElement(); // worksheet
    writer.endDocument();
}

void Worksheet::generateXMLToFile(const std::string& filename) const {
    xml::XMLStreamWriter writer(filename);
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 工作表属性
    if (sheet_view_.right_to_left) {
        writer.startElement("sheetPr");
        writer.writeAttribute("rightToLeft", "1");
        writer.endElement(); // sheetPr
    }
    
    // 尺寸信息
    auto [max_row, max_col] = getUsedRange();
    if (max_row >= 0 && max_col >= 0) {
        writer.startElement("dimension");
        std::string ref = "A1:" + cellReference(max_row, max_col);
        writer.writeAttribute("ref", ref.c_str());
        writer.endElement(); // dimension
    }
    
    // 使用回调模式生成各部分
    auto file_callback = [&writer](const char* data, size_t size) {
        writer.writeRaw(std::string(data, size));
    };
    
    // 工作表视图
    generateSheetViewsXML(file_callback);
    
    // 工作表格式信息
    std::string sheet_format_xml = "<sheetFormatPr defaultRowHeight=\"" +
                                  std::to_string(default_row_height_) +
                                  "\" defaultColWidth=\"" +
                                  std::to_string(default_col_width_) + "\"/>";
    writer.writeRaw(sheet_format_xml.c_str());
    
    // 列信息
    generateColumnsXML(file_callback);
    
    // 工作表数据
    generateSheetDataXML(file_callback);
    
    // 工作表保护
    generateSheetProtectionXML(file_callback);
    
    // 自动筛选
    generateAutoFilterXML(file_callback);
    
     // 合并单元格
    generateMergeCellsXML(file_callback);
    
    // 打印选项
    generatePrintOptionsXML(file_callback);
    
    // 页边距
    generatePageMarginsXML(file_callback);
    
    // 页面设置
    generatePageSetupXML(file_callback);
    
    writer.endElement(); // worksheet
    writer.endDocument();
}

void Worksheet::generateRelsXML(const std::function<void(const char*, size_t)>& callback) const {
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

std::string Worksheet::columnToLetter(int col) const {
    std::string result;
    while (col >= 0) {
        result = static_cast<char>('A' + (col % 26)) + result;
        col = col / 26 - 1;
    }
    return result;
}

std::string Worksheet::cellReference(int row, int col) const {
    return columnToLetter(col) + std::to_string(row + 1);
}

std::string Worksheet::rangeReference(int first_row, int first_col, int last_row, int last_col) const {
    return cellReference(first_row, first_col) + ":" + cellReference(last_row, last_col);
}

void Worksheet::validateCellPosition(int row, int col) const {
    if (row < 0 || col < 0) {
        throw std::invalid_argument("Cell position cannot be negative");
    }
    if (row > 1048575 || col > 16383) { // Excel 2007+ limits
        throw std::invalid_argument("Cell position exceeds Excel limits");
    }
}

void Worksheet::validateRange(int first_row, int first_col, int last_row, int last_col) const {
    validateCellPosition(first_row, first_col);
    validateCellPosition(last_row, last_col);
    if (first_row > last_row || first_col > last_col) {
        throw std::invalid_argument("Invalid range: first position must be before last position");
    }
}

// ========== XML生成辅助方法 ==========

void Worksheet::generateSheetDataXML(const std::function<void(const char*, size_t)>& callback) const {
    xml::XMLStreamWriter writer(callback);
    writer.startElement("sheetData");
    
    // 按行排序输出单元格数据
    std::map<int, std::map<int, const Cell*>> sorted_cells;
    // 关键修复 #1: 只要单元格有值或有格式，就必须处理
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
            writer.writeAttribute("r", cellReference(row_num, col_num).c_str());
            
            // 关键修复 #2: 应用单元格格式
            if (cell->hasFormat()) {
                writer.writeAttribute("s", std::to_string(cell->getFormat()->getXfIndex()).c_str());
            }
            
            // 只有在单元格不为空时才写入值
            if (!cell->isEmpty()) {
                if (cell->isFormula()) {
                    writer.writeAttribute("t", "str");
                    writer.startElement("f");
                    writer.writeText(cell->getFormula().c_str());
                    writer.endElement(); // f
                } else if (cell->isString()) {
                    // 关键修复 #3: 根据是否使用SST，决定写入共享字符串还是内联字符串
                    if (sst_) {
                        writer.writeAttribute("t", "s");
                        writer.startElement("v");
                        writer.writeText(std::to_string(sst_->getStringId(cell->getStringValue())).c_str());
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
}

void Worksheet::generateColumnsXML(const std::function<void(const char*, size_t)>& callback) const {
    if (column_info_.empty()) {
        return;
    }
    
    xml::XMLStreamWriter writer(callback);
    writer.startElement("cols");
    
    for (const auto& [col_num, col_info] : column_info_) {
        writer.startElement("col");
        writer.writeAttribute("min", std::to_string(col_num + 1).c_str());
        writer.writeAttribute("max", std::to_string(col_num + 1).c_str());
        
        if (col_info.width > 0) {
            writer.writeAttribute("width", std::to_string(col_info.width).c_str());
            writer.writeAttribute("customWidth", "1");
        }
        
        if (col_info.hidden) {
            writer.writeAttribute("hidden", "1");
        }
        
        writer.endElement(); // col
    }
    
    writer.endElement(); // cols
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
        std::string ref = rangeReference(range.first_row, range.first_col, range.last_row, range.last_col);
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
    std::string ref = rangeReference(autofilter_->first_row, autofilter_->first_col,
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
    
    // 选择区域
    if (!selection_.empty()) {
        writer.startElement("selection");
        writer.writeAttribute("sqref", selection_.c_str());
        if (!active_cell_.empty()) {
            writer.writeAttribute("activeCell", active_cell_.c_str());
        }
        writer.endElement(); // selection
    }
    
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
            std::string top_left = cellReference(freeze_panes_->top_left_row, freeze_panes_->top_left_col);
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
    // 这个方法在写入数据时被调用，用于跟踪使用范围
    // 实际实现中可以优化性能
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
    if (current_row_->height > 0 || current_row_->format || current_row_->hidden) {
        RowInfo& row_info = row_info_[row_num];
        if (current_row_->height > 0) {
            row_info.height = current_row_->height;
        }
        if (current_row_->format) {
            row_info.format = current_row_->format;
        }
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
    
    if (format_pool_) {
        stats.unique_formats = format_pool_->getFormatCount();
        auto dedup_stats = format_pool_->getDeduplicationStats();
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

void Worksheet::writeOptimizedCell(int row, int col, Cell&& cell, std::shared_ptr<Format> format) {
    updateUsedRangeOptimized(row, col);
    
    // 处理格式
    if (format) {
        if (format_pool_) {
            // 使用格式池去重
            size_t format_index = format_pool_->getFormatIndex(format.get());
            cell.setFormat(format);
        } else {
            cell.setFormat(format);
        }
    }
    
    ensureCurrentRow(row);
    current_row_->cells[col] = std::move(cell);
    current_row_->data_changed = true;
}

void Worksheet::updateUsedRangeOptimized(int row, int col) {
    min_row_ = std::min(min_row_, row);
    max_row_ = std::max(max_row_, row);
    min_col_ = std::min(min_col_, col);
    max_col_ = std::max(max_col_, col);
}

}} // namespace fastexcel::core
