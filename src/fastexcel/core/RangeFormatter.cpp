#include "RangeFormatter.hpp"
#include "Worksheet.hpp"
#include "fastexcel/utils/AddressParser.hpp"
#include <stdexcept>
#include <sstream>
#include <fmt/format.h>
#include <algorithm>
#include <regex>

namespace fastexcel {
namespace core {

RangeFormatter::RangeFormatter(Worksheet* worksheet) 
    : worksheet_(worksheet) {
    if (!worksheet_) {
        throw std::invalid_argument("工作表指针不能为空");
    }
}

// 范围设置

RangeFormatter& RangeFormatter::setRange(int start_row, int start_col, int end_row, int end_col) {
    if (start_row < 0 || start_col < 0 || end_row < 0 || end_col < 0) {
        throw std::invalid_argument("行列坐标不能为负数");
    }
    if (start_row > end_row || start_col > end_col) {
        throw std::invalid_argument("起始坐标不能大于结束坐标");
    }
    
    start_row_ = start_row;
    start_col_ = start_col;
    end_row_ = end_row;
    end_col_ = end_col;
    
    return *this;
}

RangeFormatter& RangeFormatter::setRange(const std::string& range) {
    if (!parseRange(range)) {
        throw std::invalid_argument(fmt::format("无效的Excel地址格式: {}", range));
    }
    return *this;
}

RangeFormatter& RangeFormatter::setRow(int row, int start_col, int end_col) {
    if (row < 0) {
        throw std::invalid_argument("行号不能为负数");
    }
    
    start_row_ = end_row_ = row;
    start_col_ = std::max(0, start_col);
    
    if (end_col == -1) {
        // 自动确定结束列（使用工作表实际使用的最大列）
        auto [max_row, max_col] = worksheet_->getUsedRange();
        end_col_ = (max_col >= 0) ? max_col : 1023; // 如果没有数据，使用Excel最大列数
    } else {
        end_col_ = end_col;
    }
    
    return *this;
}

RangeFormatter& RangeFormatter::setColumn(int col, int start_row, int end_row) {
    if (col < 0) {
        throw std::invalid_argument("列号不能为负数");
    }
    
    start_col_ = end_col_ = col;
    start_row_ = std::max(0, start_row);
    
    if (end_row == -1) {
        // 自动确定结束行（使用工作表实际使用的最大行）
        auto [max_row, max_col] = worksheet_->getUsedRange();
        end_row_ = (max_row >= 0) ? max_row : 1048575; // 如果没有数据，使用Excel最大行数
    } else {
        end_row_ = end_row;
    }
    
    return *this;
}

// 格式应用

RangeFormatter& RangeFormatter::applyFormat(const FormatDescriptor& format) {
    pending_format_ = std::make_unique<FormatDescriptor>(format);
    return *this;
}

RangeFormatter& RangeFormatter::applyStyle(const StyleBuilder& builder) {
    pending_format_ = std::make_unique<FormatDescriptor>(builder.build());
    return *this;
}

RangeFormatter& RangeFormatter::applySharedFormat(std::shared_ptr<const FormatDescriptor> format) {
    if (format) {
        pending_format_ = std::make_unique<FormatDescriptor>(*format);
    }
    return *this;
}

// 表格样式

RangeFormatter& RangeFormatter::asTable(const std::string& style_name) {
    table_style_name_ = style_name;
    return *this;
}

RangeFormatter& RangeFormatter::withHeaders(bool has_headers) {
    has_headers_ = has_headers;
    return *this;
}

RangeFormatter& RangeFormatter::withBanding(bool row_banding, bool col_banding) {
    row_banding_ = row_banding;
    col_banding_ = col_banding;
    return *this;
}

// 边框快捷方法

RangeFormatter& RangeFormatter::allBorders(BorderStyle style, core::Color color) {
    border_target_ = BorderTarget::All;
    border_style_ = style;
    border_color_ = color;
    return *this;
}

RangeFormatter& RangeFormatter::outsideBorders(BorderStyle style, core::Color color) {
    border_target_ = BorderTarget::Outside;
    border_style_ = style;
    border_color_ = color;
    return *this;
}

RangeFormatter& RangeFormatter::insideBorders(BorderStyle style, core::Color color) {
    border_target_ = BorderTarget::Inside;
    border_style_ = style;
    border_color_ = color;
    return *this;
}

RangeFormatter& RangeFormatter::noBorders() {
    border_target_ = BorderTarget::All;
    border_style_ = BorderStyle::None;
    border_color_ = core::Color::BLACK;
    return *this;
}

// 快捷格式方法

RangeFormatter& RangeFormatter::backgroundColor(core::Color color) {
    if (!pending_format_) {
        pending_format_ = std::make_unique<FormatDescriptor>(
            StyleBuilder().backgroundColor(color).build()
        );
    } else {
        // 基于现有格式创建新格式
        auto builder = StyleBuilder(*pending_format_).backgroundColor(color);
        pending_format_ = std::make_unique<FormatDescriptor>(builder.build());
    }
    return *this;
}

RangeFormatter& RangeFormatter::fontColor(core::Color color) {
    if (!pending_format_) {
        pending_format_ = std::make_unique<FormatDescriptor>(
            StyleBuilder().fontColor(color).build()
        );
    } else {
        auto builder = StyleBuilder(*pending_format_).fontColor(color);
        pending_format_ = std::make_unique<FormatDescriptor>(builder.build());
    }
    return *this;
}

RangeFormatter& RangeFormatter::bold(bool bold) {
    if (!pending_format_) {
        pending_format_ = std::make_unique<FormatDescriptor>(
            StyleBuilder().bold(bold).build()
        );
    } else {
        auto builder = StyleBuilder(*pending_format_).bold(bold);
        pending_format_ = std::make_unique<FormatDescriptor>(builder.build());
    }
    return *this;
}

RangeFormatter& RangeFormatter::align(HorizontalAlign horizontal, VerticalAlign vertical) {
    if (!pending_format_) {
        pending_format_ = std::make_unique<FormatDescriptor>(
            StyleBuilder().horizontalAlign(horizontal).verticalAlign(vertical).build()
        );
    } else {
        auto builder = StyleBuilder(*pending_format_)
            .horizontalAlign(horizontal)
            .verticalAlign(vertical);
        pending_format_ = std::make_unique<FormatDescriptor>(builder.build());
    }
    return *this;
}

RangeFormatter& RangeFormatter::centerAlign() {
    return align(HorizontalAlign::Center, VerticalAlign::Center);
}

RangeFormatter& RangeFormatter::rightAlign() {
    return align(HorizontalAlign::Right, VerticalAlign::Bottom);
}

// 执行操作

int RangeFormatter::apply() {
    validateRange();
    
    int processed_cells = 0;
    int total_cells = (end_row_ - start_row_ + 1) * (end_col_ - start_col_ + 1);
    
    // 1. 应用基础格式
    if (pending_format_) {
        applyFormatToRange();
        processed_cells = total_cells;
    }
    
    // 2. 应用边框
    if (border_target_ != BorderTarget::None) {
        applyBordersToRange();
        // 如果还没有计数，边框操作也算作处理的单元格
        if (processed_cells == 0) {
            processed_cells = total_cells;
        }
    }
    
    // 3. 应用表格样式
    if (!table_style_name_.empty()) {
        applyTableStyle();
        // 如果还没有计数，表格样式操作也算作处理的单元格
        if (processed_cells == 0) {
            processed_cells = total_cells;
        }
    }
    
    return processed_cells;
}

std::string RangeFormatter::preview() const {
    std::ostringstream oss;
    oss << "RangeFormatter Preview:\n";
    oss << "  Range: (" << start_row_ << "," << start_col_ << ") to (" 
        << end_row_ << "," << end_col_ << ")\n";
    oss << "  Total cells: " << ((end_row_ - start_row_ + 1) * (end_col_ - start_col_ + 1)) << "\n";
    
    if (pending_format_) {
        oss << "  Has format: Yes\n";
    }
    
    if (border_target_ != BorderTarget::None) {
        oss << "  Border target: ";
        switch (border_target_) {
            case BorderTarget::All: oss << "All"; break;
            case BorderTarget::Outside: oss << "Outside"; break;
            case BorderTarget::Inside: oss << "Inside"; break;
            default: oss << "None"; break;
        }
        oss << " (" << static_cast<int>(border_style_) << ")\n";
    }
    
    if (!table_style_name_.empty()) {
        oss << "  Table style: " << table_style_name_;
        if (has_headers_) oss << " (with headers)";
        oss << "\n";
    }
    
    return oss.str();
}

// 静态工厂方法

RangeFormatter RangeFormatter::create(Worksheet& worksheet, const std::string& range) {
    return std::move(RangeFormatter(&worksheet).setRange(range));
}

RangeFormatter RangeFormatter::create(Worksheet& worksheet, 
                                    int start_row, int start_col, 
                                    int end_row, int end_col) {
    return std::move(RangeFormatter(&worksheet).setRange(start_row, start_col, end_row, end_col));
}

// 私有辅助方法

bool RangeFormatter::parseRange(const std::string& range) {
    // 简单的Excel地址解析，支持格式如"A1:C10"
    // 这里使用正则表达式进行基础解析
    
    std::regex range_regex(R"(([A-Z]+)([0-9]+):([A-Z]+)([0-9]+))");
    std::smatch matches;
    
    if (std::regex_match(range, matches, range_regex)) {
        // 解析起始位置
        std::string start_col_str = matches[1].str();
        int start_row = std::stoi(matches[2].str()) - 1; // 转换为0-based
        
        std::string end_col_str = matches[3].str();
        int end_row = std::stoi(matches[4].str()) - 1; // 转换为0-based
        
        // 转换列字母为数字
        int start_col = columnLetterToNumber(start_col_str);
        int end_col = columnLetterToNumber(end_col_str);
        
        if (start_col >= 0 && end_col >= 0 && start_row >= 0 && end_row >= 0) {
            start_row_ = start_row;
            start_col_ = start_col;
            end_row_ = end_row;
            end_col_ = end_col;
            return true;
        }
    }
    
    return false;
}

void RangeFormatter::validateRange() const {
    if (start_row_ < 0 || start_col_ < 0 || end_row_ < 0 || end_col_ < 0) {
        throw std::runtime_error("范围未设置或无效");
    }
    if (!worksheet_) {
        throw std::runtime_error("工作表指针无效");
    }
}

void RangeFormatter::applyFormatToRange() {
    if (!pending_format_ || !worksheet_) return;
    
    // 批量设置格式，利用智能API
    for (int row = start_row_; row <= end_row_; ++row) {
        for (int col = start_col_; col <= end_col_; ++col) {
            worksheet_->setCellFormat(row, col, *pending_format_);
        }
    }
}

void RangeFormatter::applyBordersToRange() {
    if (!worksheet_ || border_target_ == BorderTarget::None) return;
    
    // 根据边框目标应用不同的边框设置
    for (int row = start_row_; row <= end_row_; ++row) {
        for (int col = start_col_; col <= end_col_; ++col) {
            StyleBuilder builder;
            
            // 获取当前单元格格式作为基础
            auto current_format = getCellFormatDescriptor(row, col);
            if (current_format) {
                builder = StyleBuilder(*current_format);
            }
            
            switch (border_target_) {
                case BorderTarget::All:
                    builder.border(border_style_, border_color_);
                    break;
                    
                case BorderTarget::Outside:
                    // 只在边缘单元格设置相应边框
                    if (row == start_row_) builder.topBorder(border_style_, border_color_);
                    if (row == end_row_) builder.bottomBorder(border_style_, border_color_);
                    if (col == start_col_) builder.leftBorder(border_style_, border_color_);
                    if (col == end_col_) builder.rightBorder(border_style_, border_color_);
                    break;
                    
                case BorderTarget::Inside:
                    // 只在内部边界设置边框
                    if (row > start_row_) builder.topBorder(border_style_, border_color_);
                    if (row < end_row_) builder.bottomBorder(border_style_, border_color_);
                    if (col > start_col_) builder.leftBorder(border_style_, border_color_);
                    if (col < end_col_) builder.rightBorder(border_style_, border_color_);
                    break;
                    
                default:
                    continue;
            }
            
            worksheet_->setCellFormat(row, col, builder);
        }
    }
}

void RangeFormatter::applyTableStyle() {
    if (!worksheet_ || table_style_name_.empty()) return;
    
    // 实现基础的表格样式
    // 这里提供一个简单的实现，主要是演示概念
    
    StyleBuilder header_style = StyleBuilder::header()
        .backgroundColor(core::Color(173, 216, 230))  // 淡蓝色
        .bold()
        .centerAlign();
    
    StyleBuilder data_style = StyleBuilder()
        .border(BorderStyle::Thin)
        .verticalAlign(VerticalAlign::Center);
    
    StyleBuilder alt_row_style = data_style;
    if (row_banding_) {
        alt_row_style.backgroundColor(core::Color(211, 211, 211));  // 浅灰色
    }
    
    for (int row = start_row_; row <= end_row_; ++row) {
        for (int col = start_col_; col <= end_col_; ++col) {
            if (has_headers_ && row == start_row_) {
                // 标题行样式
                worksheet_->setCellFormat(row, col, header_style);
            } else if (row_banding_ && (row - start_row_ - (has_headers_ ? 1 : 0)) % 2 == 1) {
                // 交替行样式
                worksheet_->setCellFormat(row, col, alt_row_style);
            } else {
                // 普通数据行样式
                worksheet_->setCellFormat(row, col, data_style);
            }
        }
    }
}

// 辅助函数：将列字母转换为数字
int RangeFormatter::columnLetterToNumber(const std::string& col_str) {
    int result = 0;
    for (char c : col_str) {
        if (c >= 'A' && c <= 'Z') {
            result = result * 26 + (c - 'A' + 1);
        } else {
            return -1; // 无效字符
        }
    }
    return result - 1; // 转换为0-based
}

// 获取当前单元格格式
std::shared_ptr<const FormatDescriptor> RangeFormatter::getCellFormatDescriptor(int row, int col) const {
    if (!worksheet_ || !worksheet_->hasCellAt(row, col)) {
        return nullptr;
    }
    
    const auto& cell = worksheet_->getCell(row, col);
    return cell.getFormatDescriptor();
}

}} // namespace fastexcel::core
