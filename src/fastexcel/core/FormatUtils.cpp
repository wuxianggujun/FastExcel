#include "FormatUtils.hpp"
#include "StyleBuilder.hpp"
#include <sstream>
#include <regex>
#include <fmt/format.h>
#include <algorithm>
#include <stdexcept>

namespace fastexcel {
namespace core {

// 格式复制

bool FormatUtils::copyFormat(Worksheet& worksheet, 
                            int src_row, int src_col,
                            int dest_row, int dest_col) {
    if (!isValidCellPosition(src_row, src_col) || 
        !isValidCellPosition(dest_row, dest_col)) {
        return false;
    }
    
    auto src_format = getFormat(worksheet, src_row, src_col);
    if (!src_format) {
        return false; // 源单元格没有格式
    }
    
    worksheet.setCellFormat(core::Address(dest_row, dest_col), *src_format);
    return true;
}

int FormatUtils::copyFormat(Worksheet& worksheet, 
                           const std::string& src_range,
                           const std::string& dest_range) {
    try {
        auto src_bounds = parseRange(src_range);
        auto dest_bounds = parseRange(dest_range);
        
        int src_rows = src_bounds.second.first - src_bounds.first.first + 1;
        int src_cols = src_bounds.second.second - src_bounds.first.second + 1;
        int dest_rows = dest_bounds.second.first - dest_bounds.first.first + 1;
        int dest_cols = dest_bounds.second.second - dest_bounds.first.second + 1;
        
        // 检查范围大小是否匹配
        if (src_rows != dest_rows || src_cols != dest_cols) {
            throw std::invalid_argument("源范围和目标范围大小不匹配");
        }
        
        int copied_count = 0;
        for (int row = 0; row < src_rows; ++row) {
            for (int col = 0; col < src_cols; ++col) {
                int src_row = src_bounds.first.first + row;
                int src_col = src_bounds.first.second + col;
                int dest_row = dest_bounds.first.first + row;
                int dest_col = dest_bounds.first.second + col;
                
                if (copyFormat(worksheet, src_row, src_col, dest_row, dest_col)) {
                    copied_count++;
                }
            }
        }
        
        return copied_count;
    } catch (const std::exception&) {
        return 0;
    }
}

int FormatUtils::copyFormatToMultiple(Worksheet& worksheet,
                                     int src_row, int src_col,
                                     const std::vector<std::pair<int, int>>& dest_positions) {
    auto src_format = getFormat(worksheet, src_row, src_col);
    if (!src_format) {
        return 0;
    }
    
    int copied_count = 0;
    for (const auto& pos : dest_positions) {
        if (isValidCellPosition(pos.first, pos.second)) {
            worksheet.setCellFormat(core::Address(pos.first, pos.second), *src_format);
            copied_count++;
        }
    }
    
    return copied_count;
}

int FormatUtils::smartCopyFormat(Worksheet& worksheet,
                                const std::string& src_range,
                                const std::string& dest_start_cell) {
    try {
        auto src_bounds = parseRange(src_range);
        auto dest_start = parseCell(dest_start_cell);
        
        int src_rows = src_bounds.second.first - src_bounds.first.first + 1;
        int src_cols = src_bounds.second.second - src_bounds.first.second + 1;
        
        int copied_count = 0;
        for (int row = 0; row < src_rows; ++row) {
            for (int col = 0; col < src_cols; ++col) {
                int src_row = src_bounds.first.first + row;
                int src_col = src_bounds.first.second + col;
                int dest_row = dest_start.first + row;
                int dest_col = dest_start.second + col;
                
                if (copyFormat(worksheet, src_row, src_col, dest_row, dest_col)) {
                    copied_count++;
                }
            }
        }
        
        return copied_count;
    } catch (const std::exception&) {
        return 0;
    }
}

// 格式清除

void FormatUtils::clearFormat(Worksheet& worksheet, int row, int col) {
    if (isValidCellPosition(row, col)) {
        // 使用nullptr清除格式
        worksheet.setCellFormat(core::Address(row, col), std::shared_ptr<const FormatDescriptor>(nullptr));
    }
}

int FormatUtils::clearFormat(Worksheet& worksheet, const std::string& range) {
    try {
        auto bounds = parseRange(range);
        int cleared_count = 0;
        
        for (int row = bounds.first.first; row <= bounds.second.first; ++row) {
            for (int col = bounds.first.second; col <= bounds.second.second; ++col) {
                if (hasFormat(worksheet, row, col)) {
                    clearFormat(worksheet, row, col);
                    cleared_count++;
                }
            }
        }
        
        return cleared_count;
    } catch (const std::exception&) {
        return 0;
    }
}

int FormatUtils::clearAllFormats(Worksheet& worksheet) {
    // 这个实现需要访问工作表的内部数据结构
    // 简化版本：清除一个大范围
    return clearFormat(worksheet, "A1:ZZ1048576");
}

int FormatUtils::selectiveClearFormat(Worksheet& worksheet, const std::string& range,
                                     bool clear_font, bool clear_fill, bool clear_border,
                                     bool clear_alignment, bool clear_number) {
    // 如果所有选项都为true，则完全清除
    if (clear_font && clear_fill && clear_border && clear_alignment && clear_number) {
        return clearFormat(worksheet, range);
    }
    
    // 逐个单元格处理，构建只清除指定部分的新格式
    try {
        auto bounds = parseRange(range);
        int processed_count = 0;
        
        for (int row = bounds.first.first; row <= bounds.second.first; ++row) {
            for (int col = bounds.first.second; col <= bounds.second.second; ++col) {
                auto current_format = getFormat(worksheet, row, col);
                if (!current_format) continue;
                
                // 构建新格式（保留不需要清除的部分）
                StyleBuilder builder;
                
                // 字体属性：如果不清除则保留
                if (!clear_font) {
                    builder.fontName(current_format->getFontName())
                           .fontSize(current_format->getFontSize())
                           .bold(current_format->isBold())
                           .italic(current_format->isItalic())
                           .fontColor(current_format->getFontColor());
                }
                
                // 填充属性：如果不清除则保留
                if (!clear_fill) {
                    if (current_format->getPattern() != PatternType::None) {
                        builder.fill(current_format->getPattern(), 
                                    current_format->getBackgroundColor(), 
                                    current_format->getForegroundColor());
                    }
                }
                
                // 边框属性：如果不清除则保留
                if (!clear_border) {
                    builder.leftBorder(current_format->getLeftBorder(), current_format->getLeftBorderColor())
                           .rightBorder(current_format->getRightBorder(), current_format->getRightBorderColor())
                           .topBorder(current_format->getTopBorder(), current_format->getTopBorderColor())
                           .bottomBorder(current_format->getBottomBorder(), current_format->getBottomBorderColor());
                }
                
                // 对齐属性：如果不清除则保留
                if (!clear_alignment) {
                    builder.horizontalAlign(current_format->getHorizontalAlign())
                           .verticalAlign(current_format->getVerticalAlign())
                           .textWrap(current_format->isTextWrap());
                }
                
                // 数字格式：如果不清除则保留
                if (!clear_number) {
                    builder.numberFormat(current_format->getNumberFormat());
                }
                
                worksheet.setCellFormat(core::Address(row, col), builder.build());
                processed_count++;
            }
        }
        
        return processed_count;
    } catch (const std::exception&) {
        return 0;
    }
}

// 格式检查

bool FormatUtils::hasFormat(const Worksheet& worksheet, int row, int col) {
    if (!isValidCellPosition(row, col)) {
        return false;
    }
    
    try {
        const auto& cell = worksheet.getCell(core::Address(row, col));
        return cell.hasFormat();
    } catch (const std::exception&) {
        return false;
    }
}

std::optional<FormatDescriptor> FormatUtils::getFormat(const Worksheet& worksheet, 
                                                      int row, int col) {
    if (!isValidCellPosition(row, col)) {
        return std::nullopt;
    }
    
    try {
        const auto& cell = worksheet.getCell(core::Address(row, col));
        auto format_ptr = cell.getFormatDescriptor();
        if (format_ptr) {
            return *format_ptr;
        }
        return std::nullopt;
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

bool FormatUtils::hasUniformFormat(const Worksheet& worksheet, const std::string& range) {
    try {
        auto bounds = parseRange(range);
        std::optional<FormatDescriptor> first_format;
        
        for (int row = bounds.first.first; row <= bounds.second.first; ++row) {
            for (int col = bounds.first.second; col <= bounds.second.second; ++col) {
                auto current_format = getFormat(worksheet, row, col);
                
                if (!first_format.has_value() && current_format.has_value()) {
                    // 使用拷贝构造函数而不是赋值操作符
                    first_format.emplace(*current_format);
                } else if (first_format.has_value() && current_format.has_value()) {
                    if (!formatsMatch(*first_format, *current_format)) {
                        return false;
                    }
                } else if (first_format.has_value() && !current_format.has_value()) {
                    return false; // 有的有格式，有的没有格式
                } else if (!first_format.has_value() && current_format.has_value()) {
                    return false;
                }
            }
        }
        
        return true;
    } catch (const std::exception&) {
        return false;
    }
}

std::vector<FormatDescriptor> FormatUtils::getUniqueFormats(const Worksheet& worksheet,
                                                           const std::string& range) {
    std::vector<FormatDescriptor> unique_formats;
    
    try {
        auto bounds = parseRange(range);
        
        for (int row = bounds.first.first; row <= bounds.second.first; ++row) {
            for (int col = bounds.first.second; col <= bounds.second.second; ++col) {
                auto current_format = getFormat(worksheet, row, col);
                if (!current_format) continue;
                
                // 检查是否已经存在相同格式
                bool found = false;
                for (const auto& existing : unique_formats) {
                    if (formatsMatch(existing, *current_format)) {
                        found = true;
                        break;
                    }
                }
                
                if (!found) {
                    unique_formats.push_back(*current_format);
                }
            }
        }
    } catch (const std::exception&) {
        // 返回空列表
    }
    
    return unique_formats;
}

// 格式比较

bool FormatUtils::formatsMatch(const FormatDescriptor& format1, 
                              const FormatDescriptor& format2) {
    return format1 == format2;
}

FormatUtils::FormatDifference FormatUtils::compareFormats(const FormatDescriptor& format1,
                                                         const FormatDescriptor& format2) {
    FormatDifference diff;
    
    // 比较字体
    diff.font_different = (format1.getFontName() != format2.getFontName()) ||
                         (format1.getFontSize() != format2.getFontSize()) ||
                         (format1.isBold() != format2.isBold()) ||
                         (format1.isItalic() != format2.isItalic()) ||
                         (format1.getFontColor() != format2.getFontColor());
    
    // 比较填充
    diff.fill_different = (format1.getPattern() != format2.getPattern()) ||
                         (format1.getBackgroundColor() != format2.getBackgroundColor()) ||
                         (format1.getForegroundColor() != format2.getForegroundColor());
    
    // 比较边框
    diff.border_different = (format1.getLeftBorder() != format2.getLeftBorder()) ||
                           (format1.getRightBorder() != format2.getRightBorder()) ||
                           (format1.getTopBorder() != format2.getTopBorder()) ||
                           (format1.getBottomBorder() != format2.getBottomBorder());
    
    // 比较对齐
    diff.alignment_different = (format1.getHorizontalAlign() != format2.getHorizontalAlign()) ||
                              (format1.getVerticalAlign() != format2.getVerticalAlign()) ||
                              (format1.isTextWrap() != format2.isTextWrap());
    
    // 比较数字格式
    diff.number_format_different = (format1.getNumberFormat() != format2.getNumberFormat()) ||
                                  (format1.getNumberFormatIndex() != format2.getNumberFormatIndex());
    
    return diff;
}

std::string FormatUtils::FormatDifference::toString() const {
    std::ostringstream oss;
    std::vector<std::string> differences;
    
    if (font_different) differences.push_back("字体");
    if (fill_different) differences.push_back("填充");
    if (border_different) differences.push_back("边框");
    if (alignment_different) differences.push_back("对齐");
    if (number_format_different) differences.push_back("数字格式");
    
    if (differences.empty()) {
        return "无差异";
    }
    
    oss << "差异: ";
    for (size_t i = 0; i < differences.size(); ++i) {
        if (i > 0) oss << ", ";
        oss << differences[i];
    }
    
    return oss.str();
}

// 私有辅助方法

std::pair<int, int> FormatUtils::parseCell(const std::string& cell_address) {
    // 简单的Excel地址解析，如"A1" -> (0, 0)
    std::regex cell_regex(R"(([A-Z]+)([0-9]+))");
    std::smatch matches;
    
    if (std::regex_match(cell_address, matches, cell_regex)) {
        std::string col_str = matches[1].str();
        int row = std::stoi(matches[2].str()) - 1; // 转换为0-based
        
        // 转换列字母为数字
        int col = 0;
        for (char c : col_str) {
            col = col * 26 + (c - 'A' + 1);
        }
        col -= 1; // 转换为0-based
        
        return {row, col};
    }
    
    throw std::invalid_argument(fmt::format("无效的单元格地址: {}", cell_address));
}

std::pair<std::pair<int, int>, std::pair<int, int>> FormatUtils::parseRange(const std::string& range) {
    // 解析范围如"A1:C10"
    std::regex range_regex(R"(([A-Z]+[0-9]+):([A-Z]+[0-9]+))");
    std::smatch matches;
    
    if (std::regex_match(range, matches, range_regex)) {
        auto start = parseCell(matches[1].str());
        auto end = parseCell(matches[2].str());
        return {start, end};
    }
    
    throw std::invalid_argument(fmt::format("无效的范围地址: {}", range));
}

bool FormatUtils::isValidCellPosition(int row, int col) {
    return row >= 0 && col >= 0 && row < 1048576 && col < 16384; // Excel限制
}

// 格式导入导出（简化实现）

std::string FormatUtils::exportFormat(const FormatDescriptor& format) {
    // 完整的JSON导出包含所有格式属性
    std::ostringstream oss;
    oss << "{"
        // 字体属性
        << "\"fontName\":\"" << format.getFontName() << "\","
        << "\"fontSize\":" << format.getFontSize() << ","
        << "\"bold\":" << (format.isBold() ? "true" : "false") << ","
        << "\"italic\":" << (format.isItalic() ? "true" : "false") << ","
        << "\"strikeout\":" << (format.isStrikeout() ? "true" : "false") << ","
        << "\"fontColor\":\"" << format.getFontColor().toHex() << "\","
        
        // 对齐属性
        << "\"horizontalAlign\":" << static_cast<int>(format.getHorizontalAlign()) << ","
        << "\"verticalAlign\":" << static_cast<int>(format.getVerticalAlign()) << ","
        << "\"textWrap\":" << (format.isTextWrap() ? "true" : "false") << ","
        << "\"rotation\":" << format.getRotation() << ","
        << "\"indent\":" << format.getIndent() << ","
        << "\"shrink\":" << (format.isShrink() ? "true" : "false") << ","
        
        // 边框属性
        << "\"leftBorder\":" << static_cast<int>(format.getLeftBorder()) << ","
        << "\"rightBorder\":" << static_cast<int>(format.getRightBorder()) << ","
        << "\"topBorder\":" << static_cast<int>(format.getTopBorder()) << ","
        << "\"bottomBorder\":" << static_cast<int>(format.getBottomBorder()) << ","
        << "\"leftBorderColor\":\"" << format.getLeftBorderColor().toHex() << "\","
        << "\"rightBorderColor\":\"" << format.getRightBorderColor().toHex() << "\","
        << "\"topBorderColor\":\"" << format.getTopBorderColor().toHex() << "\","
        << "\"bottomBorderColor\":\"" << format.getBottomBorderColor().toHex() << "\","
        
        // 填充属性
        << "\"pattern\":" << static_cast<int>(format.getPattern()) << ","
        << "\"backgroundColor\":\"" << format.getBackgroundColor().toHex() << "\","
        << "\"foregroundColor\":\"" << format.getForegroundColor().toHex() << "\","
        
        // 数字格式
        << "\"numberFormat\":\"" << format.getNumberFormat() << "\","
        << "\"numberFormatIndex\":" << format.getNumberFormatIndex() << ","
        
        // 保护属性
        << "\"locked\":" << (format.isLocked() ? "true" : "false") << ","
        << "\"hidden\":" << (format.isHidden() ? "true" : "false")
        << "}";
    return oss.str();
}

std::optional<FormatDescriptor> FormatUtils::importFormat(const std::string& format_string) {
    if (format_string.empty()) {
        return std::nullopt;
    }
    
    try {
        // 简化的JSON解析，只处理基本属性
        StyleBuilder builder;
        
        // 解析 fontName
        std::regex font_name_regex(R"("fontName"\s*:\s*"([^"]*)");
        std::smatch match;
        if (std::regex_search(format_string, match, font_name_regex)) {
            builder.fontName(match[1].str());
        }
        
        // 解析 fontSize
        std::regex font_size_regex(R"("fontSize"\s*:\s*(\d+))");
        if (std::regex_search(format_string, match, font_size_regex)) {
            int size = std::stoi(match[1].str());
            builder.fontSize(size);
        }
        
        // 解析 bold
        std::regex bold_regex(R"("bold"\s*:\s*(true|false))");
        if (std::regex_search(format_string, match, bold_regex)) {
            bool is_bold = match[1].str() == "true";
            builder.bold(is_bold);
        }
        
        // 解析 italic
        std::regex italic_regex(R"("italic"\s*:\s*(true|false))");
        if (std::regex_search(format_string, match, italic_regex)) {
            bool is_italic = match[1].str() == "true";
            builder.italic(is_italic);
        }
        
        return builder.build();
        
    } catch (const std::exception&) {
        // JSON解析失败
        return std::nullopt;
    }
}

}} // namespace fastexcel::core
