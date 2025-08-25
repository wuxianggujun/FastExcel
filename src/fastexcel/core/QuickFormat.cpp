#include "QuickFormat.hpp"
#include "RangeFormatter.hpp"
#include "StyleBuilder.hpp"
#include "../utils/AddressParser.hpp"
#include <sstream>
#include <stdexcept>

namespace fastexcel {
namespace core {

// 财务格式化

void QuickFormat::formatAsCurrency(Worksheet& worksheet, const std::string& range, 
                                  const std::string& symbol,
                                  int decimal_places,
                                  bool thousands_separator) {
    std::string format = buildNumberFormat(symbol, decimal_places, thousands_separator);
    
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format)
        .rightAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

void QuickFormat::formatAsPercentage(Worksheet& worksheet, const std::string& range, 
                                    int decimal_places) {
    std::ostringstream format_stream;
    format_stream << "0.";
    for (int i = 0; i < decimal_places; ++i) {
        format_stream << "0";
    }
    format_stream << "%";
    
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format_stream.str())
        .rightAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

void QuickFormat::formatAsAccounting(Worksheet& worksheet, const std::string& range,
                                    const std::string& symbol) {
    // 会计格式：左对齐货币符号，右对齐数字
    std::string format = fmt::format("_({}* #,##0.00_);_({}* (#,##0.00);_({}* \"-\"??_);_(@_)", symbol, symbol, symbol);
    
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format)
        .rightAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

// 数字格式化

void QuickFormat::formatAsNumber(Worksheet& worksheet, const std::string& range,
                                int decimal_places) {
    std::ostringstream format_stream;
    format_stream << "#,##0";
    if (decimal_places > 0) {
        format_stream << ".";
        for (int i = 0; i < decimal_places; ++i) {
            format_stream << "0";
        }
    }
    
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format_stream.str())
        .rightAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

void QuickFormat::formatAsScientific(Worksheet& worksheet, const std::string& range,
                                    int decimal_places) {
    std::ostringstream format_stream;
    format_stream << "0.";
    for (int i = 0; i < decimal_places; ++i) {
        format_stream << "0";
    }
    format_stream << "E+00";
    
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format_stream.str())
        .rightAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

// 日期时间格式化

void QuickFormat::formatAsDate(Worksheet& worksheet, const std::string& range,
                              const std::string& format) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format)
        .centerAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

void QuickFormat::formatAsTime(Worksheet& worksheet, const std::string& range,
                              const std::string& format) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format)
        .centerAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

void QuickFormat::formatAsDateTime(Worksheet& worksheet, const std::string& range,
                                  const std::string& format) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .numberFormat(format)
        .centerAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

// 表格格式化

void QuickFormat::formatAsTable(Worksheet& worksheet, const std::string& range,
                               bool has_headers,
                               bool zebra_striping,
                               const std::string& style_name) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.asTable(style_name)
             .withHeaders(has_headers)
             .withBanding(zebra_striping, false);
    formatter.apply();
}

void QuickFormat::formatAsDataList(Worksheet& worksheet, const std::string& range,
                                  bool has_headers) {
    auto formatter = worksheet.rangeFormatter(range);
    
    // 应用简单的列表格式：边框 + 交替行背景
    formatter.allBorders(BorderStyle::Thin);
    
    if (has_headers) {
        // 解析range获取第一行，完整实现
        try {
            auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
            
            // 为标题行应用特殊格式
            if (start_row >= 0 && start_col >= 0 && end_col >= start_col) {
                // 构造标题行范围字符串
                std::string header_range = utils::AddressParser::indexToRange(start_row, start_col, start_row, end_col);
                auto header_formatter = worksheet.rangeFormatter(header_range);
                header_formatter.applyStyle(StyleBuilder()
                    .border(BorderStyle::Thick)  // 标题行使用较粗边框
                    .bold(true)                   // 加粗
                    .vcenterAlign()               // 垂直居中
                    .centerAlign()                // 水平居中
                    .backgroundColor(core::Color(0xE6E6E6))  // 浅灰色背景
                );
                header_formatter.apply();
                
                // 为数据行应用交替背景（如果有多行数据）
                if (end_row > start_row) {
                    for (int row = start_row + 1; row <= end_row; ++row) {
                        std::string data_range = utils::AddressParser::indexToRange(row, start_col, row, end_col);
                        auto data_formatter = worksheet.rangeFormatter(data_range);
                        if ((row - start_row - 1) % 2 == 1) { // 奇数行应用背景色
                            data_formatter.applyStyle(StyleBuilder()
                                .border(BorderStyle::Thin)
                                .backgroundColor(core::Color(0xF5F5F5))  // 更浅的背景色
                            );
                        } else {
                            data_formatter.applyStyle(StyleBuilder()
                                .border(BorderStyle::Thin)
                            );
                        }
                        data_formatter.apply();
                    }
                }
            }
        } catch (const std::exception&) {
            // 解析失败，回退到简单格式
            formatter.applyStyle(StyleBuilder()
                .border(BorderStyle::Thin)
                .vcenterAlign()
            );
        }
    }
    
    formatter.apply();
}

// 标题和文本格式化

void QuickFormat::formatAsTitle(Worksheet& worksheet, int row, int col, 
                               const std::string& text, 
                               double font_size) {
    // 如果提供了文本，先设置单元格值
    if (!text.empty()) {
        worksheet.setCellValue(row, col, text);
    }
    
    // 应用标题格式
    worksheet.setCellFormat(row, col, StyleBuilder()
        .fontSize(font_size)
        .bold()
        .centerAlign()
        .vcenterAlign()
        .fontColor(Color::BLUE)
        .build()
    );
}

void QuickFormat::formatAsHeader(Worksheet& worksheet, const std::string& range,
                                HeaderStyle style) {
    Color bg_color = getStyleColor(style);
    Color text_color = getTextColor(style);
    
    auto formatter = worksheet.rangeFormatter(range);
    
    StyleBuilder builder = StyleBuilder()
        .bold()
        .centerAlign()
        .vcenterAlign()
        .fontColor(text_color);
    
    if (bg_color != Color::WHITE) {
        builder.backgroundColor(bg_color);
    }
    
    if (style == HeaderStyle::Modern || style == HeaderStyle::Colorful) {
        builder.border(BorderStyle::Medium, Color::WHITE);
    } else {
        builder.border(BorderStyle::Thin);
    }
    
    formatter.applyStyle(builder);
    formatter.apply();
}

void QuickFormat::formatAsComment(Worksheet& worksheet, const std::string& range) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .fontSize(9.0)
        .italic()
        .fontColor(Color::GRAY)
        .leftAlign()
        .vcenterAlign()
    );
    formatter.apply();
}

// 数据突出显示

void QuickFormat::highlight(Worksheet& worksheet, const std::string& range, Color color) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.backgroundColor(color);
    formatter.apply();
}

void QuickFormat::formatAsWarning(Worksheet& worksheet, const std::string& range) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .backgroundColor(Color::ORANGE)
        .fontColor(Color::WHITE)
        .bold()
        .centerAlign()
        .vcenterAlign()
        .border(BorderStyle::Medium, Color::ORANGE)
    );
    formatter.apply();
}

void QuickFormat::formatAsError(Worksheet& worksheet, const std::string& range) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .backgroundColor(Color::RED)
        .fontColor(Color::WHITE)
        .bold()
        .centerAlign()
        .vcenterAlign()
        .border(BorderStyle::Medium, Color::RED)
    );
    formatter.apply();
}

void QuickFormat::formatAsSuccess(Worksheet& worksheet, const std::string& range) {
    auto formatter = worksheet.rangeFormatter(range);
    formatter.applyStyle(StyleBuilder()
        .backgroundColor(Color::GREEN)
        .fontColor(Color::WHITE)
        .bold()
        .centerAlign()
        .vcenterAlign()
        .border(BorderStyle::Medium, Color::GREEN)
    );
    formatter.apply();
}

// 预定义样式套餐

void QuickFormat::applyFinancialReportStyle(Worksheet& worksheet,
                                           const std::string& data_range,
                                           const std::string& header_range,
                                           const std::string& title_cell) {
    // 1. 设置主标题（如果提供）
    if (!title_cell.empty()) {
        // 简化：假设title_cell格式为"A1"
        // 实际实现应该解析单元格地址
        formatAsTitle(worksheet, 0, 0, "", 16.0);
    }
    
    // 2. 设置表头样式
    if (!header_range.empty()) {
        formatAsHeader(worksheet, header_range, HeaderStyle::Modern);
    }
    
    // 3. 设置数据区域样式
    auto formatter = worksheet.rangeFormatter(data_range);
    formatter.allBorders(BorderStyle::Thin)
             .applyStyle(StyleBuilder()
                .vcenterAlign()
                .fontSize(10.0)
             );
    formatter.apply();
}

void QuickFormat::applyModernStyle(Worksheet& worksheet,
                                  const std::string& data_range,
                                  const std::string& header_range) {
    // 现代风格：清爽的蓝白配色
    if (!header_range.empty()) {
        formatAsHeader(worksheet, header_range, HeaderStyle::Modern);
    }
    
    formatAsTable(worksheet, data_range, !header_range.empty(), true, "TableStyleLight9");
}

void QuickFormat::applyClassicStyle(Worksheet& worksheet,
                                   const std::string& data_range,
                                   const std::string& header_range) {
    // 经典风格：传统的黑白格子样式
    if (!header_range.empty()) {
        formatAsHeader(worksheet, header_range, HeaderStyle::Classic);
    }
    
    auto formatter = worksheet.rangeFormatter(data_range);
    formatter.allBorders(BorderStyle::Thin)
             .applyStyle(StyleBuilder().vcenterAlign().fontSize(11.0));
    formatter.apply();
}

// 私有辅助方法

std::string QuickFormat::buildNumberFormat(const std::string& symbol, 
                                          int decimal_places, 
                                          bool thousands_separator) {
    std::ostringstream format_stream;
    format_stream << symbol;
    
    if (thousands_separator) {
        format_stream << "#,##0";
    } else {
        format_stream << "0";
    }
    
    if (decimal_places > 0) {
        format_stream << ".";
        for (int i = 0; i < decimal_places; ++i) {
            format_stream << "0";
        }
    }
    
    return format_stream.str();
}

Color QuickFormat::getStyleColor(HeaderStyle style) {
    switch (style) {
        case HeaderStyle::Modern:
            return Color::BLUE;
        case HeaderStyle::Classic:
            return Color::GRAY;
        case HeaderStyle::Bold:
            return Color::WHITE; // 无背景色
        case HeaderStyle::Colorful:
            return Color::PURPLE;
        default:
            return Color::BLUE;
    }
}

Color QuickFormat::getTextColor(HeaderStyle style) {
    switch (style) {
        case HeaderStyle::Modern:
        case HeaderStyle::Colorful:
            return Color::WHITE;
        case HeaderStyle::Classic:
        case HeaderStyle::Bold:
            return Color::BLACK;
        default:
            return Color::WHITE;
    }
}

}} // namespace fastexcel::core
