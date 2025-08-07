//
// Created by wuxianggujun on 25-8-4.
//

#include "StylesParser.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/StyleBuilder.hpp"
#include "fastexcel/core/Color.hpp"
#include <iostream>
#include <sstream>
#include <algorithm>
#include <cctype>

namespace fastexcel {
namespace reader {

bool StylesParser::parse(const std::string& xml_content) {
    if (xml_content.empty()) {
        return true; // 空内容是正常的
    }
    
    try {
        // 清理之前的数据
        fonts_.clear();
        fills_.clear();
        borders_.clear();
        cell_xfs_.clear();
        number_formats_.clear();
        
        // 解析各个部分
        parseNumberFormats(xml_content);
        parseFonts(xml_content);
        parseFills(xml_content);
        parseBorders(xml_content);
        parseCellXfs(xml_content);
        
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "解析样式时发生错误: " << e.what() << std::endl;
        return false;
    }
}

std::shared_ptr<core::FormatDescriptor> StylesParser::getFormat(int xf_index) const {
    if (xf_index < 0 || xf_index >= static_cast<int>(cell_xfs_.size())) {
        return nullptr;
    }
    
    const auto& xf = cell_xfs_[xf_index];
    core::StyleBuilder builder;
    
    // 应用字体
    if (xf.font_id >= 0 && xf.font_id < static_cast<int>(fonts_.size())) {
        const auto& font = fonts_[xf.font_id];
        builder.fontName(font.name)
               .fontSize(font.size)
               .bold(font.bold)
               .italic(font.italic)
               .underline(font.underline ? core::UnderlineType::Single : core::UnderlineType::None)
               .strikeout(font.strikeout);
        
        if (font.color.getRed() != 0 || font.color.getGreen() != 0 || font.color.getBlue() != 0) {
            builder.fontColor(font.color);
        }
    }
    
    // 应用填充
    if (xf.fill_id >= 0 && xf.fill_id < static_cast<int>(fills_.size())) {
        const auto& fill = fills_[xf.fill_id];
        if (fill.pattern_type != "none" && (fill.fg_color.getRed() != 0 || fill.fg_color.getGreen() != 0 || fill.fg_color.getBlue() != 0)) {
            builder.backgroundColor(fill.fg_color);
        }
    }
    
    // 应用边框
    if (xf.border_id >= 0 && xf.border_id < static_cast<int>(borders_.size())) {
        const auto& border = borders_[xf.border_id];
        if (!border.left.style.empty()) {
            builder.leftBorder(getBorderStyle(border.left.style));
            if (border.left.color.getRed() != 0 || border.left.color.getGreen() != 0 || border.left.color.getBlue() != 0) {
                // Note: StyleBuilder doesn't have leftBorderColor method, use leftBorder with color
                builder.leftBorder(getBorderStyle(border.left.style), border.left.color);
            }
        }
        if (!border.right.style.empty()) {
            builder.rightBorder(getBorderStyle(border.right.style));
            if (border.right.color.getRed() != 0 || border.right.color.getGreen() != 0 || border.right.color.getBlue() != 0) {
                // Note: StyleBuilder doesn't have rightBorderColor method, use rightBorder with color
                builder.rightBorder(getBorderStyle(border.right.style), border.right.color);
            }
        }
        if (!border.top.style.empty()) {
            builder.topBorder(getBorderStyle(border.top.style));
            if (border.top.color.getRed() != 0 || border.top.color.getGreen() != 0 || border.top.color.getBlue() != 0) {
                // Note: StyleBuilder doesn't have topBorderColor method, use topBorder with color
                builder.topBorder(getBorderStyle(border.top.style), border.top.color);
            }
        }
        if (!border.bottom.style.empty()) {
            builder.bottomBorder(getBorderStyle(border.bottom.style));
            if (border.bottom.color.getRed() != 0 || border.bottom.color.getGreen() != 0 || border.bottom.color.getBlue() != 0) {
                // Note: StyleBuilder doesn't have bottomBorderColor method, use bottomBorder with color
                builder.bottomBorder(getBorderStyle(border.bottom.style), border.bottom.color);
            }
        }
    }
    
    // 应用对齐
    builder.horizontalAlign(getAlignment(xf.horizontal_alignment))
           .verticalAlign(getVerticalAlignment(xf.vertical_alignment))
           .textWrap(xf.wrap_text);
    
    // 应用数字格式
    if (xf.num_fmt_id >= 0) {
        auto it = number_formats_.find(xf.num_fmt_id);
        if (it != number_formats_.end()) {
            builder.numberFormat(it->second);
        } else {
            // 使用内置格式
            builder.numberFormat(getBuiltinNumberFormat(xf.num_fmt_id));
        }
    }
    
    // 构建FormatDescriptor并返回共享指针
    return std::make_shared<core::FormatDescriptor>(builder.build());
}

void StylesParser::parseNumberFormats(const std::string& xml_content) {
    size_t numFmts_start = xml_content.find("<numFmts");
    if (numFmts_start == std::string::npos) {
        return; // 没有自定义数字格式
    }
    
    size_t numFmts_end = xml_content.find("</numFmts>", numFmts_start);
    if (numFmts_end == std::string::npos) {
        // 尝试查找自闭合标签
        size_t self_close = xml_content.find("/>", numFmts_start);
        if (self_close != std::string::npos) {
            return; // 空的numFmts
        }
        return;
    }
    
    std::string numFmts_content = xml_content.substr(numFmts_start, numFmts_end - numFmts_start);
    
    // 解析每个numFmt
    size_t pos = 0;
    while ((pos = numFmts_content.find("<numFmt ", pos)) != std::string::npos) {
        size_t end_pos = numFmts_content.find("/>", pos);
        if (end_pos == std::string::npos) {
            break;
        }
        
        std::string numFmt_xml = numFmts_content.substr(pos, end_pos - pos + 2);
        
        // 提取numFmtId
        int numFmtId = extractIntAttribute(numFmt_xml, "numFmtId");
        std::string formatCode = extractStringAttribute(numFmt_xml, "formatCode");
        
        if (numFmtId >= 0 && !formatCode.empty()) {
            number_formats_[numFmtId] = formatCode;
        }
        
        pos = end_pos + 2;
    }
}

void StylesParser::parseFonts(const std::string& xml_content) {
    size_t fonts_start = xml_content.find("<fonts");
    if (fonts_start == std::string::npos) {
        return;
    }
    
    size_t fonts_end = xml_content.find("</fonts>", fonts_start);
    if (fonts_end == std::string::npos) {
        return;
    }
    
    std::string fonts_content = xml_content.substr(fonts_start, fonts_end - fonts_start);
    
    // 解析每个font
    size_t pos = 0;
    while ((pos = fonts_content.find("<font", pos)) != std::string::npos) {
        size_t font_end = fonts_content.find("</font>", pos);
        if (font_end == std::string::npos) {
            // 尝试查找自闭合标签
            size_t self_close = fonts_content.find("/>", pos);
            if (self_close != std::string::npos) {
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        std::string font_xml = fonts_content.substr(pos, font_end - pos + 7);
        
        FontInfo font;
        
        // 解析字体名称
        size_t name_pos = font_xml.find("<name ");
        if (name_pos != std::string::npos) {
            font.name = extractStringAttribute(font_xml.substr(name_pos), "val");
        }
        
        // 解析字体大小
        size_t sz_pos = font_xml.find("<sz ");
        if (sz_pos != std::string::npos) {
            font.size = extractDoubleAttribute(font_xml.substr(sz_pos), "val");
        }
        
        // 解析字体样式
        font.bold = font_xml.find("<b") != std::string::npos || font_xml.find("<b/>") != std::string::npos;
        font.italic = font_xml.find("<i") != std::string::npos || font_xml.find("<i/>") != std::string::npos;
        font.underline = font_xml.find("<u") != std::string::npos;
        font.strikeout = font_xml.find("<strike") != std::string::npos;
        
        // 解析字体颜色
        size_t color_pos = font_xml.find("<color ");
        if (color_pos != std::string::npos) {
            font.color = parseColor(font_xml.substr(color_pos));
        }
        
        fonts_.push_back(font);
        pos = font_end + 7;
    }
}

void StylesParser::parseFills(const std::string& xml_content) {
    size_t fills_start = xml_content.find("<fills");
    if (fills_start == std::string::npos) {
        return;
    }
    
    size_t fills_end = xml_content.find("</fills>", fills_start);
    if (fills_end == std::string::npos) {
        return;
    }
    
    std::string fills_content = xml_content.substr(fills_start, fills_end - fills_start);
    
    // 解析每个fill
    size_t pos = 0;
    while ((pos = fills_content.find("<fill", pos)) != std::string::npos) {
        size_t fill_end = fills_content.find("</fill>", pos);
        if (fill_end == std::string::npos) {
            size_t self_close = fills_content.find("/>", pos);
            if (self_close != std::string::npos) {
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        std::string fill_xml = fills_content.substr(pos, fill_end - pos + 7);
        
        FillInfo fill;
        
        // 解析patternFill
        size_t pattern_pos = fill_xml.find("<patternFill ");
        if (pattern_pos != std::string::npos) {
            fill.pattern_type = extractStringAttribute(fill_xml.substr(pattern_pos), "patternType");
            
            // 解析前景色
            size_t fgColor_pos = fill_xml.find("<fgColor ", pattern_pos);
            if (fgColor_pos != std::string::npos) {
                fill.fg_color = parseColor(fill_xml.substr(fgColor_pos));
            }
            
            // 解析背景色
            size_t bgColor_pos = fill_xml.find("<bgColor ", pattern_pos);
            if (bgColor_pos != std::string::npos) {
                fill.bg_color = parseColor(fill_xml.substr(bgColor_pos));
            }
        }
        
        fills_.push_back(fill);
        pos = fill_end + 7;
    }
}

void StylesParser::parseBorders(const std::string& xml_content) {
    size_t borders_start = xml_content.find("<borders");
    if (borders_start == std::string::npos) {
        return;
    }
    
    size_t borders_end = xml_content.find("</borders>", borders_start);
    if (borders_end == std::string::npos) {
        return;
    }
    
    std::string borders_content = xml_content.substr(borders_start, borders_end - borders_start);
    
    // 解析每个border
    size_t pos = 0;
    while ((pos = borders_content.find("<border", pos)) != std::string::npos) {
        size_t border_end = borders_content.find("</border>", pos);
        if (border_end == std::string::npos) {
            size_t self_close = borders_content.find("/>", pos);
            if (self_close != std::string::npos) {
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        std::string border_xml = borders_content.substr(pos, border_end - pos + 9);
        
        BorderInfo border;
        
        // 解析各边框
        border.left = parseBorderSide(border_xml, "left");
        border.right = parseBorderSide(border_xml, "right");
        border.top = parseBorderSide(border_xml, "top");
        border.bottom = parseBorderSide(border_xml, "bottom");
        
        borders_.push_back(border);
        pos = border_end + 9;
    }
}

void StylesParser::parseCellXfs(const std::string& xml_content) {
    size_t cellXfs_start = xml_content.find("<cellXfs");
    if (cellXfs_start == std::string::npos) {
        return;
    }
    
    size_t cellXfs_end = xml_content.find("</cellXfs>", cellXfs_start);
    if (cellXfs_end == std::string::npos) {
        return;
    }
    
    std::string cellXfs_content = xml_content.substr(cellXfs_start, cellXfs_end - cellXfs_start);
    
    // 解析每个xf
    size_t pos = 0;
    while ((pos = cellXfs_content.find("<xf ", pos)) != std::string::npos) {
        size_t xf_end = cellXfs_content.find("/>", pos);
        if (xf_end == std::string::npos) {
            // 查找完整的xf标签
            xf_end = cellXfs_content.find("</xf>", pos);
            if (xf_end == std::string::npos) {
                break;
            }
            xf_end += 5;
        } else {
            xf_end += 2;
        }
        
        std::string xf_xml = cellXfs_content.substr(pos, xf_end - pos);
        
        CellXf xf;
        xf.num_fmt_id = extractIntAttribute(xf_xml, "numFmtId");
        xf.font_id = extractIntAttribute(xf_xml, "fontId");
        xf.fill_id = extractIntAttribute(xf_xml, "fillId");
        xf.border_id = extractIntAttribute(xf_xml, "borderId");
        
        // 解析对齐信息
        size_t alignment_pos = xf_xml.find("<alignment ");
        if (alignment_pos != std::string::npos) {
            xf.horizontal_alignment = extractStringAttribute(xf_xml.substr(alignment_pos), "horizontal");
            xf.vertical_alignment = extractStringAttribute(xf_xml.substr(alignment_pos), "vertical");
            xf.wrap_text = extractStringAttribute(xf_xml.substr(alignment_pos), "wrapText") == "1";
        }
        
        cell_xfs_.push_back(xf);
        pos = xf_end;
    }
}

StylesParser::BorderSide StylesParser::parseBorderSide(const std::string& border_xml, const std::string& side) {
    BorderSide borderSide;
    
    std::string tag = "<" + side;
    size_t side_pos = border_xml.find(tag);
    if (side_pos == std::string::npos) {
        return borderSide;
    }
    
    size_t side_end = border_xml.find("</" + side + ">", side_pos);
    if (side_end == std::string::npos) {
        // 尝试查找自闭合标签
        size_t self_close = border_xml.find("/>", side_pos);
        if (self_close != std::string::npos) {
            std::string side_xml = border_xml.substr(side_pos, self_close - side_pos + 2);
            borderSide.style = extractStringAttribute(side_xml, "style");
            return borderSide;
        }
        return borderSide;
    }
    
    std::string side_xml = border_xml.substr(side_pos, side_end - side_pos + side.length() + 3);
    borderSide.style = extractStringAttribute(side_xml, "style");
    
    // 解析颜色
    size_t color_pos = side_xml.find("<color ");
    if (color_pos != std::string::npos) {
        borderSide.color = parseColor(side_xml.substr(color_pos));
    }
    
    return borderSide;
}

core::Color StylesParser::parseColor(const std::string& color_xml) {
    // 解析rgb属性
    std::string rgb = extractStringAttribute(color_xml, "rgb");
    if (!rgb.empty() && rgb.length() >= 6) {
        // 移除可能的前缀（如"FF"表示alpha）
        if (rgb.length() == 8) {
            rgb = rgb.substr(2); // 移除前两位alpha
        }
        
        try {
            unsigned int color_value = std::stoul(rgb, nullptr, 16);
            int r = (color_value >> 16) & 0xFF;
            int g = (color_value >> 8) & 0xFF;
            int b = color_value & 0xFF;
            return core::Color(static_cast<uint32_t>((r << 16) | (g << 8) | b));
        } catch (const std::exception& e) {
            // 解析失败，返回无效颜色
        }
    }
    
    // 解析theme属性
    int theme = extractIntAttribute(color_xml, "theme");
    if (theme >= 0) {
        return core::Color(static_cast<uint8_t>(theme)); // 主题颜色构造函数
    }
    
    // 解析indexed属性
    int indexed = extractIntAttribute(color_xml, "indexed");
    if (indexed >= 0) {
        return core::Color::fromIndex(static_cast<uint8_t>(indexed));
    }
    
    return core::Color(); // 返回无效颜色
}

int StylesParser::extractIntAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return -1;
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return -1;
    }
    
    try {
        return std::stoi(xml.substr(start, end - start));
    } catch (const std::exception& e) {
        return -1;
    }
}

double StylesParser::extractDoubleAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return -1.0;
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return -1.0;
    }
    
    try {
        return std::stod(xml.substr(start, end - start));
    } catch (const std::exception& e) {
        return -1.0;
    }
}

std::string StylesParser::extractStringAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return "";
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return "";
    }
    
    return xml.substr(start, end - start);
}

core::HorizontalAlign StylesParser::getAlignment(const std::string& alignment) const {
    if (alignment == "left") return core::HorizontalAlign::Left;
    if (alignment == "center") return core::HorizontalAlign::Center;
    if (alignment == "right") return core::HorizontalAlign::Right;
    if (alignment == "justify") return core::HorizontalAlign::Justify;
    if (alignment == "fill") return core::HorizontalAlign::Fill;
    return core::HorizontalAlign::None;
}

core::VerticalAlign StylesParser::getVerticalAlignment(const std::string& alignment) const {
    if (alignment == "top") return core::VerticalAlign::Top;
    if (alignment == "center") return core::VerticalAlign::Center;
    if (alignment == "bottom") return core::VerticalAlign::Bottom;
    if (alignment == "justify") return core::VerticalAlign::Justify;
    return core::VerticalAlign::Top;
}

core::BorderStyle StylesParser::getBorderStyle(const std::string& style) const {
    if (style == "thin") return core::BorderStyle::Thin;
    if (style == "medium") return core::BorderStyle::Medium;
    if (style == "thick") return core::BorderStyle::Thick;
    if (style == "double") return core::BorderStyle::Double;
    if (style == "dotted") return core::BorderStyle::Dotted;
    if (style == "dashed") return core::BorderStyle::Dashed;
    if (style == "dashDot") return core::BorderStyle::DashDot;
    if (style == "dashDotDot") return core::BorderStyle::DashDotDot;
    return core::BorderStyle::None;
}

std::string StylesParser::getBuiltinNumberFormat(int format_id) const {
    // Excel内置数字格式
    static const std::unordered_map<int, std::string> builtin_formats = {
        {0, "General"},
        {1, "0"},
        {2, "0.00"},
        {3, "#,##0"},
        {4, "#,##0.00"},
        {9, "0%"},
        {10, "0.00%"},
        {11, "0.00E+00"},
        {12, "# ?/?"},
        {13, "# ??/??"},
        {14, "mm-dd-yy"},
        {15, "d-mmm-yy"},
        {16, "d-mmm"},
        {17, "mmm-yy"},
        {18, "h:mm AM/PM"},
        {19, "h:mm:ss AM/PM"},
        {20, "h:mm"},
        {21, "h:mm:ss"},
        {22, "m/d/yy h:mm"},
        {37, "#,##0 ;(#,##0)"},
        {38, "#,##0 ;[Red](#,##0)"},
        {39, "#,##0.00;(#,##0.00)"},
        {40, "#,##0.00;[Red](#,##0.00)"}
    };
    
    auto it = builtin_formats.find(format_id);
    if (it != builtin_formats.end()) {
        return it->second;
    }
    
    return "General";
}

} // namespace reader
} // namespace fastexcel
