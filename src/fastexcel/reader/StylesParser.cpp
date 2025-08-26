#include "StylesParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/StyleBuilder.hpp"

namespace fastexcel {
namespace reader {

void StylesParser::onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int /*depth*/) {
    if (!state_.collecting_region) {
        // 顶级区域检测
        if (name == "numFmts") {
            state_.startRegion(ParseState::Region::NumFmts);
        }
        else if (name == "fonts") {
            state_.startRegion(ParseState::Region::Fonts);
        }
        else if (name == "fills") {
            state_.startRegion(ParseState::Region::Fills);
        }
        else if (name == "borders") {
            state_.startRegion(ParseState::Region::Borders);
        }
        else if (name == "cellXfs") {
            state_.startRegion(ParseState::Region::CellXfs);
        }
        return;
    }
    
    // 收集区域内的XML
    if (state_.collecting_region) {
        state_.region_xml_buffer += std::string("<") + std::string(name);
        for (const auto& attr : attributes) {
            state_.region_xml_buffer += " " + std::string(attr.name) + "=\"" + std::string(attr.value) + "\"";
        }
        state_.region_xml_buffer += '>';
        state_.region_depth++;
    }
}

void StylesParser::onEndElement(std::string_view name, int /*depth*/) {
    if (state_.collecting_region) {
        state_.region_depth--;
        
        // 检查是否到达区域结束
        if (state_.region_depth == 0) {
            // 处理完整的区域XML
            switch (state_.current_region) {
                case ParseState::Region::NumFmts:
                    processNumFmtsRegion(state_.region_xml_buffer);
                    break;
                case ParseState::Region::Fonts:
                    processFontsRegion(state_.region_xml_buffer);
                    break;
                case ParseState::Region::Fills:
                    processFillsRegion(state_.region_xml_buffer);
                    break;
                case ParseState::Region::Borders:
                    processBordersRegion(state_.region_xml_buffer);
                    break;
                case ParseState::Region::CellXfs:
                    processCellXfsRegion(state_.region_xml_buffer);
                    break;
                default:
                    break;
            }
            state_.endRegion();
        } else {
            // 区域内部元素结束
            state_.region_xml_buffer += "</" + std::string(name) + ">";
        }
    }
}

void StylesParser::onText(std::string_view text, int /*depth*/) {
    if (state_.collecting_region && !text.empty()) {
        // XML实体转义处理
        std::string escaped_text(text);
        // 基本XML实体转义
        size_t pos = 0;
        while ((pos = escaped_text.find('&', pos)) != std::string::npos) {
            escaped_text.replace(pos, 1, "&amp;");
            pos += 5;
        }
        pos = 0;
        while ((pos = escaped_text.find('<', pos)) != std::string::npos) {
            escaped_text.replace(pos, 1, "&lt;");
            pos += 4;
        }
        pos = 0;
        while ((pos = escaped_text.find('>', pos)) != std::string::npos) {
            escaped_text.replace(pos, 1, "&gt;");
            pos += 4;
        }
        
        state_.region_xml_buffer += escaped_text;
    }
}

// ==================== 区域处理方法 ====================

void StylesParser::processNumFmtsRegion(std::string_view region_xml) {
    const char* p = region_xml.data();
    const char* end = p + region_xml.size();
    
    // 查找所有 <numFmt> 元素
    while (p < end) {
        const char* numfmt_start = std::search(p, end, 
            reinterpret_cast<const char*>("<numFmt"), 
            reinterpret_cast<const char*>("<numFmt") + 7);
        
        if (numfmt_start == end) break;
        
        const char* numfmt_end = std::find(numfmt_start, end, '>');
        if (numfmt_end == end) break;
        
        // 提取 numFmtId 和 formatCode
        std::string numfmt_id_str;
        std::string format_code;
        
        if (findAttributeInElement(numfmt_start, numfmt_end, "numFmtId", numfmt_id_str) &&
            findAttributeInElement(numfmt_start, numfmt_end, "formatCode", format_code)) {
            
            try {
                int num_fmt_id = std::stoi(numfmt_id_str);
                number_formats_[num_fmt_id] = format_code;
            } catch (const std::exception&) {
                // 忽略无效的数字格式
            }
        }
        
        p = numfmt_end + 1;
    }
}

void StylesParser::processFontsRegion(std::string_view region_xml) {
    const char* p = region_xml.data();
    const char* end = p + region_xml.size();
    
    // 预分配合理大小
    fonts_.reserve(128);
    
    // 查找所有 <font> 元素
    while (p < end) {
        const char* font_start = std::search(p, end,
            reinterpret_cast<const char*>("<font"),
            reinterpret_cast<const char*>("<font") + 5);
        
        if (font_start == end) break;
        
        // 查找对应的 </font>
        const char* font_end = std::search(font_start, end,
            reinterpret_cast<const char*>("</font>"),
            reinterpret_cast<const char*>("</font>") + 7);
        
        if (font_end == end) break;
        
        FontInfo font;
        
        // 解析字体名称
        auto name_content = extractElementContent(font_start, "name");
        if (!name_content.empty()) {
            std::string name_val;
            if (findAttributeInElement(name_content.data(), name_content.data() + name_content.size(), "val", name_val)) {
                font.name = name_val;
            }
        }
        
        // 解析字体大小
        auto size_content = extractElementContent(font_start, "sz");
        if (!size_content.empty()) {
            std::string size_val;
            if (findAttributeInElement(size_content.data(), size_content.data() + size_content.size(), "val", size_val)) {
                try {
                    font.size = std::stod(size_val);
                } catch (const std::exception&) {}
            }
        }
        
        // 解析字体样式（bold, italic, underline, strikeout）
        font.bold = std::search(font_start, font_end,
            reinterpret_cast<const char*>("<b"),
            reinterpret_cast<const char*>("<b") + 2) != font_end;
            
        font.italic = std::search(font_start, font_end,
            reinterpret_cast<const char*>("<i"),
            reinterpret_cast<const char*>("<i") + 2) != font_end;
            
        font.underline = std::search(font_start, font_end,
            reinterpret_cast<const char*>("<u"),
            reinterpret_cast<const char*>("<u") + 2) != font_end;
            
        font.strikeout = std::search(font_start, font_end,
            reinterpret_cast<const char*>("<strike"),
            reinterpret_cast<const char*>("<strike") + 7) != font_end;
        
        // 解析颜色
        auto color_content = extractElementContent(font_start, "color");
        if (!color_content.empty()) {
            const char* color_p = color_content.data();
            const char* color_end_p = color_p + color_content.size();
            parseColorAttribute(color_p, color_end_p, font.color);
        }
        
        fonts_.emplace_back(std::move(font));
        p = font_end + 7; // 跳过 </font>
    }
}

void StylesParser::processFillsRegion(std::string_view region_xml) {
    const char* p = region_xml.data();
    const char* end = p + region_xml.size();
    
    // 预分配合理大小
    fills_.reserve(128);
    
    // 查找所有 <fill> 元素
    while (p < end) {
        const char* fill_start = std::search(p, end,
            reinterpret_cast<const char*>("<fill"),
            reinterpret_cast<const char*>("<fill") + 5);
        
        if (fill_start == end) break;
        
        // 查找对应的 </fill>
        const char* fill_end = std::search(fill_start, end,
            reinterpret_cast<const char*>("</fill>"),
            reinterpret_cast<const char*>("</fill>") + 7);
        
        if (fill_end == end) break;
        
        FillInfo fill;
        
        // 解析 patternFill
        auto pattern_content = extractElementContent(fill_start, "patternFill");
        if (!pattern_content.empty()) {
            const char* pattern_start = pattern_content.data();
            const char* pattern_end = pattern_start + pattern_content.size();
            
            // 提取 patternType
            std::string pattern_type;
            if (findAttributeInElement(pattern_start, pattern_end, "patternType", pattern_type)) {
                fill.pattern_type = pattern_type;
            }
            
            // 解析前景色
            auto fg_content = extractElementContent(pattern_start, "fgColor");
            if (!fg_content.empty()) {
                const char* fg_p = fg_content.data();
                const char* fg_end_p = fg_p + fg_content.size();
                parseColorAttribute(fg_p, fg_end_p, fill.fg_color);
            }
            
            // 解析背景色
            auto bg_content = extractElementContent(pattern_start, "bgColor");
            if (!bg_content.empty()) {
                const char* bg_p = bg_content.data();
                const char* bg_end_p = bg_p + bg_content.size();
                parseColorAttribute(bg_p, bg_end_p, fill.bg_color);
            }
        }
        
        fills_.emplace_back(std::move(fill));
        p = fill_end + 7; // 跳过 </fill>
    }
}

void StylesParser::processBordersRegion(std::string_view region_xml) {
    const char* p = region_xml.data();
    const char* end = p + region_xml.size();
    
    // 预分配合理大小
    borders_.reserve(128);
    
    // 查找所有 <border> 元素
    while (p < end) {
        const char* border_start = std::search(p, end,
            reinterpret_cast<const char*>("<border"),
            reinterpret_cast<const char*>("<border") + 7);
        
        if (border_start == end) break;
        
        // 查找对应的 </border>
        const char* border_end = std::search(border_start, end,
            reinterpret_cast<const char*>("</border>"),
            reinterpret_cast<const char*>("</border>") + 9);
        
        if (border_end == end) break;
        
        BorderInfo border;
        
        // 解析各个边框
        auto parseBoderSide = [&](const char* side_name, BorderSide& side) {
            auto side_content = extractElementContent(border_start, side_name);
            if (!side_content.empty()) {
                const char* side_start = side_content.data();
                const char* side_end = side_start + side_content.size();
                
                // 提取 style
                std::string style;
                if (findAttributeInElement(side_start, side_end, "style", style)) {
                    side.style = style;
                }
                
                // 解析颜色
                auto color_content = extractElementContent(side_start, "color");
                if (!color_content.empty()) {
                    const char* color_p = color_content.data();
                    const char* color_end_p = color_p + color_content.size();
                    parseColorAttribute(color_p, color_end_p, side.color);
                }
            }
        };
        
        parseBoderSide("left", border.left);
        parseBoderSide("right", border.right);
        parseBoderSide("top", border.top);
        parseBoderSide("bottom", border.bottom);
        parseBoderSide("diagonal", border.diagonal);
        
        borders_.emplace_back(std::move(border));
        p = border_end + 9; // 跳过 </border>
    }
}

void StylesParser::processCellXfsRegion(std::string_view region_xml) {
    const char* p = region_xml.data();
    const char* end = p + region_xml.size();
    
    // 预分配合理大小
    cell_xfs_.reserve(1024);
    
    // 查找所有 <xf> 元素
    while (p < end) {
        const char* xf_start = std::search(p, end,
            reinterpret_cast<const char*>("<xf"),
            reinterpret_cast<const char*>("<xf") + 3);
        
        if (xf_start == end) break;
        
        const char* xf_end = std::find(xf_start, end, '>');
        if (xf_end == end) break;
        
        CellXf xf;
        
        // 提取基本属性
        std::string attr_val;
        if (findAttributeInElement(xf_start, xf_end, "numFmtId", attr_val)) {
            try {
                xf.num_fmt_id = std::stoi(attr_val);
            } catch (const std::exception&) {}
        }
        
        if (findAttributeInElement(xf_start, xf_end, "fontId", attr_val)) {
            try {
                xf.font_id = std::stoi(attr_val);
            } catch (const std::exception&) {}
        }
        
        if (findAttributeInElement(xf_start, xf_end, "fillId", attr_val)) {
            try {
                xf.fill_id = std::stoi(attr_val);
            } catch (const std::exception&) {}
        }
        
        if (findAttributeInElement(xf_start, xf_end, "borderId", attr_val)) {
            try {
                xf.border_id = std::stoi(attr_val);
            } catch (const std::exception&) {}
        }
        
        // 查找对应的 </xf> 或者检查是否是自闭合标签
        const char* close_tag = std::search(xf_end + 1, end,
            reinterpret_cast<const char*>("</xf>"),
            reinterpret_cast<const char*>("</xf>") + 5);
        
        if (close_tag != end) {
            // 解析对齐方式
            auto alignment_content = extractElementContent(xf_start, "alignment");
            if (!alignment_content.empty()) {
                const char* align_start = alignment_content.data();
                const char* align_end = align_start + alignment_content.size();
                
                if (findAttributeInElement(align_start, align_end, "horizontal", attr_val)) {
                    xf.horizontal_alignment = attr_val;
                }
                
                if (findAttributeInElement(align_start, align_end, "vertical", attr_val)) {
                    xf.vertical_alignment = attr_val;
                }
                
                if (findAttributeInElement(align_start, align_end, "wrapText", attr_val)) {
                    xf.wrap_text = (attr_val == "1" || attr_val == "true");
                }
                
                if (findAttributeInElement(align_start, align_end, "indent", attr_val)) {
                    try {
                        xf.indent = std::stoi(attr_val);
                    } catch (const std::exception&) {}
                }
                
                if (findAttributeInElement(align_start, align_end, "textRotation", attr_val)) {
                    try {
                        xf.text_rotation = std::stoi(attr_val);
                    } catch (const std::exception&) {}
                }
            }
            
            p = close_tag + 5; // 跳过 </xf>
        } else {
            p = xf_end + 1;
        }
        
        cell_xfs_.emplace_back(std::move(xf));
    }
}

// ==================== 指针扫描工具方法 ====================

void StylesParser::parseColorAttribute(const char*& p, const char* end, core::Color& color) {
    std::string attr_val;
    
    // 查找 rgb 属性
    if (findAttributeInElement(p, end, "rgb", attr_val) && !attr_val.empty()) {
        color = core::Color::fromHex(attr_val);
        return;
    }
    
    // 查找 theme 属性
    if (findAttributeInElement(p, end, "theme", attr_val) && !attr_val.empty()) {
        try {
            int theme_id = std::stoi(attr_val);
            color = core::Color::fromTheme(static_cast<uint8_t>(theme_id));
            return;
        } catch (const std::exception&) {}
    }
    
    // 查找 indexed 属性
    if (findAttributeInElement(p, end, "indexed", attr_val) && !attr_val.empty()) {
        try {
            int index = std::stoi(attr_val);
            color = core::Color::fromIndex(static_cast<uint8_t>(index));
            return;
        } catch (const std::exception&) {}
    }
    
    // 默认颜色
    color = core::Color();
}

std::string_view StylesParser::extractElementContent(const char* xml, const char* element_name) {
    // 查找开始标签
    std::string start_tag = "<" + std::string(element_name);
    const char* tag_start = std::search(xml, xml + std::strlen(xml),
        start_tag.c_str(), start_tag.c_str() + start_tag.length());
    
    if (tag_start == xml + std::strlen(xml)) {
        return std::string_view();
    }
    
    // 找到开始标签的结束
    const char* content_start = std::find(tag_start, xml + std::strlen(xml), '>');
    if (content_start == xml + std::strlen(xml)) {
        return std::string_view();
    }
    content_start++; // 跳过 '>'
    
    // 查找结束标签
    std::string end_tag = "</" + std::string(element_name) + ">";
    const char* content_end = std::search(content_start, xml + std::strlen(xml),
        end_tag.c_str(), end_tag.c_str() + end_tag.length());
    
    if (content_end == xml + std::strlen(xml)) {
        // 可能是自闭合标签，从开始标签返回
        const char* self_close = std::find(tag_start, xml + std::strlen(xml), '>');
        if (self_close != xml + std::strlen(xml) && *(self_close - 1) == '/') {
            return std::string_view(tag_start, self_close + 1 - tag_start);
        }
        return std::string_view();
    }
    
    return std::string_view(tag_start, content_end + end_tag.length() - tag_start);
}

bool StylesParser::findAttributeInElement(const char* element_start, const char* element_end, 
                                        const char* attr_name, std::string& out_value) {
    size_t attr_len = std::strlen(attr_name);
    
    // 查找属性名
    const char* attr_pos = std::search(element_start, element_end,
        attr_name, attr_name + attr_len);
    
    if (attr_pos == element_end) {
        return false;
    }
    
    // 跳过属性名和可能的空格，查找 '='
    const char* eq_pos = attr_pos + attr_len;
    while (eq_pos < element_end && (*eq_pos == ' ' || *eq_pos == '\t' || *eq_pos == '\n' || *eq_pos == '\r')) {
        eq_pos++;
    }
    
    if (eq_pos >= element_end || *eq_pos != '=') {
        return false;
    }
    
    eq_pos++; // 跳过 '='
    
    // 跳过等号后的空格，查找引号
    while (eq_pos < element_end && (*eq_pos == ' ' || *eq_pos == '\t' || *eq_pos == '\n' || *eq_pos == '\r')) {
        eq_pos++;
    }
    
    if (eq_pos >= element_end || (*eq_pos != '"' && *eq_pos != '\'')) {
        return false;
    }
    
    char quote_char = *eq_pos;
    const char* value_start = eq_pos + 1;
    
    // 查找结束引号
    const char* value_end = std::find(value_start, element_end, quote_char);
    if (value_end == element_end) {
        return false;
    }
    
    out_value.assign(value_start, value_end - value_start);
    return true;
}

// ==================== 公共接口实现 ====================

std::shared_ptr<core::FormatDescriptor> StylesParser::getFormat(int xf_index) const {
    if (xf_index < 0 || xf_index >= static_cast<int>(cell_xfs_.size())) {
        return nullptr;
    }
    
    const CellXf& xf = cell_xfs_[xf_index];
    core::StyleBuilder builder;
    
    // 设置数字格式
    if (xf.num_fmt_id >= 0) {
        auto fmt_it = number_formats_.find(xf.num_fmt_id);
        if (fmt_it != number_formats_.end()) {
            builder.numberFormat(fmt_it->second);
        } else {
            builder.numberFormat(getBuiltinNumberFormat(xf.num_fmt_id));
        }
    }
    
    // 设置字体
    if (xf.font_id >= 0 && xf.font_id < static_cast<int>(fonts_.size())) {
        const FontInfo& font = fonts_[xf.font_id];
        builder.fontName(font.name)
               .fontSize(font.size)
               .bold(font.bold)
               .italic(font.italic)
               .strikeout(font.strikeout)
               .fontColor(font.color);
        
        if (font.underline) {
            builder.underline(core::UnderlineType::Single);
        }
    }
    
    // 设置填充
    if (xf.fill_id >= 0 && xf.fill_id < static_cast<int>(fills_.size())) {
        const FillInfo& fill = fills_[xf.fill_id];
        builder.fill(getPatternType(fill.pattern_type), fill.fg_color, fill.bg_color);
    }
    
    // 设置边框
    if (xf.border_id >= 0 && xf.border_id < static_cast<int>(borders_.size())) {
        const BorderInfo& border = borders_[xf.border_id];
        
        if (!border.left.style.empty()) {
            builder.leftBorder(getBorderStyle(border.left.style), border.left.color);
        }
        if (!border.right.style.empty()) {
            builder.rightBorder(getBorderStyle(border.right.style), border.right.color);
        }
        if (!border.top.style.empty()) {
            builder.topBorder(getBorderStyle(border.top.style), border.top.color);
        }
        if (!border.bottom.style.empty()) {
            builder.bottomBorder(getBorderStyle(border.bottom.style), border.bottom.color);
        }
    }
    
    // 设置对齐
    builder.horizontalAlign(getAlignment(xf.horizontal_alignment))
           .verticalAlign(getVerticalAlignment(xf.vertical_alignment))
           .textWrap(xf.wrap_text)
           .indent(static_cast<uint8_t>(std::max(0, xf.indent)))
           .rotation(static_cast<int16_t>(xf.text_rotation));
    
    return std::make_shared<core::FormatDescriptor>(builder.build());
}

// ==================== 枚举转换方法 ====================

core::HorizontalAlign StylesParser::getAlignment(const std::string& alignment) const {
    if (alignment == "left") return core::HorizontalAlign::Left;
    if (alignment == "center") return core::HorizontalAlign::Center;
    if (alignment == "right") return core::HorizontalAlign::Right;
    if (alignment == "fill") return core::HorizontalAlign::Fill;
    if (alignment == "justify") return core::HorizontalAlign::Justify;
    if (alignment == "centerContinuous") return core::HorizontalAlign::CenterAcross;
    if (alignment == "distributed") return core::HorizontalAlign::Distributed;
    return core::HorizontalAlign::None;
}

core::VerticalAlign StylesParser::getVerticalAlignment(const std::string& alignment) const {
    if (alignment == "top") return core::VerticalAlign::Top;
    if (alignment == "center") return core::VerticalAlign::Center;
    if (alignment == "bottom") return core::VerticalAlign::Bottom;
    if (alignment == "justify") return core::VerticalAlign::Justify;
    if (alignment == "distributed") return core::VerticalAlign::Distributed;
    return core::VerticalAlign::Bottom;
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
    if (style == "slantDashDot") return core::BorderStyle::SlantDashDot;
    return core::BorderStyle::None;
}

core::PatternType StylesParser::getPatternType(const std::string& pattern) const {
    if (pattern == "solid") return core::PatternType::Solid;
    if (pattern == "darkGray") return core::PatternType::DarkGray;
    if (pattern == "mediumGray") return core::PatternType::MediumGray;
    if (pattern == "lightGray") return core::PatternType::LightGray;
    if (pattern == "gray125") return core::PatternType::Gray125;
    if (pattern == "gray0625") return core::PatternType::Gray0625;
    if (pattern == "darkHorizontal") return core::PatternType::DarkHorizontal;
    if (pattern == "darkVertical") return core::PatternType::DarkVertical;
    if (pattern == "darkDown") return core::PatternType::DarkDown;
    if (pattern == "darkUp") return core::PatternType::DarkUp;
    if (pattern == "darkGrid") return core::PatternType::DarkGrid;
    if (pattern == "darkTrellis") return core::PatternType::DarkTrellis;
    if (pattern == "lightHorizontal") return core::PatternType::LightHorizontal;
    if (pattern == "lightVertical") return core::PatternType::LightVertical;
    if (pattern == "lightDown") return core::PatternType::LightDown;
    if (pattern == "lightUp") return core::PatternType::LightUp;
    if (pattern == "lightGrid") return core::PatternType::LightGrid;
    if (pattern == "lightTrellis") return core::PatternType::LightTrellis;
    return core::PatternType::None;
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
        {40, "#,##0.00;[Red](#,##0.00)"},
        {45, "mm:ss"},
        {46, "[h]:mm:ss"},
        {47, "mmss.0"},
        {48, "##0.0E+0"},
        {49, "@"}
    };
    
    auto it = builtin_formats.find(format_id);
    return (it != builtin_formats.end()) ? it->second : "General";
}

} // namespace reader
} // namespace fastexcel