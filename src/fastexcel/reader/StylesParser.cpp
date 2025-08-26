#include "StylesParser.hpp"
#include "fastexcel/core/StyleBuilder.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <cctype>

namespace fastexcel {
namespace reader {

void StylesParser::onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
    using Context = ParseState::Context;
    
    if (name == "numFmts") {
        parse_state_.pushContext(Context::NumFmts);
    }
    else if (name == "numFmt" && parse_state_.getCurrentContext() == Context::NumFmts) {
        // 解析数字格式
        auto numFmtId = findIntAttribute(attributes, "numFmtId");
        auto formatCode = findAttribute(attributes, "formatCode");
        
        if (numFmtId && formatCode && *numFmtId >= 0) {
            number_formats_[*numFmtId] = *formatCode;
        }
    }
    
    // === 字体解析 ===
    else if (name == "fonts") {
        parse_state_.pushContext(Context::Fonts);
    }
    else if (name == "font" && parse_state_.getCurrentContext() == Context::Fonts) {
        fonts_.emplace_back();
        parse_state_.current_font = &fonts_.back();
        parse_state_.pushContext(Context::Font);
    }
    else if (parse_state_.getCurrentContext() == Context::Font) {
        if (name == "name") {
            parse_state_.pushContext(Context::FontName);
        }
        else if (name == "sz") {
            parse_state_.pushContext(Context::FontSize);
        }
        else if (name == "color") {
            parse_state_.pushContext(Context::FontColor);
        }
        else if (name == "b") {
            parse_state_.current_font->bold = true;
        }
        else if (name == "i") {
            parse_state_.current_font->italic = true;
        }
        else if (name == "u") {
            parse_state_.current_font->underline = true;
        }
        else if (name == "strike") {
            parse_state_.current_font->strikeout = true;
        }
    }
    
    // === 填充解析 ===
    else if (name == "fills") {
        parse_state_.pushContext(Context::Fills);
    }
    else if (name == "fill" && parse_state_.getCurrentContext() == Context::Fills) {
        fills_.emplace_back();
        parse_state_.current_fill = &fills_.back();
        parse_state_.pushContext(Context::Fill);
    }
    else if (name == "patternFill" && parse_state_.getCurrentContext() == Context::Fill) {
        parse_state_.pushContext(Context::PatternFill);
        if (auto patternType = findAttribute(attributes, "patternType")) {
            parse_state_.current_fill->pattern_type = *patternType;
        }
    }
    else if (parse_state_.getCurrentContext() == Context::PatternFill) {
        if (name == "fgColor") {
            parse_state_.pushContext(Context::FgColor);
            // 解析前景色
            if (auto rgb = findAttribute(attributes, "rgb")) {
                parse_state_.current_fill->fg_color = core::Color::fromHex(*rgb);
            }
            else if (auto theme = findIntAttribute(attributes, "theme")) {
                // 处理主题色
                parse_state_.current_fill->fg_color = core::Color::fromTheme(*theme);
            }
            else if (auto indexed = findIntAttribute(attributes, "indexed")) {
                // 处理索引色
                parse_state_.current_fill->fg_color = core::Color::fromIndex(*indexed);
            }
        }
        else if (name == "bgColor") {
            parse_state_.pushContext(Context::BgColor);
            // 解析背景色 
            if (auto rgb = findAttribute(attributes, "rgb")) {
                parse_state_.current_fill->bg_color = core::Color::fromHex(*rgb);
            }
            else if (auto theme = findIntAttribute(attributes, "theme")) {
                parse_state_.current_fill->bg_color = core::Color::fromTheme(*theme);
            }
            else if (auto indexed = findIntAttribute(attributes, "indexed")) {
                parse_state_.current_fill->bg_color = core::Color::fromIndex(*indexed);
            }
        }
    }
    
    // === 边框解析 ===
    else if (name == "borders") {
        parse_state_.pushContext(Context::Borders);
    }
    else if (name == "border" && parse_state_.getCurrentContext() == Context::Borders) {
        borders_.emplace_back();
        parse_state_.current_border = &borders_.back();
        parse_state_.pushContext(Context::Border);
    }
    else if (parse_state_.getCurrentContext() == Context::Border) {
        if (name == "left") {
            parse_state_.pushContext(Context::BorderLeft);
            parse_state_.current_border_side = &parse_state_.current_border->left;
            if (auto style = findAttribute(attributes, "style")) {
                parse_state_.current_border_side->style = *style;
            }
        }
        else if (name == "right") {
            parse_state_.pushContext(Context::BorderRight);
            parse_state_.current_border_side = &parse_state_.current_border->right;
            if (auto style = findAttribute(attributes, "style")) {
                parse_state_.current_border_side->style = *style;
            }
        }
        else if (name == "top") {
            parse_state_.pushContext(Context::BorderTop);
            parse_state_.current_border_side = &parse_state_.current_border->top;
            if (auto style = findAttribute(attributes, "style")) {
                parse_state_.current_border_side->style = *style;
            }
        }
        else if (name == "bottom") {
            parse_state_.pushContext(Context::BorderBottom);
            parse_state_.current_border_side = &parse_state_.current_border->bottom;
            if (auto style = findAttribute(attributes, "style")) {
                parse_state_.current_border_side->style = *style;
            }
        }
        else if (name == "diagonal") {
            parse_state_.pushContext(Context::BorderDiagonal);
            parse_state_.current_border_side = &parse_state_.current_border->diagonal;
            if (auto style = findAttribute(attributes, "style")) {
                parse_state_.current_border_side->style = *style;
            }
        }
    }
    else if (name == "color" && 
             (parse_state_.getCurrentContext() == Context::BorderLeft ||
              parse_state_.getCurrentContext() == Context::BorderRight ||
              parse_state_.getCurrentContext() == Context::BorderTop ||
              parse_state_.getCurrentContext() == Context::BorderBottom ||
              parse_state_.getCurrentContext() == Context::BorderDiagonal)) {
        parse_state_.pushContext(Context::BorderColor);
        // 解析边框颜色
        if (parse_state_.current_border_side) {
            if (auto rgb = findAttribute(attributes, "rgb")) {
                parse_state_.current_border_side->color = core::Color::fromHex(*rgb);
            }
            else if (auto theme = findIntAttribute(attributes, "theme")) {
                parse_state_.current_border_side->color = core::Color::fromTheme(*theme);
            }
            else if (auto indexed = findIntAttribute(attributes, "indexed")) {
                parse_state_.current_border_side->color = core::Color::fromIndex(*indexed);
            }
        }
    }
    
    // === CellXfs 解析 ===
    else if (name == "cellXfs") {
        parse_state_.pushContext(Context::CellXfs);
    }
    else if (name == "xf" && parse_state_.getCurrentContext() == Context::CellXfs) {
        cell_xfs_.emplace_back();
        parse_state_.current_xf = &cell_xfs_.back();
        
        // 解析 xf 属性
        if (auto numFmtId = findIntAttribute(attributes, "numFmtId")) {
            parse_state_.current_xf->num_fmt_id = *numFmtId;
        }
        if (auto fontId = findIntAttribute(attributes, "fontId")) {
            parse_state_.current_xf->font_id = *fontId;
        }
        if (auto fillId = findIntAttribute(attributes, "fillId")) {
            parse_state_.current_xf->fill_id = *fillId;
        }
        if (auto borderId = findIntAttribute(attributes, "borderId")) {
            parse_state_.current_xf->border_id = *borderId;
        }
    }
    else if (name == "alignment" && parse_state_.current_xf) {
        // 解析对齐属性
        if (auto horizontal = findAttribute(attributes, "horizontal")) {
            parse_state_.current_xf->horizontal_alignment = *horizontal;
        }
        if (auto vertical = findAttribute(attributes, "vertical")) {
            parse_state_.current_xf->vertical_alignment = *vertical;
        }
        if (auto wrapText = findAttribute(attributes, "wrapText")) {
            parse_state_.current_xf->wrap_text = (*wrapText == "1" || *wrapText == "true");
        }
        if (auto indent = findIntAttribute(attributes, "indent")) {
            parse_state_.current_xf->indent = *indent;
        }
        if (auto textRotation = findIntAttribute(attributes, "textRotation")) {
            parse_state_.current_xf->text_rotation = *textRotation;
        }
    }
}

void StylesParser::onEndElement(const std::string& name, int depth) {
    using Context = ParseState::Context;
    
    // 根据元素名称弹出对应的上下文
    if (name == "numFmts" || name == "fonts" || name == "fills" || 
        name == "borders" || name == "cellXfs") {
        parse_state_.popContext();
    }
    else if (name == "font") {
        parse_state_.current_font = nullptr;
        parse_state_.popContext();
    }
    else if (name == "fill") {
        parse_state_.current_fill = nullptr;
        parse_state_.popContext();
    }
    else if (name == "border") {
        parse_state_.current_border = nullptr;
        parse_state_.popContext();
    }
    else if (name == "left" || name == "right" || name == "top" || 
             name == "bottom" || name == "diagonal") {
        parse_state_.current_border_side = nullptr;
        parse_state_.popContext();
    }
    else if (name == "patternFill" || name == "name" || name == "sz" || 
             name == "color" || name == "fgColor" || name == "bgColor") {
        parse_state_.popContext();
    }
}

void StylesParser::onText(const std::string& text, int depth) {
    using Context = ParseState::Context;
    
    if (parse_state_.getCurrentContext() == Context::FontName && parse_state_.current_font) {
        parse_state_.current_font->name = text;
    }
    else if (parse_state_.getCurrentContext() == Context::FontSize && parse_state_.current_font) {
        try {
            parse_state_.current_font->size = std::stod(text);
        } catch (...) {
            parse_state_.current_font->size = 11.0; // 默认值
        }
    }
}

std::shared_ptr<core::FormatDescriptor> StylesParser::getFormat(int xf_index) const {
    if (xf_index < 0 || xf_index >= static_cast<int>(cell_xfs_.size())) {
        return nullptr;
    }
    
    const CellXf& xf = cell_xfs_[xf_index];
    core::StyleBuilder builder;
    
    // 设置数字格式
    if (xf.num_fmt_id >= 0) {
        auto it = number_formats_.find(xf.num_fmt_id);
        if (it != number_formats_.end()) {
            builder.numberFormat(it->second);
        } else {
            // 使用内置格式
            std::string builtin = getBuiltinNumberFormat(xf.num_fmt_id);
            if (!builtin.empty()) {
                builder.numberFormat(builtin);
            }
        }
    }
    
    // 设置字体
    if (xf.font_id >= 0 && xf.font_id < static_cast<int>(fonts_.size())) {
        const FontInfo& font = fonts_[xf.font_id];
        builder.font(font.name, font.size)
               .bold(font.bold)
               .italic(font.italic)
               .underline(font.underline ? core::UnderlineType::Single : core::UnderlineType::None)
               .strikeout(font.strikeout)
               .fontColor(font.color);
    }
    
    // 设置填充
    if (xf.fill_id >= 0 && xf.fill_id < static_cast<int>(fills_.size())) {
        const FillInfo& fill = fills_[xf.fill_id];
        if (fill.pattern_type != "none") {
            builder.fill(getPatternType(fill.pattern_type), fill.bg_color, fill.fg_color);
        }
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
    if (!xf.horizontal_alignment.empty()) {
        builder.horizontalAlign(getAlignment(xf.horizontal_alignment));
    }
    if (!xf.vertical_alignment.empty()) {
        builder.verticalAlign(getVerticalAlignment(xf.vertical_alignment));
    }
    if (xf.wrap_text) {
        builder.textWrap(true);
    }
    
    return std::make_shared<core::FormatDescriptor>(builder.build());
}

// 保留原有的枚举转换方法（这些方法逻辑简单，不需要修改）
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
    return core::VerticalAlign::Bottom;
}

core::BorderStyle StylesParser::getBorderStyle(const std::string& style) const {
    if (style == "thin") return core::BorderStyle::Thin;
    if (style == "thick") return core::BorderStyle::Thick;
    if (style == "medium") return core::BorderStyle::Medium;
    if (style == "dashed") return core::BorderStyle::Dashed;
    if (style == "dotted") return core::BorderStyle::Dotted;
    if (style == "double") return core::BorderStyle::Double;
    return core::BorderStyle::None;
}

core::PatternType StylesParser::getPatternType(const std::string& pattern) const {
    if (pattern == "solid") return core::PatternType::Solid;
    if (pattern == "darkGray") return core::PatternType::DarkGray;
    if (pattern == "mediumGray") return core::PatternType::MediumGray;
    if (pattern == "lightGray") return core::PatternType::LightGray;
    if (pattern == "gray125") return core::PatternType::Gray125;
    if (pattern == "gray0625") return core::PatternType::Gray0625;
    return core::PatternType::None;
}

std::string StylesParser::getBuiltinNumberFormat(int format_id) const {
    // Excel内置数字格式映射
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
        {22, "m/d/yy h:mm"}
    };
    
    auto it = builtin_formats.find(format_id);
    return (it != builtin_formats.end()) ? it->second : std::string();
}

} // namespace reader
} // namespace fastexcel