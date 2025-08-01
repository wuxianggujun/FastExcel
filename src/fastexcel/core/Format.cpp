#include "fastexcel/core/Format.hpp"
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <functional>

namespace fastexcel {
namespace core {

// ========== 字体设置 ==========

void Format::setFontName(const std::string& name) {
    font_name_ = name;
    markFontChanged();
}

void Format::setFontSize(double size) {
    // 限制字体大小范围
    font_size_ = std::max(1.0, std::min(409.0, size));
    markFontChanged();
}

void Format::setFontColor(Color color) {
    font_color_ = color & 0xFFFFFF; // 确保只有RGB部分
    markFontChanged();
}

void Format::setBold(bool bold) {
    bold_ = bold;
    markFontChanged();
}

void Format::setItalic(bool italic) {
    italic_ = italic;
    markFontChanged();
}

void Format::setUnderline(UnderlineType underline) {
    underline_ = underline;
    markFontChanged();
}

void Format::setStrikeout(bool strikeout) {
    strikeout_ = strikeout;
    markFontChanged();
}

void Format::setFontOutline(bool outline) {
    outline_ = outline;
    markFontChanged();
}

void Format::setFontShadow(bool shadow) {
    shadow_ = shadow;
    markFontChanged();
}

void Format::setFontScript(FontScript script) {
    script_ = script;
    markFontChanged();
}

void Format::setFontFamily(uint8_t family) {
    font_family_ = family;
    markFontChanged();
}

void Format::setFontCharset(uint8_t charset) {
    font_charset_ = charset;
    markFontChanged();
}

void Format::setFontCondense(bool condense) {
    font_condense_ = condense;
    markFontChanged();
}

void Format::setFontExtend(bool extend) {
    font_extend_ = extend;
    markFontChanged();
}

void Format::setFontScheme(const std::string& scheme) {
    font_scheme_ = scheme;
    markFontChanged();
}

void Format::setTheme(uint8_t theme) {
    theme_ = theme;
    markFontChanged();
}

// ========== 对齐设置 ==========

void Format::setHorizontalAlign(HorizontalAlign align) {
    horizontal_align_ = align;
    markAlignmentChanged();
}

void Format::setVerticalAlign(VerticalAlign align) {
    vertical_align_ = align;
    markAlignmentChanged();
}

void Format::setAlign(uint8_t alignment) {
    // 兼容libxlsxwriter的对齐值
    if (alignment >= static_cast<uint8_t>(VerticalAlign::Top)) {
        setVerticalAlign(static_cast<VerticalAlign>(alignment));
    } else {
        setHorizontalAlign(static_cast<HorizontalAlign>(alignment));
    }
}

void Format::setTextWrap(bool wrap) {
    text_wrap_ = wrap;
    markAlignmentChanged();
}

void Format::setRotation(int16_t angle) {
    // 限制角度范围
    if (angle == 270 || (angle >= -90 && angle <= 90)) {
        rotation_ = angle;
        markAlignmentChanged();
    }
}

void Format::setIndent(uint8_t level) {
    indent_ = level;
    markAlignmentChanged();
}

void Format::setShrink(bool shrink) {
    shrink_ = shrink;
    markAlignmentChanged();
}

void Format::setReadingOrder(uint8_t order) {
    reading_order_ = order;
    markAlignmentChanged();
}

// ========== 边框设置 ==========

void Format::setBorder(BorderStyle style) {
    left_border_ = right_border_ = top_border_ = bottom_border_ = style;
    markBorderChanged();
}

void Format::setLeftBorder(BorderStyle style) {
    left_border_ = style;
    markBorderChanged();
}

void Format::setRightBorder(BorderStyle style) {
    right_border_ = style;
    markBorderChanged();
}

void Format::setTopBorder(BorderStyle style) {
    top_border_ = style;
    markBorderChanged();
}

void Format::setBottomBorder(BorderStyle style) {
    bottom_border_ = style;
    markBorderChanged();
}

void Format::setBorderColor(Color color) {
    Color masked_color = color & 0xFFFFFF;
    left_border_color_ = right_border_color_ = top_border_color_ = bottom_border_color_ = masked_color;
    markBorderChanged();
}

void Format::setLeftBorderColor(Color color) {
    left_border_color_ = color & 0xFFFFFF;
    markBorderChanged();
}

void Format::setRightBorderColor(Color color) {
    right_border_color_ = color & 0xFFFFFF;
    markBorderChanged();
}

void Format::setTopBorderColor(Color color) {
    top_border_color_ = color & 0xFFFFFF;
    markBorderChanged();
}

void Format::setBottomBorderColor(Color color) {
    bottom_border_color_ = color & 0xFFFFFF;
    markBorderChanged();
}

void Format::setDiagType(DiagonalBorderType type) {
    diag_type_ = type;
    markBorderChanged();
}

void Format::setDiagBorder(BorderStyle style) {
    diag_border_ = style;
    markBorderChanged();
}

void Format::setDiagColor(Color color) {
    diag_border_color_ = color & 0xFFFFFF;
    markBorderChanged();
}

// ========== 填充设置 ==========

void Format::setPattern(PatternType pattern) {
    pattern_ = pattern;
    markFillChanged();
}

void Format::setBackgroundColor(Color color) {
    bg_color_ = color & 0xFFFFFF;
    if (pattern_ == PatternType::None) {
        pattern_ = PatternType::Solid;
    }
    markFillChanged();
}

void Format::setForegroundColor(Color color) {
    fg_color_ = color & 0xFFFFFF;
    markFillChanged();
}

// ========== 数字格式设置 ==========

void Format::setNumberFormat(const std::string& format) {
    num_format_ = format;
}

void Format::setNumberFormatIndex(uint16_t index) {
    num_format_index_ = index;
}

// ========== 保护设置 ==========

void Format::setUnlocked(bool unlocked) {
    locked_ = !unlocked;
    markProtectionChanged();
}

void Format::setHidden(bool hidden) {
    hidden_ = hidden;
    markProtectionChanged();
}

// ========== 其他设置 ==========

void Format::setQuotePrefix(bool prefix) {
    quote_prefix_ = prefix;
}

void Format::setHyperlink(bool hyperlink) {
    hyperlink_ = hyperlink;
}

void Format::setColorIndexed(uint8_t index) {
    color_indexed_ = index;
}

void Format::setFontOnly(bool font_only) {
    font_only_ = font_only;
}

// ========== 格式检查 ==========

bool Format::hasAnyFormatting() const {
    return has_font_ || has_fill_ || has_border_ || has_alignment_ || has_protection_ ||
           !num_format_.empty() || num_format_index_ != 0 || quote_prefix_ || hyperlink_;
}

// ========== XML生成 ==========

std::string Format::generateFontXML() const {
    if (!has_font_) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<font>";
    
    // 字体大小
    if (font_size_ != 11.0) {
        oss << "<sz val=\"" << font_size_ << "\"/>";
    }
    
    // 字体名称
    if (!font_name_.empty() && font_name_ != "Calibri") {
        oss << "<name val=\"" << font_name_ << "\"/>";
    }
    
    // 字体族
    if (font_family_ != 2) {
        oss << "<family val=\"" << static_cast<int>(font_family_) << "\"/>";
    }
    
    // 字符集
    if (font_charset_ != 1) {
        oss << "<charset val=\"" << static_cast<int>(font_charset_) << "\"/>";
    }
    
    // 粗体
    if (bold_) {
        oss << "<b/>";
    }
    
    // 斜体
    if (italic_) {
        oss << "<i/>";
    }
    
    // 删除线
    if (strikeout_) {
        oss << "<strike/>";
    }
    
    // 轮廓
    if (outline_) {
        oss << "<outline/>";
    }
    
    // 阴影
    if (shadow_) {
        oss << "<shadow/>";
    }
    
    // 压缩
    if (font_condense_) {
        oss << "<condense/>";
    }
    
    // 扩展
    if (font_extend_) {
        oss << "<extend/>";
    }
    
    // 下划线
    if (underline_ != UnderlineType::None) {
        std::string underline_str;
        switch (underline_) {
            case UnderlineType::Single: underline_str = "single"; break;
            case UnderlineType::Double: underline_str = "double"; break;
            case UnderlineType::SingleAccounting: underline_str = "singleAccounting"; break;
            case UnderlineType::DoubleAccounting: underline_str = "doubleAccounting"; break;
            default: underline_str = "single"; break;
        }
        oss << "<u val=\"" << underline_str << "\"/>";
    }
    
    // 上标/下标
    if (script_ != FontScript::None) {
        std::string script_str = (script_ == FontScript::Superscript) ? "1" : "2";
        oss << "<vertAlign val=\"" << script_str << "\"/>";
    }
    
    // 字体颜色
    if (font_color_ != COLOR_BLACK) {
        oss << "<color rgb=\"" << colorToHex(font_color_) << "\"/>";
    }
    
    // 字体方案
    if (!font_scheme_.empty()) {
        oss << "<scheme val=\"" << font_scheme_ << "\"/>";
    }
    
    oss << "</font>";
    return oss.str();
}

std::string Format::generateFillXML() const {
    if (!has_fill_) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<fill>";
    
    std::string pattern_str = patternTypeToString(pattern_);
    oss << "<patternFill patternType=\"" << pattern_str << "\">";
    
    if (pattern_ != PatternType::None) {
        // 前景色
        if (fg_color_ != COLOR_BLACK || pattern_ == PatternType::Solid) {
            Color color = (pattern_ == PatternType::Solid) ? bg_color_ : fg_color_;
            oss << "<fgColor rgb=\"" << colorToHex(color) << "\"/>";
        }
        
        // 背景色
        if (bg_color_ != COLOR_WHITE && pattern_ != PatternType::Solid) {
            oss << "<bgColor rgb=\"" << colorToHex(bg_color_) << "\"/>";
        }
    }
    
    oss << "</patternFill>";
    oss << "</fill>";
    
    return oss.str();
}

std::string Format::generateBorderXML() const {
    if (!has_border_) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<border>";
    
    // 左边框
    if (left_border_ != BorderStyle::None) {
        oss << "<left style=\"" << borderStyleToString(left_border_) << "\">";
        oss << "<color rgb=\"" << colorToHex(left_border_color_) << "\"/>";
        oss << "</left>";
    } else {
        oss << "<left/>";
    }
    
    // 右边框
    if (right_border_ != BorderStyle::None) {
        oss << "<right style=\"" << borderStyleToString(right_border_) << "\">";
        oss << "<color rgb=\"" << colorToHex(right_border_color_) << "\"/>";
        oss << "</right>";
    } else {
        oss << "<right/>";
    }
    
    // 上边框
    if (top_border_ != BorderStyle::None) {
        oss << "<top style=\"" << borderStyleToString(top_border_) << "\">";
        oss << "<color rgb=\"" << colorToHex(top_border_color_) << "\"/>";
        oss << "</top>";
    } else {
        oss << "<top/>";
    }
    
    // 下边框
    if (bottom_border_ != BorderStyle::None) {
        oss << "<bottom style=\"" << borderStyleToString(bottom_border_) << "\">";
        oss << "<color rgb=\"" << colorToHex(bottom_border_color_) << "\"/>";
        oss << "</bottom>";
    } else {
        oss << "<bottom/>";
    }
    
    // 对角线边框
    if (diag_border_ != BorderStyle::None && diag_type_ != DiagonalBorderType::None) {
        oss << "<diagonal style=\"" << borderStyleToString(diag_border_) << "\">";
        oss << "<color rgb=\"" << colorToHex(diag_border_color_) << "\"/>";
        oss << "</diagonal>";
    } else {
        oss << "<diagonal/>";
    }
    
    oss << "</border>";
    return oss.str();
}

std::string Format::generateAlignmentXML() const {
    if (!has_alignment_) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<alignment";
    
    // 水平对齐
    if (horizontal_align_ != HorizontalAlign::None) {
        std::string align_str;
        switch (horizontal_align_) {
            case HorizontalAlign::Left: align_str = "left"; break;
            case HorizontalAlign::Center: align_str = "center"; break;
            case HorizontalAlign::Right: align_str = "right"; break;
            case HorizontalAlign::Fill: align_str = "fill"; break;
            case HorizontalAlign::Justify: align_str = "justify"; break;
            case HorizontalAlign::CenterAcross: align_str = "centerContinuous"; break;
            case HorizontalAlign::Distributed: align_str = "distributed"; break;
            default: break;
        }
        if (!align_str.empty()) {
            oss << " horizontal=\"" << align_str << "\"";
        }
    }
    
    // 垂直对齐
    if (vertical_align_ != VerticalAlign::Bottom) {
        std::string align_str;
        switch (vertical_align_) {
            case VerticalAlign::Top: align_str = "top"; break;
            case VerticalAlign::Center: align_str = "center"; break;
            case VerticalAlign::Justify: align_str = "justify"; break;
            case VerticalAlign::Distributed: align_str = "distributed"; break;
            default: align_str = "bottom"; break;
        }
        oss << " vertical=\"" << align_str << "\"";
    }
    
    // 文本换行
    if (text_wrap_) {
        oss << " wrapText=\"1\"";
    }
    
    // 文本旋转
    if (rotation_ != 0) {
        oss << " textRotation=\"" << rotation_ << "\"";
    }
    
    // 缩进
    if (indent_ > 0) {
        oss << " indent=\"" << static_cast<int>(indent_) << "\"";
    }
    
    // 收缩以适应
    if (shrink_) {
        oss << " shrinkToFit=\"1\"";
    }
    
    // 阅读顺序
    if (reading_order_ > 0) {
        oss << " readingOrder=\"" << static_cast<int>(reading_order_) << "\"";
    }
    
    oss << "/>";
    return oss.str();
}

std::string Format::generateProtectionXML() const {
    if (!has_protection_) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<protection";
    
    if (!locked_) {
        oss << " locked=\"0\"";
    }
    
    if (hidden_) {
        oss << " hidden=\"1\"";
    }
    
    oss << "/>";
    return oss.str();
}

std::string Format::generateNumberFormatXML() const {
    if (num_format_.empty()) {
        return "";
    }
    
    std::ostringstream oss;
    oss << "<numFmt numFmtId=\"" << num_format_index_ << "\" formatCode=\"" << num_format_ << "\"/>";
    return oss.str();
}

// ========== 格式比较和哈希 ==========

bool Format::equals(const Format& other) const {
    return font_name_ == other.font_name_ &&
           font_size_ == other.font_size_ &&
           bold_ == other.bold_ &&
           italic_ == other.italic_ &&
           underline_ == other.underline_ &&
           strikeout_ == other.strikeout_ &&
           font_color_ == other.font_color_ &&
           horizontal_align_ == other.horizontal_align_ &&
           vertical_align_ == other.vertical_align_ &&
           text_wrap_ == other.text_wrap_ &&
           rotation_ == other.rotation_ &&
           indent_ == other.indent_ &&
           left_border_ == other.left_border_ &&
           right_border_ == other.right_border_ &&
           top_border_ == other.top_border_ &&
           bottom_border_ == other.bottom_border_ &&
           left_border_color_ == other.left_border_color_ &&
           right_border_color_ == other.right_border_color_ &&
           top_border_color_ == other.top_border_color_ &&
           bottom_border_color_ == other.bottom_border_color_ &&
           pattern_ == other.pattern_ &&
           bg_color_ == other.bg_color_ &&
           fg_color_ == other.fg_color_ &&
           num_format_ == other.num_format_ &&
           locked_ == other.locked_ &&
           hidden_ == other.hidden_;
}

size_t Format::hash() const {
    size_t h1 = std::hash<std::string>{}(font_name_);
    size_t h2 = std::hash<double>{}(font_size_);
    size_t h3 = std::hash<bool>{}(bold_);
    size_t h4 = std::hash<bool>{}(italic_);
    size_t h5 = std::hash<uint32_t>{}(font_color_);
    
    // 简单的哈希组合
    return h1 ^ (h2 << 1) ^ (h3 << 2) ^ (h4 << 3) ^ (h5 << 4);
}

// ========== 兼容性方法 ==========

void Format::setBorderStyle(const std::string& style, Color color) {
    BorderStyle border_style = BorderStyle::None;
    
    if (style == "thin") border_style = BorderStyle::Thin;
    else if (style == "medium") border_style = BorderStyle::Medium;
    else if (style == "thick") border_style = BorderStyle::Thick;
    else if (style == "double") border_style = BorderStyle::Double;
    else if (style == "dashed") border_style = BorderStyle::Dashed;
    else if (style == "dotted") border_style = BorderStyle::Dotted;
    else if (style == "hair") border_style = BorderStyle::Hair;
    
    setBorder(border_style);
    setBorderColor(color);
}

void Format::setPattern(const std::string& pattern, Color color) {
    PatternType pattern_type = PatternType::None;
    
    if (pattern == "solid") pattern_type = PatternType::Solid;
    else if (pattern == "mediumGray") pattern_type = PatternType::MediumGray;
    else if (pattern == "darkGray") pattern_type = PatternType::DarkGray;
    else if (pattern == "lightGray") pattern_type = PatternType::LightGray;
    
    setPattern(pattern_type);
    if (pattern_type == PatternType::Solid) {
        setBackgroundColor(color);
    } else {
        setForegroundColor(color);
    }
}

// ========== 内部辅助方法 ==========

std::string Format::borderStyleToString(BorderStyle style) const {
    switch (style) {
        case BorderStyle::Thin: return "thin";
        case BorderStyle::Medium: return "medium";
        case BorderStyle::Dashed: return "dashed";
        case BorderStyle::Dotted: return "dotted";
        case BorderStyle::Thick: return "thick";
        case BorderStyle::Double: return "double";
        case BorderStyle::Hair: return "hair";
        case BorderStyle::MediumDashed: return "mediumDashed";
        case BorderStyle::DashDot: return "dashDot";
        case BorderStyle::MediumDashDot: return "mediumDashDot";
        case BorderStyle::DashDotDot: return "dashDotDot";
        case BorderStyle::MediumDashDotDot: return "mediumDashDotDot";
        case BorderStyle::SlantDashDot: return "slantDashDot";
        default: return "none";
    }
}

std::string Format::patternTypeToString(PatternType pattern) const {
    switch (pattern) {
        case PatternType::Solid: return "solid";
        case PatternType::MediumGray: return "mediumGray";
        case PatternType::DarkGray: return "darkGray";
        case PatternType::LightGray: return "lightGray";
        case PatternType::DarkHorizontal: return "darkHorizontal";
        case PatternType::DarkVertical: return "darkVertical";
        case PatternType::DarkDown: return "darkDown";
        case PatternType::DarkUp: return "darkUp";
        case PatternType::DarkGrid: return "darkGrid";
        case PatternType::DarkTrellis: return "darkTrellis";
        case PatternType::LightHorizontal: return "lightHorizontal";
        case PatternType::LightVertical: return "lightVertical";
        case PatternType::LightDown: return "lightDown";
        case PatternType::LightUp: return "lightUp";
        case PatternType::LightGrid: return "lightGrid";
        case PatternType::LightTrellis: return "lightTrellis";
        case PatternType::Gray125: return "gray125";
        case PatternType::Gray0625: return "gray0625";
        default: return "none";
    }
}

std::string Format::colorToHex(Color color) const {
    std::ostringstream oss;
    oss << std::hex << std::uppercase << std::setw(6) << std::setfill('0') << (color & 0xFFFFFF);
    return oss.str();
}

}} // namespace fastexcel::core