#include "fastexcel/core/Format.hpp"
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <functional>

namespace fastexcel {
namespace core {

// ========== 私有辅助方法 ==========

// 模板化的属性设置方法，减少重复代码
template<typename T>
void Format::setPropertyWithMarker(T& property, const T& value, void (Format::*marker)()) {
    property = value;
    (this->*marker)();
}

// 带验证的属性设置方法
template<typename T>
void Format::setValidatedProperty(T& property, const T& value,
                                 std::function<bool(const T&)> validator,
                                 void (Format::*marker)()) {
    if (!validator || validator(value)) {
        property = value;
        (this->*marker)();
    }
}

// 重载版本，用于不需要验证器的情况
template<typename T>
void Format::setValidatedProperty(T& property, const T& value) {
    property = value;
}

// ========== 字体设置 ==========

void Format::setFontName(const std::string& name) {
    setPropertyWithMarker(font_name_, name, &Format::markFontChanged);
}

void Format::setFontSize(double size) {
    // 直接验证字体大小范围
    if (size >= 1.0 && size <= 409.0) {
        setPropertyWithMarker(font_size_, size, &Format::markFontChanged);
    }
}

void Format::setFontColor(Color color) {
    setPropertyWithMarker(font_color_, color, &Format::markFontChanged);
}

void Format::setBold(bool bold) {
    setPropertyWithMarker(bold_, bold, &Format::markFontChanged);
}

void Format::setItalic(bool italic) {
    setPropertyWithMarker(italic_, italic, &Format::markFontChanged);
}

void Format::setUnderline(UnderlineType underline) {
    setPropertyWithMarker(underline_, underline, &Format::markFontChanged);
}

void Format::setStrikeout(bool strikeout) {
    setPropertyWithMarker(strikeout_, strikeout, &Format::markFontChanged);
}

void Format::setFontOutline(bool outline) {
    setPropertyWithMarker(outline_, outline, &Format::markFontChanged);
}

void Format::setFontShadow(bool shadow) {
    setPropertyWithMarker(shadow_, shadow, &Format::markFontChanged);
}

void Format::setFontScript(FontScript script) {
    setPropertyWithMarker(script_, script, &Format::markFontChanged);
}

void Format::setFontFamily(uint8_t family) {
    setPropertyWithMarker(font_family_, family, &Format::markFontChanged);
}

void Format::setFontCharset(uint8_t charset) {
    setPropertyWithMarker(font_charset_, charset, &Format::markFontChanged);
}

void Format::setFontCondense(bool condense) {
    setPropertyWithMarker(font_condense_, condense, &Format::markFontChanged);
}

void Format::setFontExtend(bool extend) {
    setPropertyWithMarker(font_extend_, extend, &Format::markFontChanged);
}

void Format::setFontScheme(const std::string& scheme) {
    setPropertyWithMarker(font_scheme_, scheme, &Format::markFontChanged);
}

void Format::setTheme(uint8_t theme) {
    setPropertyWithMarker(theme_, theme, &Format::markFontChanged);
}

void Format::setSuperscript(bool superscript) {
    FontScript newScript = superscript ? FontScript::Superscript :
                          (script_ == FontScript::Superscript ? FontScript::None : script_);
    setPropertyWithMarker(script_, newScript, &Format::markFontChanged);
}

void Format::setSubscript(bool subscript) {
    FontScript newScript = subscript ? FontScript::Subscript :
                          (script_ == FontScript::Subscript ? FontScript::None : script_);
    setPropertyWithMarker(script_, newScript, &Format::markFontChanged);
}

// ========== 对齐设置 ==========

void Format::setHorizontalAlign(HorizontalAlign align) {
    setPropertyWithMarker(horizontal_align_, align, &Format::markAlignmentChanged);
}

void Format::setVerticalAlign(VerticalAlign align) {
    setPropertyWithMarker(vertical_align_, align, &Format::markAlignmentChanged);
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
    setPropertyWithMarker(text_wrap_, wrap, &Format::markAlignmentChanged);
}

void Format::setRotation(int16_t angle) {
    // 直接验证角度范围
    if (angle == 270 || (angle >= -90 && angle <= 90)) {
        setPropertyWithMarker(rotation_, angle, &Format::markAlignmentChanged);
    }
}

void Format::setIndent(uint8_t level) {
    setPropertyWithMarker(indent_, level, &Format::markAlignmentChanged);
}

void Format::setShrink(bool shrink) {
    setPropertyWithMarker(shrink_, shrink, &Format::markAlignmentChanged);
}

void Format::setShrinkToFit(bool shrink) {
    setShrink(shrink);
}

void Format::setReadingOrder(uint8_t order) {
    setPropertyWithMarker(reading_order_, order, &Format::markAlignmentChanged);
}

// ========== 边框设置 ==========

void Format::setBorder(BorderStyle style) {
    left_border_ = right_border_ = top_border_ = bottom_border_ = style;
    markBorderChanged();
}

void Format::setLeftBorder(BorderStyle style) {
    setPropertyWithMarker(left_border_, style, &Format::markBorderChanged);
}

void Format::setRightBorder(BorderStyle style) {
    setPropertyWithMarker(right_border_, style, &Format::markBorderChanged);
}

void Format::setTopBorder(BorderStyle style) {
    setPropertyWithMarker(top_border_, style, &Format::markBorderChanged);
}

void Format::setBottomBorder(BorderStyle style) {
    setPropertyWithMarker(bottom_border_, style, &Format::markBorderChanged);
}

void Format::setBorderColor(Color color) {
    left_border_color_ = right_border_color_ = top_border_color_ = bottom_border_color_ = color;
    markBorderChanged();
}

void Format::setLeftBorderColor(Color color) {
    setPropertyWithMarker(left_border_color_, color, &Format::markBorderChanged);
}

void Format::setRightBorderColor(Color color) {
    setPropertyWithMarker(right_border_color_, color, &Format::markBorderChanged);
}

void Format::setTopBorderColor(Color color) {
    setPropertyWithMarker(top_border_color_, color, &Format::markBorderChanged);
}

void Format::setBottomBorderColor(Color color) {
    setPropertyWithMarker(bottom_border_color_, color, &Format::markBorderChanged);
}

void Format::setDiagType(DiagonalBorderType type) {
    setPropertyWithMarker(diag_type_, type, &Format::markBorderChanged);
}

void Format::setDiagBorder(BorderStyle style) {
    setPropertyWithMarker(diag_border_, style, &Format::markBorderChanged);
}

void Format::setDiagColor(Color color) {
    setPropertyWithMarker(diag_border_color_, color, &Format::markBorderChanged);
}

void Format::setDiagonalBorder(BorderStyle style) {
    setDiagBorder(style);
}

void Format::setDiagonalBorderColor(Color color) {
    setDiagColor(color);
}

void Format::setDiagonalType(DiagonalType type) {
    setDiagType(type);
}

// ========== 填充设置 ==========

void Format::setPattern(PatternType pattern) {
    setPropertyWithMarker(pattern_, pattern, &Format::markFillChanged);
}

void Format::setBackgroundColor(Color color) {
    bg_color_ = color;
    if (pattern_ == PatternType::None) {
        pattern_ = PatternType::Solid;
    }
    markFillChanged();
}

void Format::setForegroundColor(Color color) {
    setPropertyWithMarker(fg_color_, color, &Format::markFillChanged);
}

// ========== 数字格式设置 ==========

void Format::setNumberFormat(const std::string& format) {
    num_format_ = format;
}

void Format::setNumberFormatIndex(uint16_t index) {
    num_format_index_ = index;
}

void Format::setNumberFormat(NumberFormatType type) {
    switch (type) {
        case NumberFormatType::General:
            num_format_ = "General";
            num_format_index_ = 0;
            break;
        case NumberFormatType::Number:
            num_format_ = "0";
            num_format_index_ = 1;
            break;
        case NumberFormatType::Decimal:
            num_format_ = "0.00";
            num_format_index_ = 2;
            break;
        case NumberFormatType::Currency:
            num_format_ = "$#,##0.00";
            num_format_index_ = 7;
            break;
        case NumberFormatType::Accounting:
            num_format_ = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
            num_format_index_ = 44;
            break;
        case NumberFormatType::Percentage:
            num_format_ = "0%";
            num_format_index_ = 9;
            break;
        case NumberFormatType::Fraction:
            num_format_ = "# ?/?";
            num_format_index_ = 12;
            break;
        case NumberFormatType::Scientific:
            num_format_ = "0.00E+00";
            num_format_index_ = 11;
            break;
        case NumberFormatType::Date:
            num_format_ = "m/d/yy";
            num_format_index_ = 14;
            break;
        case NumberFormatType::Time:
            num_format_ = "h:mm:ss AM/PM";
            num_format_index_ = 21;
            break;
        case NumberFormatType::Text:
            num_format_ = "@";
            num_format_index_ = 49;
            break;
        default:
            num_format_ = "General";
            num_format_index_ = 0;
            break;
    }
}

// ========== 保护设置 ==========

void Format::setUnlocked(bool unlocked) {
    setPropertyWithMarker(locked_, !unlocked, &Format::markProtectionChanged);
}

void Format::setLocked(bool locked) {
    setPropertyWithMarker(locked_, locked, &Format::markProtectionChanged);
}

void Format::setHidden(bool hidden) {
    setPropertyWithMarker(hidden_, hidden, &Format::markProtectionChanged);
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
    
    // 字体大小 - 总是包含
    oss << "<sz val=\"" << font_size_ << "\"/>";
    
    // 字体名称 - 总是包含
    oss << "<name val=\"" << font_name_ << "\"/>";
    
    // 字体族 - 总是包含
    oss << "<family val=\"" << static_cast<int>(font_family_) << "\"/>";
    
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
        std::string script_str = (script_ == FontScript::Superscript) ? "superscript" : "subscript";
        oss << "<vertAlign val=\"" << script_str << "\"/>";
    }
    
    // 字体颜色
    if (font_color_ != Color::BLACK) {
        oss << font_color_.toXML();
    }
    
    // 字体方案 - 总是包含
    std::string scheme = font_scheme_.empty() ? "minor" : font_scheme_;
    oss << "<scheme val=\"" << scheme << "\"/>";
    
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
        if (fg_color_ != Color::BLACK || pattern_ == PatternType::Solid) {
            Color color = (pattern_ == PatternType::Solid) ? bg_color_ : fg_color_;
            oss << "<fgColor " << color.toXML().substr(7, color.toXML().length() - 9) << "/>";
        }
        
        // 背景色
        if (bg_color_ != Color::WHITE && pattern_ != PatternType::Solid) {
            oss << "<bgColor " << bg_color_.toXML().substr(7, bg_color_.toXML().length() - 9) << "/>";
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
        oss << left_border_color_.toXML();
        oss << "</left>";
    } else {
        oss << "<left/>";
    }
    
    // 右边框
    if (right_border_ != BorderStyle::None) {
        oss << "<right style=\"" << borderStyleToString(right_border_) << "\">";
        oss << right_border_color_.toXML();
        oss << "</right>";
    } else {
        oss << "<right/>";
    }
    
    // 上边框
    if (top_border_ != BorderStyle::None) {
        oss << "<top style=\"" << borderStyleToString(top_border_) << "\">";
        oss << top_border_color_.toXML();
        oss << "</top>";
    } else {
        oss << "<top/>";
    }
    
    // 下边框
    if (bottom_border_ != BorderStyle::None) {
        oss << "<bottom style=\"" << borderStyleToString(bottom_border_) << "\">";
        oss << bottom_border_color_.toXML();
        oss << "</bottom>";
    } else {
        oss << "<bottom/>";
    }
    
    // 对角线边框
    if (diag_border_ != BorderStyle::None && diag_type_ != DiagonalBorderType::None) {
        oss << "<diagonal style=\"" << borderStyleToString(diag_border_) << "\">";
        oss << diag_border_color_.toXML();
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
        // 处理旋转角度映射，与libxlsxwriter保持一致
        int16_t rotation = rotation_;
        if (rotation == 270) {
            rotation = 255;
        } else if (rotation < 0) {
            rotation = -rotation + 90;
        }
        oss << " textRotation=\"" << rotation << "\"";
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
    
    // 返回自闭合标签
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

std::string Format::generateXML() const {
    std::ostringstream oss;
    oss << "<xf";
    
    // 数字格式ID
    oss << " numFmtId=\"" << num_format_index_ << "\"";
    
    // 字体ID
    oss << " fontId=\"" << font_index_ << "\"";
    
    // 填充ID
    oss << " fillId=\"" << fill_index_ << "\"";
    
    // 边框ID
    oss << " borderId=\"" << border_index_ << "\"";
    
    // 父样式ID
    oss << " xfId=\"0\"";
    
    // 应用标志
    if (has_font_) {
        oss << " applyFont=\"1\"";
    }
    if (has_fill_) {
        oss << " applyFill=\"1\"";
    }
    if (has_border_) {
        oss << " applyBorder=\"1\"";
    }
    if (has_alignment_) {
        oss << " applyAlignment=\"1\"";
    }
    if (has_protection_) {
        oss << " applyProtection=\"1\"";
    }
    if (!num_format_.empty()) {
        oss << " applyNumberFormat=\"1\"";
    }
    
    // 如果有对齐或保护设置，需要添加子元素
    if (has_alignment_ || has_protection_) {
        oss << ">";
        
        if (has_alignment_) {
            // 生成alignment元素作为xf的子元素（自闭合标签）
            oss << generateAlignmentXML();
        }
        
        if (has_protection_) {
            oss << generateProtectionXML();
        }
        
        oss << "</xf>";
    } else {
        oss << "/>";
    }
    
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


}} // namespace fastexcel::core
