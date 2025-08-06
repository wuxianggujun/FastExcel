#pragma once

#include "FormatDescriptor.hpp"
#include <string>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 样式构建器 - 流式API创建格式
 * 
 * 提供链式调用的流畅API，用于创建不可变的格式描述符。
 * 这是用户主要接触的API，隐藏了底层的复杂性。
 */
class StyleBuilder {
private:
    // 构建过程中的可变状态
    std::string font_name_ = "Calibri";
    double font_size_ = 11.0;
    bool bold_ = false;
    bool italic_ = false;
    UnderlineType underline_ = UnderlineType::None;
    bool strikeout_ = false;
    FontScript script_ = FontScript::None;
    core::Color font_color_ = core::Color::BLACK;
    uint8_t font_family_ = 2;
    uint8_t font_charset_ = 1;
    
    HorizontalAlign horizontal_align_ = HorizontalAlign::None;
    VerticalAlign vertical_align_ = VerticalAlign::Bottom;
    bool text_wrap_ = false;
    int16_t rotation_ = 0;
    uint8_t indent_ = 0;
    bool shrink_ = false;
    
    BorderStyle left_border_ = BorderStyle::None;
    BorderStyle right_border_ = BorderStyle::None;
    BorderStyle top_border_ = BorderStyle::None;
    BorderStyle bottom_border_ = BorderStyle::None;
    BorderStyle diag_border_ = BorderStyle::None;
    DiagonalBorderType diag_type_ = DiagonalBorderType::None;
    
    core::Color left_border_color_ = core::Color::BLACK;
    core::Color right_border_color_ = core::Color::BLACK;
    core::Color top_border_color_ = core::Color::BLACK;
    core::Color bottom_border_color_ = core::Color::BLACK;
    core::Color diag_border_color_ = core::Color::BLACK;
    
    PatternType pattern_ = PatternType::None;
    core::Color bg_color_ = core::Color::WHITE;
    core::Color fg_color_ = core::Color::BLACK;
    
    std::string num_format_;
    uint16_t num_format_index_ = 0;
    
    bool locked_ = true;
    bool hidden_ = false;

public:
    StyleBuilder() = default;
    
    // 从现有格式创建Builder（用于修改现有格式）
    explicit StyleBuilder(const FormatDescriptor& format);
    
    // ========== 字体设置（链式调用） ==========
    
    /**
     * @brief 设置字体名称
     * @param name 字体名称
     * @return Builder引用，支持链式调用
     */
    StyleBuilder& fontName(const std::string& name) {
        font_name_ = name;
        return *this;
    }
    
    /**
     * @brief 设置字体大小
     * @param size 字体大小（1.0-409.0）
     * @return Builder引用
     */
    StyleBuilder& fontSize(double size) {
        if (size >= 1.0 && size <= 409.0) {
            font_size_ = size;
        }
        return *this;
    }
    
    /**
     * @brief 设置字体（名称+大小）
     * @param name 字体名称
     * @param size 字体大小
     * @return Builder引用
     */
    StyleBuilder& font(const std::string& name, double size) {
        return fontName(name).fontSize(size);
    }
    
    /**
     * @brief 设置字体（名称+大小+粗体）
     * @param name 字体名称
     * @param size 字体大小
     * @param is_bold 是否粗体
     * @return Builder引用
     */
    StyleBuilder& font(const std::string& name, double size, bool is_bold) {
        return font(name, size).bold(is_bold);
    }
    
    /**
     * @brief 设置字体颜色
     * @param color 字体颜色
     * @return Builder引用
     */
    StyleBuilder& fontColor(core::Color color) {
        font_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置粗体
     * @param is_bold 是否粗体（默认true）
     * @return Builder引用
     */
    StyleBuilder& bold(bool is_bold = true) {
        bold_ = is_bold;
        return *this;
    }
    
    /**
     * @brief 设置斜体
     * @param is_italic 是否斜体（默认true）
     * @return Builder引用
     */
    StyleBuilder& italic(bool is_italic = true) {
        italic_ = is_italic;
        return *this;
    }
    
    /**
     * @brief 设置下划线
     * @param type 下划线类型
     * @return Builder引用
     */
    StyleBuilder& underline(UnderlineType type = UnderlineType::Single) {
        underline_ = type;
        return *this;
    }
    
    /**
     * @brief 设置删除线
     * @param is_strikeout 是否删除线（默认true）
     * @return Builder引用
     */
    StyleBuilder& strikeout(bool is_strikeout = true) {
        strikeout_ = is_strikeout;
        return *this;
    }
    
    /**
     * @brief 设置上标
     * @param is_super 是否上标（默认true）
     * @return Builder引用
     */
    StyleBuilder& superscript(bool is_super = true) {
        script_ = is_super ? FontScript::Superscript : FontScript::None;
        return *this;
    }
    
    /**
     * @brief 设置下标
     * @param is_sub 是否下标（默认true）
     * @return Builder引用
     */
    StyleBuilder& subscript(bool is_sub = true) {
        script_ = is_sub ? FontScript::Subscript : FontScript::None;
        return *this;
    }
    
    // ========== 对齐设置 ==========
    
    /**
     * @brief 设置水平对齐
     * @param align 对齐方式
     * @return Builder引用
     */
    StyleBuilder& horizontalAlign(HorizontalAlign align) {
        horizontal_align_ = align;
        return *this;
    }
    
    /**
     * @brief 设置垂直对齐
     * @param align 对齐方式
     * @return Builder引用
     */
    StyleBuilder& verticalAlign(VerticalAlign align) {
        vertical_align_ = align;
        return *this;
    }
    
    /**
     * @brief 左对齐
     * @return Builder引用
     */
    StyleBuilder& leftAlign() {
        return horizontalAlign(HorizontalAlign::Left);
    }
    
    /**
     * @brief 居中对齐
     * @return Builder引用
     */
    StyleBuilder& centerAlign() {
        return horizontalAlign(HorizontalAlign::Center);
    }
    
    /**
     * @brief 右对齐
     * @return Builder引用
     */
    StyleBuilder& rightAlign() {
        return horizontalAlign(HorizontalAlign::Right);
    }
    
    /**
     * @brief 垂直居中
     * @return Builder引用
     */
    StyleBuilder& vcenterAlign() {
        return verticalAlign(VerticalAlign::Center);
    }
    
    /**
     * @brief 设置文本换行
     * @param wrap 是否换行（默认true）
     * @return Builder引用
     */
    StyleBuilder& textWrap(bool wrap = true) {
        text_wrap_ = wrap;
        return *this;
    }
    
    /**
     * @brief 设置文本旋转角度
     * @param angle 旋转角度（-90到90，或270）
     * @return Builder引用
     */
    StyleBuilder& rotation(int16_t angle) {
        if ((angle >= -90 && angle <= 90) || angle == 270) {
            rotation_ = angle;
        }
        return *this;
    }
    
    /**
     * @brief 设置缩进级别
     * @param level 缩进级别
     * @return Builder引用
     */
    StyleBuilder& indent(uint8_t level) {
        indent_ = level;
        return *this;
    }
    
    /**
     * @brief 设置收缩以适应
     * @param shrink_to_fit 是否收缩（默认true）
     * @return Builder引用
     */
    StyleBuilder& shrinkToFit(bool shrink_to_fit = true) {
        shrink_ = shrink_to_fit;
        return *this;
    }
    
    // ========== 边框设置 ==========
    
    /**
     * @brief 设置所有边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& border(BorderStyle style, core::Color color = core::Color::BLACK) {
        left_border_ = right_border_ = top_border_ = bottom_border_ = style;
        left_border_color_ = right_border_color_ = top_border_color_ = bottom_border_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置左边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& leftBorder(BorderStyle style, core::Color color = core::Color::BLACK) {
        left_border_ = style;
        left_border_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置右边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& rightBorder(BorderStyle style, core::Color color = core::Color::BLACK) {
        right_border_ = style;
        right_border_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置上边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& topBorder(BorderStyle style, core::Color color = core::Color::BLACK) {
        top_border_ = style;
        top_border_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置下边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& bottomBorder(BorderStyle style, core::Color color = core::Color::BLACK) {
        bottom_border_ = style;
        bottom_border_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置对角线边框
     * @param style 边框样式
     * @param type 对角线类型
     * @param color 边框颜色（可选）
     * @return Builder引用
     */
    StyleBuilder& diagonalBorder(BorderStyle style, 
                               DiagonalBorderType type = DiagonalBorderType::Both,
                               core::Color color = core::Color::BLACK) {
        diag_border_ = style;
        diag_type_ = type;
        diag_border_color_ = color;
        return *this;
    }
    
    // ========== 填充设置 ==========
    
    /**
     * @brief 设置填充（纯色）
     * @param color 背景色
     * @return Builder引用
     */
    StyleBuilder& fill(core::Color color) {
        pattern_ = PatternType::Solid;
        bg_color_ = color;
        return *this;
    }
    
    /**
     * @brief 设置填充模式
     * @param pattern 填充模式
     * @param bg_color 背景色
     * @param fg_color 前景色（可选）
     * @return Builder引用
     */
    StyleBuilder& fill(PatternType pattern, 
                      core::Color bg_color, 
                      core::Color fg_color = core::Color::BLACK) {
        pattern_ = pattern;
        bg_color_ = bg_color;
        fg_color_ = fg_color;
        return *this;
    }
    
    /**
     * @brief 设置背景色
     * @param color 背景色
     * @return Builder引用
     */
    StyleBuilder& backgroundColor(core::Color color) {
        if (pattern_ == PatternType::None) {
            pattern_ = PatternType::Solid;
        }
        bg_color_ = color;
        return *this;
    }
    
    // ========== 数字格式设置 ==========
    
    /**
     * @brief 设置数字格式
     * @param format 格式字符串
     * @return Builder引用
     */
    StyleBuilder& numberFormat(const std::string& format) {
        num_format_ = format;
        num_format_index_ = 0;  // 自定义格式
        return *this;
    }
    
    /**
     * @brief 设置数字格式索引
     * @param index 内置格式索引
     * @return Builder引用
     */
    StyleBuilder& numberFormatIndex(uint16_t index) {
        num_format_index_ = index;
        num_format_.clear();  // 清除自定义格式
        return *this;
    }
    
    // 常用数字格式的便捷方法
    StyleBuilder& currency() { return numberFormatIndex(7); }           // ¤#,##0.00
    StyleBuilder& percentage() { return numberFormatIndex(10); }        // 0.00%
    StyleBuilder& date() { return numberFormatIndex(14); }            // m/d/yyyy
    StyleBuilder& time() { return numberFormatIndex(21); }             // h:mm:ss AM/PM
    StyleBuilder& dateTime() { return numberFormat("m/d/yyyy h:mm"); }
    StyleBuilder& scientific() { return numberFormatIndex(11); }        // 0.00E+00
    StyleBuilder& text() { return numberFormatIndex(49); }             // @
    
    // ========== 保护设置 ==========
    
    /**
     * @brief 设置单元格锁定
     * @param locked 是否锁定（默认true）
     * @return Builder引用
     */
    StyleBuilder& locked(bool locked = true) {
        locked_ = locked;
        return *this;
    }
    
    /**
     * @brief 设置单元格解锁
     * @param unlocked 是否解锁（默认true）
     * @return Builder引用
     */
    StyleBuilder& unlocked(bool unlocked = true) {
        locked_ = !unlocked;
        return *this;
    }
    
    /**
     * @brief 设置公式隐藏
     * @param hidden 是否隐藏（默认true）
     * @return Builder引用
     */
    StyleBuilder& hidden(bool hidden = true) {
        hidden_ = hidden;
        return *this;
    }
    
    // ========== 构建最终对象 ==========
    
    /**
     * @brief 构建不可变的格式描述符
     * @return 格式描述符对象
     */
    FormatDescriptor build() const;
    
    // ========== 预定义样式的静态工厂方法 ==========
    
    static StyleBuilder header() {
        return StyleBuilder()
            .bold()
            .fontSize(14)
            .centerAlign()
            .vcenterAlign();
    }
    
    static StyleBuilder title() {
        return StyleBuilder()
            .bold()
            .fontSize(16)
            .centerAlign()
            .vcenterAlign();
    }
    
    static StyleBuilder money() {
        return StyleBuilder()
            .currency()
            .rightAlign()
            .vcenterAlign();
    }
    
    static StyleBuilder percent() {
        return StyleBuilder()
            .percentage()
            .rightAlign()
            .vcenterAlign();
    }
    
    static StyleBuilder dateStyle() {
        return StyleBuilder()
            .date()
            .centerAlign()
            .vcenterAlign();
    }
};

/**
 * @brief 命名样式 - 带名称的格式描述符
 * 
 * 用于表示具有名称的样式，便于样式管理和重用。
 */
class NamedStyle {
private:
    std::string name_;
    FormatDescriptor format_;
    
public:
    NamedStyle(const std::string& name, const FormatDescriptor& format)
        : name_(name), format_(format) {}
    
    NamedStyle(const std::string& name, const StyleBuilder& builder)
        : name_(name), format_(builder.build()) {}
    
    const std::string& getName() const { return name_; }
    const FormatDescriptor& getFormat() const { return format_; }
    
    bool operator==(const NamedStyle& other) const {
        return name_ == other.name_ && format_ == other.format_;
    }
    
    bool operator!=(const NamedStyle& other) const {
        return !(*this == other);
    }
    
    size_t hash() const {
        std::hash<std::string> str_hasher;
        return str_hasher(name_) ^ (format_.hash() << 1);
    }
};

}} // namespace fastexcel::core

// 为NamedStyle提供std::hash特化
namespace std {
template<>
struct hash<fastexcel::core::NamedStyle> {
    size_t operator()(const fastexcel::core::NamedStyle& style) const {
        return style.hash();
    }
};
} // namespace std