#pragma once

#include "fastexcel/core/Color.hpp"
#include "FormatDescriptor.hpp"  // 包含枚举定义
#include "FormatTypes.hpp"       // 包含NumberFormatType
#include <string>
#include <memory>
#include <cstdint>
#include <functional>

namespace fastexcel {
namespace core {

/**
 * @brief Format类 - Excel单元格格式化
 * 
 * 向后兼容的Format类，负责管理Excel单元格的所有格式属性。
 * 在新架构中，建议使用FormatDescriptor和StyleBuilder。
 * 
 * 提供与libxlsxwriter兼容的格式化功能，包括：
 * - 字体设置（名称、大小、颜色、样式）
 * - 对齐设置（水平、垂直、换行、旋转、缩进）
 * - 边框设置（样式、颜色、对角线）
 * - 填充设置（背景色、前景色、模式）
 * - 数字格式设置
 * - 保护设置（锁定、隐藏）
 */
class Format {
private:
    // 字体属性
    std::string font_name_ = "Calibri";
    double font_size_ = 11.0;
    bool bold_ = false;
    bool italic_ = false;
    UnderlineType underline_ = UnderlineType::None;
    bool strikeout_ = false;
    bool outline_ = false;
    bool shadow_ = false;
    FontScript script_ = FontScript::None;
    Color font_color_ = Color::BLACK;
    uint8_t font_family_ = 2;
    uint8_t font_charset_ = 1;
    bool font_condense_ = false;
    bool font_extend_ = false;
    std::string font_scheme_;
    uint8_t theme_ = 1;
    
    // 对齐属性
    HorizontalAlign horizontal_align_ = HorizontalAlign::None;
    VerticalAlign vertical_align_ = VerticalAlign::Bottom;
    bool text_wrap_ = false;
    int16_t rotation_ = 0;
    uint8_t indent_ = 0;
    bool shrink_ = false;
    uint8_t reading_order_ = 0;
    bool just_distrib_ = false;
    
    // 边框属性
    BorderStyle left_border_ = BorderStyle::None;
    BorderStyle right_border_ = BorderStyle::None;
    BorderStyle top_border_ = BorderStyle::None;
    BorderStyle bottom_border_ = BorderStyle::None;
    BorderStyle diag_border_ = BorderStyle::None;
    DiagonalBorderType diag_type_ = DiagonalBorderType::None;
    
    Color left_border_color_ = Color::BLACK;
    Color right_border_color_ = Color::BLACK;
    Color top_border_color_ = Color::BLACK;
    Color bottom_border_color_ = Color::BLACK;
    Color diag_border_color_ = Color::BLACK;
    
    // 填充属性
    PatternType pattern_ = PatternType::None;
    Color bg_color_ = Color::WHITE;
    Color fg_color_ = Color::BLACK;
    
    // 数字格式
    std::string num_format_;
    uint16_t num_format_index_ = 0;
    
    // 保护属性
    bool locked_ = true;
    bool hidden_ = false;
    
    // 其他属性
    bool quote_prefix_ = false;
    bool hyperlink_ = false;
    uint8_t color_indexed_ = 0;
    bool font_only_ = false;
    
    // 格式索引
    int32_t xf_index_ = -1;
    int32_t dxf_index_ = -1;
    int32_t font_index_ = -1;
    int32_t fill_index_ = -1;
    int32_t border_index_ = -1;
    
    // 标记是否有相应的格式设置
    bool has_font_ = false;
    bool has_fill_ = false;
    bool has_border_ = false;
    bool has_alignment_ = false;
    bool has_protection_ = false;

public:
    Format() = default;
    ~Format() = default;
    
    // 允许拷贝构造和赋值（为了测试）
    Format(const Format&) = default;
    Format& operator=(const Format&) = default;
    
    // 允许移动构造和赋值
    Format(Format&&) = default;
    Format& operator=(Format&&) = default;
    
    // ========== 字体设置 ==========
    
    /**
     * @brief 设置字体名称
     * @param name 字体名称，如"Calibri", "Arial"等
     */
    void setFontName(const std::string& name);
    
    /**
     * @brief 设置字体大小
     * @param size 字体大小，范围1.0-409.0
     */
    void setFontSize(double size);
    
    /**
     * @brief 设置字体颜色
     * @param color RGB颜色值
     */
    void setFontColor(Color color);
    
    /**
     * @brief 设置粗体
     * @param bold 是否粗体
     */
    void setBold(bool bold = true);
    
    /**
     * @brief 设置斜体
     * @param italic 是否斜体
     */
    void setItalic(bool italic = true);
    
    /**
     * @brief 设置下划线
     * @param underline 下划线类型
     */
    void setUnderline(UnderlineType underline);
    
    /**
     * @brief 设置删除线
     * @param strikeout 是否删除线
     */
    void setStrikeout(bool strikeout = true);
    
    /**
     * @brief 设置字体轮廓
     * @param outline 是否轮廓
     */
    void setFontOutline(bool outline = true);
    
    /**
     * @brief 设置字体阴影
     * @param shadow 是否阴影
     */
    void setFontShadow(bool shadow = true);
    
    /**
     * @brief 设置字体脚本（上标/下标）
     * @param script 脚本类型
     */
    void setFontScript(FontScript script);
    
    /**
     * @brief 设置上标
     * @param superscript 是否上标
     */
    void setSuperscript(bool superscript = true);
    
    /**
     * @brief 设置下标
     * @param subscript 是否下标
     */
    void setSubscript(bool subscript = true);
    
    /**
     * @brief 设置字体族
     * @param family 字体族索引
     */
    void setFontFamily(uint8_t family);
    
    /**
     * @brief 设置字体字符集
     * @param charset 字符集
     */
    void setFontCharset(uint8_t charset);
    
    /**
     * @brief 设置字体压缩
     * @param condense 是否压缩
     */
    void setFontCondense(bool condense = true);
    
    /**
     * @brief 设置字体扩展
     * @param extend 是否扩展
     */
    void setFontExtend(bool extend = true);
    
    /**
     * @brief 设置字体方案
     * @param scheme 字体方案
     */
    void setFontScheme(const std::string& scheme);
    
    /**
     * @brief 设置主题
     * @param theme 主题索引
     */
    void setTheme(uint8_t theme);
    
    // ========== 对齐设置 ==========
    
    /**
     * @brief 设置水平对齐
     * @param align 水平对齐方式
     */
    void setHorizontalAlign(HorizontalAlign align);
    
    /**
     * @brief 设置垂直对齐
     * @param align 垂直对齐方式
     */
    void setVerticalAlign(VerticalAlign align);
    
    /**
     * @brief 设置对齐（兼容libxlsxwriter）
     * @param alignment 对齐方式
     */
    void setAlign(uint8_t alignment);
    
    /**
     * @brief 设置文本换行
     * @param wrap 是否换行
     */
    void setTextWrap(bool wrap = true);
    
    /**
     * @brief 设置文本旋转
     * @param angle 旋转角度（-90到90，或270）
     */
    void setRotation(int16_t angle);
    
    /**
     * @brief 设置缩进级别
     * @param level 缩进级别
     */
    void setIndent(uint8_t level);
    
    /**
     * @brief 设置收缩以适应
     * @param shrink 是否收缩
     */
    void setShrink(bool shrink = true);
    
    /**
     * @brief 设置收缩以适应（别名）
     * @param shrink 是否收缩
     */
    void setShrinkToFit(bool shrink = true);
    
    /**
     * @brief 设置阅读顺序
     * @param order 阅读顺序
     */
    void setReadingOrder(uint8_t order);
    
    // ========== 边框设置 ==========
    
    /**
     * @brief 设置所有边框
     * @param style 边框样式
     */
    void setBorder(BorderStyle style);
    
    /**
     * @brief 设置左边框
     * @param style 边框样式
     */
    void setLeftBorder(BorderStyle style);
    
    /**
     * @brief 设置右边框
     * @param style 边框样式
     */
    void setRightBorder(BorderStyle style);
    
    /**
     * @brief 设置上边框
     * @param style 边框样式
     */
    void setTopBorder(BorderStyle style);
    
    /**
     * @brief 设置下边框
     * @param style 边框样式
     */
    void setBottomBorder(BorderStyle style);
    
    /**
     * @brief 设置所有边框颜色
     * @param color 边框颜色
     */
    void setBorderColor(Color color);
    
    /**
     * @brief 设置左边框颜色
     * @param color 边框颜色
     */
    void setLeftBorderColor(Color color);
    
    /**
     * @brief 设置右边框颜色
     * @param color 边框颜色
     */
    void setRightBorderColor(Color color);
    
    /**
     * @brief 设置上边框颜色
     * @param color 边框颜色
     */
    void setTopBorderColor(Color color);
    
    /**
     * @brief 设置下边框颜色
     * @param color 边框颜色
     */
    void setBottomBorderColor(Color color);
    
    /**
     * @brief 设置对角线边框类型
     * @param type 对角线类型
     */
    void setDiagType(DiagonalBorderType type);
    
    /**
     * @brief 设置对角线边框样式
     * @param style 边框样式
     */
    void setDiagBorder(BorderStyle style);
    
    /**
     * @brief 设置对角线边框颜色
     * @param color 边框颜色
     */
    void setDiagColor(Color color);
    
    /**
     * @brief 设置对角线边框（别名）
     * @param style 边框样式
     */
    void setDiagonalBorder(BorderStyle style);
    
    /**
     * @brief 设置对角线边框颜色（别名）
     * @param color 边框颜色
     */
    void setDiagonalBorderColor(Color color);
    
    /**
     * @brief 设置对角线类型（别名）
     * @param type 对角线类型
     */
    void setDiagonalType(DiagonalType type);
    
    // ========== 填充设置 ==========
    
    /**
     * @brief 设置填充模式
     * @param pattern 填充模式
     */
    void setPattern(PatternType pattern);
    
    /**
     * @brief 设置背景色
     * @param color 背景色
     */
    void setBackgroundColor(Color color);
    
    /**
     * @brief 设置前景色
     * @param color 前景色
     */
    void setForegroundColor(Color color);
    
    // ========== 数字格式设置 ==========
    
    /**
     * @brief 设置数字格式
     * @param format 格式字符串
     */
    void setNumberFormat(const std::string& format);
    
    /**
     * @brief 设置数字格式索引
     * @param index 内置格式索引
     */
    void setNumberFormatIndex(uint16_t index);
    
    /**
     * @brief 设置数字格式（使用预定义类型）
     * @param type 数字格式类型
     */
    void setNumberFormat(NumberFormatType type);
    
    // ========== 保护设置 ==========
    
    /**
     * @brief 设置单元格解锁
     * @param unlocked 是否解锁
     */
    void setUnlocked(bool unlocked = true);
    
    /**
     * @brief 设置单元格锁定
     * @param locked 是否锁定
     */
    void setLocked(bool locked = true);
    
    /**
     * @brief 设置公式隐藏
     * @param hidden 是否隐藏
     */
    void setHidden(bool hidden = true);
    
    // ========== 其他设置 ==========
    
    /**
     * @brief 设置引号前缀
     * @param prefix 是否添加引号前缀
     */
    void setQuotePrefix(bool prefix = true);
    
    /**
     * @brief 设置超链接格式
     * @param hyperlink 是否为超链接
     */
    void setHyperlink(bool hyperlink = true);
    
    /**
     * @brief 设置颜色索引
     * @param index 颜色索引
     */
    void setColorIndexed(uint8_t index);
    
    /**
     * @brief 设置仅字体格式
     * @param font_only 是否仅字体
     */
    void setFontOnly(bool font_only = true);
    
    // ========== 获取属性 ==========
    
    const std::string& getFontName() const { return font_name_; }
    double getFontSize() const { return font_size_; }
    Color getFontColor() const { return font_color_; }
    bool isBold() const { return bold_; }
    bool isItalic() const { return italic_; }
    UnderlineType getUnderline() const { return underline_; }
    bool isStrikeout() const { return strikeout_; }
    FontScript getFontScript() const { return script_; }
    bool isSuperscript() const { return script_ == FontScript::Superscript; }
    bool isSubscript() const { return script_ == FontScript::Subscript; }
    
    HorizontalAlign getHorizontalAlign() const { return horizontal_align_; }
    VerticalAlign getVerticalAlign() const { return vertical_align_; }
    bool isTextWrap() const { return text_wrap_; }
    int16_t getRotation() const { return rotation_; }
    uint8_t getIndent() const { return indent_; }
    bool isShrink() const { return shrink_; }
    bool isShrinkToFit() const { return shrink_; }
    
    BorderStyle getLeftBorder() const { return left_border_; }
    BorderStyle getRightBorder() const { return right_border_; }
    BorderStyle getTopBorder() const { return top_border_; }
    BorderStyle getBottomBorder() const { return bottom_border_; }
    BorderStyle getDiagBorder() const { return diag_border_; }
    DiagonalBorderType getDiagType() const { return diag_type_; }
    BorderStyle getDiagonalBorder() const { return diag_border_; }
    DiagonalType getDiagonalType() const { return diag_type_; }
    
    Color getLeftBorderColor() const { return left_border_color_; }
    Color getRightBorderColor() const { return right_border_color_; }
    Color getTopBorderColor() const { return top_border_color_; }
    Color getBottomBorderColor() const { return bottom_border_color_; }
    Color getDiagBorderColor() const { return diag_border_color_; }
    Color getDiagonalBorderColor() const { return diag_border_color_; }
    
    PatternType getPattern() const { return pattern_; }
    Color getBackgroundColor() const { return bg_color_; }
    Color getForegroundColor() const { return fg_color_; }
    
    const std::string& getNumberFormat() const { return num_format_; }
    uint16_t getNumberFormatIndex() const { return num_format_index_; }
    
    bool isLocked() const { return locked_; }
    bool isHidden() const { return hidden_; }
    bool hasQuotePrefix() const { return quote_prefix_; }
    
    // ========== 格式索引管理 ==========
    
    void setXfIndex(int32_t index) { xf_index_ = index; }
    int32_t getXfIndex() const { return xf_index_; }
    
    void setDxfIndex(int32_t index) { dxf_index_ = index; }
    int32_t getDxfIndex() const { return dxf_index_; }
    
    void setFontIndex(int32_t index) { font_index_ = index; }
    int32_t getFontIndex() const { return font_index_; }
    
    void setFillIndex(int32_t index) { fill_index_ = index; }
    int32_t getFillIndex() const { return fill_index_; }
    
    void setBorderIndex(int32_t index) { border_index_ = index; }
    int32_t getBorderIndex() const { return border_index_; }
    
    // ========== 格式检查 ==========
    
    bool hasFont() const { return has_font_; }
    bool hasFill() const { return has_fill_; }
    bool hasBorder() const { return has_border_; }
    bool hasAlignment() const { return has_alignment_; }
    bool hasProtection() const { return has_protection_; }
    bool hasAnyFormatting() const;
    
    // ========== XML生成 ==========
    
    std::string generateFontXML() const;
    std::string generateFillXML() const;
    std::string generateBorderXML() const;
    std::string generateAlignmentXML() const;
    std::string generateProtectionXML() const;
    std::string generateNumberFormatXML() const;
    std::string generateXML() const;
    
    // ========== 格式比较和哈希 ==========
    
    bool equals(const Format& other) const;
    size_t hash() const;
    
    // ========== 兼容性方法 ==========
    
    // 为了与现有代码兼容而保留的方法
    void setWrapText(bool wrap = true) { setTextWrap(wrap); }
    void setHorizontalAlignment(HorizontalAlign align) { setHorizontalAlign(align); }
    void setVerticalAlignment(VerticalAlign align) { setVerticalAlign(align); }
    
private:
    // 内部辅助方法
    void markFontChanged() { has_font_ = true; }
    void markFillChanged() { has_fill_ = true; }
    void markBorderChanged() { has_border_ = true; }
    void markAlignmentChanged() { has_alignment_ = true; }
    void markProtectionChanged() { has_protection_ = true; }
    
    // 模板化的属性设置辅助方法
    template<typename T>
    void setPropertyWithMarker(T& property, const T& value, void (Format::*marker)());
    
    template<typename T>
    void setValidatedProperty(T& property, const T& value,
                             std::function<bool(const T&)> validator,
                             void (Format::*marker)());
    
    // 重载版本，用于不需要验证器的情况
    template<typename T>
    void setValidatedProperty(T& property, const T& value);
    
    std::string borderStyleToString(BorderStyle style) const;
    std::string patternTypeToString(PatternType pattern) const;
};

}} // namespace fastexcel::core