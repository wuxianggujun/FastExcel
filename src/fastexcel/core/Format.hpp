#pragma once

#include <string>
#include <memory>

namespace fastexcel {
namespace core {

enum class FontBold {
    None,
    Bold
};

enum class FontItalic {
    None,
    Italic
};

enum class FontUnderline {
    None,
    Single,
    Double
};

enum class HorizontalAlignment {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed
};

enum class VerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed
};

class Format {
private:
    // 字体属性
    std::string font_name_;
    int font_size_ = 11;
    FontBold font_bold_ = FontBold::None;
    FontItalic font_italic_ = FontItalic::None;
    FontUnderline font_underline_ = FontUnderline::None;
    uint32_t font_color_ = 0x000000; // 黑色
    
    // 对齐属性
    HorizontalAlignment horizontal_align_ = HorizontalAlignment::General;
    VerticalAlignment vertical_align_ = VerticalAlignment::Bottom;
    bool wrap_text_ = false;
    
    // 背景属性
    bool has_background_ = false;
    uint32_t background_color_ = 0xFFFFFF; // 白色
    uint32_t pattern_color_ = 0x000000; // 黑色
    std::string pattern_type_ = "none";
    
    // 边框属性
    bool has_border_ = false;
    uint32_t border_color_ = 0x000000; // 黑色
    std::string border_style_ = "none";
    
    // 数字格式
    std::string number_format_;
    
    // 格式ID
    int format_id_ = -1;
    
public:
    Format() = default;
    ~Format() = default;
    
    // 字体设置
    void setFontName(const std::string& name) { font_name_ = name; }
    void setFontSize(int size) { font_size_ = size; }
    void setBold(bool bold = true) { font_bold_ = bold ? FontBold::Bold : FontBold::None; }
    void setItalic(bool italic = true) { font_italic_ = italic ? FontItalic::Italic : FontItalic::None; }
    void setUnderline(FontUnderline underline) { font_underline_ = underline; }
    void setFontColor(uint32_t color) { font_color_ = color; }
    
    // 对齐设置
    void setHorizontalAlignment(HorizontalAlignment align) { horizontal_align_ = align; }
    void setVerticalAlignment(VerticalAlignment align) { vertical_align_ = align; }
    void setWrapText(bool wrap = true) { wrap_text_ = wrap; }
    
    // 背景设置
    void setBackgroundColor(uint32_t color) { 
        has_background_ = true; 
        background_color_ = color; 
    }
    void setPattern(const std::string& pattern, uint32_t color = 0x000000) {
        pattern_type_ = pattern;
        pattern_color_ = color;
    }
    
    // 边框设置
    void setBorderStyle(const std::string& style, uint32_t color = 0x000000) {
        has_border_ = true;
        border_style_ = style;
        border_color_ = color;
    }
    
    // 数字格式设置
    void setNumberFormat(const std::string& format) { number_format_ = format; }
    
    // 获取属性
    std::string getFontName() const { return font_name_; }
    int getFontSize() const { return font_size_; }
    FontBold getBold() const { return font_bold_; }
    FontItalic getItalic() const { return font_italic_; }
    FontUnderline getUnderline() const { return font_underline_; }
    uint32_t getFontColor() const { return font_color_; }
    
    HorizontalAlignment getHorizontalAlignment() const { return horizontal_align_; }
    VerticalAlignment getVerticalAlignment() const { return vertical_align_; }
    bool getWrapText() const { return wrap_text_; }
    
    bool hasBackground() const { return has_background_; }
    uint32_t getBackgroundColor() const { return background_color_; }
    std::string getPatternType() const { return pattern_type_; }
    uint32_t getPatternColor() const { return pattern_color_; }
    
    bool hasBorder() const { return has_border_; }
    uint32_t getBorderColor() const { return border_color_; }
    std::string getBorderStyle() const { return border_style_; }
    
    std::string getNumberFormat() const { return number_format_; }
    
    // 格式ID管理
    void setFormatId(int id) { format_id_ = id; }
    int getFormatId() const { return format_id_; }
    
    // 生成XML格式字符串
    std::string generateFontXML() const;
    std::string generateAlignmentXML() const;
    std::string generateFillXML() const;
    std::string generateBorderXML() const;
    std::string generateNumberFormatXML() const;
    
    // 检查是否有任何格式设置
    bool hasAnyFormatting() const;
};

}} // namespace fastexcel::core