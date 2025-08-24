#pragma once

#include "Color.hpp"
#include "FormatTypes.hpp"  // 使用独立的格式类型定义
#include <string>
#include <cstdint>
#include <functional>

namespace fastexcel {

// 前向声明
namespace memory {
    template<typename T, size_t PoolSize> class FixedSizePool;
}

namespace core {

/**
 * @brief 不可变的格式描述符 - 值对象模式
 * 
 * 这是新架构的核心类，采用值对象模式，一旦创建就不可修改。
 * 所有格式属性都在构造时确定，确保线程安全和缓存友好。
 * 
 * 设计原则：
 * - 不可变性：所有字段都是const，创建后无法修改  
 * - 值语义：支持复制、比较、哈希等操作
 * - 线程安全：不可变对象天然线程安全
 * - 缓存友好：可以安全地被多个工作簿共享
 */
class FormatDescriptor {
private:
    // 字体属性
    const std::string font_name_;
    const double font_size_;
    const bool bold_;
    const bool italic_;
    const UnderlineType underline_;
    const bool strikeout_;
    const FontScript script_;
    const core::Color font_color_;
    const uint8_t font_family_;
    const uint8_t font_charset_;
    
    // 对齐属性
    const HorizontalAlign horizontal_align_;
    const VerticalAlign vertical_align_;
    const bool text_wrap_;
    const int16_t rotation_;
    const uint8_t indent_;
    const bool shrink_;
    
    // 边框属性
    const BorderStyle left_border_;
    const BorderStyle right_border_;
    const BorderStyle top_border_;
    const BorderStyle bottom_border_;
    const BorderStyle diag_border_;
    const DiagonalBorderType diag_type_;
    
    const core::Color left_border_color_;
    const core::Color right_border_color_;
    const core::Color top_border_color_;
    const core::Color bottom_border_color_;
    const core::Color diag_border_color_;
    
    // 填充属性
    const PatternType pattern_;
    const core::Color bg_color_;
    const core::Color fg_color_;
    
    // 数字格式
    const std::string num_format_;
    const uint16_t num_format_index_;
    
    // 保护属性
    const bool locked_;
    const bool hidden_;
    
    // 预计算的哈希值（用于性能优化）
    const size_t hash_value_;
    
    // 私有构造函数，只能通过Builder创建
    FormatDescriptor(
        const std::string& font_name,
        double font_size,
        bool bold,
        bool italic,
        UnderlineType underline,
        bool strikeout,
        FontScript script,
        core::Color font_color,
        uint8_t font_family,
        uint8_t font_charset,
        HorizontalAlign horizontal_align,
        VerticalAlign vertical_align,
        bool text_wrap,
        int16_t rotation,
        uint8_t indent,
        bool shrink,
        BorderStyle left_border,
        BorderStyle right_border,
        BorderStyle top_border,
        BorderStyle bottom_border,
        BorderStyle diag_border,
        DiagonalBorderType diag_type,
        core::Color left_border_color,
        core::Color right_border_color,
        core::Color top_border_color,
        core::Color bottom_border_color,
        core::Color diag_border_color,
        PatternType pattern,
        core::Color bg_color,
        core::Color fg_color,
        const std::string& num_format,
        uint16_t num_format_index,
        bool locked,
        bool hidden
    );
    
    // 计算哈希值
    size_t calculateHash() const;

public:
    // 友元类Builder可以访问私有构造函数
    friend class StyleBuilder;
    // 允许内存池访问私有构造函数
    template<typename T, size_t PoolSize> friend class memory::FixedSizePool;
    
    // 拷贝和移动构造函数（由于const成员需要显式定义）
    FormatDescriptor(const FormatDescriptor& other) = default;
    FormatDescriptor(FormatDescriptor&& other) = default;
    
    // Add virtual destructor for memory pool compatibility
    virtual ~FormatDescriptor() = default;
    
    // 赋值操作符无法实现（因为所有成员都是const的）
    FormatDescriptor& operator=(const FormatDescriptor& other) = delete;
    FormatDescriptor& operator=(FormatDescriptor&& other) = delete;
    
    // 默认格式的静态创建方法
    static const FormatDescriptor& getDefault();
    
    // 获取属性（只读）
    
    const std::string& getFontName() const { return font_name_; }
    double getFontSize() const { return font_size_; }
    bool isBold() const { return bold_; }
    bool isItalic() const { return italic_; }
    UnderlineType getUnderline() const { return underline_; }
    bool isStrikeout() const { return strikeout_; }
    FontScript getFontScript() const { return script_; }
    core::Color getFontColor() const { return font_color_; }
    uint8_t getFontFamily() const { return font_family_; }
    uint8_t getFontCharset() const { return font_charset_; }
    
    HorizontalAlign getHorizontalAlign() const { return horizontal_align_; }
    VerticalAlign getVerticalAlign() const { return vertical_align_; }
    bool isTextWrap() const { return text_wrap_; }
    int16_t getRotation() const { return rotation_; }
    uint8_t getIndent() const { return indent_; }
    bool isShrink() const { return shrink_; }
    
    BorderStyle getLeftBorder() const { return left_border_; }
    BorderStyle getRightBorder() const { return right_border_; }
    BorderStyle getTopBorder() const { return top_border_; }
    BorderStyle getBottomBorder() const { return bottom_border_; }
    BorderStyle getDiagBorder() const { return diag_border_; }
    DiagonalBorderType getDiagType() const { return diag_type_; }
    
    core::Color getLeftBorderColor() const { return left_border_color_; }
    core::Color getRightBorderColor() const { return right_border_color_; }
    core::Color getTopBorderColor() const { return top_border_color_; }
    core::Color getBottomBorderColor() const { return bottom_border_color_; }
    core::Color getDiagBorderColor() const { return diag_border_color_; }
    
    PatternType getPattern() const { return pattern_; }
    core::Color getBackgroundColor() const { return bg_color_; }
    core::Color getForegroundColor() const { return fg_color_; }
    
    const std::string& getNumberFormat() const { return num_format_; }
    uint16_t getNumberFormatIndex() const { return num_format_index_; }
    
    bool isLocked() const { return locked_; }
    bool isHidden() const { return hidden_; }
    
    // 格式检查
    
    bool hasFont() const;
    bool hasFill() const; 
    bool hasBorder() const;
    bool hasAlignment() const;
    bool hasProtection() const;
    bool hasAnyFormatting() const;
    
    // 值对象操作
    
    /**
     * @brief 比较两个格式是否相等
     */
    bool operator==(const FormatDescriptor& other) const;
    bool operator!=(const FormatDescriptor& other) const { return !(*this == other); }
    
    /**
     * @brief 获取哈希值（预计算）
     */
    size_t hash() const { return hash_value_; }
    
    /**
     * @brief 创建修改版本（返回新对象）
     * 使用Builder模式创建基于当前格式的修改版本
     */
    class ModificationBuilder;
    ModificationBuilder modify() const;
};

}} // namespace fastexcel::core

// 为FormatDescriptor提供std::hash特化
namespace std {
template<>
struct hash<fastexcel::core::FormatDescriptor> {
    size_t operator()(const fastexcel::core::FormatDescriptor& desc) const {
        return desc.hash();
    }
};
} // namespace std
