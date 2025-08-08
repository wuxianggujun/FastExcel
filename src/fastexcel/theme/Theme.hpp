#pragma once

#include <array>
#include <memory>
#include <optional>
#include <string>
#include <unordered_map>

#include "fastexcel/core/Color.hpp"

namespace fastexcel {
namespace theme {

// 颜色方案（12项）
class ThemeColorScheme {
public:
    enum class ColorType : uint8_t {
        Background1 = 0,
        Text1 = 1,
        Background2 = 2,
        Text2 = 3,
        Accent1 = 4,
        Accent2 = 5,
        Accent3 = 6,
        Accent4 = 7,
        Accent5 = 8,
        Accent6 = 9,
        Hyperlink = 10,
        FollowedHyperlink = 11,
    };

private:
    std::array<core::Color, 12> colors_{};

public:
    ThemeColorScheme();

    core::Color getColor(ColorType type) const;
    void setColor(ColorType type, const core::Color& color);

    // 按名称（如 "accent1"）获取/设置
    core::Color getColorByName(const std::string& name) const;
    bool setColorByName(const std::string& name, const core::Color& color);
};

// 字体方案
class ThemeFontScheme {
public:
    struct FontSet {
        std::string latin;
        std::string eastAsia;
        std::string complexScript;
    };

private:
    FontSet major_fonts_{}; // 标题字体
    FontSet minor_fonts_{}; // 正文字体

public:
    const FontSet& getMajorFonts() const { return major_fonts_; }
    const FontSet& getMinorFonts() const { return minor_fonts_; }
    void setMajorFontLatin(const std::string& name) { major_fonts_.latin = name; }
    void setMajorFontEastAsia(const std::string& name) { major_fonts_.eastAsia = name; }
    void setMajorFontComplex(const std::string& name) { major_fonts_.complexScript = name; }
    void setMinorFontLatin(const std::string& name) { minor_fonts_.latin = name; }
    void setMinorFontEastAsia(const std::string& name) { minor_fonts_.eastAsia = name; }
    void setMinorFontComplex(const std::string& name) { minor_fonts_.complexScript = name; }
};

// 完整主题
class Theme {
private:
    std::string name_ {"Office Theme"};
    ThemeColorScheme color_scheme_{};
    ThemeFontScheme font_scheme_{};

public:
    Theme() = default;
    explicit Theme(const std::string& name) : name_(name) {}

    const std::string& getName() const { return name_; }
    void setName(const std::string& name) { name_ = name; }

    ThemeColorScheme& colors() { return color_scheme_; }
    const ThemeColorScheme& colors() const { return color_scheme_; }

    ThemeFontScheme& fonts() { return font_scheme_; }
    const ThemeFontScheme& fonts() const { return font_scheme_; }

    // 生成最小可用的 OOXML 主题XML
    std::string toXML() const;
};

} // namespace theme
} // namespace fastexcel


