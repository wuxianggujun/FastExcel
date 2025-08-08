#pragma once

#include "fastexcel/core/Color.hpp"
#include "fastexcel/theme/Theme.hpp"

namespace fastexcel { namespace theme {

class ThemeUtils {
public:
    // 将 Color 根据主题对象解析为实际 RGB（应用 tint）
    static uint32_t resolveRGB(const core::Color& color, const Theme* theme_ptr) {
        using core::Color;
        if (color.getType() != Color::Type::Theme || theme_ptr == nullptr) {
            return color.getRGB();
        }
        // 主题索引到 ThemeColorScheme::ColorType 映射
        const uint8_t idx = static_cast<uint8_t>(color.getValue());
        ThemeColorScheme::ColorType type = ThemeColorScheme::ColorType::Text1;
        switch (idx) {
            case 0: type = ThemeColorScheme::ColorType::Text1; break;           // dk1
            case 1: type = ThemeColorScheme::ColorType::Background1; break;     // lt1
            case 2: type = ThemeColorScheme::ColorType::Text2; break;           // dk2
            case 3: type = ThemeColorScheme::ColorType::Background2; break;     // lt2
            case 4: type = ThemeColorScheme::ColorType::Accent1; break;
            case 5: type = ThemeColorScheme::ColorType::Accent2; break;
            case 6: type = ThemeColorScheme::ColorType::Accent3; break;
            case 7: type = ThemeColorScheme::ColorType::Accent4; break;
            case 8: type = ThemeColorScheme::ColorType::Accent5; break;
            case 9: type = ThemeColorScheme::ColorType::Accent6; break;
            case 10: type = ThemeColorScheme::ColorType::Hyperlink; break;
            case 11: type = ThemeColorScheme::ColorType::FollowedHyperlink; break;
            default: type = ThemeColorScheme::ColorType::Text1; break;
        }
        core::Color base = theme_ptr->colors().getColor(type);
        // 应用 tint
        core::Color tinted = base;
        tinted.setTint(color.getTint());
        return tinted.getRGB();
    }
};

}} // namespace fastexcel::theme


