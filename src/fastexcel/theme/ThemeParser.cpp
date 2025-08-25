#include "fastexcel/theme/ThemeParser.hpp"

namespace fastexcel {
namespace theme {

using xml::XMLStreamReader;

static inline uint32_t parseRGB(const std::string& srgb) {
    std::string hex = srgb;
    if (hex.size() == 8) hex = hex.substr(2);
    if (hex.size() < 6) return 0;
    try {
        return static_cast<uint32_t>(std::stoul(hex, nullptr, 16));
    } catch (const std::invalid_argument& e) {
        FASTEXCEL_LOG_DEBUG("Invalid RGB hex format '{}': {}", srgb, e.what());
        return 0;
    } catch (const std::out_of_range& e) {
        FASTEXCEL_LOG_DEBUG("RGB hex value out of range '{}': {}", srgb, e.what());
        return 0;
    }
}

std::unique_ptr<Theme> ThemeParser::parseFromXML(const std::string& xml) {
    XMLStreamReader reader;
    auto dom = reader.parseToDOM(xml);
    if (!dom) return nullptr;

    // 查找 a:themeElements
    XMLStreamReader::SimpleElement* themeEl = nullptr;
    if (dom->name.find(":theme") != std::string::npos || dom->name == "theme") {
        themeEl = dom.get();
    } else {
        for (const auto& child_up : dom->children) {
            const auto* child = child_up.get();
            if (!child) continue;
            if (child->name.find(":theme") != std::string::npos || child->name == "theme") {
                themeEl = const_cast<XMLStreamReader::SimpleElement*>(child);
                break;
            }
        }
    }
    if (!themeEl) return nullptr;

    auto result = std::make_unique<Theme>();
    for (const auto& child_up : themeEl->children) {
        const auto* child = child_up.get();
        if (!child) continue;
        if (child->name.find("themeElements") != std::string::npos) {
            for (const auto& e_up : child->children) {
                const auto* e = e_up.get();
                if (!e) continue;
                if (e->name.find("clrScheme") != std::string::npos) {
                    parseColorScheme(const_cast<XMLStreamReader::SimpleElement*>(e), *result);
                } else if (e->name.find("fontScheme") != std::string::npos) {
                    parseFontScheme(const_cast<XMLStreamReader::SimpleElement*>(e), *result);
                }
            }
        }
    }
    return result;
}

void ThemeParser::parseColorScheme(XMLStreamReader::SimpleElement* clrScheme, Theme& out) {
    if (!clrScheme) return;
    // 子元素如 a:dk1, a:lt1, a:accent1...
    for (const auto& c_up : clrScheme->children) {
        const auto* c = c_up.get();
        if (!c) continue;
        std::string tag = c->name;
        // 在子节点下查找 a:srgbClr 或 a:sysClr
        core::Color color(static_cast<uint32_t>(0x000000));
        for (const auto& inner_up : c->children) {
            const auto* inner = inner_up.get();
            if (!inner) continue;
            if (inner->name.find("srgbClr") != std::string::npos) {
                auto it = inner->attributes.find("val");
                if (it != inner->attributes.end()) {
                    color = core::Color(static_cast<uint32_t>(parseRGB(it->second)));
                }
            } else if (inner->name.find("sysClr") != std::string::npos) {
                // 简化：sysClr 映射为常见黑白
                auto it = inner->attributes.find("lastClr");
                if (it != inner->attributes.end()) {
                    color = core::Color(static_cast<uint32_t>(parseRGB(it->second)));
                }
            }
        }

        // 将颜色写入方案
        if (tag.find("dk1") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Text1, color);
        else if (tag.find("lt1") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Background1, color);
        else if (tag.find("dk2") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Text2, color);
        else if (tag.find("lt2") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Background2, color);
        else if (tag.find("accent1") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent1, color);
        else if (tag.find("accent2") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent2, color);
        else if (tag.find("accent3") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent3, color);
        else if (tag.find("accent4") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent4, color);
        else if (tag.find("accent5") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent5, color);
        else if (tag.find("accent6") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Accent6, color);
        else if (tag.find("hlink") != std::string::npos && tag.find("folHlink") == std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::Hyperlink, color);
        else if (tag.find("folHlink") != std::string::npos) out.colors().setColor(ThemeColorScheme::ColorType::FollowedHyperlink, color);
    }
}

void ThemeParser::parseFontScheme(XMLStreamReader::SimpleElement* fontScheme, Theme& out) {
    if (!fontScheme) return;
    for (const auto& e_up : fontScheme->children) {
        const auto* e = e_up.get();
        if (!e) continue;
        if (e->name.find("majorFont") != std::string::npos) {
            for (const auto& f_up : e->children) {
                const auto* f = f_up.get();
                if (!f) continue;
                auto it = f->attributes.find("typeface");
                if (it == f->attributes.end()) continue;
                if (f->name.find("latin") != std::string::npos) out.fonts().setMajorFontLatin(it->second);
                else if (f->name.find("ea") != std::string::npos) out.fonts().setMajorFontEastAsia(it->second);
                else if (f->name.find("cs") != std::string::npos) out.fonts().setMajorFontComplex(it->second);
            }
        } else if (e->name.find("minorFont") != std::string::npos) {
            for (const auto& f_up : e->children) {
                const auto* f = f_up.get();
                if (!f) continue;
                auto it = f->attributes.find("typeface");
                if (it == f->attributes.end()) continue;
                if (f->name.find("latin") != std::string::npos) out.fonts().setMinorFontLatin(it->second);
                else if (f->name.find("ea") != std::string::npos) out.fonts().setMinorFontEastAsia(it->second);
                else if (f->name.find("cs") != std::string::npos) out.fonts().setMinorFontComplex(it->second);
            }
        }
    }
}

} // namespace theme
} // namespace fastexcel


