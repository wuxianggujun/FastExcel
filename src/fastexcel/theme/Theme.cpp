#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"

namespace fastexcel {
namespace theme {

ThemeColorScheme::ThemeColorScheme() {
    // 简要默认：文本黑、背景白，其余用常见Office主题色（可后续由解析覆盖）
    setColor(ColorType::Text1, core::Color(static_cast<uint32_t>(0x000000)));
    setColor(ColorType::Background1, core::Color(static_cast<uint32_t>(0xFFFFFF)));
    setColor(ColorType::Text2, core::Color(static_cast<uint32_t>(0x1F497D)));
    setColor(ColorType::Background2, core::Color(static_cast<uint32_t>(0xEEECE1)));
    setColor(ColorType::Accent1, core::Color(static_cast<uint32_t>(0x4F81BD)));
    setColor(ColorType::Accent2, core::Color(static_cast<uint32_t>(0xC0504D)));
    setColor(ColorType::Accent3, core::Color(static_cast<uint32_t>(0x9BBB59)));
    setColor(ColorType::Accent4, core::Color(static_cast<uint32_t>(0x8064A2)));
    setColor(ColorType::Accent5, core::Color(static_cast<uint32_t>(0x4BACC6)));
    setColor(ColorType::Accent6, core::Color(static_cast<uint32_t>(0xF79646)));
    setColor(ColorType::Hyperlink, core::Color(static_cast<uint32_t>(0x0000FF)));
    setColor(ColorType::FollowedHyperlink, core::Color(static_cast<uint32_t>(0x800080)));
}

core::Color ThemeColorScheme::getColor(ColorType type) const {
    return colors_[static_cast<size_t>(type)];
}

void ThemeColorScheme::setColor(ColorType type, const core::Color& color) {
    colors_[static_cast<size_t>(type)] = color;
}

static inline bool iequals(const std::string& a, const std::string& b) {
    if (a.size() != b.size()) return false;
    for (size_t i = 0; i < a.size(); ++i) {
        char ca = static_cast<char>(::tolower(a[i]));
        char cb = static_cast<char>(::tolower(b[i]));
        if (ca != cb) return false;
    }
    return true;
}

core::Color ThemeColorScheme::getColorByName(const std::string& name) const {
    // 接受 lt1/dk1/lt2/dk2/accent1..6/hlink/folHlink 或背景/文本别名
    auto toType = [](const std::string& n) -> std::optional<ColorType> {
        if (iequals(n, "lt1") || iequals(n, "background1")) return ColorType::Background1;
        if (iequals(n, "dk1") || iequals(n, "text1")) return ColorType::Text1;
        if (iequals(n, "lt2") || iequals(n, "background2")) return ColorType::Background2;
        if (iequals(n, "dk2") || iequals(n, "text2")) return ColorType::Text2;
        if (iequals(n, "accent1")) return ColorType::Accent1;
        if (iequals(n, "accent2")) return ColorType::Accent2;
        if (iequals(n, "accent3")) return ColorType::Accent3;
        if (iequals(n, "accent4")) return ColorType::Accent4;
        if (iequals(n, "accent5")) return ColorType::Accent5;
        if (iequals(n, "accent6")) return ColorType::Accent6;
        if (iequals(n, "hlink") || iequals(n, "hyperlink")) return ColorType::Hyperlink;
        if (iequals(n, "folHlink") || iequals(n, "followedhyperlink")) return ColorType::FollowedHyperlink;
        return std::nullopt;
    };

    if (auto t = toType(name)) {
        return getColor(*t);
    }
    return core::Color(static_cast<uint32_t>(0x000000));
}

bool ThemeColorScheme::setColorByName(const std::string& name, const core::Color& color) {
    auto toType = [](const std::string& n) -> std::optional<ColorType> {
        if (iequals(n, "lt1") || iequals(n, "background1")) return ColorType::Background1;
        if (iequals(n, "dk1") || iequals(n, "text1")) return ColorType::Text1;
        if (iequals(n, "lt2") || iequals(n, "background2")) return ColorType::Background2;
        if (iequals(n, "dk2") || iequals(n, "text2")) return ColorType::Text2;
        if (iequals(n, "accent1")) return ColorType::Accent1;
        if (iequals(n, "accent2")) return ColorType::Accent2;
        if (iequals(n, "accent3")) return ColorType::Accent3;
        if (iequals(n, "accent4")) return ColorType::Accent4;
        if (iequals(n, "accent5")) return ColorType::Accent5;
        if (iequals(n, "accent6")) return ColorType::Accent6;
        if (iequals(n, "hlink") || iequals(n, "hyperlink")) return ColorType::Hyperlink;
        if (iequals(n, "folHlink") || iequals(n, "followedhyperlink")) return ColorType::FollowedHyperlink;
        return std::nullopt;
    };

    if (auto t = toType(name)) {
        setColor(*t, color);
        return true;
    }
    return false;
}

static void writeSchemeColor(xml::XMLStreamWriter& w, const char* tag, const core::Color& c) {
    w.startElement(tag);
    w.startElement("a:srgbClr");
    auto hex = c.toHex(false);
    // 确保为6位
    if (hex.size() == 8) hex = hex.substr(2);
    if (hex.size() < 6) hex = std::string(6 - hex.size(), '0') + hex;
    w.writeAttribute("val", hex);
    w.endElement(); // a:srgbClr
    w.endElement(); // tag
}

std::string Theme::toXML() const {
    std::string out;
    xml::XMLStreamWriter writer([&](const std::string& data) { out.append(data); });

    writer.startDocument();
    writer.startElement("a:theme");
    writer.writeAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    writer.writeAttribute("name", name_);

    writer.startElement("a:themeElements");

    // 颜色方案
    writer.startElement("a:clrScheme");
    writer.writeAttribute("name", name_);

    writeSchemeColor(writer, "a:dk1", colors().getColor(ThemeColorScheme::ColorType::Text1));
    writeSchemeColor(writer, "a:lt1", colors().getColor(ThemeColorScheme::ColorType::Background1));
    writeSchemeColor(writer, "a:dk2", colors().getColor(ThemeColorScheme::ColorType::Text2));
    writeSchemeColor(writer, "a:lt2", colors().getColor(ThemeColorScheme::ColorType::Background2));
    writeSchemeColor(writer, "a:accent1", colors().getColor(ThemeColorScheme::ColorType::Accent1));
    writeSchemeColor(writer, "a:accent2", colors().getColor(ThemeColorScheme::ColorType::Accent2));
    writeSchemeColor(writer, "a:accent3", colors().getColor(ThemeColorScheme::ColorType::Accent3));
    writeSchemeColor(writer, "a:accent4", colors().getColor(ThemeColorScheme::ColorType::Accent4));
    writeSchemeColor(writer, "a:accent5", colors().getColor(ThemeColorScheme::ColorType::Accent5));
    writeSchemeColor(writer, "a:accent6", colors().getColor(ThemeColorScheme::ColorType::Accent6));
    writeSchemeColor(writer, "a:hlink", colors().getColor(ThemeColorScheme::ColorType::Hyperlink));
    writeSchemeColor(writer, "a:folHlink", colors().getColor(ThemeColorScheme::ColorType::FollowedHyperlink));

    writer.endElement(); // a:clrScheme

    // 字体方案
    writer.startElement("a:fontScheme");
    writer.writeAttribute("name", name_);

    // majorFont
    writer.startElement("a:majorFont");
    writer.startElement("a:latin");
    writer.writeAttribute("typeface", fonts().getMajorFonts().latin);
    writer.endElement();
    writer.startElement("a:ea");
    writer.writeAttribute("typeface", fonts().getMajorFonts().eastAsia);
    writer.endElement();
    writer.startElement("a:cs");
    writer.writeAttribute("typeface", fonts().getMajorFonts().complexScript);
    writer.endElement();
    writer.endElement(); // a:majorFont

    // minorFont
    writer.startElement("a:minorFont");
    writer.startElement("a:latin");
    writer.writeAttribute("typeface", fonts().getMinorFonts().latin);
    writer.endElement();
    writer.startElement("a:ea");
    writer.writeAttribute("typeface", fonts().getMinorFonts().eastAsia);
    writer.endElement();
    writer.startElement("a:cs");
    writer.writeAttribute("typeface", fonts().getMinorFonts().complexScript);
    writer.endElement();
    writer.endElement(); // a:minorFont

    writer.endElement(); // a:fontScheme

    writer.endElement(); // a:themeElements
    writer.endElement(); // a:theme
    writer.endDocument();

    return out;
}

} // namespace theme
} // namespace fastexcel


