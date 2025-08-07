#include "StyleSerializer.hpp"
#include <unordered_set>

namespace fastexcel {

namespace xml {

void StyleSerializer::serialize(const core::FormatRepository& repository, 
                               xml::XMLStreamWriter& writer) {
    writeStyleSheet(repository, writer);
}

void StyleSerializer::serialize(const core::FormatRepository& repository,
                               const std::function<void(const char*, size_t)>& callback) {
    xml::XMLStreamWriter writer(callback);
    serialize(repository, writer);
}

void StyleSerializer::serializeToFile(const core::FormatRepository& repository,
                                     const std::string& filename) {
    xml::XMLStreamWriter writer(filename);
    serialize(repository, writer);
}

void StyleSerializer::writeStyleSheet(const core::FormatRepository& repository,
                                     xml::XMLStreamWriter& writer) {
    writer.startDocument();
    
    writer.startElement("styleSheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    writer.writeAttribute("mc:Ignorable", "x14ac x16r2 xr");
    writer.writeAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
    writer.writeAttribute("xmlns:x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
    writer.writeAttribute("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
    
    // 写入各个部分
    writeNumberFormats(repository, writer);
    writeFonts(repository, writer);
    writeFills(repository, writer);
    writeBorders(repository, writer);
    writeCellXfs(repository, writer);
    
    writer.endElement(); // styleSheet
    writer.endDocument();
}

void StyleSerializer::writeNumberFormats(const core::FormatRepository& repository,
                                        xml::XMLStreamWriter& writer) {
    // 收集自定义数字格式
    std::vector<std::string> unique_numfmts;
    std::vector<int> format_to_numfmt_id;
    collectUniqueNumberFormats(repository, unique_numfmts, format_to_numfmt_id);
    
    if (unique_numfmts.empty()) {
        return; // 没有自定义数字格式
    }
    
    writer.startElement("numFmts");
    writer.writeAttribute("count", std::to_string(unique_numfmts.size()));
    
    int custom_id = 164; // Excel自定义格式从164开始
    for (const auto& numfmt : unique_numfmts) {
        writer.startElement("numFmt");
        writer.writeAttribute("numFmtId", std::to_string(custom_id++));
        writer.writeAttribute("formatCode", numfmt);
        writer.endElement(); // numFmt
    }
    
    writer.endElement(); // numFmts
}

void StyleSerializer::writeFonts(const core::FormatRepository& repository,
                                xml::XMLStreamWriter& writer) {
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_fonts;
    std::vector<int> format_to_font_id;
    collectUniqueFonts(repository, unique_fonts, format_to_font_id);
    
    writer.startElement("fonts");
    writer.writeAttribute("count", std::to_string(unique_fonts.size()));
    writer.writeAttribute("x14ac:knownFonts", "1");
    
    for (const auto& font : unique_fonts) {
        writeFont(*font, writer);
    }
    
    writer.endElement(); // fonts
}

void StyleSerializer::writeFills(const core::FormatRepository& repository,
                                xml::XMLStreamWriter& writer) {
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_fills;
    std::vector<int> format_to_fill_id;
    collectUniqueFills(repository, unique_fills, format_to_fill_id);
    
    writer.startElement("fills");
    writer.writeAttribute("count", std::to_string(unique_fills.size()));
    
    for (const auto& fill : unique_fills) {
        writeFill(*fill, writer);
    }
    
    writer.endElement(); // fills
}

void StyleSerializer::writeBorders(const core::FormatRepository& repository,
                                  xml::XMLStreamWriter& writer) {
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_borders;
    std::vector<int> format_to_border_id;
    collectUniqueBorders(repository, unique_borders, format_to_border_id);
    
    writer.startElement("borders");
    writer.writeAttribute("count", std::to_string(unique_borders.size()));
    
    for (const auto& border : unique_borders) {
        writeBorder(*border, writer);
    }
    
    writer.endElement(); // borders
}

void StyleSerializer::writeCellXfs(const core::FormatRepository& repository,
                                  xml::XMLStreamWriter& writer) {
    // 创建组件映射
    std::vector<int> font_mapping, fill_mapping, border_mapping, numfmt_mapping;
    createComponentMappings(repository, font_mapping, fill_mapping, border_mapping, numfmt_mapping);
    
    writer.startElement("cellXfs");
    writer.writeAttribute("count", std::to_string(repository.getFormatCount()));
    
    // 遍历所有格式
    for (const auto& format_pair : repository) {
        int format_id = format_pair.id;
        const auto& format = format_pair.format;
        
        int font_id = font_mapping[format_id];
        int fill_id = fill_mapping[format_id];
        int border_id = border_mapping[format_id];
        int numfmt_id = numfmt_mapping[format_id];
        
        writeCellXf(*format, font_id, fill_id, border_id, numfmt_id, writer);
    }
    
    writer.endElement(); // cellXfs
}

void StyleSerializer::writeFont(const core::FormatDescriptor& format,
                               xml::XMLStreamWriter& writer) {
    writer.startElement("font");
    
    if (format.isBold()) {
        writer.writeEmptyElement("b");
    }
    
    if (format.isItalic()) {
        writer.writeEmptyElement("i");
    }
    
    if (format.getUnderline() != core::UnderlineType::None) {
        writer.startElement("u");
        if (format.getUnderline() != core::UnderlineType::Single) {
            writer.writeAttribute("val", underlineTypeToXml(format.getUnderline()));
        }
        writer.endElement(); // u
    }
    
    if (format.isStrikeout()) {
        writer.writeEmptyElement("strike");
    }
    
    // 字体大小
    writer.startElement("sz");
    writer.writeAttribute("val", std::to_string(format.getFontSize()));
    writer.endElement(); // sz
    
    // 字体颜色
    writer.startElement("color");
    writer.writeAttribute("rgb", colorToXml(format.getFontColor()));
    writer.endElement(); // color
    
    // 字体名称
    writer.startElement("name");
    writer.writeAttribute("val", format.getFontName());
    writer.endElement(); // name
    
    // 字体族
    writer.startElement("family");
    writer.writeAttribute("val", std::to_string(format.getFontFamily()));
    writer.endElement(); // family
    
    // 字符集
    writer.startElement("charset");
    writer.writeAttribute("val", std::to_string(format.getFontCharset()));
    writer.endElement(); // charset
    
    writer.endElement(); // font
}

void StyleSerializer::writeFill(const core::FormatDescriptor& format,
                               xml::XMLStreamWriter& writer) {
    writer.startElement("fill");
    
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", patternTypeToXml(format.getPattern()));
    
    if (format.getPattern() != core::PatternType::None) {
        if (format.getPattern() == core::PatternType::Solid) {
            writer.startElement("fgColor");
            writer.writeAttribute("rgb", colorToXml(format.getBackgroundColor()));
            writer.endElement(); // fgColor
        } else {
            writer.startElement("fgColor");
            writer.writeAttribute("rgb", colorToXml(format.getForegroundColor()));
            writer.endElement(); // fgColor
            
            writer.startElement("bgColor");
            writer.writeAttribute("rgb", colorToXml(format.getBackgroundColor()));
            writer.endElement(); // bgColor
        }
    }
    
    writer.endElement(); // patternFill
    writer.endElement(); // fill
}

void StyleSerializer::writeBorder(const core::FormatDescriptor& format,
                                 xml::XMLStreamWriter& writer) {
    writer.startElement("border");
    
    // 左边框
    writer.startElement("left");
    if (format.getLeftBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getLeftBorder()));
        writer.startElement("color");
        writer.writeAttribute("rgb", colorToXml(format.getLeftBorderColor()));
        writer.endElement(); // color
    }
    writer.endElement(); // left
    
    // 右边框
    writer.startElement("right");
    if (format.getRightBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getRightBorder()));
        writer.startElement("color");
        writer.writeAttribute("rgb", colorToXml(format.getRightBorderColor()));
        writer.endElement(); // color
    }
    writer.endElement(); // right
    
    // 上边框
    writer.startElement("top");
    if (format.getTopBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getTopBorder()));
        writer.startElement("color");
        writer.writeAttribute("rgb", colorToXml(format.getTopBorderColor()));
        writer.endElement(); // color
    }
    writer.endElement(); // top
    
    // 下边框
    writer.startElement("bottom");
    if (format.getBottomBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getBottomBorder()));
        writer.startElement("color");
        writer.writeAttribute("rgb", colorToXml(format.getBottomBorderColor()));
        writer.endElement(); // color
    }
    writer.endElement(); // bottom
    
    // 对角线边框
    writer.startElement("diagonal");
    if (format.getDiagBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getDiagBorder()));
        writer.startElement("color");
        writer.writeAttribute("rgb", colorToXml(format.getDiagBorderColor()));
        writer.endElement(); // color
    }
    writer.endElement(); // diagonal
    
    writer.endElement(); // border
}

void StyleSerializer::writeCellXf(const core::FormatDescriptor& format,
                                 int font_id, int fill_id, int border_id, int num_fmt_id,
                                 xml::XMLStreamWriter& writer) {
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", std::to_string(num_fmt_id));
    writer.writeAttribute("fontId", std::to_string(font_id));
    writer.writeAttribute("fillId", std::to_string(fill_id));
    writer.writeAttribute("borderId", std::to_string(border_id));
    
    // 应用标志
    if (num_fmt_id > 0) {
        writer.writeAttribute("applyNumberFormat", "1");
    }
    if (format.hasFont()) {
        writer.writeAttribute("applyFont", "1");
    }
    if (format.hasFill()) {
        writer.writeAttribute("applyFill", "1");
    }
    if (format.hasBorder()) {
        writer.writeAttribute("applyBorder", "1");
    }
    if (needsAlignment(format)) {
        writer.writeAttribute("applyAlignment", "1");
        writeAlignment(format, writer);
    }
    if (needsProtection(format)) {
        writer.writeAttribute("applyProtection", "1");
        writeProtection(format, writer);
    }
    
    writer.endElement(); // xf
}

void StyleSerializer::writeAlignment(const core::FormatDescriptor& format,
                                    xml::XMLStreamWriter& writer) {
    writer.startElement("alignment");
    
    if (format.getHorizontalAlign() != core::HorizontalAlign::None) {
        writer.writeAttribute("horizontal", horizontalAlignToXml(format.getHorizontalAlign()));
    }
    
    if (format.getVerticalAlign() != core::VerticalAlign::Bottom) {
        writer.writeAttribute("vertical", verticalAlignToXml(format.getVerticalAlign()));
    }
    
    if (format.getRotation() != 0) {
        writer.writeAttribute("textRotation", std::to_string(format.getRotation()));
    }
    
    if (format.getIndent() > 0) {
        writer.writeAttribute("indent", std::to_string(format.getIndent()));
    }
    
    if (format.isTextWrap()) {
        writer.writeAttribute("wrapText", "1");
    }
    
    if (format.isShrink()) {
        writer.writeAttribute("shrinkToFit", "1");
    }
    
    writer.endElement(); // alignment
}

void StyleSerializer::writeProtection(const core::FormatDescriptor& format,
                                     xml::XMLStreamWriter& writer) {
    writer.startElement("protection");
    
    if (!format.isLocked()) {
        writer.writeAttribute("locked", "0");
    }
    
    if (format.isHidden()) {
        writer.writeAttribute("hidden", "1");
    }
    
    writer.endElement(); // protection
}

// ========== 辅助方法实现 ==========

std::string StyleSerializer::borderStyleToXml(core::BorderStyle style) {
    switch (style) {
        case core::BorderStyle::None: return "none";
        case core::BorderStyle::Thin: return "thin";
        case core::BorderStyle::Medium: return "medium";
        case core::BorderStyle::Thick: return "thick";
        case core::BorderStyle::Double: return "double";
        case core::BorderStyle::Hair: return "hair";
        case core::BorderStyle::Dotted: return "dotted";
        case core::BorderStyle::Dashed: return "dashed";
        case core::BorderStyle::DashDot: return "dashDot";
        case core::BorderStyle::DashDotDot: return "dashDotDot";
        case core::BorderStyle::MediumDashed: return "mediumDashed";
        case core::BorderStyle::MediumDashDot: return "mediumDashDot";
        case core::BorderStyle::MediumDashDotDot: return "mediumDashDotDot";
        case core::BorderStyle::SlantDashDot: return "slantDashDot";
        default: return "none";
    }
}

std::string StyleSerializer::patternTypeToXml(core::PatternType pattern) {
    switch (pattern) {
        case core::PatternType::None: return "none";
        case core::PatternType::Solid: return "solid";
        case core::PatternType::MediumGray: return "mediumGray";
        case core::PatternType::DarkGray: return "darkGray";
        case core::PatternType::LightGray: return "lightGray";
        case core::PatternType::DarkHorizontal: return "darkHorizontal";
        case core::PatternType::DarkVertical: return "darkVertical";
        case core::PatternType::DarkDown: return "darkDown";
        case core::PatternType::DarkUp: return "darkUp";
        case core::PatternType::DarkGrid: return "darkGrid";
        case core::PatternType::DarkTrellis: return "darkTrellis";
        case core::PatternType::LightHorizontal: return "lightHorizontal";
        case core::PatternType::LightVertical: return "lightVertical";
        case core::PatternType::LightDown: return "lightDown";
        case core::PatternType::LightUp: return "lightUp";
        case core::PatternType::LightGrid: return "lightGrid";
        case core::PatternType::LightTrellis: return "lightTrellis";
        case core::PatternType::Gray125: return "gray125";
        case core::PatternType::Gray0625: return "gray0625";
        default: return "none";
    }
}

std::string StyleSerializer::underlineTypeToXml(core::UnderlineType underline) {
    switch (underline) {
        case core::UnderlineType::Single: return "single";
        case core::UnderlineType::Double: return "double";
        case core::UnderlineType::SingleAccounting: return "singleAccounting";
        case core::UnderlineType::DoubleAccounting: return "doubleAccounting";
        default: return "none";
    }
}

std::string StyleSerializer::horizontalAlignToXml(core::HorizontalAlign align) {
    switch (align) {
        case core::HorizontalAlign::Left: return "left";
        case core::HorizontalAlign::Center: return "center";
        case core::HorizontalAlign::Right: return "right";
        case core::HorizontalAlign::Fill: return "fill";
        case core::HorizontalAlign::Justify: return "justify";
        case core::HorizontalAlign::CenterAcross: return "centerContinuous";
        case core::HorizontalAlign::Distributed: return "distributed";
        default: return "general";
    }
}

std::string StyleSerializer::verticalAlignToXml(core::VerticalAlign align) {
    switch (align) {
        case core::VerticalAlign::Top: return "top";
        case core::VerticalAlign::Center: return "center";
        case core::VerticalAlign::Bottom: return "bottom";
        case core::VerticalAlign::Justify: return "justify";
        case core::VerticalAlign::Distributed: return "distributed";
        default: return "bottom";
    }
}

std::string StyleSerializer::colorToXml(const core::Color& color) {
    // 返回带有Alpha通道的ARGB格式（Excel XML标准格式）
    std::string hex = color.toHex(false);  // 不包含#前缀
    
    // 确保格式为8位十六进制（包含Alpha通道）
    if (hex.length() == 6) {
        return "FF" + hex;  // 添加完全不透明的Alpha通道
    }
    return hex;
}

bool StyleSerializer::needsAlignment(const core::FormatDescriptor& format) {
    return format.hasAlignment();
}

bool StyleSerializer::needsProtection(const core::FormatDescriptor& format) {
    return format.hasProtection();
}

void StyleSerializer::createComponentMappings(
    const core::FormatRepository& repository,
    std::vector<int>& font_mapping,
    std::vector<int>& fill_mapping,
    std::vector<int>& border_mapping,
    std::vector<int>& numfmt_mapping) {
    
    // 收集去重后的组件
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_fonts;
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_fills;
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_borders;
    std::vector<std::string> unique_numfmts;
    
    collectUniqueFonts(repository, unique_fonts, font_mapping);
    collectUniqueFills(repository, unique_fills, fill_mapping);
    collectUniqueBorders(repository, unique_borders, border_mapping);
    collectUniqueNumberFormats(repository, unique_numfmts, numfmt_mapping);
}

void StyleSerializer::collectUniqueFonts(
    const core::FormatRepository& repository,
    std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_fonts,
    std::vector<int>& format_to_font_id) {
    
    // 使用lambda比较字体属性是否相同
    auto fontEquals = [](const core::FormatDescriptor& a, const core::FormatDescriptor& b) {
        return a.getFontName() == b.getFontName() &&
               a.getFontSize() == b.getFontSize() &&
               a.isBold() == b.isBold() &&
               a.isItalic() == b.isItalic() &&
               a.getUnderline() == b.getUnderline() &&
               a.isStrikeout() == b.isStrikeout() &&
               a.getFontScript() == b.getFontScript() &&
               a.getFontColor() == b.getFontColor() &&
               a.getFontFamily() == b.getFontFamily() &&
               a.getFontCharset() == b.getFontCharset();
    };
    
    format_to_font_id.reserve(repository.getFormatCount());
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 查找是否已存在相同的字体
        bool found = false;
        for (size_t i = 0; i < unique_fonts.size(); ++i) {
            if (fontEquals(*format, *unique_fonts[i])) {
                format_to_font_id.push_back(static_cast<int>(i));
                found = true;
                break;
            }
        }
        
        // 如果没有找到相同字体，添加新的唯一字体
        if (!found) {
            unique_fonts.push_back(format);
            format_to_font_id.push_back(static_cast<int>(unique_fonts.size() - 1));
        }
    }
}

void StyleSerializer::collectUniqueFills(
    const core::FormatRepository& repository,
    std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_fills,
    std::vector<int>& format_to_fill_id) {
    
    // 使用lambda比较填充属性是否相同
    auto fillEquals = [](const core::FormatDescriptor& a, const core::FormatDescriptor& b) {
        return a.getPattern() == b.getPattern() &&
               a.getBackgroundColor() == b.getBackgroundColor() &&
               a.getForegroundColor() == b.getForegroundColor();
    };
    
    format_to_fill_id.reserve(repository.getFormatCount());
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 查找是否已存在相同的填充
        bool found = false;
        for (size_t i = 0; i < unique_fills.size(); ++i) {
            if (fillEquals(*format, *unique_fills[i])) {
                format_to_fill_id.push_back(static_cast<int>(i));
                found = true;
                break;
            }
        }
        
        // 如果没有找到相同填充，添加新的唯一填充
        if (!found) {
            unique_fills.push_back(format);
            format_to_fill_id.push_back(static_cast<int>(unique_fills.size() - 1));
        }
    }
}

void StyleSerializer::collectUniqueBorders(
    const core::FormatRepository& repository,
    std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_borders,
    std::vector<int>& format_to_border_id) {
    
    // 使用lambda比较边框属性是否相同
    auto borderEquals = [](const core::FormatDescriptor& a, const core::FormatDescriptor& b) {
        return a.getLeftBorder() == b.getLeftBorder() &&
               a.getRightBorder() == b.getRightBorder() &&
               a.getTopBorder() == b.getTopBorder() &&
               a.getBottomBorder() == b.getBottomBorder() &&
               a.getDiagBorder() == b.getDiagBorder() &&
               a.getDiagType() == b.getDiagType() &&
               a.getLeftBorderColor() == b.getLeftBorderColor() &&
               a.getRightBorderColor() == b.getRightBorderColor() &&
               a.getTopBorderColor() == b.getTopBorderColor() &&
               a.getBottomBorderColor() == b.getBottomBorderColor() &&
               a.getDiagBorderColor() == b.getDiagBorderColor();
    };
    
    format_to_border_id.reserve(repository.getFormatCount());
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 查找是否已存在相同的边框
        bool found = false;
        for (size_t i = 0; i < unique_borders.size(); ++i) {
            if (borderEquals(*format, *unique_borders[i])) {
                format_to_border_id.push_back(static_cast<int>(i));
                found = true;
                break;
            }
        }
        
        // 如果没有找到相同边框，添加新的唯一边框
        if (!found) {
            unique_borders.push_back(format);
            format_to_border_id.push_back(static_cast<int>(unique_borders.size() - 1));
        }
    }
}

void StyleSerializer::collectUniqueNumberFormats(
    const core::FormatRepository& repository,
    std::vector<std::string>& unique_numfmts,
    std::vector<int>& format_to_numfmt_id) {
    
    std::unordered_set<std::string> seen;
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        const std::string& numfmt = format->getNumberFormat();
        
        if (!numfmt.empty() && seen.find(numfmt) == seen.end()) {
            unique_numfmts.push_back(numfmt);
            seen.insert(numfmt);
        }
        
        if (numfmt.empty()) {
            format_to_numfmt_id.push_back(format->getNumberFormatIndex());
        } else {
            auto it = std::find(unique_numfmts.begin(), unique_numfmts.end(), numfmt);
            format_to_numfmt_id.push_back(164 + static_cast<int>(it - unique_numfmts.begin()));
        }
    }
}

}} // namespace fastexcel::xml