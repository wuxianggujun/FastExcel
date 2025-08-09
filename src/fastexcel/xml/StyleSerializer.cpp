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

    // 关键关联部分：确保最小合法的 cellStyleXfs 和 cellStyles
    writer.startElement("cellStyleXfs");
    writer.writeAttribute("count", "1");
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.endElement(); // xf
    writer.endElement(); // cellStyleXfs

    writeCellXfs(repository, writer);

    writer.startElement("cellStyles");
    writer.writeAttribute("count", "1");
    writer.startElement("cellStyle");
    writer.writeAttribute("name", "Normal");
    writer.writeAttribute("xfId", "0");
    writer.writeAttribute("builtinId", "0");
    writer.endElement(); // cellStyle
    writer.endElement(); // cellStyles
    
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
    
    // 确保count至少为2（Excel标准要求），+2是因为我们强制输出了none和gray125
    size_t fill_count = std::max<size_t>(2, unique_fills.size() + 2);
    
    writer.startElement("fills");
    writer.writeAttribute("count", std::to_string(fill_count));
    
    // 强制输出Excel标准的前两个填充
    // fillId=0: none 填充
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "none");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    // fillId=1: gray125 填充（Excel标准默认）
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "gray125");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    // 输出其余的自定义填充（从索引0开始，对应fillId=2+）
    for (size_t i = 0; i < unique_fills.size(); ++i) {
        writeFill(*unique_fills[i], writer);
    }
    
    writer.endElement(); // fills
}

void StyleSerializer::writeBorders(const core::FormatRepository& repository,
                                  xml::XMLStreamWriter& writer) {
    std::vector<std::shared_ptr<const core::FormatDescriptor>> unique_borders;
    std::vector<int> format_to_border_id;
    collectUniqueBorders(repository, unique_borders, format_to_border_id);
    
    // 确保至少一个边框组（Excel 最低要求）
    if (unique_borders.empty()) {
        writer.startElement("borders");
        writer.writeAttribute("count", "1");
        // 缺省空 border
        writer.startElement("border");
        writer.writeEmptyElement("left");
        writer.writeEmptyElement("right");
        writer.writeEmptyElement("top");
        writer.writeEmptyElement("bottom");
        writer.writeEmptyElement("diagonal");
        writer.endElement(); // border
        writer.endElement(); // borders
        return;
    }
    
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
    writeColorAttributes(format.getFontColor(), writer);
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
        if (format.getPattern() == core::PatternType::Gray125) {
            // gray125模式不需要颜色，就像原始Excel文件中的格式
        } else if (format.getPattern() == core::PatternType::Solid) {
            writer.startElement("fgColor");
            writeColorAttributes(format.getBackgroundColor(), writer);
            writer.endElement(); // fgColor
        } else {
            writer.startElement("fgColor");
            writeColorAttributes(format.getForegroundColor(), writer);
            writer.endElement(); // fgColor
            
            writer.startElement("bgColor");
            writeColorAttributes(format.getBackgroundColor(), writer);
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
        writeColorAttributes(format.getLeftBorderColor(), writer);
        writer.endElement(); // color
    }
    writer.endElement(); // left
    
    // 右边框
    writer.startElement("right");
    if (format.getRightBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getRightBorder()));
        writer.startElement("color");
        writeColorAttributes(format.getRightBorderColor(), writer);
        writer.endElement(); // color
    }
    writer.endElement(); // right
    
    // 上边框
    writer.startElement("top");
    if (format.getTopBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getTopBorder()));
        writer.startElement("color");
        writeColorAttributes(format.getTopBorderColor(), writer);
        writer.endElement(); // color
    }
    writer.endElement(); // top
    
    // 下边框
    writer.startElement("bottom");
    if (format.getBottomBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getBottomBorder()));
        writer.startElement("color");
        writeColorAttributes(format.getBottomBorderColor(), writer);
        writer.endElement(); // color
    }
    writer.endElement(); // bottom
    
    // 对角线边框
    writer.startElement("diagonal");
    if (format.getDiagBorder() != core::BorderStyle::None) {
        writer.writeAttribute("style", borderStyleToXml(format.getDiagBorder()));
        writer.startElement("color");
        writeColorAttributes(format.getDiagBorderColor(), writer);
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
    // 关联到缺省的 cellStyleXfs[0]
    writer.writeAttribute("xfId", "0");
    
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

void StyleSerializer::writeColorAttributes(const core::Color& color, xml::XMLStreamWriter& writer) {
    // 根据颜色类型写入正确的XML属性
    switch (color.getType()) {
        case core::Color::Type::Theme:
            writer.writeAttribute("theme", std::to_string(color.getValue()));
            if (color.getTint() != 0.0) {
                writer.writeAttribute("tint", std::to_string(color.getTint()));
            }
            break;
            
        case core::Color::Type::Indexed:
            writer.writeAttribute("indexed", std::to_string(color.getValue()));
            break;
            
        case core::Color::Type::Auto:
            writer.writeAttribute("auto", "1");
            break;
            
        case core::Color::Type::RGB:
        default:
            {
                // 返回带有Alpha通道的ARGB格式（Excel XML标准格式）
                std::string hex = color.toHex(false);  // 不包含#前缀
                // 确保格式为8位十六进制（包含Alpha通道）
                if (hex.length() == 6) {
                    hex = "FF" + hex;  // 添加完全不透明的Alpha通道
                }
                writer.writeAttribute("rgb", hex);
            }
            break;
    }
}

// 保留旧版本兼容性（已弃用）
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
    
    // 使用unordered_map实现O(1)查找优化，替换原来的O(n²)算法
    std::unordered_map<std::string, int> font_hash_to_id;
    
    // 预分配内存以减少重新分配
    format_to_font_id.reserve(repository.getFormatCount());
    unique_fonts.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    font_hash_to_id.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 使用哈希表实现O(1)查找
        std::string font_key = createFontHashKey(*format);
        
        auto it = font_hash_to_id.find(font_key);
        if (it != font_hash_to_id.end()) {
            // 找到匹配的字体，使用现有ID
            format_to_font_id.push_back(it->second);
        } else {
            // 添加新的唯一字体
            unique_fonts.push_back(format);
            int new_font_id = static_cast<int>(unique_fonts.size() - 1);
            font_hash_to_id[font_key] = new_font_id;
            format_to_font_id.push_back(new_font_id);
        }
    }
}

void StyleSerializer::collectUniqueFills(
    const core::FormatRepository& repository,
    std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_fills,
    std::vector<int>& format_to_fill_id) {
    
    // 使用unordered_map实现O(1)查找优化，替换原来的O(n²)算法
    std::unordered_map<std::string, int> fill_hash_to_id;
    
    // 预分配内存以减少重新分配
    format_to_fill_id.reserve(repository.getFormatCount());
    unique_fills.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    fill_hash_to_id.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 特殊处理：None模式映射到fillId=0
        if (format->getPattern() == core::PatternType::None) {
            format_to_fill_id.push_back(0);
            continue;
        }
        
        // 特殊处理：Gray125模式映射到fillId=1
        if (format->getPattern() == core::PatternType::Gray125) {
            format_to_fill_id.push_back(1);
            continue;
        }
        
        // 其他模式：使用哈希表实现O(1)查找
        std::string fill_key = createFillHashKey(*format);
        
        auto it = fill_hash_to_id.find(fill_key);
        if (it != fill_hash_to_id.end()) {
            // 找到匹配的填充，使用现有ID
            format_to_fill_id.push_back(it->second);
        } else {
            // 添加新的唯一填充
            unique_fills.push_back(format);
            int new_fill_id = static_cast<int>(unique_fills.size() - 1 + 2);  // +2 偏移
            fill_hash_to_id[fill_key] = new_fill_id;
            format_to_fill_id.push_back(new_fill_id);
        }
    }
}

void StyleSerializer::collectUniqueBorders(
    const core::FormatRepository& repository,
    std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_borders,
    std::vector<int>& format_to_border_id) {
    
    // 使用unordered_map实现O(1)查找优化，替换原来的O(n²)算法
    std::unordered_map<std::string, int> border_hash_to_id;
    
    // 预分配内存以减少重新分配
    format_to_border_id.reserve(repository.getFormatCount());
    unique_borders.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    border_hash_to_id.reserve(std::min(repository.getFormatCount(), size_t(1000)));
    
    for (const auto& format_pair : repository) {
        const auto& format = format_pair.format;
        
        // 使用哈希表实现O(1)查找
        std::string border_key = createBorderHashKey(*format);
        
        auto it = border_hash_to_id.find(border_key);
        if (it != border_hash_to_id.end()) {
            // 找到匹配的边框，使用现有ID
            format_to_border_id.push_back(it->second);
        } else {
            // 添加新的唯一边框
            unique_borders.push_back(format);
            int new_border_id = static_cast<int>(unique_borders.size() - 1);
            border_hash_to_id[border_key] = new_border_id;
            format_to_border_id.push_back(new_border_id);
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

// ========== 哈希键生成函数（性能优化）==========

std::string StyleSerializer::createFillHashKey(const core::FormatDescriptor& format) {
    // Gray125是特殊模式，单独处理
    if (format.getPattern() == core::PatternType::Gray125) {
        return "gray125";
    }
    
    // 生成填充哈希键：模式|背景色|前景色
    return std::to_string(static_cast<int>(format.getPattern())) + "|" +
           format.getBackgroundColor().toHex(false) + "|" +
           format.getForegroundColor().toHex(false);
}

std::string StyleSerializer::createFontHashKey(const core::FormatDescriptor& format) {
    // 生成字体哈希键：名称|大小|粗体|斜体|下划线|删除线|脚本|颜色|字族|字符集
    return format.getFontName() + "|" +
           std::to_string(format.getFontSize()) + "|" +
           std::to_string(format.isBold()) + "|" +
           std::to_string(format.isItalic()) + "|" +
           std::to_string(static_cast<int>(format.getUnderline())) + "|" +
           std::to_string(format.isStrikeout()) + "|" +
           std::to_string(static_cast<int>(format.getFontScript())) + "|" +
           format.getFontColor().toHex(false) + "|" +
           std::to_string(format.getFontFamily()) + "|" +
           std::to_string(format.getFontCharset());
}

std::string StyleSerializer::createBorderHashKey(const core::FormatDescriptor& format) {
    // 生成边框哈希键：左|右|上|下|对角线|对角线类型|左色|右色|上色|下色|对角线色
    return std::to_string(static_cast<int>(format.getLeftBorder())) + "|" +
           std::to_string(static_cast<int>(format.getRightBorder())) + "|" +
           std::to_string(static_cast<int>(format.getTopBorder())) + "|" +
           std::to_string(static_cast<int>(format.getBottomBorder())) + "|" +
           std::to_string(static_cast<int>(format.getDiagBorder())) + "|" +
           std::to_string(static_cast<int>(format.getDiagType())) + "|" +
           format.getLeftBorderColor().toHex(false) + "|" +
           format.getRightBorderColor().toHex(false) + "|" +
           format.getTopBorderColor().toHex(false) + "|" +
           format.getBottomBorderColor().toHex(false) + "|" +
           format.getDiagBorderColor().toHex(false);
}

}} // namespace fastexcel::xml