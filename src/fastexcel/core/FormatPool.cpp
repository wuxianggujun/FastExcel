#include "fastexcel/core/FormatPool.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/core/Exception.hpp"
#include <set>
#include <map>

namespace fastexcel {
namespace core {

// FormatKey实现
FormatKey::FormatKey() 
    : font_size(11.0), bold(false), italic(false), underline(false), strikethrough(false),
      font_color(0x000000), horizontal_align(0), vertical_align(0), text_wrap(false),
      text_rotation(0), border_style(0), border_color(0x000000), pattern(0),
      bg_color(0xFFFFFF), fg_color(0x000000), locked(true), hidden(false) {
    font_name = "Calibri";
    number_format = "General";
}

FormatKey::FormatKey(const Format& format) : FormatKey() {
    // 从Format对象提取属性
    font_name = format.getFontName();
    font_size = format.getFontSize();
    bold = format.isBold();
    italic = format.isItalic();
    underline = (format.getUnderline() != UnderlineType::None);
    strikethrough = format.isStrikeout();
    font_color = format.getFontColor().getRGB();
    
    // 对齐属性
    horizontal_align = static_cast<int>(format.getHorizontalAlign());
    vertical_align = static_cast<int>(format.getVerticalAlign());
    text_wrap = format.isTextWrap();
    text_rotation = format.getRotation();
    
    // 边框属性（简化为左边框）
    border_style = static_cast<int>(format.getLeftBorder());
    border_color = format.getLeftBorderColor().getRGB();
    
    // 填充属性
    pattern = static_cast<int>(format.getPattern());
    bg_color = format.getBackgroundColor().getRGB();
    fg_color = format.getForegroundColor().getRGB();
    
    // 数字格式
    number_format = format.getNumberFormat();
    
    // 保护属性
    locked = format.isLocked();
    hidden = format.isHidden();
}

bool FormatKey::operator==(const FormatKey& other) const {
    return font_name == other.font_name &&
           font_size == other.font_size &&
           bold == other.bold &&
           italic == other.italic &&
           underline == other.underline &&
           strikethrough == other.strikethrough &&
           font_color == other.font_color &&
           horizontal_align == other.horizontal_align &&
           vertical_align == other.vertical_align &&
           text_wrap == other.text_wrap &&
           text_rotation == other.text_rotation &&
           border_style == other.border_style &&
           border_color == other.border_color &&
           pattern == other.pattern &&
           bg_color == other.bg_color &&
           fg_color == other.fg_color &&
           number_format == other.number_format &&
           locked == other.locked &&
           hidden == other.hidden;
}

}} // namespace fastexcel::core

// std::hash特化实现
namespace std {
size_t hash<fastexcel::core::FormatKey>::operator()(const fastexcel::core::FormatKey& key) const {
    size_t h1 = hash<string>{}(key.font_name);
    size_t h2 = hash<double>{}(key.font_size);
    size_t h3 = hash<bool>{}(key.bold);
    size_t h4 = hash<bool>{}(key.italic);
    size_t h5 = hash<uint32_t>{}(key.font_color);
    size_t h6 = hash<int>{}(key.horizontal_align);
    size_t h7 = hash<uint32_t>{}(key.bg_color);
    size_t h8 = hash<string>{}(key.number_format);
    
    // 组合哈希值
    return h1 ^ (h2 << 1) ^ (h3 << 2) ^ (h4 << 3) ^ 
           (h5 << 4) ^ (h6 << 5) ^ (h7 << 6) ^ (h8 << 7);
}
}

namespace fastexcel {
namespace core {

FormatPool::FormatPool() : next_index_(0), total_requests_(0), cache_hits_(0) {
    // 创建默认格式
    default_format_ = std::make_unique<Format>();
    
    // 预留空间
    formats_.reserve(100);
    format_cache_.reserve(100);
    format_to_index_.reserve(100);
    
    // 添加默认格式到池中
    Format* default_ptr = default_format_.get();
    
    // 创建默认格式的副本添加到formats_向量中
    auto default_copy = std::make_unique<Format>(*default_format_);
    formats_.push_back(std::move(default_copy));
    
    // 添加到映射中
    format_to_index_[default_ptr] = 0;
    FormatKey default_key(*default_format_);
    format_cache_[default_key] = default_ptr;
    
    next_index_ = 1;
}

Format* FormatPool::getOrCreateFormat(const FormatKey& key) {
    total_requests_++;
    
    // 检查缓存
    auto it = format_cache_.find(key);
    if (it != format_cache_.end()) {
        cache_hits_++;
        return it->second;
    }
    
    // 创建新格式
    Format* format = createFormatFromKey(key);
    format_cache_[key] = format;
    
    return format;
}

Format* FormatPool::getOrCreateFormat(const Format& format) {
    FormatKey key(format);
    return getOrCreateFormat(key);
}

Format* FormatPool::addFormat(std::unique_ptr<Format> format) {
    Format* format_ptr = format.get();
    
    // 检查是否已存在相同格式
    FormatKey key(*format);
    auto cache_it = format_cache_.find(key);
    if (cache_it != format_cache_.end()) {
        return cache_it->second;  // 返回已存在的格式
    }
    
    // 添加新格式
    size_t index = next_index_++;
    formats_.push_back(std::move(format));
    format_cache_[key] = format_ptr;
    format_to_index_[format_ptr] = index;
    
    return format_ptr;
}

void FormatPool::importStyles(const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
    // 批量导入样式，用于格式复制功能
    // 新策略：同时使用去重机制和原始样式保存
    size_t formats_before = formats_.size();
    size_t actually_added = 0;
    
    // 保存原始样式用于XML生成
    setRawStylesForCopy(styles);
    
    // 传统的去重导入（用于运行时格式管理）
    for (const auto& [index, format] : styles) {
        if (format) {
            auto format_copy = std::make_unique<Format>(*format);
            addFormat(std::move(format_copy));
            if (formats_.size() > formats_before + actually_added) {
                actually_added++;
            }
        }
    }
    
    LOG_DEBUG("ImportStyles统计: 输入{}个样式, 导入前{}个格式, 导入后{}个格式, 实际新增{}个格式, 原始样式保存{}个", 
              styles.size(), formats_before, formats_.size(), actually_added, raw_styles_for_copy_.size());
}

void FormatPool::setRawStylesForCopy(const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
    raw_styles_for_copy_ = styles;
    LOG_DEBUG("FormatPool保存了{}个原始样式用于XML生成", raw_styles_for_copy_.size());
}

size_t FormatPool::getFormatIndex(Format* format) const {
    if (format == default_format_.get()) {
        return 0;  // 默认格式索引为0
    }
    
    auto it = format_to_index_.find(format);
    if (it != format_to_index_.end()) {
        return it->second;
    }
    
    FASTEXCEL_THROW_PARAM("Format not found in pool");
}

Format* FormatPool::getFormatByIndex(size_t index) const {
    if (index == 0) {
        return default_format_.get();
    }
    
    // 在formats_中查找对应索引的格式
    for (const auto& [format_ptr, format_index] : format_to_index_) {
        if (format_index == index) {
            return format_ptr;
        }
    }
    
    FASTEXCEL_THROW_PARAM("Invalid format index: " + std::to_string(index));
}

double FormatPool::getCacheHitRate() const {
    if (total_requests_ == 0) {
        return 0.0;
    }
    return static_cast<double>(cache_hits_) / total_requests_;
}

void FormatPool::clear() {
    formats_.clear();
    format_cache_.clear();
    format_to_index_.clear();
    next_index_ = 1;  // 保留默认格式的索引0
    total_requests_ = 0;
    cache_hits_ = 0;
    
    // 重新设置默认格式
    format_to_index_[default_format_.get()] = 0;
}


void FormatPool::generateStylesXMLInternal(xml::XMLStreamWriter& writer) const {
    // 🔧 优化：提取通用的样式XML生成逻辑，避免代码重复
    writer.startElement("styleSheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    // 添加调试日志来检查格式数量
    LOG_DEBUG("GenerateStylesXML: formats_.size()={}, getFormatCount()={}, hasRawStylesForCopy()={}, raw_styles_count={}", 
              formats_.size(), getFormatCount(), hasRawStylesForCopy(), 
              hasRawStylesForCopy() ? raw_styles_for_copy_.size() : 0);
    
    // 优先使用原始样式数据（格式复制场景）
    if (hasRawStylesForCopy() && !raw_styles_for_copy_.empty()) {
        LOG_DEBUG("使用Excel标准去重+索引架构生成XML，包含{}个样式", raw_styles_for_copy_.size());
        
        // 将样式按索引排序以便生成
        std::vector<std::pair<int, std::shared_ptr<core::Format>>> sorted_styles(
            raw_styles_for_copy_.begin(), raw_styles_for_copy_.end());
        std::sort(sorted_styles.begin(), sorted_styles.end(), 
                  [](const auto& a, const auto& b) { return a.first < b.first; });
        
        // === 第一阶段：收集和去重所有样式元素 ===
        
        // 1. 收集自定义数字格式（去重）
        std::vector<std::pair<int, std::string>> unique_numfmts; // (id, formatCode)
        std::set<std::string> seen_numfmt_codes;
        int next_numfmt_id = 176; // Excel自定义格式从176开始
        
        for (const auto& [index, format] : sorted_styles) {
            if (format) {
                std::string formatCode = format->getNumberFormat();
                // 只处理非标准格式
                if (!formatCode.empty() && formatCode != "General" && 
                    seen_numfmt_codes.find(formatCode) == seen_numfmt_codes.end()) {
                    unique_numfmts.push_back({next_numfmt_id++, formatCode});
                    seen_numfmt_codes.insert(formatCode);
                }
            }
        }
        
        // 2. 收集字体（去重）
        std::vector<std::string> unique_fonts;
        std::map<std::string, int> font_to_index; // 字体XML -> 索引
        unique_fonts.push_back("<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"); // 默认字体
        font_to_index["<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>"] = 0;
        
        for (const auto& [index, format] : sorted_styles) {
            if (format) {
                std::string fontXML = format->generateFontXML();
                if (!fontXML.empty() && font_to_index.find(fontXML) == font_to_index.end()) {
                    font_to_index[fontXML] = unique_fonts.size();
                    unique_fonts.push_back(fontXML);
                }
            }
        }
        
        // 3. 收集填充（去重）
        std::vector<std::string> unique_fills;
        std::map<std::string, int> fill_to_index;
        // 添加Excel标准的前两个填充
        unique_fills.push_back("<fill><patternFill patternType=\"none\"/></fill>");
        unique_fills.push_back("<fill><patternFill patternType=\"gray125\"/></fill>");
        fill_to_index["<fill><patternFill patternType=\"none\"/></fill>"] = 0;
        fill_to_index["<fill><patternFill patternType=\"gray125\"/></fill>"] = 1;
        
        for (const auto& [index, format] : sorted_styles) {
            if (format) {
                std::string fillXML = format->generateFillXML();
                if (!fillXML.empty() && fill_to_index.find(fillXML) == fill_to_index.end()) {
                    fill_to_index[fillXML] = unique_fills.size();
                    unique_fills.push_back(fillXML);
                }
            }
        }
        
        // 4. 收集边框（去重）
        std::vector<std::string> unique_borders;
        std::map<std::string, int> border_to_index;
        unique_borders.push_back("<border><left/><right/><top/><bottom/><diagonal/></border>"); // 默认边框
        border_to_index["<border><left/><right/><top/><bottom/><diagonal/></border>"] = 0;
        
        for (const auto& [index, format] : sorted_styles) {
            if (format) {
                std::string borderXML = format->generateBorderXML();
                if (!borderXML.empty() && border_to_index.find(borderXML) == border_to_index.end()) {
                    border_to_index[borderXML] = unique_borders.size();
                    unique_borders.push_back(borderXML);
                }
            }
        }
        
        LOG_DEBUG("去重统计: 自定义数字格式={}个, 字体={}个, 填充={}个, 边框={}个", 
                  unique_numfmts.size(), unique_fonts.size(), unique_fills.size(), unique_borders.size());
        
        // === 第二阶段：生成XML ===
        
        // 1. 生成数字格式
        if (!unique_numfmts.empty()) {
            writer.startElement("numFmts");
            writer.writeAttribute("count", std::to_string(unique_numfmts.size()).c_str());
            for (const auto& [id, formatCode] : unique_numfmts) {
                writer.startElement("numFmt");
                writer.writeAttribute("numFmtId", std::to_string(id).c_str());
                writer.writeAttribute("formatCode", formatCode.c_str());
                writer.endElement(); // numFmt
            }
            writer.endElement(); // numFmts
        }
        
        // 2. 生成字体（使用去重后的字体）
        writer.startElement("fonts");
        writer.writeAttribute("count", std::to_string(unique_fonts.size()).c_str());
        // 先关闭fonts开始标签，然后写入字体内容
        writer.writeText("");  // 这将强制关闭开始标签
        
        for (const auto& fontXML : unique_fonts) {
            if (!fontXML.empty()) {
                // 确保XML片段格式正确
                writer.writeRaw(fontXML);
            }
        }
        writer.endElement(); // fonts
        
        // 3. 生成填充（使用去重后的填充）
        writer.startElement("fills");
        writer.writeAttribute("count", std::to_string(unique_fills.size()).c_str());
        // 先关闭fills开始标签，然后写入填充内容
        writer.writeText("");  // 这将强制关闭开始标签
        
        for (const auto& fillXML : unique_fills) {
            if (!fillXML.empty()) {
                writer.writeRaw(fillXML);
            }
        }
        writer.endElement(); // fills
        
        // 4. 生成边框（使用去重后的边框）
        writer.startElement("borders");
        writer.writeAttribute("count", std::to_string(unique_borders.size()).c_str());
        // 先关闭borders开始标签，然后写入边框内容
        writer.writeText("");  // 这将强制关闭开始标签
        
        for (const auto& borderXML : unique_borders) {
            if (!borderXML.empty()) {
                writer.writeRaw(borderXML);
            }
        }
        writer.endElement(); // borders
        
        // 单元格样式格式
        writer.startElement("cellStyleXfs");
        writer.writeAttribute("count", "1");
        
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.endElement(); // xf
        
        writer.endElement(); // cellStyleXfs
        
        // 单元格格式
        writer.startElement("cellXfs");
        writer.writeAttribute("count", std::to_string(sorted_styles.size() + 1).c_str());
        
        // 默认格式
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.writeAttribute("xfId", "0");
        writer.endElement(); // xf
        
        // 所有原始样式的cellXf（使用正确的去重索引）
        for (size_t i = 0; i < sorted_styles.size(); ++i) {
            const auto& [index, format] = sorted_styles[i];
            if (format) {
                // 根据去重映射查找正确的索引
                std::string fontXML = format->generateFontXML();
                std::string fillXML = format->generateFillXML();
                std::string borderXML = format->generateBorderXML();
                
                int fontIndex = 0; // 默认字体
                int fillIndex = 0; // 默认填充
                int borderIndex = 0; // 默认边框
                
                // 查找字体索引
                if (!fontXML.empty() && font_to_index.find(fontXML) != font_to_index.end()) {
                    fontIndex = font_to_index[fontXML];
                }
                
                // 查找填充索引
                if (!fillXML.empty() && fill_to_index.find(fillXML) != fill_to_index.end()) {
                    fillIndex = fill_to_index[fillXML];
                }
                
                // 查找边框索引
                if (!borderXML.empty() && border_to_index.find(borderXML) != border_to_index.end()) {
                    borderIndex = border_to_index[borderXML];
                }
                
                // 设置正确的索引
                format->setFontIndex(fontIndex);
                format->setFillIndex(fillIndex);
                format->setBorderIndex(borderIndex);
                
                std::string xfXML = format->generateXML();
                writer.writeRaw(xfXML);
            }
        }
        writer.endElement(); // cellXfs
        
        // 单元格样式
        writer.startElement("cellStyles");
        writer.writeAttribute("count", "1");
        
        writer.startElement("cellStyle");
        writer.writeAttribute("name", "Normal");
        writer.writeAttribute("xfId", "0");
        writer.writeAttribute("builtinId", "0");
        writer.endElement(); // cellStyle
        
        writer.endElement(); // cellStyles
        
        // 添加dxfs元素（差异格式，即使为空也需要）
        writer.startElement("dxfs");
        writer.writeAttribute("count", "0");
        writer.endElement(); // dxfs
        
        // 添加tableStyles元素
        writer.startElement("tableStyles");
        writer.writeAttribute("count", "0");
        writer.writeAttribute("defaultTableStyle", "TableStyleMedium2");
        writer.writeAttribute("defaultPivotStyle", "PivotStyleLight16");
        writer.endElement(); // tableStyles
    }
    // 如果没有导入的样式，使用简化的默认样式
    else if (formats_.size() <= 1) {  // 只有默认格式或完全没有格式
        LOG_DEBUG("使用简化的默认样式生成");
        // 默认字体
        writer.startElement("fonts");
        writer.writeAttribute("count", "1");
        writer.startElement("font");
        writer.startElement("sz");
        writer.writeAttribute("val", "11");
        writer.endElement(); // sz
        writer.startElement("color");
        writer.writeAttribute("theme", "1");
        writer.endElement(); // color
        writer.startElement("name");
        writer.writeAttribute("val", "Calibri");
        writer.endElement(); // name
        writer.startElement("family");
        writer.writeAttribute("val", "2");
        writer.endElement(); // family
        writer.startElement("scheme");
        writer.writeAttribute("val", "minor");
        writer.endElement(); // scheme
        writer.endElement(); // font
        writer.endElement(); // fonts
        
        // 默认填充
        writer.startElement("fills");
        writer.writeAttribute("count", "2");
        writer.startElement("fill");
        writer.startElement("patternFill");
        writer.writeAttribute("patternType", "none");
        writer.endElement(); // patternFill
        writer.endElement(); // fill
        writer.startElement("fill");
        writer.startElement("patternFill");
        writer.writeAttribute("patternType", "gray125");
        writer.endElement(); // patternFill
        writer.endElement(); // fill
        writer.endElement(); // fills
        
        // 默认边框
        writer.startElement("borders");
        writer.writeAttribute("count", "1");
        writer.startElement("border");
        writer.startElement("left");
        writer.endElement(); // left
        writer.startElement("right");
        writer.endElement(); // right
        writer.startElement("top");
        writer.endElement(); // top
        writer.startElement("bottom");
        writer.endElement(); // bottom
        writer.startElement("diagonal");
        writer.endElement(); // diagonal
        writer.endElement(); // border
        writer.endElement(); // borders
        
        // 单元格样式
        writer.startElement("cellStyleXfs");
        writer.writeAttribute("count", "1");
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.endElement(); // xf
        writer.endElement(); // cellStyleXfs
        
        writer.startElement("cellXfs");
        writer.writeAttribute("count", "1");
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.writeAttribute("xfId", "0");
        writer.endElement(); // xf
        writer.endElement(); // cellXfs
        
        writer.startElement("cellStyles");
        writer.writeAttribute("count", "1");
        writer.startElement("cellStyle");
        writer.writeAttribute("name", "Normal");
        writer.writeAttribute("xfId", "0");
        writer.writeAttribute("builtinId", "0");
        writer.endElement(); // cellStyle
        writer.endElement(); // cellStyles
        
        writer.startElement("dxfs");
        writer.writeAttribute("count", "0");
        writer.endElement(); // dxfs
        
        writer.startElement("tableStyles");
        writer.writeAttribute("count", "0");
        writer.writeAttribute("defaultTableStyle", "TableStyleMedium2");
        writer.writeAttribute("defaultPivotStyle", "PivotStyleLight16");
        writer.endElement(); // tableStyles
    } else {
        // 使用formats_向量的标准生成逻辑（保持向后兼容）
        LOG_DEBUG("使用formats_向量生成，包含{}个格式", formats_.size());
        // ... 此处可以保留原有的else分支代码用于向后兼容
        // 为简化示例，暂时使用简化版本
        writer.startElement("fonts");
        writer.writeAttribute("count", "1");
        writer.startElement("font");
        writer.startElement("sz");
        writer.writeAttribute("val", "11");
        writer.endElement(); // sz
        writer.startElement("name");
        writer.writeAttribute("val", "Calibri");
        writer.endElement(); // name
        writer.startElement("family");
        writer.writeAttribute("val", "2");
        writer.endElement(); // family
        writer.startElement("scheme");
        writer.writeAttribute("val", "minor");
        writer.endElement(); // scheme
        writer.endElement(); // font
        writer.endElement(); // fonts
        
        writer.startElement("fills");
        writer.writeAttribute("count", "2");
        writer.startElement("fill");
        writer.startElement("patternFill");
        writer.writeAttribute("patternType", "none");
        writer.endElement();
        writer.endElement();
        writer.startElement("fill");
        writer.startElement("patternFill");
        writer.writeAttribute("patternType", "gray125");
        writer.endElement();
        writer.endElement();
        writer.endElement();
        
        writer.startElement("borders");
        writer.writeAttribute("count", "1");
        writer.startElement("border");
        writer.startElement("left");
        writer.endElement();
        writer.startElement("right");
        writer.endElement();
        writer.startElement("top");
        writer.endElement();
        writer.startElement("bottom");
        writer.endElement();
        writer.startElement("diagonal");
        writer.endElement();
        writer.endElement();
        writer.endElement();
        
        writer.startElement("cellStyleXfs");
        writer.writeAttribute("count", "1");
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.endElement();
        writer.endElement();
        
        writer.startElement("cellXfs");
        writer.writeAttribute("count", "1");
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", "0");
        writer.writeAttribute("fontId", "0");
        writer.writeAttribute("fillId", "0");
        writer.writeAttribute("borderId", "0");
        writer.writeAttribute("xfId", "0");
        writer.endElement();
        writer.endElement();
        
        writer.startElement("cellStyles");
        writer.writeAttribute("count", "1");
        writer.startElement("cellStyle");
        writer.writeAttribute("name", "Normal");
        writer.writeAttribute("xfId", "0");
        writer.writeAttribute("builtinId", "0");
        writer.endElement();
        writer.endElement();
        
        writer.startElement("dxfs");
        writer.writeAttribute("count", "0");
        writer.endElement();
        
        writer.startElement("tableStyles");
        writer.writeAttribute("count", "0");
        writer.writeAttribute("defaultTableStyle", "TableStyleMedium2");
        writer.writeAttribute("defaultPivotStyle", "PivotStyleLight16");
        writer.endElement();
    }
    
    writer.endElement(); // styleSheet
}

void FormatPool::generateStylesXML(const std::function<void(const char*, size_t)>& callback) const {
    // 🔧 修复：使用通用的样式生成逻辑，支持导入的样式
    xml::XMLStreamWriter writer(callback);
    writer.startDocument();
    generateStylesXMLInternal(writer);
    writer.endDocument();
}

void FormatPool::generateStylesXMLToFile(const std::string& filename) const {
    // 🔧 优化：使用通用的样式生成逻辑
    xml::XMLStreamWriter writer(filename);
    writer.startDocument();
    generateStylesXMLInternal(writer);
    writer.endDocument();
    
    // 数字格式部分
    std::vector<std::string> custom_formats;
    for (const auto& format : formats_) {
        std::string numFmt = format->generateNumberFormatXML();
        if (!numFmt.empty()) {
            custom_formats.push_back(numFmt);
        }
    }
    
    if (!custom_formats.empty()) {
        writer.startElement("numFmts");
        writer.writeAttribute("count", std::to_string(custom_formats.size()).c_str());
        for (const auto& fmt : custom_formats) {
            writer.writeRaw(fmt);
        }
        writer.endElement(); // numFmts
    }
    
    // 字体部分
    writer.startElement("fonts");
    writer.writeAttribute("count", std::to_string(getFormatCount() + 1).c_str());
    
    // 默认字体 - 完整的Calibri 11定义
    writer.startElement("font");
    writer.startElement("sz");
    writer.writeAttribute("val", "11");
    writer.endElement(); // sz
    writer.startElement("name");
    writer.writeAttribute("val", "Calibri");
    writer.endElement(); // name
    writer.startElement("family");
    writer.writeAttribute("val", "2");
    writer.endElement(); // family
    writer.startElement("scheme");
    writer.writeAttribute("val", "minor");
    writer.endElement(); // scheme
    writer.endElement(); // font
    
    // 其他字体
    for (const auto& format : formats_) {
        if (format->hasFont()) {
            std::string fontXML = format->generateFontXML();
            if (!fontXML.empty()) {
                writer.writeRaw(fontXML);
            } else {
                // 如果生成失败，使用默认字体
                writer.startElement("font");
                writer.startElement("sz");
                writer.writeAttribute("val", "11");
                writer.endElement(); // sz
                writer.startElement("name");
                writer.writeAttribute("val", "Calibri");
                writer.endElement(); // name
                writer.startElement("family");
                writer.writeAttribute("val", "2");
                writer.endElement(); // family
                writer.startElement("scheme");
                writer.writeAttribute("val", "minor");
                writer.endElement(); // scheme
                writer.endElement(); // font
            }
        } else {
            // 没有字体设置，使用默认字体
            writer.startElement("font");
            writer.startElement("sz");
            writer.writeAttribute("val", "11");
            writer.endElement(); // sz
            writer.startElement("name");
            writer.writeAttribute("val", "Calibri");
            writer.endElement(); // name
            writer.startElement("family");
            writer.writeAttribute("val", "2");
            writer.endElement(); // family
            writer.startElement("scheme");
            writer.writeAttribute("val", "minor");
            writer.endElement(); // scheme
            writer.endElement(); // font
        }
    }
    
    writer.endElement(); // fonts
    
    // 填充部分
    writer.startElement("fills");
    writer.writeAttribute("count", std::to_string(getFormatCount() + 2).c_str());
    
    // 默认填充
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "none");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "gray125");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    // 其他填充
    for (const auto& format : formats_) {
        std::string fillXML = format->generateFillXML();
        if (!fillXML.empty()) {
            writer.writeRaw(fillXML);
        } else {
            writer.startElement("fill");
            writer.startElement("patternFill");
            writer.writeAttribute("patternType", "none");
            writer.endElement(); // patternFill
            writer.endElement(); // fill
        }
    }
    
    writer.endElement(); // fills
    
    // 边框部分
    writer.startElement("borders");
    writer.writeAttribute("count", std::to_string(getFormatCount() + 1).c_str());
    
    // 默认边框
    writer.startElement("border");
    writer.startElement("left");
    writer.endElement(); // left
    writer.startElement("right");
    writer.endElement(); // right
    writer.startElement("top");
    writer.endElement(); // top
    writer.startElement("bottom");
    writer.endElement(); // bottom
    writer.startElement("diagonal");
    writer.endElement(); // diagonal
    writer.endElement(); // border
    
    // 其他边框
    for (const auto& format : formats_) {
        std::string borderXML = format->generateBorderXML();
        if (!borderXML.empty()) {
            writer.writeRaw(borderXML);
        } else {
            writer.startElement("border");
            writer.startElement("left");
            writer.endElement(); // left
            writer.startElement("right");
            writer.endElement(); // right
            writer.startElement("top");
            writer.endElement(); // top
            writer.startElement("bottom");
            writer.endElement(); // bottom
            writer.startElement("diagonal");
            writer.endElement(); // diagonal
            writer.endElement(); // border
        }
    }
    
    writer.endElement(); // borders
    
    // 单元格样式格式
    writer.startElement("cellStyleXfs");
    writer.writeAttribute("count", "1");
    
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.endElement(); // xf
    
    writer.endElement(); // cellStyleXfs
    
    // 单元格格式
    writer.startElement("cellXfs");
    writer.writeAttribute("count", std::to_string(getFormatCount() + 1).c_str());
    
    // 默认格式
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.writeAttribute("xfId", "0");
    writer.endElement(); // xf
    
    // 其他格式
    for (size_t i = 0; i < formats_.size(); ++i) {
        const auto& format = formats_[i];
        // 关键修复：正确设置格式索引
        // 字体索引：默认字体(0) + 当前格式索引
        format->setFontIndex(format->hasFont() ? (i + 1) : 0);
        // 填充索引：默认填充(0,1) + 当前格式索引
        format->setFillIndex(format->hasFill() ? (i + 2) : 0);
        // 边框索引：默认边框(0) + 当前格式索引
        format->setBorderIndex(format->hasBorder() ? (i + 1) : 0);
        
        std::string xfXML = format->generateXML();
        writer.writeRaw(xfXML);
    }
    
    writer.endElement(); // cellXfs
    
    // 单元格样式
    writer.startElement("cellStyles");
    writer.writeAttribute("count", "1");
    
    writer.startElement("cellStyle");
    writer.writeAttribute("name", "Normal");
    writer.writeAttribute("xfId", "0");
    writer.writeAttribute("builtinId", "0");
    writer.endElement(); // cellStyle
    
    writer.endElement(); // cellStyles
    
    // 添加dxfs元素（差异格式，即使为空也需要）
    writer.startElement("dxfs");
    writer.writeAttribute("count", "0");
    writer.endElement(); // dxfs
    
    // 添加tableStyles元素
    writer.startElement("tableStyles");
    writer.writeAttribute("count", "0");
    writer.writeAttribute("defaultTableStyle", "TableStyleMedium2");
    writer.writeAttribute("defaultPivotStyle", "PivotStyleLight16");
    writer.endElement(); // tableStyles
    
    writer.endElement(); // styleSheet
    writer.endDocument();
}

size_t FormatPool::getMemoryUsage() const {
    size_t usage = sizeof(FormatPool);
    
    // 格式对象内存
    usage += formats_.capacity() * sizeof(std::unique_ptr<Format>);
    for (const auto& format : formats_) {
        usage += sizeof(Format);  // 简化计算
    }
    
    // 缓存内存
    usage += format_cache_.bucket_count() * sizeof(std::pair<FormatKey, Format*>);
    usage += format_to_index_.bucket_count() * sizeof(std::pair<Format*, size_t>);
    
    return usage;
}

FormatPool::DeduplicationStats FormatPool::getDeduplicationStats() const {
    DeduplicationStats stats;
    stats.total_requests = total_requests_;
    stats.unique_formats = format_cache_.size();
    
    if (stats.total_requests > 0) {
        stats.deduplication_ratio = 1.0 - (static_cast<double>(stats.unique_formats) / stats.total_requests);
    } else {
        stats.deduplication_ratio = 0.0;
    }
    
    return stats;
}

Format* FormatPool::createFormatFromKey(const FormatKey& key) {
    auto format = std::make_unique<Format>();
    
    // 根据key设置格式属性
    // 字体属性
    format->setFontName(key.font_name);
    format->setFontSize(key.font_size);
    format->setBold(key.bold);
    format->setItalic(key.italic);
    if (key.underline) {
        format->setUnderline(UnderlineType::Single);
    }
    format->setStrikeout(key.strikethrough);
    format->setFontColor(key.font_color);
    
    // 对齐属性
    format->setHorizontalAlign(static_cast<HorizontalAlign>(key.horizontal_align));
    format->setVerticalAlign(static_cast<VerticalAlign>(key.vertical_align));
    format->setTextWrap(key.text_wrap);
    format->setRotation(key.text_rotation);
    
    // 边框属性
    if (key.border_style != 0) {
        format->setBorder(static_cast<BorderStyle>(key.border_style));
        format->setBorderColor(key.border_color);
    }
    
    // 填充属性
    if (key.pattern != 0) {
        format->setPattern(static_cast<PatternType>(key.pattern));
        format->setBackgroundColor(key.bg_color);
        format->setForegroundColor(key.fg_color);
    }
    
    // 数字格式
    if (!key.number_format.empty() && key.number_format != "General") {
        format->setNumberFormat(key.number_format);
    }
    
    // 保护属性
    format->setLocked(key.locked);
    format->setHidden(key.hidden);
    
    Format* format_ptr = format.get();
    size_t index = next_index_++;
    
    formats_.push_back(std::move(format));
    format_to_index_[format_ptr] = index;
    
    return format_ptr;
}

void FormatPool::updateFormatIndex(Format* format, size_t index) {
    format_to_index_[format] = index;
}

}} // namespace fastexcel::core
