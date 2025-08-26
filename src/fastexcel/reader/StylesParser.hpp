#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/Color.hpp"
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace reader {

// 数据结构定义 - 必须在类定义之前
struct FontInfo {
    std::string name = "Calibri";
    double size = 11.0;
    bool bold = false;
    bool italic = false;
    bool underline = false;
    bool strikeout = false;
    core::Color color;
};

struct FillInfo {
    std::string pattern_type = "none";
    core::Color fg_color;
    core::Color bg_color;
};

struct BorderSide {
    std::string style = "none";
    core::Color color;
};

struct BorderInfo {
    BorderSide left, right, top, bottom, diagonal;
};

struct CellXf {
    int num_fmt_id = 0;
    int font_id = 0;
    int fill_id = 0;
    int border_id = 0;
    bool apply_number_format = false;
    bool apply_font = false;
    bool apply_fill = false;
    bool apply_border = false;
    bool apply_alignment = false;
    bool apply_protection = false;
    
    // 对齐属性
    std::string horizontal_alignment = "general";
    std::string vertical_alignment = "bottom";
    int text_rotation = 0;
    bool wrap_text = false;
    int indent = 0;
    bool shrink_to_fit = false;
};

/**
 * @brief 极高性能混合架构样式解析器
 * 
 * 混合架构策略：
 * 1. SAX处理顶级区域切换 (numFmts, fonts, fills, borders, cellXfs)
 * 2. 指针扫描处理区域内部复杂嵌套 (避免深层回调风暴)
 * 3. 批量构建样式对象，减少内存分配
 * 
 * 性能优化：
 * - 零字符串查找：基于指针扫描
 * - 减少SAX回调：从数千次减少到数十次
 * - 内存效率：预分配和批量处理
 * - 保持完整功能：支持所有Excel样式特性
 */
class StylesParser : public BaseSAXParser {
public:
    StylesParser() = default;
    ~StylesParser() = default;
    
    /**
     * @brief 解析样式XML内容
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content) {
        clear();
        return parseXML(xml_content);
    }
    
    /**
     * @brief 根据XF索引获取格式描述符
     * @param xf_index XF索引
     * @return 格式描述符指针，如果索引无效返回nullptr
     */
    std::shared_ptr<core::FormatDescriptor> getFormat(int xf_index) const;
    
    /**
     * @brief 获取解析的格式数量
     * @return 格式数量
     */
    /**
     * @brief 获取默认字体信息（通常是cellXfs中第一个字体）
     * @return 默认字体信息，包含name和size
     */
    std::pair<std::string, double> getDefaultFontInfo() const {
        // 通常第一个cellXf对应Normal样式，使用其fontId
        if (!cell_xfs_.empty()) {
            int font_id = cell_xfs_[0].font_id;
            if (font_id >= 0 && font_id < static_cast<int>(fonts_.size())) {
                const auto& font = fonts_[font_id];
                return {font.name, font.size};
            }
        }
        
        // 如果有字体但cellXfs为空，使用第一个字体
        if (!fonts_.empty()) {
            const auto& font = fonts_[0];
            return {font.name, font.size};
        }
        
        // 默认值
        return {"Calibri", 11.0};
    }
    
    /**
     * @brief 获取字体列表
     * @return 字体信息列表
     */
    const std::vector<FontInfo>& getFonts() const { return fonts_; }
    
    size_t getFormatCount() const { return cell_xfs_.size(); }

private:
    // 简化的解析状态 - 只跟踪区域级别
    struct ParseState {
        enum class Region {
            None,
            NumFmts,    // 收集 <numFmts>...</numFmts>
            Fonts,      // 收集 <fonts>...</fonts>  
            Fills,      // 收集 <fills>...</fills>
            Borders,    // 收集 <borders>...</borders>
            CellXfs     // 收集 <cellXfs>...</cellXfs>
        };
        
        Region current_region = Region::None;
        std::string region_xml_buffer;  // 收集当前区域的完整XML
        bool collecting_region = false;
        int region_depth = 0;
        
        void startRegion(Region region) {
            current_region = region;
            collecting_region = true;
            region_xml_buffer.clear();
            region_depth = 0;
        }
        
        void endRegion() {
            current_region = Region::None;
            collecting_region = false;
            region_xml_buffer.clear();
            region_depth = 0;
        }
        
        void reset() {
            endRegion();
        }
    } state_;
    
    // 解析后的数据
    std::vector<FontInfo> fonts_;
    std::vector<FillInfo> fills_;
    std::vector<BorderInfo> borders_;
    std::vector<CellXf> cell_xfs_;
    std::unordered_map<int, std::string> number_formats_;
    
    // 混合架构实现
    void onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) override;
    void onEndElement(const std::string& name, int depth) override;
    void onText(const std::string& text, int depth) override;
    
    // 区域处理方法 - 使用指针扫描
    void processNumFmtsRegion(std::string_view region_xml);
    void processFontsRegion(std::string_view region_xml);
    void processFillsRegion(std::string_view region_xml);
    void processBordersRegion(std::string_view region_xml);
    void processCellXfsRegion(std::string_view region_xml);
    
    // 指针扫描工具方法
    static void parseColorAttribute(const char*& p, const char* end, core::Color& color);
    static std::string_view extractElementContent(const char* xml, const char* element_name);
    static bool findAttributeInElement(const char* element_start, const char* element_end, 
                                     const char* attr_name, std::string& out_value);
    
    // 清理数据
    void clear() {
        fonts_.clear();
        fills_.clear();
        borders_.clear();
        cell_xfs_.clear();
        number_formats_.clear();
        state_.reset();
    }
    
    // 枚举转换方法保持不变 - 确保完整功能
    core::HorizontalAlign getAlignment(const std::string& alignment) const;
    core::VerticalAlign getVerticalAlignment(const std::string& alignment) const;
    core::BorderStyle getBorderStyle(const std::string& style) const;
    core::PatternType getPatternType(const std::string& pattern) const;
    std::string getBuiltinNumberFormat(int format_id) const;
};

} // namespace reader
} // namespace fastexcel

