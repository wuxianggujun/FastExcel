//
// Created by wuxianggujun on 25-8-4.
//

#pragma once


#include "fastexcel/core/Format.hpp"
#include "fastexcel/core/Color.hpp"
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace reader {

/**
 * @brief 样式解析器
 * 
 * 负责解析Excel文件中的styles.xml文件，
 * 提取字体、填充、边框、对齐等样式信息
 */
class StylesParser {
public:
    StylesParser() = default;
    ~StylesParser() = default;
    
    /**
     * @brief 解析样式XML内容
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content);
    
    /**
     * @brief 根据XF索引获取格式对象
     * @param xf_index XF索引
     * @return 格式对象指针，如果索引无效返回nullptr
     */
    std::shared_ptr<core::Format> getFormat(int xf_index) const;
    
    /**
     * @brief 获取解析的格式数量
     * @return 格式数量
     */
    size_t getFormatCount() const { return cell_xfs_.size(); }

private:
    // 字体信息结构
    struct FontInfo {
        std::string name = "Calibri";
        double size = 11.0;
        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool strikeout = false;
        core::Color color;
    };
    
    // 填充信息结构
    struct FillInfo {
        std::string pattern_type = "none";
        core::Color fg_color;
        core::Color bg_color;
    };
    
    // 边框边信息结构
    struct BorderSide {
        std::string style;
        core::Color color;
    };
    
    // 边框信息结构
    struct BorderInfo {
        BorderSide left;
        BorderSide right;
        BorderSide top;
        BorderSide bottom;
        BorderSide diagonal;
    };
    
    // 单元格XF信息结构
    struct CellXf {
        int num_fmt_id = -1;
        int font_id = -1;
        int fill_id = -1;
        int border_id = -1;
        std::string horizontal_alignment;
        std::string vertical_alignment;
        bool wrap_text = false;
        int indent = 0;
        int text_rotation = 0;
    };
    
    // 解析后的数据
    std::vector<FontInfo> fonts_;
    std::vector<FillInfo> fills_;
    std::vector<BorderInfo> borders_;
    std::vector<CellXf> cell_xfs_;
    std::unordered_map<int, std::string> number_formats_;
    
    // 解析方法
    void parseNumberFormats(const std::string& xml_content);
    void parseFonts(const std::string& xml_content);
    void parseFills(const std::string& xml_content);
    void parseBorders(const std::string& xml_content);
    void parseCellXfs(const std::string& xml_content);
    
    // 辅助解析方法
    BorderSide parseBorderSide(const std::string& border_xml, const std::string& side);
    core::Color parseColor(const std::string& color_xml);
    
    // XML属性提取方法
    int extractIntAttribute(const std::string& xml, const std::string& attr_name);
    double extractDoubleAttribute(const std::string& xml, const std::string& attr_name);
    std::string extractStringAttribute(const std::string& xml, const std::string& attr_name);
    
    // 枚举转换方法
    core::HorizontalAlign getAlignment(const std::string& alignment) const;
    core::VerticalAlign getVerticalAlignment(const std::string& alignment) const;
    core::BorderStyle getBorderStyle(const std::string& style) const;
    
    // 内置数字格式
    std::string getBuiltinNumberFormat(int format_id) const;
};

} // namespace reader
} // namespace fastexcel

