#pragma once

#include <string>
#include <memory>

namespace fastexcel {
namespace utils {

/**
 * @brief Excel样式文件解析器
 * 
 * 解析styles.xml文件，获取真实的默认字体信息，
 * 为ColumnWidthCalculator提供准确的MDW计算基础
 */
class StylesParser {
public:
    /**
     * @brief 默认字体信息
     */
    struct DefaultFontInfo {
        std::string name = "Calibri";           // 字体名称
        double size = 11.0;                     // 字体大小(pt)
        int mdw = 7;                           // 计算得出的MDW
        bool is_parsed = false;                // 是否成功解析
        
        DefaultFontInfo() = default;
        DefaultFontInfo(const std::string& font_name, double font_size) 
            : name(font_name), size(font_size), is_parsed(true) {
            mdw = calculateMDW(font_name, font_size);
        }
        
    private:
        static int calculateMDW(const std::string& font_name, double font_size);
    };
    
    /**
     * @brief 从styles.xml文件解析默认字体
     * @param styles_xml_path styles.xml文件路径
     * @return 默认字体信息
     */
    static DefaultFontInfo parseDefaultFont(const std::string& styles_xml_path);
    
    /**
     * @brief 从styles.xml内容解析默认字体
     * @param xml_content XML内容字符串
     * @return 默认字体信息
     */
    static DefaultFontInfo parseDefaultFontFromContent(const std::string& xml_content);
    
    /**
     * @brief 从工作簿解析默认字体（推荐方法）
     * @param workbook_dir 工作簿目录路径（包含xl子目录）
     * @return 默认字体信息
     */
    static DefaultFontInfo parseFromWorkbook(const std::string& workbook_dir);

private:
    /**
     * @brief 解析字体节点
     * @param xml_content XML内容
     * @param font_id 字体ID
     * @return 字体信息 (name, size)
     */
    static std::pair<std::string, double> parseFontById(const std::string& xml_content, int font_id);
    
    /**
     * @brief 查找Normal样式的字体ID
     * @param xml_content XML内容
     * @return 字体ID，-1表示未找到
     */
    static int findNormalStyleFontId(const std::string& xml_content);
    
    /**
     * @brief 简单的XML标签值提取
     * @param xml XML内容
     * @param tag_name 标签名
     * @return 标签值，未找到返回空字符串
     */
    static std::string extractTagValue(const std::string& xml, const std::string& tag_name);
    
    /**
     * @brief 提取XML标签属性值
     * @param xml XML内容
     * @param tag_name 标签名
     * @param attr_name 属性名
     * @return 属性值，未找到返回空字符串
     */
    static std::string extractTagAttribute(const std::string& xml, const std::string& tag_name, const std::string& attr_name);
};

}} // namespace fastexcel::utils