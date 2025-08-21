#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/core/FormatRepository.hpp"
#include "XMLStreamWriter.hpp"
#include <string>
#include <functional>

namespace fastexcel {
namespace xml {

/**
 * @brief XLSX样式序列化器
 * 
 * 负责将FormatRepository中的格式信息序列化为XLSX格式的XML。
 * 这是基础设施层的一部分，将领域模型转换为特定的输出格式。
 */
class StyleSerializer {
public:
    /**
     * @brief 序列化样式信息到XML流
     * @param repository 格式仓储（只读）
     * @param writer XML写入器
     */
    static void serialize(const core::FormatRepository& repository, 
                         xml::XMLStreamWriter& writer);
    
    /**
     * @brief 序列化样式信息到回调函数（流式输出）
     * @param repository 格式仓储（只读）
     * @param callback 数据输出回调函数
     */
    static void serialize(const core::FormatRepository& repository,
                         const std::function<void(const char*, size_t)>& callback);
    
    /**
     * @brief 序列化样式信息到文件
     * @param repository 格式仓储（只读）
     * @param filename 输出文件路径
     */
    static void serializeToFile(const core::FormatRepository& repository,
                               const std::string& filename);

private:
    /**
     * @brief 写入样式XML的根元素
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeStyleSheet(const core::FormatRepository& repository,
                               xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入数字格式部分
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeNumberFormats(const core::FormatRepository& repository,
                                  xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入字体部分
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeFonts(const core::FormatRepository& repository,
                          xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入填充部分
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeFills(const core::FormatRepository& repository,
                          xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入边框部分
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeBorders(const core::FormatRepository& repository,
                            xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入单元格样式交叉引用部分
     * @param repository 格式仓储
     * @param writer XML写入器
     */
    static void writeCellXfs(const core::FormatRepository& repository,
                            xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入单个字体元素
     * @param format 格式描述符
     * @param writer XML写入器
     */
    static void writeFont(const core::FormatDescriptor& format,
                         xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入单个填充元素
     * @param format 格式描述符
     * @param writer XML写入器
     */
    static void writeFill(const core::FormatDescriptor& format,
                         xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入单个边框元素
     * @param format 格式描述符
     * @param writer XML写入器
     */
    static void writeBorder(const core::FormatDescriptor& format,
                           xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入单个单元格格式元素
     * @param format 格式描述符
     * @param font_id 字体ID
     * @param fill_id 填充ID
     * @param border_id 边框ID
     * @param num_fmt_id 数字格式ID
     * @param writer XML写入器
     */
    static void writeCellXf(const core::FormatDescriptor& format,
                           int font_id, int fill_id, int border_id, int num_fmt_id,
                           xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入对齐信息
     * @param format 格式描述符
     * @param writer XML写入器
     */
    static void writeAlignment(const core::FormatDescriptor& format,
                              xml::XMLStreamWriter& writer);
    
    /**
     * @brief 写入保护信息
     * @param format 格式描述符
     * @param writer XML写入器
     */
    static void writeProtection(const core::FormatDescriptor& format,
                               xml::XMLStreamWriter& writer);
    
    // ========== 辅助方法 ==========
    
    /**
     * @brief 获取边框样式的XML属性值
     * @param style 边框样式
     * @return XML属性值
     */
    static std::string borderStyleToXml(core::BorderStyle style);
    
    /**
     * @brief 获取填充模式的XML属性值
     * @param pattern 填充模式
     * @return XML属性值
     */
    static std::string patternTypeToXml(core::PatternType pattern);
    
    /**
     * @brief 获取下划线类型的XML属性值
     * @param underline 下划线类型
     * @return XML属性值
     */
    static std::string underlineTypeToXml(core::UnderlineType underline);
    
    /**
     * @brief 获取水平对齐的XML属性值
     * @param align 对齐方式
     * @return XML属性值
     */
    static std::string horizontalAlignToXml(core::HorizontalAlign align);
    
    /**
     * @brief 获取垂直对齐的XML属性值
     * @param align 对齐方式
     * @return XML属性值
     */
    static std::string verticalAlignToXml(core::VerticalAlign align);
    
    /**
     * @brief 获取颜色的XML属性值
     * @param color 颜色对象
     * @return XML属性值
     */
    static std::string colorToXml(const core::Color& color);
    
    /**
     * @brief 写入颜色属性到XML元素
     * @param color 颜色对象
     * @param writer XML写入器
     */
    static void writeColorAttributes(const core::Color& color, xml::XMLStreamWriter& writer);
    
    /**
     * @brief 检查格式是否需要对齐信息
     * @param format 格式描述符
     * @return 是否需要对齐信息
     */
    static bool needsAlignment(const core::FormatDescriptor& format);
    
    /**
     * @brief 检查格式是否需要保护信息
     * @param format 格式描述符
     * @return 是否需要保护信息
     */
    static bool needsProtection(const core::FormatDescriptor& format);
    
    /**
     * @brief 创建去重映射
     * 创建字体、填充、边框等子组件的去重映射，返回格式ID到子组件ID的映射
     * @param repository 格式仓储
     * @param font_mapping 输出：格式ID到字体ID的映射
     * @param fill_mapping 输出：格式ID到填充ID的映射
     * @param border_mapping 输出：格式ID到边框ID的映射
     * @param numfmt_mapping 输出：格式ID到数字格式ID的映射
     */
    static void createComponentMappings(
        const core::FormatRepository& repository,
        std::vector<int>& font_mapping,
        std::vector<int>& fill_mapping,
        std::vector<int>& border_mapping,
        std::vector<int>& numfmt_mapping
    );
    
    /**
     * @brief 收集唯一的字体
     * @param repository 格式仓储
     * @param unique_fonts 输出：唯一字体列表
     * @param format_to_font_id 输出：格式ID到字体ID的映射
     */
    static void collectUniqueFonts(
        const core::FormatRepository& repository,
        std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_fonts,
        std::vector<int>& format_to_font_id
    );
    
    /**
     * @brief 收集唯一的填充
     * @param repository 格式仓储
     * @param unique_fills 输出：唯一填充列表
     * @param format_to_fill_id 输出：格式ID到填充ID的映射
     */
    static void collectUniqueFills(
        const core::FormatRepository& repository,
        std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_fills,
        std::vector<int>& format_to_fill_id
    );
    
    /**
     * @brief 收集唯一的边框
     * @param repository 格式仓储
     * @param unique_borders 输出：唯一边框列表
     * @param format_to_border_id 输出：格式ID到边框ID的映射
     */
    static void collectUniqueBorders(
        const core::FormatRepository& repository,
        std::vector<std::shared_ptr<const core::FormatDescriptor>>& unique_borders,
        std::vector<int>& format_to_border_id
    );
    
    /**
     * @brief 收集唯一的数字格式
     * @param repository 格式仓储
     * @param unique_numfmts 输出：唯一数字格式列表
     * @param format_to_numfmt_id 输出：格式ID到数字格式ID的映射
     */
    static void collectUniqueNumberFormats(
        const core::FormatRepository& repository,
        std::vector<std::string>& unique_numfmts,
        std::vector<int>& format_to_numfmt_id
    );
    
    /**
     * @brief 创建填充哈希键（用于O(1)查找优化）
     * @param format 格式描述符
     * @return 填充的哈希键
     */
    static std::string createFillHashKey(const core::FormatDescriptor& format);
    
    /**
     * @brief 创建字体哈希键（用于O(1)查找优化）
     * @param format 格式描述符
     * @return 字体的哈希键
     */
    static std::string createFontHashKey(const core::FormatDescriptor& format);
    
    /**
     * @brief 创建边框哈希键（用于O(1)查找优化）
     * @param format 格式描述符
     * @return 边框的哈希键
     */
    static std::string createBorderHashKey(const core::FormatDescriptor& format);
};

}} // namespace fastexcel::xml
