#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/span.hpp"
#include <string>
#include <unordered_map>
#include <vector>

using fastexcel::core::span;  // Import span into this namespace

namespace fastexcel {
namespace reader {

/**
 * @brief 高性能内容类型解析器 - 基于SAX流式解析
 * 
 * 专门解析Excel文件中的[Content_Types].xml文件，
 * 使用BaseSAXParser提供的高性能SAX解析能力，
 * 消除DOM解析开销和字符串查找操作。
 * 
 * 性能优化：
 * - 零字符串查找：基于SAX事件驱动
 * - 流式解析：不加载完整DOM到内存  
 * - 直接索引构建：解析时同步构建hash索引
 * - 内存效率：最小化临时对象创建
 */
class ContentTypesParser : public BaseSAXParser {
public:
    /**
     * @brief 默认类型结构体
     */
    struct DefaultType {
        std::string extension;     // 文件扩展名
        std::string content_type;  // 内容类型
    };

    /**
     * @brief 覆盖类型结构体
     */
    struct OverrideType {
        std::string part_name;     // 部件路径
        std::string content_type;  // 内容类型
    };

    ContentTypesParser() = default;
    ~ContentTypesParser() = default;
    
    /**
     * @brief 解析内容类型XML
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content) {
        clear();
        return parseXML(xml_content);
    }
    
    /**
     * @brief 获取默认类型列表
     * @return 默认类型列表
     */
    const std::vector<DefaultType>& getDefaults() const { return defaults_; }
    
    /**
     * @brief 获取覆盖类型列表
     * @return 覆盖类型列表
     */
    const std::vector<OverrideType>& getOverrides() const { return overrides_; }
    
    /**
     * @brief 根据扩展名查找默认类型
     * @param extension 文件扩展名
     * @return 内容类型，未找到返回空字符串
     */
    std::string findDefaultType(const std::string& extension) const;
    
    /**
     * @brief 根据部件名查找覆盖类型
     * @param part_name 部件路径
     * @return 内容类型，未找到返回空字符串
     */
    std::string findOverrideType(const std::string& part_name) const;
    
    /**
     * @brief 获取部件的内容类型（优先检查覆盖，再检查默认）
     * @param part_name 部件路径
     * @return 内容类型
     */
    std::string getContentType(const std::string& part_name) const;
    
    /**
     * @brief 获取默认类型数量
     * @return 默认类型数量
     */
    size_t getDefaultCount() const { return defaults_.size(); }
    
    /**
     * @brief 获取覆盖类型数量
     * @return 覆盖类型数量
     */
    size_t getOverrideCount() const { return overrides_.size(); }
    
    /**
     * @brief 清空解析结果
     */
    void clear();

private:
    std::vector<DefaultType> defaults_;
    std::vector<OverrideType> overrides_;
    
    // 高性能查找索引（解析时同步构建）
    std::unordered_map<std::string, std::string> default_index_;
    std::unordered_map<std::string, std::string> override_index_;
    
    // 重写基类虚函数
    void onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
};

}} // namespace fastexcel::reader