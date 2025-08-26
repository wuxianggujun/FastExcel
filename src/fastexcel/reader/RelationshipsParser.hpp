#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/span.hpp"
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>

using fastexcel::core::span;  // Import span into this namespace

namespace fastexcel {
namespace reader {

/**
 * @brief 高性能关系文件解析器 - 基于SAX流式解析
 * 
 * 专门解析Excel文件中的.rels关系文件，
 * 使用BaseSAXParser提供的高性能SAX解析能力，
 * 消除DOM解析开销和线性查找操作。
 * 
 * 性能优化：
 * - 零字符串查找：基于SAX事件驱动
 * - 流式解析：不加载完整DOM到内存
 * - O(1)查找：建立快速ID索引 
 * - 同步索引：解析时同步构建索引
 */
class RelationshipsParser : public BaseSAXParser {
public:
    /**
     * @brief 关系结构体
     */
    struct Relationship {
        std::string id;          // 如 "rId1"
        std::string type;        // 如 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        std::string target;      // 如 "worksheets/sheet1.xml"
        std::string target_mode; // 默认 "Internal"
        
        Relationship() : target_mode("Internal") {}
    };

    RelationshipsParser() = default;
    ~RelationshipsParser() = default;
    
    /**
     * @brief 解析关系XML内容
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content) {
        clear();
        return parseXML(xml_content);
    }
    
    /**
     * @brief 获取解析的关系列表
     * @return 关系列表
     */
    const std::vector<Relationship>& getRelationships() const { return relationships_; }
    
    /**
     * @brief 根据ID查找关系
     * @param id 关系ID
     * @return 关系指针，未找到返回nullptr
     */
    const Relationship* findById(const std::string& id) const;
    
    /**
     * @brief 根据类型查找关系
     * @param type 关系类型
     * @return 匹配的关系列表
     */
    std::vector<const Relationship*> findByType(const std::string& type) const;
    
    /**
     * @brief 获取关系总数
     * @return 关系总数
     */
    size_t getRelationshipCount() const { return relationships_.size(); }
    
    /**
     * @brief 清空解析结果
     */
    void clear() { 
        relationships_.clear();
        id_index_.clear();
    }

private:
    std::vector<Relationship> relationships_;
    
    // 高性能索引： ID -> 关系索引，提供O(1)查找
    std::unordered_map<std::string, size_t> id_index_;
    
    // 重写基类虚函数
    void onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
};

}} // namespace fastexcel::reader
