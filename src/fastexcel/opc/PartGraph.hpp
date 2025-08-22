#include "fastexcel/utils/Logger.hpp"
#pragma once

#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <memory>

namespace fastexcel {

// Forward declarations
namespace archive {
    class ZipReader;
}

namespace opc {

/**
 * @brief OPC部件关系图 - 管理部件之间的引用关系
 */
class PartGraph {
public:
    // 关系类型
    struct Relationship {
        std::string id;          // 如 "rId1"
        std::string type;        // 如 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        std::string target;      // 如 "worksheets/sheet1.xml"
        std::string target_mode; // 默认 "Internal"
    };
    
    // 部件信息
    struct Part {
        std::string path;                              // 部件路径
        std::string content_type;                      // 内容类型
        std::vector<Relationship> relationships;       // 该部件的关系
        std::unordered_set<std::string> references;    // 引用该部件的其他部件
        std::unordered_set<std::string> dependencies;  // 该部件依赖的其他部件
    };
    
    PartGraph();
    ~PartGraph();
    
    /**
     * 从archive::ZipReader构建关系图
     */
    bool buildFromZipReader(archive::ZipReader* reader);
    
    /**
     * 添加部件
     */
    void addPart(const std::string& path, const std::string& content_type);
    
    /**
     * 添加关系
     */
    void addRelationship(const std::string& from_part, const Relationship& rel);
    
    /**
     * 删除部件（级联删除相关关系）
     */
    void removePart(const std::string& path);
    
    /**
     * 获取部件信息
     */
    const Part* getPart(const std::string& path) const;
    
    /**
     * 获取所有部件路径
     */
    std::vector<std::string> getAllParts() const;
    
    /**
     * 获取部件的关系文件路径
     */
    std::string getRelsPath(const std::string& part_path) const;
    
    /**
     * 检查部件是否有关系
     */
    bool hasRelationships(const std::string& part_path) const;
    
    /**
     * 获取工作表相关的部件
     */
    std::vector<std::string> getSheetRelatedParts(const std::string& sheet_path) const;
    
    /**
     * 标记需要更新的关系文件
     */
    std::unordered_set<std::string> getDirtyRels(const std::unordered_set<std::string>& dirty_parts) const;
    
private:
    std::unordered_map<std::string, Part> parts_;
    
    /**
     * 解析关系文件
     */
    bool parseRels(const std::string& rels_content, const std::string& base_path);
    
    /**
     * 规范化路径
     */
    std::string normalizePath(const std::string& base, const std::string& relative) const;
};

/**
 * @brief Content Types管理器
 */
class ContentTypes {
public:
    ContentTypes();
    ~ContentTypes();
    
    /**
     * 从[Content_Types].xml解析
     */
    bool parse(const std::string& xml);
    
    /**
     * 序列化为XML
     */
    std::string serialize() const;
    
    /**
     * 添加默认类型
     */
    void addDefault(const std::string& extension, const std::string& content_type);
    
    /**
     * 添加覆盖类型
     */
    void addOverride(const std::string& part_name, const std::string& content_type);
    
    /**
     * 删除部件的覆盖
     */
    void removeOverride(const std::string& part_name);
    
    /**
     * 获取部件的内容类型
     */
    std::string getContentType(const std::string& part_name) const;
    
    /**
     * 更新工作表列表
     */
    void updateSheets(const std::vector<std::string>& sheet_names);
    
private:
    std::unordered_map<std::string, std::string> defaults_;    // 扩展名 -> 类型
    std::unordered_map<std::string, std::string> overrides_;   // 部件路径 -> 类型
};

}} // namespace fastexcel::opc
