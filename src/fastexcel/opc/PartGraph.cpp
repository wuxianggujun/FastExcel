#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/opc/PartGraph.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"  // For ZipError enum
#include "fastexcel/reader/RelationshipsParser.hpp"  // 使用专门的关系解析器，遵循单一职责原则
#include "fastexcel/reader/ContentTypesParser.hpp"   // 使用专门的内容类型解析器，提高复用性
#include "fastexcel/utils/Logger.hpp"
#include <fmt/format.h>                              // 使用fmt库优化字符串格式化性能
#include <algorithm>
#include <sstream>
#include <string_view>                               // 使用string_view减少字符串复制

namespace fastexcel {
namespace opc {

// ========== PartGraph 实现 ==========

PartGraph::PartGraph() = default;
PartGraph::~PartGraph() = default;

bool PartGraph::buildFromZipReader(archive::ZipReader* reader) {
    if (!reader || !reader->isOpen()) {
        OPC_ERROR("Invalid or closed ZipReader");
        return false;
    }
    
    // Get all files in the ZIP
    auto files = reader->listFiles();
    OPC_DEBUG("Building part graph from {} files", files.size());
    
    // Add all files as parts
    for (const auto& file : files) {
        // Skip directories
        if (!file.empty() && file.back() == '/') {
            continue;
        }
        
        // Determine content type (simplified)
        std::string content_type = "application/xml";
        if (file.find(".rels") != std::string::npos) {
            content_type = "application/vnd.openxmlformats-package.relationships+xml";
        }
        
        addPart(file, content_type);
        
        // Parse relationship files
        if (file.find(".rels") != std::string::npos) {
            std::string rels_content;
            if (reader->extractFile(file, rels_content) == archive::ZipError::Ok) {
                // Extract base path from rels file path
                std::string base_path;
                if (file == "_rels/.rels") {
                    base_path = "";
                } else {
                    // Remove _rels/ prefix and .rels suffix to get the base
                    size_t pos = file.find("_rels/");
                    if (pos != std::string::npos) {
                        base_path = file.substr(0, pos);
                    }
                }
                parseRels(rels_content, base_path);
            }
        }
    }
    
    OPC_INFO("Part graph built with {} parts", parts_.size());
    return true;
}

void PartGraph::addPart(const std::string& path, const std::string& content_type) {
    Part& part = parts_[path];
    part.path = path;
    part.content_type = content_type;
}

void PartGraph::addRelationship(const std::string& from_part, const Relationship& rel) {
    if (parts_.find(from_part) != parts_.end()) {
        parts_[from_part].relationships.push_back(rel);
        
        // Update dependencies
        std::string target_path = normalizePath(from_part, rel.target);
        parts_[from_part].dependencies.insert(target_path);
        
        // Update references
        if (parts_.find(target_path) != parts_.end()) {
            parts_[target_path].references.insert(from_part);
        }
    }
}

void PartGraph::removePart(const std::string& path) {
    // Remove the part
    parts_.erase(path);
    
    // Clean up references in other parts
    for (auto& [other_path, other_part] : parts_) {
        other_part.dependencies.erase(path);
        other_part.references.erase(path);
        
        // Remove relationships pointing to this part
        other_part.relationships.erase(
            std::remove_if(other_part.relationships.begin(), 
                          other_part.relationships.end(),
                          [&path, this, &other_path](const Relationship& rel) {
                              return normalizePath(other_path, rel.target) == path;
                          }),
            other_part.relationships.end()
        );
    }
}

const PartGraph::Part* PartGraph::getPart(const std::string& path) const {
    auto it = parts_.find(path);
    return (it != parts_.end()) ? &it->second : nullptr;
}

std::vector<std::string> PartGraph::getAllParts() const {
    std::vector<std::string> result;
    result.reserve(parts_.size());
    for (const auto& [path, part] : parts_) {
        result.push_back(path);
    }
    return result;
}

std::string PartGraph::getRelsPath(const std::string& part_path) const {
    if (part_path.empty() || part_path == "/") {
        return "_rels/.rels";
    }
    
    size_t last_slash = part_path.find_last_of('/');
    if (last_slash == std::string::npos) {
        // 使用fmt::format提高性能，避免多次字符串拼接
        return fmt::format("_rels/{}.rels", part_path);
    }
    
    // 使用string_view避免不必要的字符串复制
    std::string_view dir_view(part_path.data(), last_slash);
    std::string_view name_view(part_path.data() + last_slash + 1, part_path.size() - last_slash - 1);
    
    // 使用fmt::format一次性完成拼接，性能更优
    return fmt::format("{}/_rels/{}.rels", dir_view, name_view);
}

bool PartGraph::hasRelationships(const std::string& part_path) const {
    auto it = parts_.find(part_path);
    return (it != parts_.end()) && !it->second.relationships.empty();
}

std::vector<std::string> PartGraph::getSheetRelatedParts(const std::string& sheet_path) const {
    std::vector<std::string> result;
    
    auto part = getPart(sheet_path);
    if (!part) {
        return result;
    }
    
    // Add direct dependencies
    for (const auto& dep : part->dependencies) {
        result.push_back(dep);
    }
    
    // Add common sheet-related parts
    // This is a simplified implementation
    if (sheet_path.find("xl/worksheets/") != std::string::npos) {
        // Potentially related: drawings, comments, etc.
        // For now, just return direct dependencies
    }
    
    return result;
}

std::unordered_set<std::string> PartGraph::getDirtyRels(const std::unordered_set<std::string>& dirty_parts) const {
    std::unordered_set<std::string> dirty_rels;
    // 预分配容量，避免多次rehash
    dirty_rels.reserve(dirty_parts.size() * 2);
    
    for (const auto& part : dirty_parts) {
        // If a part is dirty, its relationship file might need updating
        std::string rels_path = getRelsPath(part);
        if (parts_.find(rels_path) != parts_.end()) {
            dirty_rels.insert(rels_path);
        }
        
        // If a part is removed or added, parent rels need updating
        size_t last_slash = part.find_last_of('/');
        if (last_slash != std::string::npos) {
            // 使用string_view避免substr的字符串复制
            std::string_view parent_view(part.data(), last_slash);
            std::string parent_rels = getRelsPath(std::string(parent_view));
            if (parts_.find(parent_rels) != parts_.end()) {
                dirty_rels.insert(std::move(parent_rels));
            }
        }
    }
    
    return dirty_rels;
}

bool PartGraph::parseRels(const std::string& rels_content, const std::string& base_path) {
    if (rels_content.empty()) {
        OPC_DEBUG("Empty rels content for base_path: {}", base_path);
        return true;
    }

    // 使用专门的RelationshipsParser，遵循单一职责原则和DRY原则
    reader::RelationshipsParser parser;
    if (!parser.parse(rels_content)) {
        OPC_ERROR("Failed to parse relationships for base_path: {}", base_path);
        return false;
    }
    
    // 将解析到的关系添加到PartGraph
    const auto& relationships = parser.getRelationships();
    for (const auto& parsed_rel : relationships) {
        // 转换为PartGraph的Relationship结构
        Relationship rel;
        rel.id = parsed_rel.id;
        rel.type = parsed_rel.type;
        rel.target = parsed_rel.target;
        rel.target_mode = parsed_rel.target_mode;
        
        addRelationship(base_path, rel);
    }
    
    OPC_DEBUG("Successfully parsed {} relationships for base_path: {}", relationships.size(), base_path);
    return true;
}

std::string PartGraph::normalizePath(const std::string& base, const std::string& relative) const {
    if (relative.empty()) {
        return base;
    }
    
    // Absolute path
    if (relative[0] == '/') {
        return relative.substr(1);  // Remove leading slash
    }
    
    // Relative path
    size_t last_slash = base.find_last_of('/');
    if (last_slash == std::string::npos) {
        return relative;
    }
    
    std::string dir = base.substr(0, last_slash);
    return dir + "/" + relative;
}

// ========== ContentTypes 实现 ==========

ContentTypes::ContentTypes() {
    // Add common defaults
    addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
    addDefault("xml", "application/xml");
}

ContentTypes::~ContentTypes() = default;

bool ContentTypes::parse(const std::string& xml) {
    if (xml.empty()) {
        OPC_DEBUG("Empty ContentTypes XML content");
        return true;
    }

    // 清理之前的数据
    defaults_.clear();
    overrides_.clear();
    
    // 使用专门的ContentTypesParser，遵循单一职责原则和DRY原则
    reader::ContentTypesParser parser;
    if (!parser.parse(xml)) {
        OPC_ERROR("Failed to parse content types XML");
        return false;
    }
    
    // 将解析结果转换为ContentTypes内部结构
    const auto& defaults = parser.getDefaults();
    for (const auto& defaultType : defaults) {
        addDefault(defaultType.extension, defaultType.content_type);
    }
    
    const auto& overrides = parser.getOverrides();
    for (const auto& overrideType : overrides) {
        addOverride(overrideType.part_name, overrideType.content_type);
    }
    
    OPC_DEBUG("ContentTypes parsed via specialized parser: {} defaults, {} overrides", defaults_.size(), overrides_.size());
    return true;
}

std::string ContentTypes::serialize() const {
    // 使用fmt::memory_buffer提高性能，避免ostringstream的开销
    fmt::memory_buffer buffer;
    buffer.reserve(1024);  // 预分配缓冲区，减少重分配
    
    fmt::format_to(std::back_inserter(buffer), 
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
                   "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n");
    
    // Write defaults - 使用fmt::format_to批量写入，性能更优
    for (const auto& [ext, type] : defaults_) {
        fmt::format_to(std::back_inserter(buffer), 
                       "  <Default Extension=\"{}\" ContentType=\"{}\"/>\r\n", ext, type);
    }
    
    // Write overrides - 使用fmt::format_to批量写入，性能更优
    for (const auto& [path, type] : overrides_) {
        fmt::format_to(std::back_inserter(buffer), 
                       "  <Override PartName=\"{}\" ContentType=\"{}\"/>\r\n", path, type);
    }
    
    fmt::format_to(std::back_inserter(buffer), "</Types>\r\n");
    
    return fmt::to_string(buffer);
}

void ContentTypes::addDefault(const std::string& extension, const std::string& content_type) {
    defaults_[extension] = content_type;
}

void ContentTypes::addOverride(const std::string& part_name, const std::string& content_type) {
    overrides_[part_name] = content_type;
}

void ContentTypes::removeOverride(const std::string& part_name) {
    overrides_.erase(part_name);
}

std::string ContentTypes::getContentType(const std::string& part_name) const {
    // Check overrides first
    auto override_it = overrides_.find(part_name);
    if (override_it != overrides_.end()) {
        return override_it->second;
    }
    
    // Check defaults by extension - 使用string_view避免substr复制
    size_t last_dot = part_name.find_last_of('.');
    if (last_dot != std::string::npos && last_dot < part_name.size() - 1) {
        std::string_view ext_view(part_name.data() + last_dot + 1, part_name.size() - last_dot - 1);
        
        // 临时创建string进行查找（C++20之前unordered_map不支持heterogeneous lookup）
        auto default_it = defaults_.find(std::string(ext_view));
        if (default_it != defaults_.end()) {
            return default_it->second;
        }
    }
    
    return "application/octet-stream";  // Default fallback
}

void ContentTypes::updateSheets(const std::vector<std::string>& sheet_names) {
    // Remove old sheet overrides - 使用remove_if提高性能
    for (auto it = overrides_.begin(); it != overrides_.end(); ) {
        if (it->first.find("/xl/worksheets/sheet") != std::string::npos) {
            it = overrides_.erase(it);
        } else {
            ++it;
        }
    }
    
    // Add new sheet overrides - 预分配空间，使用fmt::format优化字符串生成
    overrides_.reserve(overrides_.size() + sheet_names.size());
    
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        // 使用fmt::format避免多次字符串拼接
        std::string path = fmt::format("/xl/worksheets/sheet{}.xml", i + 1);
        addOverride(std::move(path), "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    }
}

}} // namespace fastexcel::opc
