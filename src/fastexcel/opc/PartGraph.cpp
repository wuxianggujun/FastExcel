#include "fastexcel/opc/PartGraph.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"  // For ZipError enum
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <sstream>

namespace fastexcel {
namespace opc {

// ========== PartGraph 实现 ==========

PartGraph::PartGraph() = default;
PartGraph::~PartGraph() = default;

bool PartGraph::buildFromZipReader(archive::ZipReader* reader) {
    if (!reader || !reader->isOpen()) {
        LOG_ERROR("Invalid or closed ZipReader");
        return false;
    }
    
    // Get all files in the ZIP
    auto files = reader->listFiles();
    LOG_DEBUG("Building part graph from {} files", files.size());
    
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
    
    LOG_INFO("Part graph built with {} parts", parts_.size());
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
        return "_rels/" + part_path + ".rels";
    }
    
    std::string dir = part_path.substr(0, last_slash);
    std::string name = part_path.substr(last_slash + 1);
    return dir + "/_rels/" + name + ".rels";
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
    
    for (const auto& part : dirty_parts) {
        // If a part is dirty, its relationship file might need updating
        std::string rels_path = getRelsPath(part);
        if (parts_.find(rels_path) != parts_.end()) {
            dirty_rels.insert(rels_path);
        }
        
        // If a part is removed or added, parent rels need updating
        size_t last_slash = part.find_last_of('/');
        if (last_slash != std::string::npos) {
            std::string parent = part.substr(0, last_slash);
            std::string parent_rels = getRelsPath(parent);
            if (parts_.find(parent_rels) != parts_.end()) {
                dirty_rels.insert(parent_rels);
            }
        }
    }
    
    return dirty_rels;
}

bool PartGraph::parseRels(const std::string& rels_content, const std::string& base_path) {
    // This is a simplified XML parsing - in production, use proper XML parser
    // For now, just return true to indicate success
    (void)rels_content;
    (void)base_path;
    
    // TODO: Implement proper relationship XML parsing
    // Parse <Relationship> elements and call addRelationship()
    
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
    // 简化的XML解析 - 在生产环境中应使用适当的XML解析器
    defaults_.clear();
    overrides_.clear();
    
    // 解析 <Default> 元素
    size_t pos = 0;
    while ((pos = xml.find("<Default", pos)) != std::string::npos) {
        size_t end = xml.find("/>", pos);
        if (end == std::string::npos) break;
        
        std::string element = xml.substr(pos, end - pos + 2);
        
        // 提取 Extension
        size_t ext_start = element.find("Extension=\"");
        if (ext_start != std::string::npos) {
            ext_start += 11;  // strlen("Extension=\"")
            size_t ext_end = element.find("\"", ext_start);
            if (ext_end != std::string::npos) {
                std::string extension = element.substr(ext_start, ext_end - ext_start);
                
                // 提取 ContentType
                size_t type_start = element.find("ContentType=\"");
                if (type_start != std::string::npos) {
                    type_start += 13;  // strlen("ContentType=\"")
                    size_t type_end = element.find("\"", type_start);
                    if (type_end != std::string::npos) {
                        std::string content_type = element.substr(type_start, type_end - type_start);
                        addDefault(extension, content_type);
                    }
                }
            }
        }
        
        pos = end + 2;
    }
    
    // 解析 <Override> 元素
    pos = 0;
    while ((pos = xml.find("<Override", pos)) != std::string::npos) {
        size_t end = xml.find("/>", pos);
        if (end == std::string::npos) break;
        
        std::string element = xml.substr(pos, end - pos + 2);
        
        // 提取 PartName
        size_t name_start = element.find("PartName=\"");
        if (name_start != std::string::npos) {
            name_start += 10;  // strlen("PartName=\"")
            size_t name_end = element.find("\"", name_start);
            if (name_end != std::string::npos) {
                std::string part_name = element.substr(name_start, name_end - name_start);
                
                // 提取 ContentType
                size_t type_start = element.find("ContentType=\"");
                if (type_start != std::string::npos) {
                    type_start += 13;  // strlen("ContentType=\"")
                    size_t type_end = element.find("\"", type_start);
                    if (type_end != std::string::npos) {
                        std::string content_type = element.substr(type_start, type_end - type_start);
                        addOverride(part_name, content_type);
                    }
                }
            }
        }
        
        pos = end + 2;
    }
    
    LOG_DEBUG("ContentTypes parsed: {} defaults, {} overrides", defaults_.size(), overrides_.size());
    return true;
}

std::string ContentTypes::serialize() const {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
    xml << "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n";
    
    // Write defaults
    for (const auto& [ext, type] : defaults_) {
        xml << "  <Default Extension=\"" << ext << "\" ContentType=\"" << type << "\"/>\r\n";
    }
    
    // Write overrides
    for (const auto& [path, type] : overrides_) {
        xml << "  <Override PartName=\"" << path << "\" ContentType=\"" << type << "\"/>\r\n";
    }
    
    xml << "</Types>\r\n";
    return xml.str();
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
    
    // Check defaults by extension
    size_t last_dot = part_name.find_last_of('.');
    if (last_dot != std::string::npos) {
        std::string ext = part_name.substr(last_dot + 1);
        auto default_it = defaults_.find(ext);
        if (default_it != defaults_.end()) {
            return default_it->second;
        }
    }
    
    return "application/octet-stream";  // Default fallback
}

void ContentTypes::updateSheets(const std::vector<std::string>& sheet_names) {
    // Remove old sheet overrides
    for (auto it = overrides_.begin(); it != overrides_.end(); ) {
        if (it->first.find("/xl/worksheets/sheet") != std::string::npos) {
            it = overrides_.erase(it);
        } else {
            ++it;
        }
    }
    
    // Add new sheet overrides
    for (size_t i = 0; i < sheet_names.size(); ++i) {
        std::string path = "/xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
        addOverride(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    }
}

}} // namespace fastexcel::opc
