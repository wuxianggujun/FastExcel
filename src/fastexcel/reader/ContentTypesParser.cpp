#include "ContentTypesParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace reader {

bool ContentTypesParser::parse(const std::string& xml_content) {
    if (xml_content.empty()) {
        LOG_DEBUG("Empty content types XML content");
        return true;
    }

    try {
        // 清理之前的数据
        clear();
        
        // 使用XMLStreamReader的DOM解析，遵循项目模式
        xml::XMLStreamReader reader;
        auto dom = reader.parseToDOM(xml_content);
        if (!dom) {
            LOG_ERROR("Failed to parse content types XML to DOM");
            return false;
        }
        
        // 查找Types根元素
        xml::XMLStreamReader::SimpleElement* typesEl = nullptr;
        if (dom->name.find("Types") != std::string::npos) {
            typesEl = dom.get();
        } else {
            // 在子元素中查找
            for (const auto& child_up : dom->children) {
                const auto* child = child_up.get();
                if (!child) continue;
                if (child->name.find("Types") != std::string::npos) {
                    typesEl = const_cast<xml::XMLStreamReader::SimpleElement*>(child);
                    break;
                }
            }
        }
        
        if (!typesEl) {
            LOG_ERROR("No Types element found in content types XML");
            return false;
        }
        
        // 解析Default和Override元素
        for (const auto& child_up : typesEl->children) {
            const auto* child = child_up.get();
            if (!child) continue;
            
            if (child->name.find("Default") != std::string::npos) {
                // 解析Default元素
                DefaultType defaultType;
                defaultType.extension = child->getAttribute("Extension");
                defaultType.content_type = child->getAttribute("ContentType");
                
                if (!defaultType.extension.empty() && !defaultType.content_type.empty()) {
                    defaults_.push_back(defaultType);
                    LOG_DEBUG("Parsed default type: .{} -> {}", defaultType.extension, defaultType.content_type);
                } else {
                    LOG_WARN("Skipping incomplete default type: extension='{}', contentType='{}'", 
                             defaultType.extension, defaultType.content_type);
                }
            } else if (child->name.find("Override") != std::string::npos) {
                // 解析Override元素
                OverrideType overrideType;
                overrideType.part_name = child->getAttribute("PartName");
                overrideType.content_type = child->getAttribute("ContentType");
                
                if (!overrideType.part_name.empty() && !overrideType.content_type.empty()) {
                    overrides_.push_back(overrideType);
                    LOG_DEBUG("Parsed override type: {} -> {}", overrideType.part_name, overrideType.content_type);
                } else {
                    LOG_WARN("Skipping incomplete override type: partName='{}', contentType='{}'", 
                             overrideType.part_name, overrideType.content_type);
                }
            }
        }
        
        // 重建索引以提高查找性能
        rebuildIndex();
        
        LOG_DEBUG("Successfully parsed {} defaults and {} overrides", defaults_.size(), overrides_.size());
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception parsing content types XML: {}", e.what());
        return false;
    }
}

std::string ContentTypesParser::findDefaultType(const std::string& extension) const {
    auto it = default_index_.find(extension);
    return (it != default_index_.end()) ? it->second : std::string();
}

std::string ContentTypesParser::findOverrideType(const std::string& part_name) const {
    auto it = override_index_.find(part_name);
    return (it != override_index_.end()) ? it->second : std::string();
}

std::string ContentTypesParser::getContentType(const std::string& part_name) const {
    // 优先检查覆盖类型
    std::string override_type = findOverrideType(part_name);
    if (!override_type.empty()) {
        return override_type;
    }
    
    // 检查默认类型（根据扩展名）
    size_t last_dot = part_name.find_last_of('.');
    if (last_dot != std::string::npos) {
        std::string extension = part_name.substr(last_dot + 1);
        std::string default_type = findDefaultType(extension);
        if (!default_type.empty()) {
            return default_type;
        }
    }
    
    // 默认返回通用类型
    return "application/octet-stream";
}

void ContentTypesParser::clear() {
    defaults_.clear();
    overrides_.clear();
    default_index_.clear();
    override_index_.clear();
}

void ContentTypesParser::rebuildIndex() {
    // 重建默认类型索引
    default_index_.clear();
    for (const auto& defaultType : defaults_) {
        default_index_[defaultType.extension] = defaultType.content_type;
    }
    
    // 重建覆盖类型索引
    override_index_.clear();
    for (const auto& overrideType : overrides_) {
        override_index_[overrideType.part_name] = overrideType.content_type;
    }
}

}} // namespace fastexcel::reader