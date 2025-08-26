#include "ContentTypesParser.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace reader {

void ContentTypesParser::onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
    if (name == "Default") {
        // 解析Default元素
        auto extension = findAttribute(attributes, "Extension");
        auto content_type = findAttribute(attributes, "ContentType");
        
        if (extension && content_type && !extension->empty() && !content_type->empty()) {
            defaults_.emplace_back();
            auto& defaultType = defaults_.back();
            defaultType.extension = std::move(*extension);
            defaultType.content_type = std::move(*content_type);
            
            // 同步构建索引，避免后续重建
            default_index_[defaultType.extension] = defaultType.content_type;
            
            FASTEXCEL_LOG_DEBUG("Parsed default type: .{} -> {}", defaultType.extension, defaultType.content_type);
        } else {
            FASTEXCEL_LOG_WARN("Skipping incomplete default type: extension='{}', contentType='{}'", 
                     extension ? *extension : "", content_type ? *content_type : "");
        }
    }
    else if (name == "Override") {
        // 解析Override元素
        auto part_name = findAttribute(attributes, "PartName");
        auto content_type = findAttribute(attributes, "ContentType");
        
        if (part_name && content_type && !part_name->empty() && !content_type->empty()) {
            overrides_.emplace_back();
            auto& overrideType = overrides_.back();
            overrideType.part_name = std::move(*part_name);
            overrideType.content_type = std::move(*content_type);
            
            // 同步构建索引，避免后续重建
            override_index_[overrideType.part_name] = overrideType.content_type;
            
            FASTEXCEL_LOG_DEBUG("Parsed override type: {} -> {}", overrideType.part_name, overrideType.content_type);
        } else {
            FASTEXCEL_LOG_WARN("Skipping incomplete override type: partName='{}', contentType='{}'", 
                     part_name ? *part_name : "", content_type ? *content_type : "");
        }
    }
    // 忽略其他元素如 <Types> 根元素
}

void ContentTypesParser::onEndElement(const std::string& name, int depth) {
    // ContentTypes解析不需要处理元素结束事件
    // 所有属性都在开始元素中处理完毕
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
    auto override_it = override_index_.find(part_name);
    if (override_it != override_index_.end()) {
        return override_it->second;
    }
    
    // 检查默认类型（根据扩展名）
    size_t last_dot = part_name.find_last_of('.');
    if (last_dot != std::string::npos) {
        std::string extension = part_name.substr(last_dot + 1);
        auto default_it = default_index_.find(extension);
        if (default_it != default_index_.end()) {
            return default_it->second;
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

} // namespace reader
} // namespace fastexcel