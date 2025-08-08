#include "RelationshipsParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace reader {

bool RelationshipsParser::parse(const std::string& xml_content) {
    if (xml_content.empty()) {
        LOG_DEBUG("Empty relationships XML content");
        return true;
    }

    try {
        // 清理之前的数据
        clear();
        
        // 使用XMLStreamReader的DOM解析，遵循ThemeParser的模式
        xml::XMLStreamReader reader;
        auto dom = reader.parseToDOM(xml_content);
        if (!dom) {
            LOG_ERROR("Failed to parse relationships XML to DOM");
            return false;
        }
        
        // 查找Relationships根元素
        xml::XMLStreamReader::SimpleElement* relationshipsEl = nullptr;
        if (dom->name.find("Relationships") != std::string::npos) {
            relationshipsEl = dom.get();
        } else {
            // 在子元素中查找
            for (const auto& child_up : dom->children) {
                const auto* child = child_up.get();
                if (!child) continue;
                if (child->name.find("Relationships") != std::string::npos) {
                    relationshipsEl = const_cast<xml::XMLStreamReader::SimpleElement*>(child);
                    break;
                }
            }
        }
        
        if (!relationshipsEl) {
            LOG_ERROR("No Relationships element found in XML");
            return false;
        }
        
        // 解析每个Relationship元素
        for (const auto& child_up : relationshipsEl->children) {
            const auto* child = child_up.get();
            if (!child) continue;
            
            if (child->name.find("Relationship") != std::string::npos) {
                Relationship rel;
                
                // 提取属性
                rel.id = child->getAttribute("Id");
                rel.type = child->getAttribute("Type");
                rel.target = child->getAttribute("Target");
                rel.target_mode = child->getAttribute("TargetMode", "Internal");
                
                // 验证必需属性
                if (!rel.id.empty() && !rel.type.empty() && !rel.target.empty()) {
                    relationships_.push_back(rel);
                    LOG_DEBUG("Parsed relationship: {} -> {} ({})", rel.id, rel.target, rel.type);
                } else {
                    LOG_WARN("Skipping incomplete relationship: id='{}', type='{}', target='{}'", 
                             rel.id, rel.type, rel.target);
                }
            }
        }
        
        LOG_DEBUG("Successfully parsed {} relationships", relationships_.size());
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception parsing relationships XML: {}", e.what());
        return false;
    }
}

const RelationshipsParser::Relationship* RelationshipsParser::findById(const std::string& id) const {
    auto it = std::find_if(relationships_.begin(), relationships_.end(),
                          [&id](const Relationship& rel) {
                              return rel.id == id;
                          });
    return (it != relationships_.end()) ? &(*it) : nullptr;
}

std::vector<const RelationshipsParser::Relationship*> RelationshipsParser::findByType(const std::string& type) const {
    std::vector<const Relationship*> result;
    for (const auto& rel : relationships_) {
        if (rel.type == type) {
            result.push_back(&rel);
        }
    }
    return result;
}

}} // namespace fastexcel::reader