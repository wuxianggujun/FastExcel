#include "RelationshipsParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace reader {

void RelationshipsParser::onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
    if (name == "Relationship") {
        // 解析Relationship元素
        auto id = findAttribute(attributes, "Id");
        auto type = findAttribute(attributes, "Type");
        auto target = findAttribute(attributes, "Target");
        auto target_mode = findAttribute(attributes, "TargetMode");
        
        // 验证必需属性
        if (id && type && target && !id->empty() && !type->empty() && !target->empty()) {
            Relationship rel;
            rel.id = *id;
            rel.type = *type;
            rel.target = *target;
            rel.target_mode = target_mode ? *target_mode : "Internal";  // 默认值
            
            // 添加到集合并建立ID索引
            size_t index = relationships_.size();
            relationships_.push_back(std::move(rel));
            id_index_[*id] = index;
            
            FASTEXCEL_LOG_DEBUG("Parsed relationship: {} -> {} ({})", *id, *target, *type);
        } else {
            FASTEXCEL_LOG_WARN("Skipping incomplete relationship: id='{}', type='{}', target='{}'", 
                     id ? *id : "", type ? *type : "", target ? *target : "");
        }
    }
    // 忽略其他元素如 <Relationships> 根元素
}

void RelationshipsParser::onEndElement(const std::string& name, int depth) {
    // RelationshipsParser不需要处理元素结束事件
    // 所有属性都在开始元素中处理完毕
}

const RelationshipsParser::Relationship* RelationshipsParser::findById(const std::string& id) const {
    // 使用O(1)索引查找，替代线性查找
    auto it = id_index_.find(id);
    if (it != id_index_.end()) {
        size_t index = it->second;
        if (index < relationships_.size()) {
            return &relationships_[index];
        }
    }
    return nullptr;
}

std::vector<const RelationshipsParser::Relationship*> RelationshipsParser::findByType(const std::string& type) const {
    std::vector<const Relationship*> result;
    result.reserve(4); // 预分配空间，大多数情况下关系数量不多
    
    for (const auto& rel : relationships_) {
        if (rel.type == type) {
            result.push_back(&rel);
        }
    }
    return result;
}

} // namespace reader
} // namespace fastexcel