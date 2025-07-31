#pragma once

#include "XMLWriter.h"
#include <string>
#include <vector>

namespace fastexcel {
namespace xml {

struct Relationship {
    std::string id;
    std::string type;
    std::string target;
    std::string target_mode; // "Internal" or "External"
};

class Relationships {
public:
    Relationships() = default;
    ~Relationships() = default;
    
    // 添加关系
    void addRelationship(const std::string& id, const std::string& type, const std::string& target);
    void addRelationship(const std::string& id, const std::string& type, const std::string& target, const std::string& target_mode);
    
    // 生成XML内容
    std::string generate() const;
    
    // 清空关系
    void clear();
    
    // 获取关系数量
    size_t size() const { return relationships_.size(); }
    
private:
    std::vector<Relationship> relationships_;
    
    // 生成唯一ID
    std::string generateId() const;
};

}} // namespace fastexcel::xml