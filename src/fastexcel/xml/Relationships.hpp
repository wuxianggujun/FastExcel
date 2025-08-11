#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <string>
#include <vector>
#include <functional>

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
    
    // 生成XML内容到回调函数（流式写入）
    void generate(const std::function<void(const char*, size_t)>& callback) const;
    
    // 生成XML内容到文件（流式写入）
    void generateToFile(const std::string& filename) const;
    
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
