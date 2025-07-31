#pragma once

#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <string>
#include <vector>
#include <unordered_map>

namespace fastexcel {
namespace xml {

class SharedStrings {
public:
    SharedStrings() = default;
    ~SharedStrings() = default;
    
    // 添加共享字符串
    int addString(const std::string& str);
    
    // 获取字符串索引
    int getStringIndex(const std::string& str) const;
    
    // 获取字符串
    std::string getString(int index) const;
    
    // 生成XML内容
    std::string generate() const;
    
    // 清空字符串
    void clear();
    
    // 获取字符串数量
    size_t size() const { return strings_.size(); }
    
private:
    std::vector<std::string> strings_;
    std::unordered_map<std::string, int> string_map_;
    
    // 转义XML特殊字符
    std::string escapeString(const std::string& str) const;
};

}} // namespace fastexcel::xml
