#pragma once

#include <string>
#include <sstream>
#include <stack>

namespace fastexcel {
namespace xml {

class XMLWriter {
private:
    std::ostringstream buffer_;
    std::stack<std::string> element_stack_;
    bool in_element_ = false;
    
public:
    XMLWriter() = default;
    ~XMLWriter() = default;
    
    // 文档操作
    void startDocument();
    void endDocument();
    
    // 元素操作
    void startElement(const std::string& name);
    void endElement();
    void writeEmptyElement(const std::string& name);
    
    // 属性和文本
    void writeAttribute(const std::string& name, const std::string& value);
    void writeAttribute(const std::string& name, int value);
    void writeAttribute(const std::string& name, double value);
    void writeText(const std::string& text);
    
    // 获取结果
    std::string toString() const;
    void clear();
    
private:
    void escapeText(std::string& text) const;
    void escapeAttribute(std::string& text) const;
};

}} // namespace fastexcel::xml