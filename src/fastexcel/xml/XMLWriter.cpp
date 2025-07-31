#include "fastexcel/xml/XMLWriter.hpp"
#include <sstream>
#include <iomanip>

namespace fastexcel {
namespace xml {

void XMLWriter::startDocument() {
    buffer_.str("");
    buffer_.clear();
    buffer_ << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
}

void XMLWriter::endDocument() {
    while (!element_stack_.empty()) {
        endElement();
    }
}

void XMLWriter::startElement(const std::string& name) {
    if (in_element_) {
        buffer_ << ">";
        in_element_ = false;
    }
    
    buffer_ << "<" << name;
    element_stack_.push(name);
    in_element_ = true;
}

void XMLWriter::endElement() {
    if (element_stack_.empty()) {
        return;
    }
    
    std::string name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        buffer_ << "/>";
        in_element_ = false;
    } else {
        buffer_ << "</" << name << ">";
    }
}

void XMLWriter::writeEmptyElement(const std::string& name) {
    if (in_element_) {
        buffer_ << ">";
        in_element_ = false;
    }
    
    buffer_ << "<" << name << "/>";
}

void XMLWriter::writeAttribute(const std::string& name, const std::string& value) {
    if (!in_element_) {
        return;
    }
    
    std::string escaped_value = value;
    escapeAttribute(escaped_value);
    
    buffer_ << " " << name << "=\"" << escaped_value << "\"";
}

void XMLWriter::writeAttribute(const std::string& name, int value) {
    writeAttribute(name, std::to_string(value));
}

void XMLWriter::writeAttribute(const std::string& name, double value) {
    std::ostringstream oss;
    oss << std::fixed << value;
    writeAttribute(name, oss.str());
}

void XMLWriter::writeText(const std::string& text) {
    if (in_element_) {
        buffer_ << ">";
        in_element_ = false;
    }
    
    std::string escaped_text = text;
    escapeText(escaped_text);
    
    buffer_ << escaped_text;
}

std::string XMLWriter::toString() const {
    return buffer_.str();
}

void XMLWriter::clear() {
    buffer_.str("");
    buffer_.clear();
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    in_element_ = false;
}

void XMLWriter::escapeText(std::string& text) const {
    size_t pos = 0;
    while ((pos = text.find('&', pos)) != std::string::npos) {
        text.replace(pos, 1, "&");
        pos += 5;
    }
    
    pos = 0;
    while ((pos = text.find('<', pos)) != std::string::npos) {
        text.replace(pos, 1, "<");
        pos += 4;
    }
    
    pos = 0;
    while ((pos = text.find('>', pos)) != std::string::npos) {
        text.replace(pos, 1, ">");
        pos += 4;
    }
}

void XMLWriter::escapeAttribute(std::string& text) const {
    escapeText(text);
    
    size_t pos = 0;
    while ((pos = text.find('\"', pos)) != std::string::npos) {
        text.replace(pos, 1, "&quot;");
        pos += 6;
    }
    
    pos = 0;
    while ((pos = text.find('\'', pos)) != std::string::npos) {
        text.replace(pos, 1, "&apos;");
        pos += 6;
    }
}

}} // namespace fastexcel::xml