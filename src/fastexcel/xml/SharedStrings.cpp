#include "fastexcel/xml/SharedStrings.hpp"

namespace fastexcel {
namespace xml {

int SharedStrings::addString(const std::string& str) {
    auto it = string_map_.find(str);
    if (it != string_map_.end()) {
        return it->second;
    }
    
    int index = static_cast<int>(strings_.size());
    strings_.push_back(str);
    string_map_[str] = index;
    
    return index;
}

int SharedStrings::getStringIndex(const std::string& str) const {
    auto it = string_map_.find(str);
    if (it != string_map_.end()) {
        return it->second;
    }
    return -1;
}

std::string SharedStrings::getString(int index) const {
    if (index >= 0 && index < static_cast<int>(strings_.size())) {
        return strings_[index];
    }
    return "";
}

std::string SharedStrings::generate() const {
    XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", std::to_string(strings_.size()));
    writer.writeAttribute("uniqueCount", std::to_string(strings_.size()));
    
    for (const auto& str : strings_) {
        writer.startElement("si");
        writer.startElement("t");
        writer.writeText(str);
        writer.endElement(); // t
        writer.endElement(); // si
    }
    
    writer.endElement(); // sst
    writer.endDocument();
    
    return writer.toString();
}

void SharedStrings::clear() {
    strings_.clear();
    string_map_.clear();
}

std::string SharedStrings::escapeString(const std::string& str) const {
    std::string result = str;
    
    // 转义XML特殊字符
    size_t pos = 0;
    while ((pos = result.find('&', pos)) != std::string::npos) {
        result.replace(pos, 1, "&");
        pos += 5;
    }
    
    pos = 0;
    while ((pos = result.find('<', pos)) != std::string::npos) {
        result.replace(pos, 1, "<");
        pos += 4;
    }
    
    pos = 0;
    while ((pos = result.find('>', pos)) != std::string::npos) {
        result.replace(pos, 1, ">");
        pos += 4;
    }
    
    return result;
}

}} // namespace fastexcel::xml