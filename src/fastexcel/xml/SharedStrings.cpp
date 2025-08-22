#include "fastexcel/xml/SharedStrings.hpp"
#include <fmt/format.h>

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

void SharedStrings::generate(const std::function<void(const char*, size_t)>& callback) const {
    XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", fmt::format("{}", strings_.size()).c_str());
    writer.writeAttribute("uniqueCount", fmt::format("{}", strings_.size()).c_str());
    
    for (const auto& str : strings_) {
        writer.startElement("si");
        writer.startElement("t");
        writer.writeText(str.c_str());
        writer.endElement(); // t
        writer.endElement(); // si
    }
    
    writer.endElement(); // sst
    writer.endDocument();
}

void SharedStrings::generateToFile(const std::string& filename) const {
    XMLStreamWriter writer(filename);
    writer.startDocument();
    writer.startElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", fmt::format("{}", strings_.size()).c_str());
    writer.writeAttribute("uniqueCount", fmt::format("{}", strings_.size()).c_str());
    
    for (const auto& str : strings_) {
        writer.startElement("si");
        writer.startElement("t");
        writer.writeText(str.c_str());
        writer.endElement(); // t
        writer.endElement(); // si
    }
    
    writer.endElement(); // sst
    writer.endDocument();
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
        result.replace(pos, 1, "&amp;");
        pos += 5;
    }
    
    pos = 0;
    while ((pos = result.find('<', pos)) != std::string::npos) {
        result.replace(pos, 1, "&lt;");
        pos += 4;
    }
    
    pos = 0;
    while ((pos = result.find('>', pos)) != std::string::npos) {
        result.replace(pos, 1, "&gt;");
        pos += 4;
    }
    
    return result;
}

}} // namespace fastexcel::xml