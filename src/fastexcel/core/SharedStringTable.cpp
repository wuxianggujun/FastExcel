#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/core/Exception.hpp"
#include <fmt/format.h>

namespace fastexcel {
namespace core {

SharedStringTable::SharedStringTable() : next_id_(0) {
    // 预留一些空间以减少重新分配
    string_to_id_.reserve(1000);
    id_to_string_.reserve(1000);
}

int32_t SharedStringTable::addString(const std::string& str) {
    // 检查字符串是否已存在
    auto it = string_to_id_.find(str);
    if (it != string_to_id_.end()) {
        return it->second;  // 返回已存在的ID
    }
    
    // 添加新字符串
    int32_t id = next_id_++;
    string_to_id_[str] = id;
    id_to_string_.push_back(str);
    
    return id;
}

const std::string& SharedStringTable::getString(int32_t id) const {
    if (id < 0 || static_cast<size_t>(id) >= id_to_string_.size()) {
        FASTEXCEL_THROW_PARAM(fmt::format("Invalid string ID: {}", id));
    }
    return id_to_string_[id];
}

bool SharedStringTable::hasString(const std::string& str) const {
    return string_to_id_.find(str) != string_to_id_.end();
}

int32_t SharedStringTable::getStringId(const std::string& str) const {
    auto it = string_to_id_.find(str);
    return (it != string_to_id_.end()) ? it->second : -1;
}

int32_t SharedStringTable::addStringWithId(const std::string& str, int32_t original_id) {
    // 检查字符串是否已存在
    auto it = string_to_id_.find(str);
    if (it != string_to_id_.end()) {
        return it->second;  // 返回已存在的ID
    }
    
    // 确保id_to_string_数组足够大
    if (static_cast<size_t>(original_id) >= id_to_string_.size()) {
        id_to_string_.resize(original_id + 1);
    }
    
    // 检查指定的ID位置是否已被占用
    if (!id_to_string_[original_id].empty()) {
        // 如果原始ID位置已被占用，使用新的ID
        int32_t new_id = next_id_++;
        string_to_id_[str] = new_id;
        id_to_string_.push_back(str);
        return new_id;
    }
    
    // 使用原始ID
    string_to_id_[str] = original_id;
    id_to_string_[original_id] = str;
    
    // 更新next_id_以确保不会重复使用已占用的ID
    if (original_id >= next_id_) {
        next_id_ = original_id + 1;
    }
    
    return original_id;
}

void SharedStringTable::clear() {
    string_to_id_.clear();
    id_to_string_.clear();
    next_id_ = 0;
}

void SharedStringTable::generateXML(const std::function<void(const std::string&)>& callback) const {
    xml::XMLStreamWriter writer(callback);
    writer.startDocument();
    writer.startElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", std::to_string(id_to_string_.size()).c_str());
    writer.writeAttribute("uniqueCount", std::to_string(string_to_id_.size()).c_str());
    
    // 即使没有字符串，也要生成完整的XML结构
    for (const auto& str : id_to_string_) {
        writer.startElement("si");
        writer.startElement("t");
        
        // 处理特殊字符转义
        std::string escaped_str = str;
        // 简单的XML转义处理
        size_t pos = 0;
        while ((pos = escaped_str.find('&', pos)) != std::string::npos) {
            escaped_str.replace(pos, 1, "&amp;");
            pos += 5;
        }
        pos = 0;
        while ((pos = escaped_str.find('<', pos)) != std::string::npos) {
            escaped_str.replace(pos, 1, "&lt;");
            pos += 4;
        }
        pos = 0;
        while ((pos = escaped_str.find('>', pos)) != std::string::npos) {
            escaped_str.replace(pos, 1, "&gt;");
            pos += 4;
        }
        
        writer.writeText(escaped_str.c_str());
        writer.endElement(); // t
        
        // 添加phoneticPr标签以提高Excel兼容性
        writer.startElement("phoneticPr");
        writer.writeAttribute("fontId", "1");
        writer.writeAttribute("type", "noConversion");
        writer.endElement(); // phoneticPr
        
        writer.endElement(); // si
    }
    
    writer.endElement(); // sst
    writer.endDocument();
}


size_t SharedStringTable::getMemoryUsage() const {
    size_t usage = sizeof(SharedStringTable);
    
    // 计算unordered_map的内存使用
    usage += string_to_id_.bucket_count() * sizeof(std::pair<std::string, int32_t>);
    for (const auto& [str, id] : string_to_id_) {
        usage += str.capacity();
    }
    
    // 计算vector的内存使用
    usage += id_to_string_.capacity() * sizeof(std::string);
    for (const auto& str : id_to_string_) {
        usage += str.capacity();
    }
    
    return usage;
}

SharedStringTable::CompressionStats SharedStringTable::getCompressionStats() const {
    CompressionStats stats;
    
    // 计算原始大小（如果每个字符串都单独存储）
    stats.original_size = 0;
    std::unordered_map<std::string, size_t> string_counts;
    
    // 统计每个字符串的使用次数
    for (const auto& [str, id] : string_to_id_) {
        string_counts[str] = 1;  // 至少使用一次
    }
    
    // 计算原始总大小
    for (const auto& [str, count] : string_counts) {
        stats.original_size += str.size() * count;
    }
    
    // 计算压缩后大小（每个唯一字符串只存储一次）
    stats.compressed_size = 0;
    for (const auto& str : id_to_string_) {
        stats.compressed_size += str.size();
    }
    
    // 计算压缩率
    if (stats.original_size > 0) {
        stats.compression_ratio = 1.0 - (static_cast<double>(stats.compressed_size) / stats.original_size);
    } else {
        stats.compression_ratio = 0.0;
    }
    
    return stats;
}

}} // namespace fastexcel::core
