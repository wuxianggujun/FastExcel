#include "fastexcel/utils/ModuleLoggers.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#include "SharedStringsParser.hpp"
#include <iostream>
#include <sstream>

namespace fastexcel {
namespace reader {

bool SharedStringsParser::parse(const std::string& xml_content) {
    if (xml_content.empty()) {
        return false;
    }
    
    try {
        clear();
        
        // 查找所有的 <si> 标签（String Item）
        size_t pos = 0;
        int index = 0;
        
        while ((pos = xml_content.find("<si", pos)) != std::string::npos) {
            // 找到 <si> 标签的结束位置
            size_t si_end = xml_content.find("</si>", pos);
            if (si_end == std::string::npos) {
                // 尝试查找自闭合标签
                size_t self_close = xml_content.find("/>", pos);
                if (self_close != std::string::npos && self_close < xml_content.find("<si", pos + 1)) {
                    // 空字符串项
                    strings_[index++] = "";
                    pos = self_close + 2;
                    continue;
                } else {
                    break; // 格式错误
                }
            }
            
            // 提取 <si> 标签内的内容
            size_t si_start = xml_content.find(">", pos) + 1;
            std::string si_content = xml_content.substr(si_start, si_end - si_start);
            
            // 解析字符串内容
            std::string text_content = extractTextContent(si_content, 0, si_content.length());
            
            // 解码XML实体
            text_content = decodeXMLEntities(text_content);
            
            strings_[index++] = text_content;
            pos = si_end + 5; // 跳过 </si>
        }
        
        return true;
        
    } catch (const std::exception& e) {
        READER_ERROR("解析共享字符串时发生错误: {}", e.what());
        return false;
    }
}

std::string SharedStringsParser::getString(int index) const {
    auto it = strings_.find(index);
    if (it != strings_.end()) {
        return it->second;
    }
    return "";
}

void SharedStringsParser::clear() {
    strings_.clear();
}

std::string SharedStringsParser::extractTextContent(const std::string& xml, size_t start_pos, size_t end_pos) {
    std::string result;
    size_t pos = start_pos;
    
    while (pos < end_pos) {
        // 查找 <t> 标签（Text）
        size_t t_start = xml.find("<t", pos);
        if (t_start == std::string::npos || t_start >= end_pos) {
            break;
        }
        
        // 找到 <t> 标签的开始内容位置
        size_t content_start = xml.find(">", t_start);
        if (content_start == std::string::npos || content_start >= end_pos) {
            break;
        }
        content_start++;
        
        // 找到 </t> 标签
        size_t t_end = xml.find("</t>", content_start);
        if (t_end == std::string::npos || t_end > end_pos) {
            // 尝试查找自闭合标签
            size_t self_close = xml.find("/>", t_start);
            if (self_close != std::string::npos && self_close < content_start) {
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        // 提取文本内容
        std::string text = xml.substr(content_start, t_end - content_start);
        result += text;
        
        pos = t_end + 4; // 跳过 </t>
    }
    
    // 如果没有找到 <t> 标签，可能是简单的文本内容
    if (result.empty() && start_pos < end_pos) {
        // 检查是否有其他格式的文本，如富文本格式
        size_t r_start = xml.find("<r>", start_pos);
        if (r_start != std::string::npos && r_start < end_pos) {
            // 处理富文本格式 <r><t>text</t></r>
            while (r_start != std::string::npos && r_start < end_pos) {
                size_t r_end = xml.find("</r>", r_start);
                if (r_end == std::string::npos || r_end > end_pos) {
                    break;
                }
                
                std::string r_content = xml.substr(r_start, r_end - r_start + 4);
                std::string r_text = extractTextContent(r_content, 0, r_content.length());
                result += r_text;
                
                r_start = xml.find("<r>", r_end);
            }
        }
    }
    
    return result;
}

std::string SharedStringsParser::decodeXMLEntities(const std::string& text) {
    std::string result = text;
    
    // 替换常见的XML实体
    size_t pos = 0;
    
    // &lt; -> <
    while ((pos = result.find("&lt;", pos)) != std::string::npos) {
        result.replace(pos, 4, "<");
        pos += 1;
    }
    
    pos = 0;
    // &gt; -> >
    while ((pos = result.find("&gt;", pos)) != std::string::npos) {
        result.replace(pos, 4, ">");
        pos += 1;
    }
    
    pos = 0;
    // &amp; -> &
    while ((pos = result.find("&amp;", pos)) != std::string::npos) {
        result.replace(pos, 5, "&");
        pos += 1;
    }
    
    pos = 0;
    // &quot; -> "
    while ((pos = result.find("&quot;", pos)) != std::string::npos) {
        result.replace(pos, 6, "\"");
        pos += 1;
    }
    
    pos = 0;
    // &apos; -> '
    while ((pos = result.find("&apos;", pos)) != std::string::npos) {
        result.replace(pos, 6, "'");
        pos += 1;
    }
    
    return result;
}

} // namespace reader
} // namespace fastexcel
