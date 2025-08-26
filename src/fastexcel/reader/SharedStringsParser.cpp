//
// Created by wuxianggujun on 25-8-4.
//

#include "SharedStringsParser.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace reader {

void SharedStringsParser::onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
    if (name == "si") {
        // 开始新的共享字符串项
        parse_state_.startNewString();
    }
    else if (name == "t") {
        // 开始文本元素
        parse_state_.in_text_element = true;
        startCollectingText();
    }
    else if (name == "r") {
        // 富文本运行元素
        parse_state_.in_rich_text = true;
    }
    // 其他元素（如 <rPh>, <phoneticPr> 等）我们忽略
}

void SharedStringsParser::onEndElement(const std::string& name, int depth) {
    if (name == "si") {
        // 结束共享字符串项，保存收集到的文本
        std::string final_text = parse_state_.current_text;
        if (!final_text.empty()) {
            // 解码XML实体
            final_text = decodeXMLEntities(final_text);
        }
        
        strings_[parse_state_.current_string_index++] = std::move(final_text);
        parse_state_.endString();
    }
    else if (name == "t") {
        // 结束文本元素，将收集的文本加入当前字符串
        if (parse_state_.in_text_element) {
            parse_state_.current_text += getCurrentText();
            parse_state_.in_text_element = false;
            stopCollectingText();
        }
    }
    else if (name == "r") {
        // 结束富文本运行
        parse_state_.in_rich_text = false;
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
    parse_state_.reset();
}

} // namespace reader
} // namespace fastexcel