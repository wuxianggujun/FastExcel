//
// Created by wuxianggujun on 25-8-4.
//

#include "SharedStringsParser.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"

namespace fastexcel {
namespace reader {

void SharedStringsParser::onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int /*depth*/) {
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

void SharedStringsParser::onEndElement(std::string_view name, int /*depth*/) {
    if (name == "si") {
        // 结束共享字符串项，保存收集到的文本
        if (!parse_state_.current_text.empty()) {
            // 解码XML实体
            parse_state_.current_text = decodeXMLEntities(parse_state_.current_text);
        }
        
        strings_[parse_state_.current_string_index++] = std::move(parse_state_.current_text);
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

bool SharedStringsParser::parseStream(archive::ZipReader* zip_reader, const std::string& internal_path) {
    if (!zip_reader) return false;

    clear();

    xml::XMLStreamReader reader;

    reader.setStartElementCallback([this](std::string_view name, span<const xml::XMLAttribute> attributes, int depth){
        this->handleStartElement(name, attributes, depth);
    });
    reader.setEndElementCallback([this](std::string_view name, int depth){
        this->handleEndElement(name, depth);
    });
    reader.setTextCallback([this](std::string_view text, int depth){
        this->handleText(text, depth);
    });
    reader.setErrorCallback([this](xml::XMLParseError /*err*/, const std::string& msg, int line, int col){
        FASTEXCEL_LOG_ERROR("SST stream parse error at {}:{} -> {}", line, col, msg);
    });

    if (reader.beginParsing() != xml::XMLParseError::Ok) {
        return false;
    }

    auto err = zip_reader->streamFile(internal_path,
        [&reader](const uint8_t* data, size_t size){
            return reader.feedData(reinterpret_cast<const char*>(data), size) == xml::XMLParseError::Ok;
        },
        1 << 16);

    if (archive::isError(err)) {
        FASTEXCEL_LOG_ERROR("Zip stream failed: {}", internal_path);
        return false;
    }

    return reader.endParsing() == xml::XMLParseError::Ok;
}

} // namespace reader
} // namespace fastexcel
