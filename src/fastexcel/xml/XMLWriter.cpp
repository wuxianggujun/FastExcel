#include "fastexcel/xml/XMLWriter.hpp"
#include <algorithm>
#include <cstring>
#include <cstdio>

namespace fastexcel {
namespace xml {

XMLWriter::XMLWriter() {
    // 预分配缓冲区
    buffer_.resize(INITIAL_BUFFER_SIZE);
    buffer_pos_ = 0;
    in_element_ = false;
    output_file_ = nullptr;
    owns_file_ = false;
}

XMLWriter::~XMLWriter() {
    clear();
    if (output_file_ && owns_file_) {
        fclose(output_file_);
    }
}

void XMLWriter::startDocument() {
    resetBuffer();
    writeRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
}

void XMLWriter::endDocument() {
    while (!element_stack_.empty()) {
        endElement();
    }
}

void XMLWriter::startElement(const std::string& name) {
    if (in_element_) {
        writeChar('>');
        in_element_ = false;
    }
    
    writeChar('<');
    writeRaw(name);
    element_stack_.push(name);
    in_element_ = true;
}

void XMLWriter::endElement() {
    if (element_stack_.empty()) {
        LOG_WARN("Attempted to end element when stack is empty");
        return;
    }
    
    std::string name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        writeRaw("/>");
        in_element_ = false;
    } else {
        writeRaw("</");
        writeRaw(name);
        writeChar('>');
    }
}

void XMLWriter::writeEmptyElement(const std::string& name) {
    if (in_element_) {
        writeChar('>');
        in_element_ = false;
    }
    
    writeChar('<');
    writeRaw(name);
    writeRaw("/>");
}

void XMLWriter::writeAttribute(const std::string& name, const std::string& value) {
    if (!in_element_) {
        LOG_WARN("Attempted to write attribute '{}' outside of element", name);
        return;
    }
    
    writeChar(' ');
    writeRaw(name);
    writeRaw("=\"");
    
    // 使用优化的属性转义，参考libxlsxwriter
    if (needsAttributeEscaping(value)) {
        std::string escaped_value = escapeAttributes(value);
        writeRaw(escaped_value);
    } else {
        writeRaw(value);
    }
    
    writeChar('\"');
}

void XMLWriter::writeAttribute(const std::string& name, int value) {
    // 使用高效的数字格式化
    writeAttribute(name, formatInt(value));
}

void XMLWriter::writeAttribute(const std::string& name, double value) {
    // 使用高效的数字格式化
    writeAttribute(name, formatDouble(value));
}

void XMLWriter::writeText(const std::string& text) {
    if (in_element_) {
        writeChar('>');
        in_element_ = false;
    }
    
    // 使用优化的文本转义，参考libxlsxwriter
    if (needsDataEscaping(text)) {
        std::string escaped_text = escapeData(text);
        writeRaw(escaped_text);
    } else {
        writeRaw(text);
    }
}

std::string XMLWriter::toString() const {
    return std::string(buffer_.data(), buffer_pos_);
}

void XMLWriter::clear() {
    resetBuffer();
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    in_element_ = false;
}

void XMLWriter::writeToBuffer(const char* data, size_t length) {
    if (output_file_) {
        // 如果设置了输出文件，直接写入文件
        fwrite(data, 1, length, output_file_);
    } else {
        // 否则写入缓冲区
        ensureCapacity(buffer_pos_ + length);
        memcpy(buffer_.data() + buffer_pos_, data, length);
        buffer_pos_ += length;
    }
}

void XMLWriter::ensureCapacity(size_t required) {
    if (required <= buffer_.size()) {
        return;
    }
    
    // 使用指数增长策略减少频繁重新分配
    size_t new_size = std::max(required, buffer_.size() * BUFFER_GROWTH_FACTOR);
    resizeBuffer(new_size);
    
    LOG_DEBUG("XMLWriter buffer resized to {} bytes", new_size);
}

std::string XMLWriter::escapeAttributes(const std::string& text) const {
    std::string result;
    result.reserve(text.size() * 6);  // 最坏情况下每个字符变成6个字符
    
    for (char c : text) {
        switch (c) {
            case '&':
                result.append(AMP_ESCAPE, strlen(AMP_ESCAPE));
                break;
            case '<':
                result.append(LT_ESCAPE, strlen(LT_ESCAPE));
                break;
            case '>':
                result.append(GT_ESCAPE, strlen(GT_ESCAPE));
                break;
            case '\"':
                result.append(QUOT_ESCAPE, strlen(QUOT_ESCAPE));
                break;
            case '\n':
                result.append(NL_ESCAPE, strlen(NL_ESCAPE));
                break;
            default:
                result.push_back(c);
                break;
        }
    }
    
    return result;
}

std::string XMLWriter::escapeData(const std::string& text) const {
    std::string result;
    result.reserve(text.size() * 5);  // 最坏情况下每个字符变成5个字符
    
    for (char c : text) {
        switch (c) {
            case '&':
                result.append(AMP_ESCAPE, strlen(AMP_ESCAPE));
                break;
            case '<':
                result.append(LT_ESCAPE, strlen(LT_ESCAPE));
                break;
            case '>':
                result.append(GT_ESCAPE, strlen(GT_ESCAPE));
                break;
            default:
                result.push_back(c);
                break;
        }
    }
    
    return result;
}

void XMLWriter::writeRaw(const char* data, size_t length) {
    writeToBuffer(data, length);
}

void XMLWriter::writeRaw(const std::string& str) {
    writeToBuffer(str.data(), str.length());
}

void XMLWriter::writeChar(char c) {
    if (output_file_) {
        fputc(c, output_file_);
    } else {
        ensureCapacity(buffer_pos_ + 1);
        buffer_[buffer_pos_++] = c;
    }
}

void XMLWriter::resizeBuffer(size_t new_size) {
    buffer_.resize(new_size);
}

void XMLWriter::resetBuffer() {
    buffer_pos_ = 0;
    // 不释放内存，保留已分配的缓冲区供后续使用
}

bool XMLWriter::writeToFile(const std::string& filename) {
    FILE* file = fopen(filename.c_str(), "wb");
    if (!file) {
        LOG_ERROR("Failed to open file '{}' for writing", filename);
        return false;
    }
    
    // 将当前的缓冲区内容写入文件
    if (buffer_pos_ > 0) {
        fwrite(buffer_.data(), 1, buffer_pos_, file);
    }
    
    // 设置文件输出
    output_file_ = file;
    owns_file_ = true;
    
    // 重置缓冲区，后续内容直接写入文件
    resetBuffer();
    
    LOG_INFO("XMLWriter now writing to file '{}'", filename);
    return true;
}

bool XMLWriter::setOutputFile(FILE* file) {
    if (!file) {
        LOG_ERROR("Invalid file pointer provided");
        return false;
    }
    
    // 将当前的缓冲区内容写入文件
    if (buffer_pos_ > 0) {
        fwrite(buffer_.data(), 1, buffer_pos_, file);
    }
    
    // 设置文件输出
    output_file_ = file;
    owns_file_ = false;
    
    // 重置缓冲区，后续内容直接写入文件
    resetBuffer();
    
    LOG_DEBUG("XMLWriter now writing to provided file stream");
    return true;
}

bool XMLWriter::needsAttributeEscaping(const std::string& text) const {
    // 快速检查是否需要转义，参考libxlsxwriter的优化
    return text.find_first_of("&<>\"\n") != std::string::npos;
}

bool XMLWriter::needsDataEscaping(const std::string& text) const {
    // 快速检查是否需要转义，参考libxlsxwriter的优化
    return text.find_first_of("&<>") != std::string::npos;
}

std::string XMLWriter::formatInt(int value) const {
    char buffer[32];
    int length = snprintf(buffer, sizeof(buffer), "%d", value);
    return std::string(buffer, length);
}

std::string XMLWriter::formatDouble(double value) const {
    char buffer[64];
    int length = snprintf(buffer, sizeof(buffer), "%.6g", value);
    return std::string(buffer, length);
}

}} // namespace fastexcel::xml