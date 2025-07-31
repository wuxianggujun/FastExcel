#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <algorithm>
#include <cstdio>
#include <cstring>

namespace fastexcel {
namespace xml {

XMLStreamWriter::XMLStreamWriter() {
    buffer_pos_ = 0;
    in_element_ = false;
    output_file_ = nullptr;
    owns_file_ = false;
    direct_file_mode_ = false;
}

XMLStreamWriter::~XMLStreamWriter() {
    clear();
    if (output_file_ && owns_file_) {
        fclose(output_file_);
    }
}

void XMLStreamWriter::setDirectFileMode(FILE* file, bool take_ownership) {
    // 刷新当前缓冲区内容
    flushBuffer();
    
    output_file_ = file;
    owns_file_ = take_ownership;
    direct_file_mode_ = true;
    
    LOG_DEBUG("XMLStreamWriter switched to direct file mode");
}

void XMLStreamWriter::setBufferedMode() {
    direct_file_mode_ = false;
    LOG_DEBUG("XMLStreamWriter switched to buffered mode");
}

void XMLStreamWriter::flushBuffer() {
    if (buffer_pos_ > 0 && output_file_) {
        fwrite(buffer_, 1, buffer_pos_, output_file_);
        buffer_pos_ = 0;
    }
}

void XMLStreamWriter::startDocument() {
    buffer_pos_ = 0;
    writeRawDirect("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n", 48);
}

void XMLStreamWriter::endDocument() {
    while (!element_stack_.empty()) {
        endElement();
    }
    flushBuffer();
}

void XMLStreamWriter::startElement(const char* name) {
    if (in_element_) {
        writeRawDirect(">", 1);
        in_element_ = false;
    }
    
    writeRawDirect("<", 1);
    writeRawDirect(name, strlen(name));
    element_stack_.push(name);
    in_element_ = true;
}

void XMLStreamWriter::endElement() {
    if (element_stack_.empty()) {
        LOG_WARN("Attempted to end element when stack is empty");
        return;
    }
    
    std::string name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        writeRawDirect("/>", 2);
        in_element_ = false;
    } else {
        writeRawDirect("</", 2);
        writeRawDirect(name.c_str(), name.length());
        writeRawDirect(">", 1);
    }
}

void XMLStreamWriter::writeEmptyElement(const char* name) {
    if (in_element_) {
        writeRawDirect(">", 1);
        in_element_ = false;
    }
    
    writeRawDirect("<", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("/>", 2);
}

void XMLStreamWriter::writeAttribute(const char* name, const char* value) {
    if (!in_element_) {
        LOG_WARN("Attempted to write attribute '{}' outside of element", name);
        return;
    }
    
    writeRawDirect(" ", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("=\"", 2);
    
    size_t value_len = strlen(value);
    if (needsAttributeEscaping(value, value_len)) {
        if (direct_file_mode_ && output_file_) {
            escapeAttributesToFile(value, value_len);
        } else {
            escapeAttributesToBuffer(value, value_len);
        }
    } else {
        writeRawDirect(value, value_len);
    }
    
    writeRawDirect("\"", 1);
}

void XMLStreamWriter::writeAttribute(const char* name, int value) {
    char buffer[32];
    int length = snprintf(buffer, sizeof(buffer), "%d", value);
    
    writeRawDirect(" ", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("=\"", 2);
    writeRawDirect(buffer, length);
    writeRawDirect("\"", 1);
}

void XMLStreamWriter::writeAttribute(const char* name, double value) {
    char buffer[64];
    int length = snprintf(buffer, sizeof(buffer), "%.6g", value);
    
    writeRawDirect(" ", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("=\"", 2);
    writeRawDirect(buffer, length);
    writeRawDirect("\"", 1);
}

void XMLStreamWriter::writeText(const char* text) {
    if (in_element_) {
        writeRawDirect(">", 1);
        in_element_ = false;
    }
    
    size_t text_len = strlen(text);
    if (needsDataEscaping(text, text_len)) {
        if (direct_file_mode_ && output_file_) {
            escapeDataToFile(text, text_len);
        } else {
            escapeDataToBuffer(text, text_len);
        }
    } else {
        writeRawDirect(text, text_len);
    }
}

std::string XMLStreamWriter::toString() const {
    if (direct_file_mode_) {
        LOG_WARN("toString() called in direct file mode, result may be incomplete");
    }
    return std::string(buffer_, buffer_pos_);
}

void XMLStreamWriter::clear() {
    buffer_pos_ = 0;
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    in_element_ = false;
    pending_attributes_.clear();
}

bool XMLStreamWriter::writeToFile(const std::string& filename) {
    FILE* file = fopen(filename.c_str(), "wb");
    if (!file) {
        LOG_ERROR("Failed to open file '{}' for writing", filename);
        return false;
    }
    
    // 将当前的缓冲区内容写入文件
    if (buffer_pos_ > 0) {
        fwrite(buffer_, 1, buffer_pos_, file);
        buffer_pos_ = 0;
    }
    
    // 设置文件输出
    output_file_ = file;
    owns_file_ = true;
    direct_file_mode_ = true;
    
    LOG_INFO("XMLStreamWriter now writing to file '{}'", filename);
    return true;
}

bool XMLStreamWriter::setOutputFile(FILE* file, bool take_ownership) {
    if (!file) {
        LOG_ERROR("Invalid file pointer provided");
        return false;
    }
    
    // 将当前的缓冲区内容写入文件
    if (buffer_pos_ > 0) {
        fwrite(buffer_, 1, buffer_pos_, file);
        buffer_pos_ = 0;
    }
    
    // 设置文件输出
    output_file_ = file;
    owns_file_ = take_ownership;
    direct_file_mode_ = true;
    
    LOG_DEBUG("XMLStreamWriter now writing to provided file stream");
    return true;
}

void XMLStreamWriter::startAttributeBatch() {
    // 批处理模式的实现可以在这里添加
}

void XMLStreamWriter::endAttributeBatch() {
    // 批处理模式的实现可以在这里添加
}

// 内部实现方法

void XMLStreamWriter::writeRawToBuffer(const char* data, size_t length) {
    if (buffer_pos_ + length > BUFFER_SIZE) {
        flushBuffer();
        
        // 如果数据仍然大于缓冲区，直接写入文件
        if (length > BUFFER_SIZE && output_file_) {
            fwrite(data, 1, length, output_file_);
            return;
        }
    }
    
    memcpy(buffer_ + buffer_pos_, data, length);
    buffer_pos_ += length;
}

void XMLStreamWriter::writeRawToFile(const char* data, size_t length) {
    if (output_file_) {
        fwrite(data, 1, length, output_file_);
    }
}

void XMLStreamWriter::writeRawDirect(const char* data, size_t length) {
    if (direct_file_mode_ && output_file_) {
        writeRawToFile(data, length);
    } else {
        writeRawToBuffer(data, length);
    }
}

void XMLStreamWriter::escapeAttributesToBuffer(const char* text, size_t length) {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
                writeRawToBuffer(AMP_REPLACEMENT, AMP_REPLACEMENT_LEN);
                break;
            case '<':
                writeRawToBuffer(LT_REPLACEMENT, LT_REPLACEMENT_LEN);
                break;
            case '>':
                writeRawToBuffer(GT_REPLACEMENT, GT_REPLACEMENT_LEN);
                break;
            case '\"':
                writeRawToBuffer(QUOT_REPLACEMENT, QUOT_REPLACEMENT_LEN);
                break;
            case '\n':
                writeRawToBuffer(NL_REPLACEMENT, NL_REPLACEMENT_LEN);
                break;
            default:
                writeRawToBuffer(&text[i], 1);
                break;
        }
    }
}

void XMLStreamWriter::escapeDataToBuffer(const char* text, size_t length) {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
                writeRawToBuffer(AMP_REPLACEMENT, AMP_REPLACEMENT_LEN);
                break;
            case '<':
                writeRawToBuffer(LT_REPLACEMENT, LT_REPLACEMENT_LEN);
                break;
            case '>':
                writeRawToBuffer(GT_REPLACEMENT, GT_REPLACEMENT_LEN);
                break;
            default:
                writeRawToBuffer(&text[i], 1);
                break;
        }
    }
}

void XMLStreamWriter::escapeAttributesToFile(const char* text, size_t length) {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
                fwrite(AMP_REPLACEMENT, 1, AMP_REPLACEMENT_LEN, output_file_);
                break;
            case '<':
                fwrite(LT_REPLACEMENT, 1, LT_REPLACEMENT_LEN, output_file_);
                break;
            case '>':
                fwrite(GT_REPLACEMENT, 1, GT_REPLACEMENT_LEN, output_file_);
                break;
            case '\"':
                fwrite(QUOT_REPLACEMENT, 1, QUOT_REPLACEMENT_LEN, output_file_);
                break;
            case '\n':
                fwrite(NL_REPLACEMENT, 1, NL_REPLACEMENT_LEN, output_file_);
                break;
            default:
                fwrite(&text[i], 1, 1, output_file_);
                break;
        }
    }
}

void XMLStreamWriter::escapeDataToFile(const char* text, size_t length) {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
                fwrite(AMP_REPLACEMENT, 1, AMP_REPLACEMENT_LEN, output_file_);
                break;
            case '<':
                fwrite(LT_REPLACEMENT, 1, LT_REPLACEMENT_LEN, output_file_);
                break;
            case '>':
                fwrite(GT_REPLACEMENT, 1, GT_REPLACEMENT_LEN, output_file_);
                break;
            default:
                fwrite(&text[i], 1, 1, output_file_);
                break;
        }
    }
}

bool XMLStreamWriter::needsAttributeEscaping(const char* text, size_t length) const {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
            case '<':
            case '>':
            case '\"':
            case '\n':
                return true;
        }
    }
    return false;
}

bool XMLStreamWriter::needsDataEscaping(const char* text, size_t length) const {
    for (size_t i = 0; i < length; i++) {
        switch (text[i]) {
            case '&':
            case '<':
            case '>':
                return true;
        }
    }
    return false;
}

}} // namespace fastexcel::xml