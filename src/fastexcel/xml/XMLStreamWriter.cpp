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
    // 如果缓冲区中有数据且用户没有处理，记录警告
    if ((buffer_pos_ > 0 || !whole_.empty()) && !direct_file_mode_) {
        LOG_WARN("XMLStreamWriter destroyed with {} bytes in buffer and {} bytes in whole_", buffer_pos_, whole_.size());
    }
    
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
    if (buffer_pos_ == 0) return;

    if (direct_file_mode_ && output_file_) {
        fwrite(buffer_, 1, buffer_pos_, output_file_);
    } else {
        // 缓冲模式：把数据拼到 whole_，避免丢失
        whole_.append(buffer_, buffer_pos_);
    }
    buffer_pos_ = 0;
}

void XMLStreamWriter::startDocument() {
    buffer_pos_ = 0;
    whole_.clear();  // 清空累积的缓冲区，确保每次开始新文档时都是干净的
    const char* xml_decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    size_t xml_decl_len = strlen(xml_decl);
    writeRawDirect(xml_decl, xml_decl_len);
}

void XMLStreamWriter::endDocument() {
    while (!element_stack_.empty()) {
        endElement();
    }
    
    // 确保所有数据都被写入
    flushBuffer();     // 现在缓冲模式也安全
}

void XMLStreamWriter::startElement(const char* name) {
    if (in_element_) {
        writeRawDirect(">", 1);
        in_element_ = false;
    }
    
    writeRawDirect("<", 1);
    size_t name_len = strlen(name);
    writeRawDirect(name, name_len);
    element_stack_.push(name);
    in_element_ = true;
}

void XMLStreamWriter::endElement() {
    if (element_stack_.empty()) {
        LOG_WARN("Attempted to end element when stack is empty");
        return;
    }
    
    const char* name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        writeRawDirect("/>", 2);
        in_element_ = false;
    } else {
        writeRawDirect("</", 2);
        writeRawDirect(name, strlen(name));
        writeRawDirect(">", 1);
    }
}

void XMLStreamWriter::writeEmptyElement(const char* name) {
    if (in_element_) {
        writeRawDirect(">", 1);
        in_element_ = false;
    }
    
    writeRawDirect("<", 1);
    size_t name_len = strlen(name);
    writeRawDirect(name, name_len);
    writeRawDirect("/>", 2);
}

void XMLStreamWriter::writeAttribute(const char* name, const char* value) {
    if (!in_element_) {
        LOG_WARN("Attempted to write attribute '{}' outside of element", name);
        return;
    }
    
    writeRawDirect(" ", 1);
    size_t name_len = strlen(name);
    writeRawDirect(name, name_len);
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
        return std::string();
    }
    
    // 在缓冲模式下，返回 whole_ + buffer_ 的完整内容
    std::string result;
    result.reserve(whole_.size() + buffer_pos_);
    result.append(whole_);
    result.append(buffer_, buffer_pos_);
    return result;
}

void XMLStreamWriter::clear() {
    buffer_pos_ = 0;
    whole_.clear();
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
    // 开始批处理模式，暂存属性直到endAttributeBatch调用
    // 这里我们只是设置一个标志，实际属性会暂存在pending_attributes_中
    // 这个方法主要是为了API完整性，实际优化在endAttributeBatch中
}

void XMLStreamWriter::endAttributeBatch() {
    // 结束批处理模式，将所有暂存的属性一次性写入
    if (!pending_attributes_.empty()) {
        for (const auto& attr : pending_attributes_) {
            writeRawDirect(" ", 1);
            writeRawDirect(attr.key.c_str(), attr.key.length());
            writeRawDirect("=\"", 2);
            
            size_t value_len = attr.value.length();
            if (needsAttributeEscaping(attr.value.c_str(), value_len)) {
                if (direct_file_mode_ && output_file_) {
                    escapeAttributesToFile(attr.value.c_str(), value_len);
                } else {
                    escapeAttributesToBuffer(attr.value.c_str(), value_len);
                }
            } else {
                writeRawDirect(attr.value.c_str(), value_len);
            }
            
            writeRawDirect("\"", 1);
        }
        pending_attributes_.clear();
    }
}

// 内部实现方法

void XMLStreamWriter::writeRawToBuffer(const char* data, size_t length) {
    // 如果数据太大，分批处理
    while (length > 0) {
        size_t available_space = BUFFER_SIZE - buffer_pos_;
        
        if (available_space == 0) {
            // 缓冲区已满，需要刷新
            flushBuffer();
            available_space = BUFFER_SIZE - buffer_pos_;
            
            // 如果刷新后仍然没有可用空间，直接返回
            if (available_space == 0) {
                return;
            }
        }
        
        size_t chunk_size = std::min(length, available_space);
        memcpy(buffer_ + buffer_pos_, data, chunk_size);
        buffer_pos_ += chunk_size;
        data += chunk_size;
        length -= chunk_size;
    }
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
            case '\'':
                writeRawToBuffer(APOS_REPLACEMENT, APOS_REPLACEMENT_LEN);
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
            case '\'':
                fwrite(APOS_REPLACEMENT, 1, APOS_REPLACEMENT_LEN, output_file_);
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
            case '\'':
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
