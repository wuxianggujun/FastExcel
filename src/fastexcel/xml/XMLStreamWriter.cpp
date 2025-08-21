#include "fastexcel/utils/ModuleLoggers.hpp"
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
    callback_mode_ = false;
    auto_flush_ = true;
}

XMLStreamWriter::XMLStreamWriter(const std::function<void(const char*, size_t)>& callback) : XMLStreamWriter() {
    // 设置回调模式
    callback_mode_ = true;
    write_callback_ = [callback](const std::string& chunk) {
        callback(chunk.c_str(), chunk.size());
    };
}

XMLStreamWriter::XMLStreamWriter(const std::string& filename) : XMLStreamWriter() {
    // 设置直接文件模式（统一通过 Path 处理跨平台打开）
    fastexcel::core::Path path(filename);
    FILE* file = path.openForWrite(true);
    if (file) {
        setDirectFileMode(file, true);
    } else {
        XML_ERROR("Failed to open file for writing: {}", filename);
    }
}

XMLStreamWriter::~XMLStreamWriter() {
    // 只有在缓冲区中有数据且没有被正常处理时才记录警告
    if (buffer_pos_ > 0 && !direct_file_mode_ && !callback_mode_) {
        // 只有在数据量较大时才记录警告，避免正常使用时的噪音
        if (buffer_pos_ > 100) {
            XML_WARN("XMLStreamWriter destroyed with {} bytes in buffer", buffer_pos_);
        }
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
    
    XML_DEBUG("XMLStreamWriter switched to direct file mode");
}

// 不再提供 setBufferedMode() 接口，统一使用流式写入
// 现在只支持直接文件模式和回调模式

void XMLStreamWriter::setCallbackMode(WriteCallback callback, bool auto_flush) {
    if (!callback) {
        XML_ERROR("Invalid callback provided to setCallbackMode");
        return;
    }
    
    // 刷新当前缓冲区内容
    flushBuffer();
    
    direct_file_mode_ = false;
    callback_mode_ = true;
    write_callback_ = std::move(callback);
    auto_flush_ = auto_flush;
    
    XML_DEBUG("XMLStreamWriter switched to callback mode with auto_flush={}", auto_flush);
}

void XMLStreamWriter::flushBuffer() {
    if (buffer_pos_ == 0) return;

    if (direct_file_mode_ && output_file_) {
        fwrite(buffer_, 1, buffer_pos_, output_file_);
        buffer_pos_ = 0;
    } else if (callback_mode_ && write_callback_) {
        // 回调模式：调用回调函数
        // 确保正确地创建字符串并清空缓冲区
        std::string chunk(buffer_, buffer_pos_);
        buffer_pos_ = 0;  // 先清空缓冲区位置
        write_callback_(chunk);
    } else {
        // 兼容性模式：如果没有设置输出目标，静默丢弃数据
        // 这主要是为了支持toString()方法的兼容性
        // 在实际使用中应该设置适当的输出目标
        buffer_pos_ = 0;
    }
}

void XMLStreamWriter::startDocument() {
    buffer_pos_ = 0;
    // 添加换行符以与常见工具的格式保持一致
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
    // 存储副本，避免外部传入临时 c_str 悬空
    element_stack_.push(std::string(name, name_len));
    in_element_ = true;
}

void XMLStreamWriter::endElement() {
    if (element_stack_.empty()) {
        XML_WARN("Attempted to end element when stack is empty");
        return;
    }
    
    const std::string& name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        writeRawDirect("/>", 2);
        in_element_ = false;
    } else {
        writeRawDirect("</", 2);
        writeRawDirect(name.c_str(), name.size());
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
        XML_WARN("Attempted to write attribute '{}' outside of element", name);
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
    if (!in_element_) {
        XML_WARN("Attempted to write attribute '{}' outside of element", name);
        return;
    }
    
    char buffer[32];
    int length = snprintf(buffer, sizeof(buffer), "%d", value);
    
    writeRawDirect(" ", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("=\"", 2);
    writeRawDirect(buffer, length);
    writeRawDirect("\"", 1);
}

void XMLStreamWriter::writeAttribute(const char* name, double value) {
    if (!in_element_) {
        XML_WARN("Attempted to write attribute '{}' outside of element", name);
        return;
    }
    
    char buffer[64];
    int length = snprintf(buffer, sizeof(buffer), "%.6g", value);
    
    writeRawDirect(" ", 1);
    writeRawDirect(name, strlen(name));
    writeRawDirect("=\"", 2);
    writeRawDirect(buffer, length);
    writeRawDirect("\"", 1);
}

void XMLStreamWriter::writeAttribute(const char* name, const std::string& value) {
    writeAttribute(name, value.c_str());
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

void XMLStreamWriter::writeText(const std::string& text) {
    writeText(text.c_str());
}

void XMLStreamWriter::writeRaw(const char* data) {
    if (data) {
        writeRawDirect(data, strlen(data));
    }
}

void XMLStreamWriter::writeRaw(const std::string& data) {
    writeRawDirect(data.c_str(), data.length());
}

// toString()方法已彻底删除 - 专注极致性能，所有XML生成都使用流式模式

void XMLStreamWriter::clear() {
    buffer_pos_ = 0;
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    in_element_ = false;
    pending_attributes_.clear();
}

bool XMLStreamWriter::writeToFile(const std::string& filename) {
    fastexcel::core::Path path(filename);
    FILE* file = path.openForWrite(true);
    if (!file) {
        XML_ERROR("Failed to open file '{}' for writing", filename);
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
    
    XML_INFO("XMLStreamWriter now writing to file '{}'", filename);
    return true;
}

bool XMLStreamWriter::setOutputFile(FILE* file, bool take_ownership) {
    if (!file) {
        XML_ERROR("Invalid file pointer provided");
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
    
    XML_DEBUG("XMLStreamWriter now writing to provided file stream");
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
            
            // 如果刷新后仍然没有可用空间，说明有问题
            if (available_space == 0) {
                XML_ERROR("Buffer flush failed, cannot write more data");
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
    } else if (callback_mode_ && write_callback_) {
        // 在回调模式下，总是使用缓冲区以确保数据的完整性
        writeRawToBuffer(data, length);
        // 如果缓冲区接近满了，自动刷新
        if (auto_flush_ && buffer_pos_ >= BUFFER_SIZE * 0.9) {
            flushBuffer();
        }
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
