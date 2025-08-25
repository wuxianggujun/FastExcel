#include "XMLStreamWriter.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/core/Exception.hpp"
#include <iomanip>
#include <fmt/format.h>
#include <algorithm>

namespace fastexcel {
namespace xml {

XMLStreamWriter::XMLStreamWriter() {
    try {
        initializeInternal();
        output_mode_ = OutputMode::MEMORY_BUFFER;
        FASTEXCEL_LOG_DEBUG("XMLStreamWriter created in memory buffer mode");
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriter::XMLStreamWriter(const std::string& filename) {
    try {
        initializeInternal();
        switchToFileMode(filename);
        FASTEXCEL_LOG_DEBUG("XMLStreamWriter created with file: {}", filename);
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriter::XMLStreamWriter(WriteCallback callback) {
    try {
        initializeInternal();
        switchToCallbackMode(std::move(callback));
        FASTEXCEL_LOG_DEBUG("XMLStreamWriter created with callback mode");
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriter::~XMLStreamWriter() {
    try {
        if (!buffer_.empty()) {
            flush();
        }
        
        // 检查是否有未关闭的元素
        if (!element_stack_.empty()) {
            FASTEXCEL_LOG_WARN("XMLStreamWriter destroyed with {} unclosed elements", 
                              element_stack_.size());
        }
        
        cleanupInternal();
        
        FASTEXCEL_LOG_DEBUG("XMLStreamWriter destroyed. Bytes written: {}, Flushes: {}", 
                           bytes_written_, flush_count_);
    } catch (...) {
        // 析构函数中不抛出异常
    }
}

XMLStreamWriter::XMLStreamWriter(XMLStreamWriter&& other) noexcept
    : buffer_(std::move(other.buffer_))
    , file_wrapper_(std::move(other.file_wrapper_))
    , output_mode_(other.output_mode_)
    , write_callback_(std::move(other.write_callback_))
    , element_stack_(std::move(other.element_stack_))
    , in_element_(other.in_element_)
    , pending_attributes_(std::move(other.pending_attributes_))
    , memory_buffer_(std::move(other.memory_buffer_))
    , bytes_written_(other.bytes_written_)
    , flush_count_(other.flush_count_) {
    
    // 重置被移动对象的状态
    other.output_mode_ = OutputMode::MEMORY_BUFFER;
    other.in_element_ = false;
    other.bytes_written_ = 0;
    other.flush_count_ = 0;
}

XMLStreamWriter& XMLStreamWriter::operator=(XMLStreamWriter&& other) noexcept {
    if (this != &other) {
        // 清理当前资源
        try {
            if (!buffer_.empty()) {
                flush();
            }
            cleanupInternal();
        } catch (...) {
            // 忽略清理时的异常
        }
        
        // 移动资源
        buffer_ = std::move(other.buffer_);
        file_wrapper_ = std::move(other.file_wrapper_);
        output_mode_ = other.output_mode_;
        write_callback_ = std::move(other.write_callback_);
        element_stack_ = std::move(other.element_stack_);
        in_element_ = other.in_element_;
        pending_attributes_ = std::move(other.pending_attributes_);
        memory_buffer_ = std::move(other.memory_buffer_);
        bytes_written_ = other.bytes_written_;
        flush_count_ = other.flush_count_;
        
        // 重置被移动对象
        other.output_mode_ = OutputMode::MEMORY_BUFFER;
        other.in_element_ = false;
        other.bytes_written_ = 0;
        other.flush_count_ = 0;
    }
    return *this;
}

void XMLStreamWriter::initializeInternal() {    
    // 设置缓冲区刷新回调
    initializeBuffer();
    
    // 预留属性容器空间
    pending_attributes_.reserve(16);
    
    // 重置状态
    in_element_ = false;
    bytes_written_ = 0;
    flush_count_ = 0;
}

void XMLStreamWriter::cleanupInternal() noexcept {
    try {
        pending_attributes_.clear();
        memory_buffer_.clear();
        file_wrapper_.reset();
    } catch (...) {
        // 清理时忽略异常
    }
}

void XMLStreamWriter::initializeBuffer() {
    auto flush_callback = [this](const char* data, size_t length) {
        flushToOutput(data, length);
    };
    
    buffer_.setFlushCallback(std::move(flush_callback));
    buffer_.setAutoFlush(true);
}

void XMLStreamWriter::flushToOutput(const char* data, size_t length) {
    if (length == 0) return;
    
    bytes_written_ += length;
    ++flush_count_;
    
    try {
        switch (output_mode_) {
        case OutputMode::FILE_DIRECT:
            if (file_wrapper_ && file_wrapper_->get()) {
                size_t written = std::fwrite(data, 1, length, file_wrapper_->get());
                if (written != length) {
                    throw core::FileException(
                        "Failed to write data to file",
                        "",
                        core::ErrorCode::FileWriteError,
                        __FILE__, __LINE__
                    );
                }
            }
            break;
            
        case OutputMode::CALLBACK:
            if (write_callback_) {
                write_callback_(std::string(data, length));
            }
            break;
            
        case OutputMode::MEMORY_BUFFER:
            memory_buffer_.append(data, length);
            break;
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Error in flushToOutput: {}", e.what());
        throw;
    }
}

void XMLStreamWriter::switchToFileMode(const std::string& filename) {
    // 先刷新当前缓冲区
    flush();
    
    try {
        file_wrapper_ = std::make_unique<utils::FileWrapper>(filename, "w");
        output_mode_ = OutputMode::FILE_DIRECT;
        write_callback_ = nullptr;
        memory_buffer_.clear();
        
        FASTEXCEL_LOG_DEBUG("Switched to file mode: {}", filename);
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to switch to file mode: {}", e.what());
        throw;
    }
}

void XMLStreamWriter::switchToCallbackMode(WriteCallback callback) {
    if (!callback) {
        throw core::ParameterException(
            "Callback cannot be null",
            "callback",
            __FILE__, __LINE__
        );
    }
    
    // 先刷新当前缓冲区
    flush();
    
    output_mode_ = OutputMode::CALLBACK;
    write_callback_ = std::move(callback);
    file_wrapper_.reset();
    memory_buffer_.clear();
    
    FASTEXCEL_LOG_DEBUG("Switched to callback mode");
}

void XMLStreamWriter::switchToMemoryMode() {
    // 先刷新当前缓冲区
    flush();
    
    output_mode_ = OutputMode::MEMORY_BUFFER;
    write_callback_ = nullptr;
    file_wrapper_.reset();
    memory_buffer_.clear();
    
    FASTEXCEL_LOG_DEBUG("Switched to memory buffer mode");
}

void XMLStreamWriter::startDocument(const std::string& encoding) {
    std::string declaration = "<?xml version=\"1.0\" encoding=\"" + encoding + "\"?>\n";
    buffer_.append(declaration);
    
    FASTEXCEL_LOG_DEBUG("Started XML document with encoding: {}", encoding);
}

void XMLStreamWriter::endDocument() {
    // 确保所有元素都已关闭
    while (!element_stack_.empty()) {
        FASTEXCEL_LOG_WARN("Auto-closing unclosed element: {}", element_stack_.top());
        endElement();
    }
    
    // 最终刷新
    flush();
    
    FASTEXCEL_LOG_DEBUG("Ended XML document");
}

void XMLStreamWriter::startElement(const std::string& name) {
    if (name.empty()) {
        throw core::ParameterException(
            "Element name cannot be empty",
            "name",
            __FILE__, __LINE__
        );
    }
    
    ensureElementClosed();
    
    buffer_.append('<');
    buffer_.append(name);
    
    element_stack_.push(name);
    in_element_ = true;
    
    FASTEXCEL_LOG_TRACE("Started element: {}", name);
}

void XMLStreamWriter::endElement() {
    if (element_stack_.empty()) {
        throw core::OperationException(
            "No element to close",
            "endElement",
            core::ErrorCode::InvalidArgument,
            __FILE__, __LINE__
        );
    }
    
    std::string element_name = element_stack_.top();
    element_stack_.pop();
    
    if (in_element_) {
        // 自闭合元素
        writeAttributesToBuffer();
        buffer_.append(" />");
        in_element_ = false;
    } else {
        // 完整关闭标签
        buffer_.append("</");
        buffer_.append(element_name);
        buffer_.append('>');
    }
    
    FASTEXCEL_LOG_TRACE("Ended element: {}", element_name);
}

void XMLStreamWriter::writeEmptyElement(const std::string& name) {
    if (name.empty()) {
        throw core::ParameterException(
            "Element name cannot be empty",
            "name",
            __FILE__, __LINE__
        );
    }
    
    ensureElementClosed();
    
    buffer_.append('<');
    buffer_.append(name);
    writeAttributesToBuffer();
    buffer_.append(" />");
    
    FASTEXCEL_LOG_TRACE("Wrote empty element: {}", name);
}

void XMLStreamWriter::writeAttribute(const std::string& name, const std::string& value) {
    if (!in_element_) {
        throw core::OperationException(
            "Cannot write attribute outside of element",
            "writeAttribute",
            core::ErrorCode::InvalidArgument,
            __FILE__, __LINE__
        );
    }
    
    if (name.empty()) {
        throw core::ParameterException(
            "Attribute name cannot be empty",
            "name",
            __FILE__, __LINE__
        );
    }
    
    pending_attributes_.emplace_back(name, value);
    
    FASTEXCEL_LOG_TRACE("Added attribute: {}=\"{}\"", name, value);
}

void XMLStreamWriter::writeAttribute(const std::string& name, int value) {
    // 直接将整数写入属性缓存，避免临时std::string
    if (!in_element_) {
        throw core::OperationException("Cannot write attribute outside of element","writeAttribute", core::ErrorCode::InvalidArgument, __FILE__, __LINE__);
    }
    pending_attributes_.emplace_back(name, fmt::format("{}", value));
}

void XMLStreamWriter::writeAttribute(const std::string& name, double value) {
    if (!in_element_) {
        throw core::OperationException("Cannot write attribute outside of element","writeAttribute", core::ErrorCode::InvalidArgument, __FILE__, __LINE__);
    }
    // 使用fmt的快速路径格式化，避免ostringstream
    std::string str = fmt::format("{}", value);
    // 可选：移除尾随零和点（保持与既有行为一致）
    auto pos = str.find_last_not_of('0');
    if (pos != std::string::npos && pos + 1 < str.size()) str.erase(pos + 1);
    if (!str.empty() && str.back() == '.') str.pop_back();
    pending_attributes_.emplace_back(name, std::move(str));
}

void XMLStreamWriter::writeAttribute(const std::string& name, bool value) {
    if (!in_element_) {
        throw core::OperationException("Cannot write attribute outside of element","writeAttribute", core::ErrorCode::InvalidArgument, __FILE__, __LINE__);
    }
    pending_attributes_.emplace_back(name, value ? std::string("1") : std::string("0"));
}

void XMLStreamWriter::writeAttribute(const std::string& name, std::string_view value) {
    if (!in_element_) {
        throw core::OperationException(
            "Cannot write attribute outside of element",
            "writeAttribute",
            core::ErrorCode::InvalidArgument,
            __FILE__, __LINE__
        );
    }
    if (name.empty()) {
        throw core::ParameterException(
            "Attribute name cannot be empty",
            "name",
            __FILE__, __LINE__
        );
    }
    // 直接存储到 pending_attributes_ 为 std::string，集中一次性分配
    pending_attributes_.emplace_back(name, std::string(value));
}

void XMLStreamWriter::writeText(const std::string& text) {
    if (text.empty()) {
        return;
    }
    
    ensureElementClosed();
    writeEscapedText(text);
    
    FASTEXCEL_LOG_TRACE("Wrote text: {}", text.substr(0, std::min(text.size(), size_t(50))));
}

void XMLStreamWriter::writeText(int value) {
    ensureElementClosed();
    auto s = fmt::format("{}", value);
    buffer_.append(s.c_str(), s.size());
}

void XMLStreamWriter::writeText(size_t value) {
    ensureElementClosed();
    auto s = fmt::format("{}", value);
    buffer_.append(s.c_str(), s.size());
}

void XMLStreamWriter::writeText(double value) {
    ensureElementClosed();
    auto s = fmt::format("{}", value);
    buffer_.append(s.c_str(), s.size());
}

void XMLStreamWriter::writeRaw(const std::string& data) {
    ensureElementClosed();
    buffer_.append(data);
    
    FASTEXCEL_LOG_TRACE("Wrote raw data: {}", data.substr(0, std::min(data.size(), size_t(50))));
}

void XMLStreamWriter::writeCDATA(const std::string& data) {
    ensureElementClosed();
    buffer_.append("<![CDATA[");
    buffer_.append(data);
    buffer_.append("]]>");
    
    FASTEXCEL_LOG_TRACE("Wrote CDATA: {}", data.substr(0, std::min(data.size(), size_t(50))));
}

void XMLStreamWriter::writeComment(const std::string& comment) {
    ensureElementClosed();
    buffer_.append("<!-- ");
    buffer_.append(comment);
    buffer_.append(" -->");
    
    FASTEXCEL_LOG_TRACE("Wrote comment: {}", comment);
}

void XMLStreamWriter::startAttributeBatch() {
    // 预留更多属性空间
    pending_attributes_.reserve(pending_attributes_.size() + 32);
}

void XMLStreamWriter::endAttributeBatch() {
    // 属性批处理结束，可以进行优化处理
    writeAttributesToBuffer();
}

void XMLStreamWriter::flush() {
    buffer_.flush();
    
    // 如果是文件模式，同时刷新文件缓冲区
    if (output_mode_ == OutputMode::FILE_DIRECT && file_wrapper_) {
        file_wrapper_->flush();
    }
}

void XMLStreamWriter::clear() {
    buffer_.clear();
    
    // 清理状态
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    in_element_ = false;
    pending_attributes_.clear();
    
    if (output_mode_ == OutputMode::MEMORY_BUFFER) {
        memory_buffer_.clear();
    }
    
    bytes_written_ = 0;
    flush_count_ = 0;
    
    FASTEXCEL_LOG_DEBUG("XMLStreamWriter cleared");
}

std::string XMLStreamWriter::toString() const {
    if (output_mode_ != OutputMode::MEMORY_BUFFER) {
        FASTEXCEL_LOG_WARN("toString() called in non-memory mode");
        return "";
    }
    
    // 如果缓冲区中还有数据，需要先获取
    std::string result = memory_buffer_;
    if (!buffer_.empty()) {
        result.append(buffer_.data(), buffer_.size());
    }
    
    return result;
}

bool XMLStreamWriter::isEmpty() const {
    return buffer_.empty() && 
           (output_mode_ != OutputMode::MEMORY_BUFFER || memory_buffer_.empty()) &&
           bytes_written_ == 0;
}

void XMLStreamWriter::setAutoFlush(bool auto_flush) {
    buffer_.setAutoFlush(auto_flush);
    FASTEXCEL_LOG_DEBUG("Auto flush set to: {}", auto_flush);
}

void XMLStreamWriter::reserveAttributeCapacity(size_t capacity) {
    pending_attributes_.reserve(capacity);
}

void XMLStreamWriter::ensureElementClosed() {
    if (in_element_) {
        writeAttributesToBuffer();
        buffer_.append('>');
        in_element_ = false;
    }
}

void XMLStreamWriter::writeAttributesToBuffer() {
    if (pending_attributes_.empty()) return;
    
    // 预估所需空间 - 避免多次内存重分配
    size_t estimated_size = 0;
    for (const auto& attr : pending_attributes_) {
        // 空格(1) + key + ="(2) + value + "(1) + 转义字符的额外空间
        estimated_size += 4 + attr.key.size() + estimateEscapedSize(attr.value);
    }
    
    // 使用临时缓冲区批量构建属性字符串
    std::string attribute_buffer;
    attribute_buffer.reserve(estimated_size);
    
    // 批量构建属性字符串
    for (const auto& attr : pending_attributes_) {
        attribute_buffer += ' ';
        attribute_buffer += attr.key;
        attribute_buffer += "=\"";
        
        // 内联XML转义 - 避免临时字符串创建
        appendEscapedInline(attribute_buffer, attr.value);
        
        attribute_buffer += '\"';
    }
    
    // 一次性写入到主缓冲区
    buffer_.append(attribute_buffer);
    pending_attributes_.clear();
}

void XMLStreamWriter::writeEscapedAttribute(const std::string& value) {
    // 直接使用XMLUtils工具类进行转义
    std::string escaped = utils::XMLUtils::escapeXML(value);
    buffer_.append(escaped);
}

// 性能优化的辅助方法实现
size_t XMLStreamWriter::estimateEscapedSize(const std::string& text) const {
    size_t extra = 0;
    for (char c : text) {
        switch (c) {
            case '<':
            case '>':
                extra += 3; // &lt; &gt; = 4字符，原来1字符，增加3
                break;
            case '&':
                extra += 4; // &amp; = 5字符，原来1字符，增加4
                break;
            case '"':
            case '\'':
                extra += 5; // &quot; &apos; = 6字符，原来1字符，增加5
                break;
        }
    }
    return text.size() + extra;
}

void XMLStreamWriter::appendEscapedInline(std::string& target, const std::string& source) const {
    for (char c : source) {
        switch (c) {
            case '<':
                target += "&lt;";
                break;
            case '>':
                target += "&gt;";
                break;
            case '&':
                target += "&amp;";
                break;
            case '"':
                target += "&quot;";
                break;
            case '\'':
                target += "&apos;";
                break;
            default:
                target += c;
                break;
        }
    }
}

void XMLStreamWriter::writeEscapedText(const std::string& text) {
    // 直接使用XMLUtils工具类进行转义
    std::string escaped = utils::XMLUtils::escapeXML(text);
    buffer_.append(escaped);
}

} // namespace xml
} // namespace fastexcel
