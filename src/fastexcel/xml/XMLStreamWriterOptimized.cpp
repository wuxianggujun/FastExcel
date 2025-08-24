#include "XMLStreamWriterOptimized.hpp"
#include "fastexcel/xml/XMLEscapeSIMD.hpp"
#include "fastexcel/core/Exception.hpp"
#include <sstream>
#include <iomanip>
#include <algorithm>

namespace fastexcel {
namespace xml {

XMLStreamWriterOptimized::XMLStreamWriterOptimized() {
    try {
        initializeInternal();
        output_mode_ = OutputMode::MEMORY_BUFFER;
        FASTEXCEL_LOG_DEBUG("XMLStreamWriterOptimized created in memory buffer mode");
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriterOptimized::XMLStreamWriterOptimized(const std::string& filename) {
    try {
        initializeInternal();
        switchToFileMode(filename);
        FASTEXCEL_LOG_DEBUG("XMLStreamWriterOptimized created with file: {}", filename);
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriterOptimized::XMLStreamWriterOptimized(WriteCallback callback) {
    try {
        initializeInternal();
        switchToCallbackMode(std::move(callback));
        FASTEXCEL_LOG_DEBUG("XMLStreamWriterOptimized created with callback mode");
    } catch (...) {
        cleanupInternal();
        throw;
    }
}

XMLStreamWriterOptimized::~XMLStreamWriterOptimized() {
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
        
        FASTEXCEL_LOG_DEBUG("XMLStreamWriterOptimized destroyed. Bytes written: {}, Flushes: {}", 
                           bytes_written_, flush_count_);
    } catch (...) {
        // 析构函数中不抛出异常
    }
}

XMLStreamWriterOptimized::XMLStreamWriterOptimized(XMLStreamWriterOptimized&& other) noexcept
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

XMLStreamWriterOptimized& XMLStreamWriterOptimized::operator=(XMLStreamWriterOptimized&& other) noexcept {
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

void XMLStreamWriterOptimized::initializeInternal() {
    // 初始化SIMD转义器
    XMLEscapeSIMD::initialize();
    
    // 设置缓冲区刷新回调
    initializeBuffer();
    
    // 预留属性容器空间
    pending_attributes_.reserve(16);
    
    // 重置状态
    in_element_ = false;
    bytes_written_ = 0;
    flush_count_ = 0;
}

void XMLStreamWriterOptimized::cleanupInternal() noexcept {
    try {
        pending_attributes_.clear();
        memory_buffer_.clear();
        file_wrapper_.reset();
    } catch (...) {
        // 清理时忽略异常
    }
}

void XMLStreamWriterOptimized::initializeBuffer() {
    auto flush_callback = [this](const char* data, size_t length) {
        flushToOutput(data, length);
    };
    
    buffer_.setFlushCallback(std::move(flush_callback));
    buffer_.setAutoFlush(true);
}

void XMLStreamWriterOptimized::flushToOutput(const char* data, size_t length) {
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
                        core::ErrorCode::WriteError,
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

void XMLStreamWriterOptimized::switchToFileMode(const std::string& filename) {
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

void XMLStreamWriterOptimized::switchToCallbackMode(WriteCallback callback) {
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

void XMLStreamWriterOptimized::switchToMemoryMode() {
    // 先刷新当前缓冲区
    flush();
    
    output_mode_ = OutputMode::MEMORY_BUFFER;
    write_callback_ = nullptr;
    file_wrapper_.reset();
    memory_buffer_.clear();
    
    FASTEXCEL_LOG_DEBUG("Switched to memory buffer mode");
}

void XMLStreamWriterOptimized::startDocument(const std::string& encoding) {
    std::string declaration = "<?xml version=\"1.0\" encoding=\"" + encoding + "\"?>\n";
    buffer_.append(declaration);
    
    FASTEXCEL_LOG_DEBUG("Started XML document with encoding: {}", encoding);
}

void XMLStreamWriterOptimized::endDocument() {
    // 确保所有元素都已关闭
    while (!element_stack_.empty()) {
        FASTEXCEL_LOG_WARN("Auto-closing unclosed element: {}", element_stack_.top());
        endElement();
    }
    
    // 最终刷新
    flush();
    
    FASTEXCEL_LOG_DEBUG("Ended XML document");
}

void XMLStreamWriterOptimized::startElement(const std::string& name) {
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

void XMLStreamWriterOptimized::endElement() {
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

void XMLStreamWriterOptimized::writeEmptyElement(const std::string& name) {
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

void XMLStreamWriterOptimized::writeAttribute(const std::string& name, const std::string& value) {
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

void XMLStreamWriterOptimized::writeAttribute(const std::string& name, int value) {
    writeAttribute(name, std::to_string(value));
}

void XMLStreamWriterOptimized::writeAttribute(const std::string& name, double value) {
    std::ostringstream oss;
    oss << std::fixed << std::setprecision(6) << value;
    std::string str = oss.str();
    
    // 移除末尾的零
    str.erase(str.find_last_not_of('0') + 1, std::string::npos);
    if (str.back() == '.') {
        str.pop_back();
    }
    
    writeAttribute(name, str);
}

void XMLStreamWriterOptimized::writeAttribute(const std::string& name, bool value) {
    writeAttribute(name, value ? "true" : "false");
}

void XMLStreamWriterOptimized::writeText(const std::string& text) {
    if (text.empty()) {
        return;
    }
    
    ensureElementClosed();
    writeEscapedText(text);
    
    FASTEXCEL_LOG_TRACE("Wrote text: {}", text.substr(0, std::min(text.size(), size_t(50))));
}

void XMLStreamWriterOptimized::writeRaw(const std::string& data) {
    ensureElementClosed();
    buffer_.append(data);
    
    FASTEXCEL_LOG_TRACE("Wrote raw data: {}", data.substr(0, std::min(data.size(), size_t(50))));
}

void XMLStreamWriterOptimized::writeCDATA(const std::string& data) {
    ensureElementClosed();
    buffer_.append("<![CDATA[");
    buffer_.append(data);
    buffer_.append("]]>");
    
    FASTEXCEL_LOG_TRACE("Wrote CDATA: {}", data.substr(0, std::min(data.size(), size_t(50))));
}

void XMLStreamWriterOptimized::writeComment(const std::string& comment) {
    ensureElementClosed();
    buffer_.append("<!-- ");
    buffer_.append(comment);
    buffer_.append(" -->");
    
    FASTEXCEL_LOG_TRACE("Wrote comment: {}", comment);
}

void XMLStreamWriterOptimized::startAttributeBatch() {
    // 预留更多属性空间
    pending_attributes_.reserve(pending_attributes_.size() + 32);
}

void XMLStreamWriterOptimized::endAttributeBatch() {
    // 属性批处理结束，可以进行优化处理
    writeAttributesToBuffer();
}

void XMLStreamWriterOptimized::flush() {
    buffer_.flush();
    
    // 如果是文件模式，同时刷新文件缓冲区
    if (output_mode_ == OutputMode::FILE_DIRECT && file_wrapper_) {
        file_wrapper_->flush();
    }
}

void XMLStreamWriterOptimized::clear() {
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

std::string XMLStreamWriterOptimized::toString() const {
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

bool XMLStreamWriterOptimized::isEmpty() const {
    return buffer_.empty() && 
           (output_mode_ != OutputMode::MEMORY_BUFFER || memory_buffer_.empty()) &&
           bytes_written_ == 0;
}

void XMLStreamWriterOptimized::setAutoFlush(bool auto_flush) {
    buffer_.setAutoFlush(auto_flush);
    FASTEXCEL_LOG_DEBUG("Auto flush set to: {}", auto_flush);
}

void XMLStreamWriterOptimized::reserveAttributeCapacity(size_t capacity) {
    pending_attributes_.reserve(capacity);
}

void XMLStreamWriterOptimized::ensureElementClosed() {
    if (in_element_) {
        writeAttributesToBuffer();
        buffer_.append('>');
        in_element_ = false;
    }
}

void XMLStreamWriterOptimized::writeAttributesToBuffer() {
    for (const auto& attr : pending_attributes_) {
        buffer_.append(' ');
        buffer_.append(attr.key);
        buffer_.append("=\"");
        writeEscapedAttribute(attr.value);
        buffer_.append('\"');
    }
    pending_attributes_.clear();
}

void XMLStreamWriterOptimized::writeEscapedAttribute(const std::string& value) {
    // 使用SIMD优化的转义
    if (XMLEscapeSIMD::isAvailable()) {
        std::string escaped = XMLEscapeSIMD::escapeAttribute(value);
        buffer_.append(escaped);
    } else {
        // 回退到标准转义
        for (char c : value) {
            switch (c) {
            case '<': buffer_.append("&lt;"); break;
            case '>': buffer_.append("&gt;"); break;
            case '&': buffer_.append("&amp;"); break;
            case '\"': buffer_.append("&quot;"); break;
            case '\'': buffer_.append("&apos;"); break;
            default: buffer_.append(c); break;
            }
        }
    }
}

void XMLStreamWriterOptimized::writeEscapedText(const std::string& text) {
    // 使用SIMD优化的转义
    if (XMLEscapeSIMD::isAvailable()) {
        std::string escaped = XMLEscapeSIMD::escapeText(text);
        buffer_.append(escaped);
    } else {
        // 回退到标准转义
        for (char c : text) {
            switch (c) {
            case '<': buffer_.append("&lt;"); break;
            case '>': buffer_.append("&gt;"); break;
            case '&': buffer_.append("&amp;"); break;
            default: buffer_.append(c); break;
            }
        }
    }
}

// XMLWriterFactory 实现

std::unique_ptr<XMLStreamWriterOptimized> XMLWriterFactory::createFileWriter(const std::string& filename) {
    return std::make_unique<XMLStreamWriterOptimized>(filename);
}

std::unique_ptr<XMLStreamWriterOptimized> XMLWriterFactory::createCallbackWriter(XMLStreamWriterOptimized::WriteCallback callback) {
    return std::make_unique<XMLStreamWriterOptimized>(std::move(callback));
}

std::unique_ptr<XMLStreamWriterOptimized> XMLWriterFactory::createMemoryWriter() {
    return std::make_unique<XMLStreamWriterOptimized>();
}

std::pair<std::unique_ptr<XMLStreamWriterOptimized>, std::string> 
XMLWriterFactory::createTempFileWriter(const std::string& prefix) {
    // 创建临时文件
    utils::TempFileWrapper temp_file(prefix, ".xml");
    std::string temp_path = temp_file.getPath();
    
    // 创建写入器
    auto writer = std::make_unique<XMLStreamWriterOptimized>(temp_path);
    
    return {std::move(writer), std::move(temp_path)};
}

} // namespace xml
} // namespace fastexcel