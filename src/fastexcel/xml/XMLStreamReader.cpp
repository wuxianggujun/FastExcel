#include "fastexcel/xml/XMLStreamReader.hpp"
#include <algorithm>
#include <cstring>
#include <fstream>
#include <iostream>
#include <fmt/format.h>
#include "fastexcel/core/Path.hpp"

namespace fastexcel {
namespace xml {

XMLStreamReader::XMLStreamReader() {
    // 为元素栈预留空间
    element_stack_slim_.reserve(MAX_DEPTH);
    
    // 为属性池预留空间
    attribute_pool_.reserve(128); // 初始预留128个属性
    
    resetState();
}

XMLStreamReader::~XMLStreamReader() {
    cleanupParser();
}

bool XMLStreamReader::initializeParser() {
    cleanupParser();
    
    // 创建expat解析器
    if (namespace_aware_) {
        parser_ = XML_ParserCreateNS(encoding_.empty() ? nullptr : encoding_.c_str(), '|');
    } else {
        parser_ = XML_ParserCreate(encoding_.empty() ? nullptr : encoding_.c_str());
    }
    
    if (!parser_) {
        handleError(XMLParseError::ParserCreateFailed, "Failed to create XML parser");
        return false;
    }
    
    // 设置用户数据指针
    XML_SetUserData(parser_, this);
    
    // 设置回调函数
    XML_SetElementHandler(parser_, startElementHandler, endElementHandler);
    XML_SetCharacterDataHandler(parser_, characterDataHandler);
    XML_SetCommentHandler(parser_, commentHandler);
    XML_SetProcessingInstructionHandler(parser_, processingInstructionHandler);
    
    FASTEXCEL_LOG_DEBUG("XML parser initialized with encoding: {}", encoding_.empty() ? "default" : encoding_);
    return true;
}

void XMLStreamReader::cleanupParser() {
    if (parser_) {
        XML_ParserFree(parser_);
        parser_ = nullptr;
    }
}

void XMLStreamReader::resetState() {
    is_parsing_ = false;
    current_depth_ = 0;
    last_error_ = XMLParseError::Ok;
    last_error_message_.clear();
    
    // 清空轻量级栈
    element_stack_slim_.clear();
    
    // 清空属性池
    attribute_pool_.clear();
    
    current_text_.clear();
    collecting_text_ = false;
    text_reserved_ = false;
    bytes_parsed_ = 0;
    elements_parsed_ = 0;
}

// 回调函数设置
void XMLStreamReader::setStartElementCallback(StartElementCallback callback) {
    start_element_callback_ = std::move(callback);
}

void XMLStreamReader::setEndElementCallback(EndElementCallback callback) {
    end_element_callback_ = std::move(callback);
}

void XMLStreamReader::setTextCallback(TextCallback callback) {
    text_callback_ = std::move(callback);
}

void XMLStreamReader::setCommentCallback(CommentCallback callback) {
    comment_callback_ = std::move(callback);
}

void XMLStreamReader::setProcessingInstructionCallback(ProcessingInstructionCallback callback) {
    pi_callback_ = std::move(callback);
}

void XMLStreamReader::setErrorCallback(ErrorCallback callback) {
    error_callback_ = std::move(callback);
}

void XMLStreamReader::setTrimWhitespace(bool trim) {
    trim_whitespace_ = trim;
}

void XMLStreamReader::setCollectText(bool collect) {
    collect_text_ = collect;
}

void XMLStreamReader::setNamespaceAware(bool aware) {
    namespace_aware_ = aware;
}

void XMLStreamReader::setEncoding(const std::string& encoding) {
    encoding_ = encoding;
}

XMLParseError XMLStreamReader::parseFromFile(const std::string& filename) {
    // 统一通过 Path 进行跨平台文件打开
    fastexcel::core::Path path(filename);
    FILE* file = path.openForRead(true);
    if (!file) {
        std::string error_msg = "Failed to open file: " + filename;
        handleError(XMLParseError::IoError, error_msg);
        return XMLParseError::IoError;
    }

    XMLParseError result = parseFromFile(file);
    fclose(file);
    return result;
}

XMLParseError XMLStreamReader::parseFromFile(FILE* file) {
    if (!file) {
        handleError(XMLParseError::InvalidInput, "Invalid file pointer");
        return XMLParseError::InvalidInput;
    }
    
    if (!initializeParser()) {
        return last_error_;
    }
    
    resetState();
    is_parsing_ = true;
    
    // 使用ParseBuffer API减少拷贝
    size_t bytes_read;
    while ((bytes_read = fread(buffer_, 1, BUFFER_SIZE, file)) > 0) {
        bytes_parsed_ += bytes_read;
        
        // 使用ParseBuffer而不是Parse，减少一次内存拷贝
        void* expat_buffer = XML_GetBuffer(parser_, static_cast<int>(bytes_read));
        if (!expat_buffer) {
            handleError(XMLParseError::MemoryError, "Failed to get Expat buffer");
            is_parsing_ = false;
            return XMLParseError::MemoryError;
        }
        
        // 直接拷贝到Expat内部缓冲区
        std::memcpy(expat_buffer, buffer_, bytes_read);
        
        int is_final = feof(file) ? 1 : 0;
        if (XML_ParseBuffer(parser_, static_cast<int>(bytes_read), is_final) == XML_STATUS_ERROR) {
            std::string error_msg = fmt::format("Parse error at line {}, column {}: {}", 
                XML_GetCurrentLineNumber(parser_),
                XML_GetCurrentColumnNumber(parser_),
                XML_ErrorString(XML_GetErrorCode(parser_)));
            handleError(XMLParseError::ParseFailed, error_msg);
            is_parsing_ = false;
            return XMLParseError::ParseFailed;
        }
        
        if (is_final) break;
    }
    
    is_parsing_ = false;
    FASTEXCEL_LOG_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
    return XMLParseError::Ok;
}

XMLParseError XMLStreamReader::parseFromString(const std::string& xml_content) {
    return parseFromBuffer(xml_content.c_str(), xml_content.size());
}

XMLParseError XMLStreamReader::parseFromBuffer(const char* buffer, size_t size) {
    if (!buffer || size == 0) {
        handleError(XMLParseError::InvalidInput, "Invalid buffer or size");
        return XMLParseError::InvalidInput;
    }
    
    if (!initializeParser()) {
        return last_error_;
    }
    
    resetState();
    is_parsing_ = true;
    bytes_parsed_ = size;
    
    // 使用ParseBuffer API
    void* expat_buffer = XML_GetBuffer(parser_, static_cast<int>(size));
    if (!expat_buffer) {
        handleError(XMLParseError::MemoryError, "Failed to get Expat buffer");
        is_parsing_ = false;
        return XMLParseError::MemoryError;
    }
    
    std::memcpy(expat_buffer, buffer, size);
    
    if (XML_ParseBuffer(parser_, static_cast<int>(size), 1) == XML_STATUS_ERROR) {
        std::string error_msg = fmt::format("Parse error at line {}, column {}: {}", 
            XML_GetCurrentLineNumber(parser_),
            XML_GetCurrentColumnNumber(parser_),
            XML_ErrorString(XML_GetErrorCode(parser_)));
        handleError(XMLParseError::ParseFailed, error_msg);
        is_parsing_ = false;
        return XMLParseError::ParseFailed;
    }
    
    is_parsing_ = false;
    FASTEXCEL_LOG_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
    return XMLParseError::Ok;
}

XMLParseError XMLStreamReader::parseChunk(const char* chunk, size_t size, bool is_final) {
    if (!parser_) {
        handleError(XMLParseError::ParserCreateFailed, "Parser not initialized");
        return XMLParseError::ParserCreateFailed;
    }
    
    if (!chunk || size == 0) {
        if (is_final) {
            // 结束解析
            if (XML_ParseBuffer(parser_, 0, 1) == XML_STATUS_ERROR) {
                std::string error_msg = fmt::format("Parse error at line {}, column {}: {}", 
                    XML_GetCurrentLineNumber(parser_),
                    XML_GetCurrentColumnNumber(parser_),
                    XML_ErrorString(XML_GetErrorCode(parser_)));
                handleError(XMLParseError::ParseFailed, error_msg);
                return XMLParseError::ParseFailed;
            }
            is_parsing_ = false;
        }
        return XMLParseError::Ok;
    }
    
    bytes_parsed_ += size;
    
    // 使用ParseBuffer API
    void* expat_buffer = XML_GetBuffer(parser_, static_cast<int>(size));
    if (!expat_buffer) {
        handleError(XMLParseError::MemoryError, "Failed to get Expat buffer");
        return XMLParseError::MemoryError;
    }
    
    std::memcpy(expat_buffer, chunk, size);
    
    if (XML_ParseBuffer(parser_, static_cast<int>(size), is_final ? 1 : 0) == XML_STATUS_ERROR) {
        std::string error_msg = fmt::format("Parse error at line {}, column {}: {}", 
            XML_GetCurrentLineNumber(parser_),
            XML_GetCurrentColumnNumber(parser_),
            XML_ErrorString(XML_GetErrorCode(parser_)));
        handleError(XMLParseError::ParseFailed, error_msg);
        return XMLParseError::ParseFailed;
    }
    
    if (is_final) {
        is_parsing_ = false;
        FASTEXCEL_LOG_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
    }
    
    return XMLParseError::Ok;
}

XMLParseError XMLStreamReader::beginParsing() {
    if (!initializeParser()) {
        return last_error_;
    }
    
    resetState();
    is_parsing_ = true;
    return XMLParseError::Ok;
}

XMLParseError XMLStreamReader::feedData(const char* data, size_t size) {
    return parseChunk(data, size, false);
}

XMLParseError XMLStreamReader::endParsing() {
    return parseChunk(nullptr, 0, true);
}

int XMLStreamReader::getCurrentLineNumber() const {
    return parser_ ? XML_GetCurrentLineNumber(parser_) : -1;
}

int XMLStreamReader::getCurrentColumnNumber() const {
    return parser_ ? XML_GetCurrentColumnNumber(parser_) : -1;
}

std::string XMLStreamReader::getParserVersion() const {
    return XML_ExpatVersion();
}

// libexpat回调函数实现
void XMLCALL XMLStreamReader::startElementHandler(void* userData, const XML_Char* name, const XML_Char** attrs) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    // 使用string_view，直接引用Expat的缓冲区（零拷贝）
    std::string_view element_name{name, std::strlen(name)};
    
    reader->elements_parsed_++;
    
    // 使用零拷贝优化的属性解析
    auto attributes = reader->parseAttributes(attrs);
    
    // 创建轻量级元素并推入栈
    if (reader->element_stack_slim_.size() < MAX_DEPTH) {
        reader->element_stack_slim_.emplace_back(
            element_name, 
            reader->current_depth_,
            static_cast<uint32_t>(reader->attribute_pool_.size() - attributes.size()),
            static_cast<uint16_t>(attributes.size())
        );
    }
    
    // 调用用户回调
    if (reader->start_element_callback_) {
        try {
            reader->start_element_callback_(element_name, attributes, reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "Start element callback error: " + std::string(e.what()));
        }
    }
    
    reader->current_depth_++;
    reader->current_text_.clear();
    reader->collecting_text_ = reader->collect_text_;
    
    // 确保文本缓冲区预留
    if (reader->collecting_text_) {
        reader->ensureTextReserve();
    }
}

void XMLCALL XMLStreamReader::endElementHandler(void* userData, const XML_Char* name) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    reader->current_depth_--;
    
    // 使用string_view，直接引用Expat的缓冲区（零拷贝）
    std::string_view element_name{name, std::strlen(name)};
    
    // 处理累积的文本内容
    if (reader->collecting_text_ && !reader->current_text_.empty()) {
        // 使用string_view trim（零拷贝）
        std::string_view text_content = reader->trim_whitespace_ ? 
            reader->trimStringView(reader->current_text_) : std::string_view{reader->current_text_};
        
        if (!text_content.empty() && reader->text_callback_) {
            try {
                reader->text_callback_(text_content, reader->current_depth_);
            } catch (const std::exception& e) {
                reader->handleError(XMLParseError::CallbackError, "Text callback error: " + std::string(e.what()));
            }
        }
    }
    
    // 调用用户回调
    if (reader->end_element_callback_) {
        try {
            reader->end_element_callback_(element_name, reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "End element callback error: " + std::string(e.what()));
        }
    }
    
    // 弹出轻量级元素栈
    if (!reader->element_stack_slim_.empty()) {
        reader->element_stack_slim_.pop_back();
    }
    
    reader->current_text_.clear();
    reader->collecting_text_ = reader->collect_text_;
}

void XMLCALL XMLStreamReader::characterDataHandler(void* userData, const XML_Char* data, int len) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    if (reader->collecting_text_ && len > 0) {
        reader->current_text_.append(data, len);
    }
}

void XMLCALL XMLStreamReader::commentHandler(void* userData, const XML_Char* data) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    if (reader->comment_callback_) {
        try {
            std::string_view comment{data, std::strlen(data)};
            reader->comment_callback_(comment, reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "Comment callback error: " + std::string(e.what()));
        }
    }
}

void XMLCALL XMLStreamReader::processingInstructionHandler(void* userData, const XML_Char* target, const XML_Char* data) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    if (reader->pi_callback_) {
        try {
            std::string_view target_view{target, std::strlen(target)};
            std::string_view data_view = data ? std::string_view{data, std::strlen(data)} : std::string_view{};
            reader->pi_callback_(target_view, data_view, reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "Processing instruction callback error: " + std::string(e.what()));
        }
    }
}

// 属性解析方法（零拷贝优化）
span<const XMLAttribute> XMLStreamReader::parseAttributes(const XML_Char** attrs) {
    size_t attr_start = attribute_pool_.size();
    
    if (attrs) {
        for (int i = 0; attrs[i]; i += 2) {
            if (attrs[i + 1]) {
                // 使用string_view，直接引用Expat内部缓冲区
                attribute_pool_.emplace_back(
                    std::string_view{attrs[i], std::strlen(attrs[i])},
                    std::string_view{attrs[i + 1], std::strlen(attrs[i + 1])}
                );
            }
        }
    }
    
    size_t attr_count = attribute_pool_.size() - attr_start;
    return span<const XMLAttribute>{
        attribute_pool_.data() + attr_start, 
        attr_count
    };
}

// 字符串视图trim方法（零拷贝）
std::string_view XMLStreamReader::trimStringView(std::string_view str) const {
    size_t start = str.find_first_not_of(" \t\n\r");
    if (start == std::string_view::npos) {
        return std::string_view{};
    }
    
    size_t end = str.find_last_not_of(" \t\n\r");
    return str.substr(start, end - start + 1);
}

// 确保文本缓冲区预留
void XMLStreamReader::ensureTextReserve() {
    if (!text_reserved_) {
        current_text_.reserve(TEXT_RESERVE_SIZE);
        text_reserved_ = true;
    }
}

std::string XMLStreamReader::trimString(const std::string& str) const {
    size_t start = str.find_first_not_of(" \t\n\r");
    if (start == std::string::npos) {
        return "";
    }
    
    size_t end = str.find_last_not_of(" \t\n\r");
    return str.substr(start, end - start + 1);
}

void XMLStreamReader::handleError(XMLParseError error, const std::string& message) {
    last_error_ = error;
    last_error_message_ = message;
    
    FASTEXCEL_LOG_ERROR("XML parse error: {}", message);
    
    if (error_callback_) {
        try {
            error_callback_(error, message, getCurrentLineNumber(), getCurrentColumnNumber());
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_ERROR("Error in error callback: {}", e.what());
        }
    }
}

// SimpleElement 实现
XMLStreamReader::SimpleElement* XMLStreamReader::SimpleElement::findChild(const std::string& element_name) const {
    for (const auto& child : children) {
        if (child->name == element_name) {
            return child.get();
        }
    }
    return nullptr;
}

std::vector<XMLStreamReader::SimpleElement*> XMLStreamReader::SimpleElement::findChildren(const std::string& element_name) const {
    std::vector<SimpleElement*> result;
    for (const auto& child : children) {
        if (child->name == element_name) {
            result.push_back(child.get());
        }
    }
    return result;
}

XMLStreamReader::SimpleElement* XMLStreamReader::SimpleElement::findChildByPath(const std::string& path) const {
    if (path.empty()) return nullptr;
    
    size_t pos = path.find('/');
    if (pos == std::string::npos) {
        return findChild(path);
    }
    
    std::string first_part = path.substr(0, pos);
    std::string remaining = path.substr(pos + 1);
    
    SimpleElement* child = findChild(first_part);
    if (child) {
        return child->findChildByPath(remaining);
    }
    return nullptr;
}

std::string XMLStreamReader::SimpleElement::getAttribute(const std::string& attr_name, const std::string& defaultValue) const {
    auto it = attributes.find(attr_name);
    return it != attributes.end() ? it->second : defaultValue;
}

bool XMLStreamReader::SimpleElement::hasAttribute(const std::string& attr_name) const {
    return attributes.find(attr_name) != attributes.end();
}

void XMLStreamReader::SimpleElement::setAttribute(const std::string& attr_name, const std::string& value) {
    attributes[attr_name] = value;
}

void XMLStreamReader::SimpleElement::removeAttribute(const std::string& attr_name) {
    attributes.erase(attr_name);
}

std::string XMLStreamReader::SimpleElement::getTextContent() const {
    return text;
}

void XMLStreamReader::SimpleElement::setTextContent(const std::string& content) {
    text = content;
}

std::string XMLStreamReader::SimpleElement::getInnerText() const {
    std::string result = text;
    for (const auto& child : children) {
        result += child->getInnerText();
    }
    return result;
}

XMLStreamReader::SimpleElement* XMLStreamReader::SimpleElement::appendChild(const std::string& element_name) {
    auto child = std::make_unique<SimpleElement>(element_name);
    SimpleElement* child_ptr = child.get();
    child_ptr->parent = this;
    children.push_back(std::move(child));
    return child_ptr;
}

XMLStreamReader::SimpleElement* XMLStreamReader::SimpleElement::prependChild(const std::string& element_name) {
    auto child = std::make_unique<SimpleElement>(element_name);
    SimpleElement* child_ptr = child.get();
    child_ptr->parent = this;
    children.insert(children.begin(), std::move(child));
    return child_ptr;
}

bool XMLStreamReader::SimpleElement::removeChild(SimpleElement* child) {
    auto it = std::find_if(children.begin(), children.end(),
                          [child](const std::unique_ptr<SimpleElement>& ptr) {
                              return ptr.get() == child;
                          });
    if (it != children.end()) {
        children.erase(it);
        return true;
    }
    return false;
}

void XMLStreamReader::SimpleElement::clear() {
    children.clear();
}

void XMLStreamReader::SimpleElement::forEach(std::function<void(SimpleElement*)> callback) {
    for (const auto& child : children) {
        callback(child.get());
    }
}

void XMLStreamReader::SimpleElement::forEachRecursive(std::function<void(SimpleElement*, int)> callback, int depth) {
    callback(this, depth);
    for (const auto& child : children) {
        child->forEachRecursive(callback, depth + 1);
    }
}

int XMLStreamReader::SimpleElement::getDepth() const {
    int depth = 0;
    const SimpleElement* current = parent;
    while (current) {
        depth++;
        current = current->parent;
    }
    return depth;
}

std::string XMLStreamReader::SimpleElement::toString(int indent) const {
    std::string result;
    std::string indentation(indent * 2, ' ');
    
    result += indentation + "<" + name;
    for (const auto& attr : attributes) {
        result += " " + attr.first + "=\"" + attr.second + "\"";
    }
    
    if (text.empty() && children.empty()) {
        result += "/>\n";
    } else {
        result += ">";
        if (!text.empty()) {
            result += text;
        }
        if (!children.empty()) {
            result += "\n";
            for (const auto& child : children) {
                result += child->toString(indent + 1);
            }
            result += indentation;
        }
        result += "</" + name + ">\n";
    }
    
    return result;
}

void XMLStreamReader::SimpleElement::print(int indent) const {
    std::cout << toString(indent);
}

// DOM解析实现
std::unique_ptr<XMLStreamReader::SimpleElement> XMLStreamReader::parseToDOM(const std::string& xml_content) {
    std::unique_ptr<SimpleElement> root;
    std::stack<SimpleElement*> element_stack;
    
    // 设置回调函数（使用优化后的零拷贝回调）
    setStartElementCallback([&](std::string_view element_name, span<const XMLAttribute> attributes, int /*depth*/) {
        auto element = std::make_unique<SimpleElement>(std::string(element_name));
        
        // 转换属性（这里需要拷贝因为DOM要保存数据）
        for (const auto& attr : attributes) {
            element->attributes[std::string(attr.name)] = std::string(attr.value);
        }
        
        SimpleElement* element_ptr = element.get();
        
        if (element_stack.empty()) {
            // 根元素
            root = std::move(element);
        } else {
            // 子元素
            element_ptr->parent = element_stack.top();
            element_stack.top()->children.push_back(std::move(element));
        }
        
        element_stack.push(element_ptr);
    });
    
    setEndElementCallback([&](std::string_view /*element_name*/, int /*depth*/) {
        if (!element_stack.empty()) {
            element_stack.pop();
        }
    });
    
    setTextCallback([&](std::string_view text, int /*depth*/) {
        if (!element_stack.empty()) {
            if (element_stack.top()->text.empty()) {
                element_stack.top()->text = std::string(text);
            } else {
                element_stack.top()->text += std::string(text);
            }
        }
    });
    
    // 解析XML
    XMLParseError result = parseFromString(xml_content);
    if (result != XMLParseError::Ok) {
        FASTEXCEL_LOG_ERROR("Failed to parse XML to DOM: {}", last_error_message_);
        return nullptr;
    }
    
    return root;
}

std::unique_ptr<XMLStreamReader::SimpleElement> XMLStreamReader::parseFileToDOM(const std::string& filename) {
    std::ifstream file(filename);
    if (!file.is_open()) {
        handleError(XMLParseError::IoError, "Failed to open file: " + filename);
        return nullptr;
    }
    
    std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    return parseToDOM(content);
}

}} // namespace fastexcel::xml
