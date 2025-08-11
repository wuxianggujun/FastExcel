#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include <algorithm>
#include <cstring>
#include <fstream>
#include <sstream>

namespace fastexcel {
namespace xml {

XMLStreamReader::XMLStreamReader() {
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
    
    XML_DEBUG("XML parser initialized with encoding: {}", encoding_.empty() ? "default" : encoding_);
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
    
    // 清空栈
    while (!element_stack_.empty()) {
        element_stack_.pop();
    }
    
    current_text_.clear();
    collecting_text_ = false;
    bytes_parsed_ = 0;
    elements_parsed_ = 0;
}

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
    FILE* file = nullptr;
    errno_t err = fopen_s(&file, filename.c_str(), "rb");
    if (err != 0 || !file) {
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
    
    size_t bytes_read;
    while ((bytes_read = fread(buffer_, 1, BUFFER_SIZE, file)) > 0) {
        bytes_parsed_ += bytes_read;
        
        int is_final = feof(file) ? 1 : 0;
        if (XML_Parse(parser_, buffer_, static_cast<int>(bytes_read), is_final) == XML_STATUS_ERROR) {
            std::ostringstream oss;
            oss << "Parse error at line " << XML_GetCurrentLineNumber(parser_)
                << ", column " << XML_GetCurrentColumnNumber(parser_)
                << ": " << XML_ErrorString(XML_GetErrorCode(parser_));
            handleError(XMLParseError::ParseFailed, oss.str());
            is_parsing_ = false;
            return XMLParseError::ParseFailed;
        }
        
        if (is_final) break;
    }
    
    is_parsing_ = false;
    XML_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
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
    
    if (XML_Parse(parser_, buffer, static_cast<int>(size), 1) == XML_STATUS_ERROR) {
        std::ostringstream oss;
        oss << "Parse error at line " << XML_GetCurrentLineNumber(parser_)
            << ", column " << XML_GetCurrentColumnNumber(parser_)
            << ": " << XML_ErrorString(XML_GetErrorCode(parser_));
        handleError(XMLParseError::ParseFailed, oss.str());
        is_parsing_ = false;
        return XMLParseError::ParseFailed;
    }
    
    is_parsing_ = false;
    XML_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
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
            if (XML_Parse(parser_, "", 0, 1) == XML_STATUS_ERROR) {
                std::ostringstream oss;
                oss << "Parse error at line " << XML_GetCurrentLineNumber(parser_)
                    << ", column " << XML_GetCurrentColumnNumber(parser_)
                    << ": " << XML_ErrorString(XML_GetErrorCode(parser_));
                handleError(XMLParseError::ParseFailed, oss.str());
                return XMLParseError::ParseFailed;
            }
            is_parsing_ = false;
        }
        return XMLParseError::Ok;
    }
    
    bytes_parsed_ += size;
    
    if (XML_Parse(parser_, chunk, static_cast<int>(size), is_final ? 1 : 0) == XML_STATUS_ERROR) {
        std::ostringstream oss;
        oss << "Parse error at line " << XML_GetCurrentLineNumber(parser_)
            << ", column " << XML_GetCurrentColumnNumber(parser_)
            << ": " << XML_ErrorString(XML_GetErrorCode(parser_));
        handleError(XMLParseError::ParseFailed, oss.str());
        return XMLParseError::ParseFailed;
    }
    
    if (is_final) {
        is_parsing_ = false;
        XML_DEBUG("Successfully parsed {} bytes, {} elements", bytes_parsed_, elements_parsed_);
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
    
    std::string element_name(name);
    std::vector<XMLAttribute> attributes = reader->parseAttributes(attrs);
    
    // 创建新元素并推入栈
    XMLElement element(element_name, reader->current_depth_);
    element.attributes = attributes;
    reader->element_stack_.push(element);
    
    reader->elements_parsed_++;
    
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
}

void XMLCALL XMLStreamReader::endElementHandler(void* userData, const XML_Char* name) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    reader->current_depth_--;
    
    std::string element_name(name);
    
    // 处理累积的文本内容
    if (reader->collecting_text_ && !reader->current_text_.empty()) {
        std::string text_content = reader->trim_whitespace_ ? 
            reader->trimString(reader->current_text_) : reader->current_text_;
        
        if (!text_content.empty() && reader->text_callback_) {
            try {
                reader->text_callback_(text_content, reader->current_depth_);
            } catch (const std::exception& e) {
                reader->handleError(XMLParseError::CallbackError, "Text callback error: " + std::string(e.what()));
            }
        }
        
        // 更新栈顶元素的文本内容
        if (!reader->element_stack_.empty()) {
            reader->element_stack_.top().text_content = text_content;
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
    
    // 弹出元素栈
    if (!reader->element_stack_.empty()) {
        reader->element_stack_.pop();
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
            reader->comment_callback_(std::string(data), reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "Comment callback error: " + std::string(e.what()));
        }
    }
}

void XMLCALL XMLStreamReader::processingInstructionHandler(void* userData, const XML_Char* target, const XML_Char* data) {
    XMLStreamReader* reader = static_cast<XMLStreamReader*>(userData);
    
    if (reader->pi_callback_) {
        try {
            reader->pi_callback_(std::string(target), data ? std::string(data) : std::string(), reader->current_depth_);
        } catch (const std::exception& e) {
            reader->handleError(XMLParseError::CallbackError, "Processing instruction callback error: " + std::string(e.what()));
        }
    }
}

// 辅助方法实现
std::vector<XMLAttribute> XMLStreamReader::parseAttributes(const XML_Char** attrs) {
    std::vector<XMLAttribute> attributes;
    
    if (attrs) {
        for (int i = 0; attrs[i]; i += 2) {
            if (attrs[i + 1]) {
                attributes.emplace_back(std::string(attrs[i]), std::string(attrs[i + 1]));
            }
        }
    }
    
    return attributes;
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
    
    XML_ERROR("XML parse error: {}", message);
    
    if (error_callback_) {
        try {
            error_callback_(error, message, getCurrentLineNumber(), getCurrentColumnNumber());
        } catch (const std::exception& e) {
            XML_ERROR("Error in error callback: {}", e.what());
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

std::string XMLStreamReader::SimpleElement::getAttribute(const std::string& attr_name, const std::string& defaultValue) const {
    auto it = attributes.find(attr_name);
    return it != attributes.end() ? it->second : defaultValue;
}

bool XMLStreamReader::SimpleElement::hasAttribute(const std::string& attr_name) const {
    return attributes.find(attr_name) != attributes.end();
}

std::string XMLStreamReader::SimpleElement::getTextContent() const {
    return text;
}

// DOM解析实现
std::unique_ptr<XMLStreamReader::SimpleElement> XMLStreamReader::parseToDOM(const std::string& xml_content) {
    std::unique_ptr<SimpleElement> root;
    std::stack<SimpleElement*> element_stack;
    
    // 设置回调函数
    setStartElementCallback([&](const std::string& element_name, const std::vector<XMLAttribute>& attributes, int /*depth*/) {
        auto element = std::make_unique<SimpleElement>(element_name);
        
        // 转换属性
        for (const auto& attr : attributes) {
            element->attributes[attr.name] = attr.value;
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
    
    setEndElementCallback([&](const std::string& /*element_name*/, int /*depth*/) {
        if (!element_stack.empty()) {
            element_stack.pop();
        }
    });
    
    setTextCallback([&](const std::string& text, int /*depth*/) {
        if (!element_stack.empty()) {
            if (element_stack.top()->text.empty()) {
                element_stack.top()->text = text;
            } else {
                element_stack.top()->text += text;
            }
        }
    });
    
    // 解析XML
    XMLParseError result = parseFromString(xml_content);
    if (result != XMLParseError::Ok) {
        XML_ERROR("Failed to parse XML to DOM: {}", last_error_message_);
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