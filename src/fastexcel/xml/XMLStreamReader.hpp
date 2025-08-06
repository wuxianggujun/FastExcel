#pragma once

#include <string>
#include <vector>
#include <stack>
#include <memory>
#include <functional>
#include <unordered_map>
#include <cstdio>
#include <expat.h>
#include "../core/Expected.hpp"
#include "../core/ErrorCode.hpp"

// 项目日志头文件
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

using namespace fastexcel::core;

/**
 * @brief 高性能流式XML解析器，基于libexpat
 * 
 * 这个类提供了与XMLStreamWriter配套的流式XML解析功能，主要特点：
 * - 基于libexpat的SAX解析，内存效率高
 * - 支持大文件流式解析，不会将整个文档加载到内存
 * - 事件驱动的回调机制，灵活处理各种XML结构
 * - 支持从文件、内存缓冲区、流等多种数据源解析
 * - 线程安全的设计
 * - 完整的错误处理和日志记录
 */

// 解析错误枚举
enum class XMLParseError {
    Ok,                    // 解析成功
    InvalidInput,          // 无效输入
    ParserCreateFailed,    // 解析器创建失败
    ParseFailed,           // 解析失败
    IoError,               // I/O错误
    EncodingError,         // 编码错误
    MemoryError,           // 内存错误
    CallbackError          // 回调函数错误
};

// 为 XMLParseError 提供 bool 转换操作符
constexpr bool operator!(XMLParseError error) noexcept {
    return error != XMLParseError::Ok;
}

constexpr bool isSuccess(XMLParseError error) noexcept {
    return error == XMLParseError::Ok;
}

constexpr bool isError(XMLParseError error) noexcept {
    return error != XMLParseError::Ok;
}

// XML属性结构
struct XMLAttribute {
    std::string name;
    std::string value;
    
    XMLAttribute(const std::string& n, const std::string& v) 
        : name(n), value(v) {}
};

// XML元素信息
struct XMLElement {
    std::string name;
    std::vector<XMLAttribute> attributes;
    std::string text_content;
    int depth;
    
    XMLElement(const std::string& n, int d) 
        : name(n), depth(d) {}
};

class XMLStreamReader {
public:
    // 事件回调函数类型定义
    using StartElementCallback = std::function<void(const std::string& name, const std::vector<XMLAttribute>& attributes, int depth)>;
    using EndElementCallback = std::function<void(const std::string& name, int depth)>;
    using TextCallback = std::function<void(const std::string& text, int depth)>;
    using CommentCallback = std::function<void(const std::string& comment, int depth)>;
    using ProcessingInstructionCallback = std::function<void(const std::string& target, const std::string& data, int depth)>;
    using ErrorCallback = std::function<void(XMLParseError error, const std::string& message, int line, int column)>;

private:
    // libexpat解析器句柄
    XML_Parser parser_ = nullptr;
    
    // 解析状态
    bool is_parsing_ = false;
    int current_depth_ = 0;
    XMLParseError last_error_ = XMLParseError::Ok;
    std::string last_error_message_;
    
    // 元素栈，用于跟踪嵌套结构
    std::stack<XMLElement> element_stack_;
    
    // 当前文本内容累积
    std::string current_text_;
    bool collecting_text_ = false;
    
    // 缓冲区设置
    static constexpr size_t BUFFER_SIZE = 8192;
    char buffer_[BUFFER_SIZE];
    
    // 回调函数
    StartElementCallback start_element_callback_;
    EndElementCallback end_element_callback_;
    TextCallback text_callback_;
    CommentCallback comment_callback_;
    ProcessingInstructionCallback pi_callback_;
    ErrorCallback error_callback_;
    
    // 解析选项
    bool trim_whitespace_ = true;
    bool collect_text_ = true;
    bool namespace_aware_ = false;
    std::string encoding_ = "UTF-8";
    
    // 统计信息
    size_t bytes_parsed_ = 0;
    size_t elements_parsed_ = 0;
    
    // libexpat回调函数（静态）
    static void XMLCALL startElementHandler(void* userData, const XML_Char* name, const XML_Char** attrs);
    static void XMLCALL endElementHandler(void* userData, const XML_Char* name);
    static void XMLCALL characterDataHandler(void* userData, const XML_Char* data, int len);
    static void XMLCALL commentHandler(void* userData, const XML_Char* data);
    static void XMLCALL processingInstructionHandler(void* userData, const XML_Char* target, const XML_Char* data);
    
    // 内部辅助方法
    bool initializeParser();
    void cleanupParser();
    void resetState();
    std::vector<XMLAttribute> parseAttributes(const XML_Char** attrs);
    std::string trimString(const std::string& str) const;
    void handleError(XMLParseError error, const std::string& message);
    
public:
    XMLStreamReader();
    ~XMLStreamReader();
    
    // 禁用拷贝构造和赋值
    XMLStreamReader(const XMLStreamReader&) = delete;
    XMLStreamReader& operator=(const XMLStreamReader&) = delete;
    
    // 回调函数设置
    void setStartElementCallback(StartElementCallback callback);
    void setEndElementCallback(EndElementCallback callback);
    void setTextCallback(TextCallback callback);
    void setCommentCallback(CommentCallback callback);
    void setProcessingInstructionCallback(ProcessingInstructionCallback callback);
    void setErrorCallback(ErrorCallback callback);
    
    // 解析选项设置
    void setTrimWhitespace(bool trim);
    void setCollectText(bool collect);
    void setNamespaceAware(bool aware);
    void setEncoding(const std::string& encoding);
    
    // 解析方法
    XMLParseError parseFromFile(const std::string& filename);
    XMLParseError parseFromFile(FILE* file);
    XMLParseError parseFromString(const std::string& xml_content);
    XMLParseError parseFromBuffer(const char* buffer, size_t size);
    XMLParseError parseChunk(const char* chunk, size_t size, bool is_final = false);
    
    // 流式解析支持
    XMLParseError beginParsing();
    XMLParseError feedData(const char* data, size_t size);
    XMLParseError endParsing();
    
    // 状态查询
    bool isParsing() const { return is_parsing_; }
    XMLParseError getLastError() const { return last_error_; }
    std::string getLastErrorMessage() const { return last_error_message_; }
    int getCurrentDepth() const { return current_depth_; }
    size_t getBytesParsed() const { return bytes_parsed_; }
    size_t getElementsParsed() const { return elements_parsed_; }
    
    // 解析器信息
    int getCurrentLineNumber() const;
    int getCurrentColumnNumber() const;
    std::string getParserVersion() const;
    
    // 便利方法：简单的DOM风格解析（适合小文档）
    struct SimpleElement {
        std::string name;
        std::unordered_map<std::string, std::string> attributes;
        std::string text;
        std::vector<std::unique_ptr<SimpleElement>> children;
        SimpleElement* parent = nullptr;
        
        SimpleElement(const std::string& n) : name(n) {}
        
        // 查找子元素
        SimpleElement* findChild(const std::string& element_name) const;
        std::vector<SimpleElement*> findChildren(const std::string& element_name) const;
        
        // 获取属性值
        std::string getAttribute(const std::string& attr_name, const std::string& defaultValue = "") const;
        bool hasAttribute(const std::string& attr_name) const;
        
        // 获取文本内容
        std::string getTextContent() const;
    };
    
    // 简单DOM解析（将整个文档解析为树结构）
    std::unique_ptr<SimpleElement> parseToDOM(const std::string& xml_content);
    std::unique_ptr<SimpleElement> parseFileToDOM(const std::string& filename);
};

}} // namespace fastexcel::xml