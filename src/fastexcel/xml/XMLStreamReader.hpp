#include "fastexcel/utils/Logger.hpp"
#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <stack>
#include <memory>
#include <functional>
#include <unordered_map>
#include <cstdio>
#include <expat.h>
#include "fastexcel/core/Constants.hpp"
#include "fastexcel/core/Expected.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/core/span.hpp"

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

// XML属性结构（零拷贝优化）
struct XMLAttribute {
    std::string_view name;
    std::string_view value;
    
    XMLAttribute(std::string_view n, std::string_view v) 
        : name(n), value(v) {}
    
    // 向后兼容的构造函数
    XMLAttribute(const std::string& n, const std::string& v) 
        : name(n), value(v) {}
};

// 轻量级XML元素信息（栈优化版本）
struct XMLElementSlim {
    std::string_view name;
    uint32_t attr_start_offset;  // 属性在全局数组中的起始位置
    uint16_t attr_count;         // 属性数量
    uint16_t depth;
    
    XMLElementSlim(std::string_view n, int d, uint32_t attr_start = 0, uint16_t attr_cnt = 0) 
        : name(n), depth(static_cast<uint16_t>(d)), attr_start_offset(attr_start), attr_count(attr_cnt) {}
};

class XMLStreamReader {
public:
    // 事件回调函数类型定义（零拷贝优化，保持原有接口）
    using StartElementCallback = std::function<void(std::string_view name, span<const XMLAttribute> attributes, int depth)>;
    using EndElementCallback = std::function<void(std::string_view name, int depth)>;
    using TextCallback = std::function<void(std::string_view text, int depth)>;
    using CommentCallback = std::function<void(std::string_view comment, int depth)>;
    using ProcessingInstructionCallback = std::function<void(std::string_view target, std::string_view data, int depth)>;
    using ErrorCallback = std::function<void(XMLParseError error, const std::string& message, int line, int column)>;

private:
    // libexpat解析器句柄
    XML_Parser parser_ = nullptr;
    
    // 解析状态
    bool is_parsing_ = false;
    int current_depth_ = 0;
    XMLParseError last_error_ = XMLParseError::Ok;
    std::string last_error_message_;
    
    // 元素栈优化：使用轻量级元素和固定容量
    std::vector<XMLElementSlim> element_stack_slim_;
    static constexpr size_t MAX_DEPTH = 256;  // 最大嵌套深度
    
    // 属性缓存池（避免重复分配）
    std::vector<XMLAttribute> attribute_pool_;
    
    // 当前文本内容累积（优化版本）
    std::string current_text_;
    bool collecting_text_ = false;
    
    // 文本缓冲区预留（减少realloc）
    static constexpr size_t TEXT_RESERVE_SIZE = 256;
    bool text_reserved_ = false;
    
    // 缓冲区设置
    static constexpr size_t BUFFER_SIZE = fastexcel::core::Constants::kIOBufferSize;
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
    span<const XMLAttribute> parseAttributes(const XML_Char** attrs);
    std::string trimString(const std::string& str) const;
    std::string_view trimStringView(std::string_view str) const;
    void handleError(XMLParseError error, const std::string& message);
    void ensureTextReserve();
    
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
        SimpleElement* findChildByPath(const std::string& path) const;  // 支持路径查找如"child/grandchild"
        
        // 属性操作
        std::string getAttribute(const std::string& attr_name, const std::string& defaultValue = "") const;
        bool hasAttribute(const std::string& attr_name) const;
        void setAttribute(const std::string& attr_name, const std::string& value);
        void removeAttribute(const std::string& attr_name);
        
        // 文本内容
        std::string getTextContent() const;
        void setTextContent(const std::string& content);
        std::string getInnerText() const;  // 递归获取所有子节点文本
        
        // 子元素操作
        SimpleElement* appendChild(const std::string& element_name);
        SimpleElement* prependChild(const std::string& element_name);
        bool removeChild(SimpleElement* child);
        void clear();  // 清除所有子元素
        
        // 遍历辅助方法
        void forEach(std::function<void(SimpleElement*)> callback);
        void forEachRecursive(std::function<void(SimpleElement*, int)> callback, int depth = 0);
        
        // 查询方法
        size_t getChildCount() const { return children.size(); }
        bool hasChildren() const { return !children.empty(); }
        bool isEmpty() const { return text.empty() && children.empty(); }
        int getDepth() const;  // 获取元素在树中的深度
        
        // 输出方法
        std::string toString(int indent = 0) const;  // 美化输出XML字符串
        void print(int indent = 0) const;  // 打印元素结构
    };
    
    // 简单DOM解析（将整个文档解析为树结构）
    std::unique_ptr<SimpleElement> parseToDOM(const std::string& xml_content);
    std::unique_ptr<SimpleElement> parseFileToDOM(const std::string& filename);
};

}} // namespace fastexcel::xml
