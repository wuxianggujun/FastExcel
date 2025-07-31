#pragma once

#include <string>
#include <vector>
#include <stack>
#include <memory>
#include <cstdio>
#include <cstring>

// 项目日志头文件
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

/**
 * @brief 高性能流式XML写入器，基于libxlsxwriter的设计模式
 * 
 * 这个类参考了libxlsxwriter的高效XML写入机制，提供了接近其性能的流式处理。
 * 主要优化：
 * - 使用固定大小缓冲区，减少动态内存分配
 * - 直接文件写入模式，避免内存拷贝
 * - 高效的字符转义算法，使用memcpy和预定义长度
 * - 属性批处理机制
 */
class XMLStreamWriter {
private:
    // 预定义的XML转义字符串和长度，参考libxlsxwriter
    static constexpr char AMP_ESCAPE[] = "&";
    static constexpr size_t AMP_ESCAPE_LEN = 1;
    
    static constexpr char LT_ESCAPE[] = "<";
    static constexpr size_t LT_ESCAPE_LEN = 1;
    
    static constexpr char GT_ESCAPE[] = ">";
    static constexpr size_t GT_ESCAPE_LEN = 1;
    
    static constexpr char QUOT_ESCAPE[] = "\"";
    static constexpr size_t QUOT_ESCAPE_LEN = 1;
    
    static constexpr char APOS_ESCAPE[] = "'";
    static constexpr size_t APOS_ESCAPE_LEN = 1;
    
    static constexpr char NL_ESCAPE[] = "\n";
    static constexpr size_t NL_ESCAPE_LEN = 1;
    
    // 转义后的字符串
    static constexpr char AMP_REPLACEMENT[] = "&";
    static constexpr size_t AMP_REPLACEMENT_LEN = 5;
    
    static constexpr char LT_REPLACEMENT[] = "<";
    static constexpr size_t LT_REPLACEMENT_LEN = 4;
    
    static constexpr char GT_REPLACEMENT[] = ">";
    static constexpr size_t GT_REPLACEMENT_LEN = 4;
    
    static constexpr char QUOT_REPLACEMENT[] = "\"";
    static constexpr size_t QUOT_REPLACEMENT_LEN = 6;
    
    static constexpr char APOS_REPLACEMENT[] = "'";
    static constexpr size_t APOS_REPLACEMENT_LEN = 6;
    
    static constexpr char NL_REPLACEMENT[] = "&#xA;";
    static constexpr size_t NL_REPLACEMENT_LEN = 5;
    
    // 固定大小缓冲区，避免动态内存分配
    static constexpr size_t BUFFER_SIZE = 8192;
    char buffer_[BUFFER_SIZE];
    size_t buffer_pos_ = 0;
    
    // 元素栈用于跟踪嵌套的XML元素
    std::stack<std::string> element_stack_;
    bool in_element_ = false;
    
    // 文件输出支持
    FILE* output_file_ = nullptr;
    bool owns_file_ = false;
    bool direct_file_mode_ = false;  // 直接文件写入模式
    
    // 属性批处理
    struct XMLAttribute {
        std::string key;
        std::string value;
        
        XMLAttribute(const std::string& k, const std::string& v) 
            : key(k), value(v) {}
    };
    
    std::vector<XMLAttribute> pending_attributes_;
    
    // 内部方法
    void flushBuffer();
    void writeRawToBuffer(const char* data, size_t length);
    void writeRawToFile(const char* data, size_t length);
    void writeRawDirect(const char* data, size_t length);
    
    // 高效转义方法
    void escapeAttributesToBuffer(const char* text, size_t length);
    void escapeDataToBuffer(const char* text, size_t length);
    void escapeAttributesToFile(const char* text, size_t length);
    void escapeDataToFile(const char* text, size_t length);
    
    // 快速检查是否需要转义
    bool needsAttributeEscaping(const char* text, size_t length) const;
    bool needsDataEscaping(const char* text, size_t length) const;
    
    // 高效的数字格式化
    void formatIntToBuffer(int value);
    void formatDoubleToBuffer(double value);
    
public:
    XMLStreamWriter();
    ~XMLStreamWriter();
    
    // 禁用拷贝构造和赋值
    XMLStreamWriter(const XMLStreamWriter&) = delete;
    XMLStreamWriter& operator=(const XMLStreamWriter&) = delete;
    
    // 模式设置
    void setDirectFileMode(FILE* file, bool take_ownership = false);
    void setBufferedMode();
    
    // 文档操作
    void startDocument();
    void endDocument();
    
    // 元素操作
    void startElement(const char* name);
    void endElement();
    void writeEmptyElement(const char* name);
    
    // 属性操作
    void writeAttribute(const char* name, const char* value);
    void writeAttribute(const char* name, int value);
    void writeAttribute(const char* name, double value);
    
    // 文本操作
    void writeText(const char* text);
    
    // 便捷方法，支持std::string
    void startElement(const std::string& name) { startElement(name.c_str()); }
    void writeEmptyElement(const std::string& name) { writeEmptyElement(name.c_str()); }
    void writeAttribute(const std::string& name, const std::string& value) { writeAttribute(name.c_str(), value.c_str()); }
    void writeAttribute(const std::string& name, int value) { writeAttribute(name.c_str(), value); }
    void writeAttribute(const std::string& name, double value) { writeAttribute(name.c_str(), value); }
    void writeText(const std::string& text) { writeText(text.c_str()); }
    
    // 获取结果（仅在缓冲模式下有效）
    std::string toString() const;
    void clear();
    
    // 文件操作
    bool writeToFile(const std::string& filename);
    bool setOutputFile(FILE* file, bool take_ownership = false);
    
    // 批处理属性
    void startAttributeBatch();
    void endAttributeBatch();
};

}} // namespace fastexcel::xml
