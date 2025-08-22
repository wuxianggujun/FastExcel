#include "fastexcel/utils/Logger.hpp"
#pragma once

#include <string>
#include <vector>
#include <stack>
#include <memory>
#include <cstdio>
#include <cstring>
#include <functional>

// 项目日志头文件
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Constants.hpp"
// 统一的 XML 转义常量
#include "fastexcel/xml/XMLEscapes.hpp"
// 需要使用 Path 进行跨平台文件打开
#include "fastexcel/core/Path.hpp"

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
 * - 支持回调函数的真正流式写入
 */
class XMLStreamWriter {
public:
    // 数据写入回调函数类型
    using WriteCallback = std::function<void(const std::string& chunk)>;
private:
    // 统一使用 fastexcel::xml::XMLEscapes 中的常量
    
    // 固定大小缓冲区，避免动态内存分配
    static constexpr size_t BUFFER_SIZE = fastexcel::core::Constants::kIOBufferSize;
    char buffer_[BUFFER_SIZE];
    size_t buffer_pos_ = 0;
    
    // 元素栈用于跟踪嵌套的XML元素
    std::stack<std::string> element_stack_;
    bool in_element_ = false;
    
    // 输出模式支持
    FILE* output_file_ = nullptr;
    bool owns_file_ = false;
    bool direct_file_mode_ = false;  // 直接文件写入模式
    bool callback_mode_ = false;     // 回调模式
    WriteCallback write_callback_;   // 写入回调函数
    bool auto_flush_ = true;         // 自动刷新
    
    // 移除缓冲模式的字符串累积，专注于极致性能
    // std::string whole_;  // 不再支持缓冲模式的字符串拼接
    
    // 属性批处理
    struct XMLAttribute {
        std::string key;
        std::string value;
        
        XMLAttribute(const std::string& k, const std::string& v) 
            : key(k), value(v) {}
    };
    
    std::vector<XMLAttribute> pending_attributes_;
    
    // 内部方法
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
    explicit XMLStreamWriter(const std::function<void(const char*, size_t)>& callback);
    explicit XMLStreamWriter(const std::string& filename);
    ~XMLStreamWriter();
    
    // 禁用拷贝构造和赋值
    XMLStreamWriter(const XMLStreamWriter&) = delete;
    XMLStreamWriter& operator=(const XMLStreamWriter&) = delete;
    
    // 模式设置 - 专注于高性能模式
    void setDirectFileMode(FILE* file, bool take_ownership = false);
    void setCallbackMode(WriteCallback callback, bool auto_flush = true);
    
    // 文档操作
    void startDocument();
    void endDocument();
    
    // 元素操作
    void startElement(const char* name);
    void endElement();
    void writeEmptyElement(const char* name);
    
    // 属性操作
    void writeAttribute(const char* name, const char* value);
    void writeAttribute(const char* name, const std::string& value);
    void writeAttribute(const char* name, int value);
    void writeAttribute(const char* name, double value);
    
    // 文本操作
    void writeText(const char* text);
    void writeText(const std::string& text);
    void writeRaw(const char* data);
    void writeRaw(const std::string& data);
    
    // 清理方法
    void clear();
    
    // toString()方法已彻底删除 - 专注极致性能
    
    // 文件操作
    bool writeToFile(const std::string& filename);
    bool setOutputFile(FILE* file, bool take_ownership = false);
    
    // 批处理属性
    void startAttributeBatch();
    void endAttributeBatch();
    
    // 缓冲区管理
    void flushBuffer();
};

}} // namespace fastexcel::xml
