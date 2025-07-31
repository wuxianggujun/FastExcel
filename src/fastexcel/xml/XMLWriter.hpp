#pragma once

#include <string>
#include <vector>
#include <stack>
#include <memory>
#include <cstdio>

// 项目日志头文件
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace xml {

/**
 * @brief 高性能流式XML写入器，基于libxlsxwriter的设计模式
 * 
 * 这个类参考了libxlsxwriter的高效XML写入机制，提供了流式处理、缓冲区优化
 * 和高效的字符转义功能。
 * 
 * 主要特性：
 * - 使用预分配缓冲区减少内存分配次数
 * - 高效的字符转义算法
 * - 流式处理支持，适合大文件生成
 * - 直接文件写入支持
 */
class XMLWriter {
private:
    // 使用预分配的缓冲区而不是字符串流，提高性能
    std::vector<char> buffer_;
    size_t buffer_pos_ = 0;
    
    // 元素栈用于跟踪嵌套的XML元素
    std::stack<std::string> element_stack_;
    bool in_element_ = false;
    
    // 预定义的XML转义字符串，避免重复创建
    static constexpr char AMP_ESCAPE[] = {'&', 'a', 'm', 'p', ';', '\0'};
    static constexpr char LT_ESCAPE[] = {'&', 'l', 't', ';', '\0'};
    static constexpr char GT_ESCAPE[] = {'&', 'g', 't', ';', '\0'};
    static constexpr char QUOT_ESCAPE[] = {'&', 'q', 'u', 'o', 't', ';', '\0'};
    static constexpr char APOS_ESCAPE[] = {'&', 'a', 'p', 'o', 's', ';', '\0'};
    static constexpr char NL_ESCAPE[] = {'&', '#', 'x', 'A', ';', '\0'};
    
    // 缓冲区初始大小和增长策略
    static constexpr size_t INITIAL_BUFFER_SIZE = 8192;
    static constexpr size_t BUFFER_GROWTH_FACTOR = 2;
    static constexpr size_t MAX_ATTRIBUTE_LENGTH = 2048;
    
public:
    XMLWriter();
    ~XMLWriter();
    
    // 禁用拷贝构造和赋值
    XMLWriter(const XMLWriter&) = delete;
    XMLWriter& operator=(const XMLWriter&) = delete;
    
    // 文档操作
    void startDocument();
    void endDocument();
    
    // 元素操作
    void startElement(const std::string& name);
    void endElement();
    void writeEmptyElement(const std::string& name);
    
    // 属性和文本
    void writeAttribute(const std::string& name, const std::string& value);
    void writeAttribute(const std::string& name, int value);
    void writeAttribute(const std::string& name, double value);
    void writeText(const std::string& text);
    
    // 获取结果
    std::string toString() const;
    void clear();
    
    // 流式写入支持
    void writeToBuffer(const char* data, size_t length);
    void ensureCapacity(size_t required);
    
    // 直接写入到文件（流式处理）
    bool writeToFile(const std::string& filename);
    bool setOutputFile(FILE* file);
    
private:
    // 高效的字符串转义方法，参考libxlsxwriter的实现
    std::string escapeAttributes(const std::string& text) const;
    std::string escapeData(const std::string& text) const;
    
    // 直接写入到缓冲区的方法，避免中间字符串拷贝
    void writeRaw(const char* data, size_t length);
    void writeRaw(const std::string& str);
    void writeChar(char c);
    
    // 优化的缓冲区管理
    void resizeBuffer(size_t new_size);
    void resetBuffer();
    
    // 文件输出支持
    FILE* output_file_ = nullptr;
    bool owns_file_ = false;
    
    // 快速检查是否需要转义
    bool needsAttributeEscaping(const std::string& text) const;
    bool needsDataEscaping(const std::string& text) const;
    
    // 高效的数字格式化
    std::string formatInt(int value) const;
    std::string formatDouble(double value) const;
};

}} // namespace fastexcel::xml