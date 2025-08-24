/**
 * @file XMLStreamWriter.hpp  
 * @brief 内存安全优化的XML流写入器
 * 
 * 主要改进：
 * - 使用RAII管理文件资源
 * - 安全的缓冲区管理，防止溢出
 * - 异常安全的构造和析构
 * - 更好的错误处理和恢复
 */

#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <stack>
#include <memory>
#include <functional>

// 优化的工具类
#include "fastexcel/utils/FileWrapper.hpp"
#include "fastexcel/utils/SafeBuffer.hpp"
#include "fastexcel/utils/Logger.hpp"

// 原有依赖
#include "fastexcel/core/Constants.hpp"
#include "fastexcel/xml/XMLEscapes.hpp"
#include "fastexcel/core/Path.hpp"

namespace fastexcel {
namespace xml {

/**
 * @brief 内存安全的高性能XML流写入器
 * 
 * 主要特性：
 * - RAII文件资源管理，防止资源泄漏
 * - 安全缓冲区，防止缓冲区溢出
 * - 异常安全的构造和析构
 * - 支持多种输出模式：文件、回调、内存
 */
class XMLStreamWriter {
public:
    using WriteCallback = std::function<void(const std::string& chunk)>;
    
    /**
     * @brief 输出模式
     */
    enum class OutputMode {
        FILE_DIRECT,    // 直接文件输出
        CALLBACK,       // 回调函数输出
        MEMORY_BUFFER   // 内存缓冲输出
    };

private:
    static constexpr size_t DEFAULT_BUFFER_SIZE = fastexcel::core::Constants::kIOBufferSize;
    
    // 安全缓冲区管理
    utils::SafeBuffer<DEFAULT_BUFFER_SIZE> buffer_;
    
    // 文件资源管理（使用RAII）
    std::unique_ptr<utils::FileWrapper> file_wrapper_;
    
    // 输出模式和回调
    OutputMode output_mode_ = OutputMode::MEMORY_BUFFER;
    WriteCallback write_callback_;
    
    // XML元素栈管理
    std::stack<std::string> element_stack_;
    bool in_element_ = false;
    
    // 属性批处理
    struct XMLAttribute {
        std::string key;
        std::string value;
        
        XMLAttribute(std::string k, std::string v) 
            : key(std::move(k)), value(std::move(v)) {}
    };
    std::vector<XMLAttribute> pending_attributes_;
    
    // 内存缓冲模式的输出
    std::string memory_buffer_;
    
    // 性能统计
    mutable size_t bytes_written_ = 0;
    mutable size_t flush_count_ = 0;
    
    // 内部方法
    void initializeBuffer();
    void flushToOutput(const char* data, size_t length);
    void writeAttributesToBuffer();
    void ensureElementClosed();
    
    // 转义方法
    void writeEscapedAttribute(const std::string& value);
    void writeEscapedText(const std::string& text);
    
public:
    /**
     * @brief 默认构造函数（内存缓冲模式）
     */
    XMLStreamWriter();
    
    /**
     * @brief 文件输出构造函数
     * @param filename 输出文件名
     * @throws FileException 文件创建失败
     */
    explicit XMLStreamWriter(const std::string& filename);
    
    /**
     * @brief 回调输出构造函数  
     * @param callback 输出回调函数
     * @throws ParameterException 回调为空
     */
    explicit XMLStreamWriter(WriteCallback callback);
    
    /**
     * @brief 析构函数，确保资源正确释放
     */
    ~XMLStreamWriter();
    
    // 禁用拷贝，允许移动
    XMLStreamWriter(const XMLStreamWriter&) = delete;
    XMLStreamWriter& operator=(const XMLStreamWriter&) = delete;
    
    XMLStreamWriter(XMLStreamWriter&& other) noexcept;
    XMLStreamWriter& operator=(XMLStreamWriter&& other) noexcept;
    
    /**
     * @brief 模式切换方法
     */
    void switchToFileMode(const std::string& filename);
    void switchToCallbackMode(WriteCallback callback);
    void switchToMemoryMode();
    
    /**
     * @brief 文档操作
     */
    void startDocument(const std::string& encoding = "UTF-8");
    void endDocument();
    
    /**
     * @brief 元素操作
     */
    void startElement(const std::string& name);
    void endElement();
    void writeEmptyElement(const std::string& name);
    
    /**
     * @brief 属性操作
     */
    void writeAttribute(const std::string& name, const std::string& value);
    void writeAttribute(const std::string& name, std::string_view value);
    void writeAttribute(const std::string& name, int value);
    void writeAttribute(const std::string& name, double value);
    void writeAttribute(const std::string& name, bool value);
    // 直接写入 string_view（避免在调用处构造 std::string）
    // 注意：内部会在写入缓冲前处理转义，不会悬挂引用
    
    /**
     * @brief 文本内容操作
     */
    void writeText(const std::string& text);
    void writeText(int value);
    void writeText(size_t value);
    void writeText(double value);
    void writeRaw(const std::string& data);
    void writeCDATA(const std::string& data);
    void writeComment(const std::string& comment);
    
    /**
     * @brief 批处理操作
     */
    void startAttributeBatch();
    void endAttributeBatch();
    
    /**
     * @brief 缓冲区管理
     */
    void flush();
    void clear();
    
    /**
     * @brief 获取输出结果（仅内存模式）
     */
    std::string toString() const;
    
    /**
     * @brief 状态查询
     */
    OutputMode getOutputMode() const { return output_mode_; }
    size_t getBytesWritten() const { return bytes_written_; }
    size_t getFlushCount() const { return flush_count_; }
    bool isEmpty() const;
    
    /**
     * @brief 性能优化设置
     */
    void setAutoFlush(bool auto_flush);
    void reserveAttributeCapacity(size_t capacity);
    
private:
    // 异常安全的初始化方法
    void initializeInternal();
    void cleanupInternal() noexcept;
};

}} // namespace fastexcel::xml
