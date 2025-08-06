#pragma once

#include <string>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 统一的文件写入接口 - 策略模式
 * 
 * 抽象了批量模式和流式模式的文件写入操作，
 * 消除了重复代码，提高了代码的可维护性。
 */
class IFileWriter {
public:
    virtual ~IFileWriter() = default;
    
    /**
     * @brief 写入完整文件内容（批量模式）
     * @param path 文件路径
     * @param content 文件内容
     * @return 是否成功
     */
    virtual bool writeFile(const std::string& path, const std::string& content) = 0;
    
    /**
     * @brief 打开流式文件写入
     * @param path 文件路径
     * @return 是否成功
     */
    virtual bool openStreamingFile(const std::string& path) = 0;
    
    /**
     * @brief 写入流式数据块
     * @param data 数据指针
     * @param size 数据大小
     * @return 是否成功
     */
    virtual bool writeStreamingChunk(const char* data, size_t size) = 0;
    
    /**
     * @brief 关闭流式文件写入
     * @return 是否成功
     */
    virtual bool closeStreamingFile() = 0;
    
    /**
     * @brief 获取写入器类型名称（用于调试）
     * @return 类型名称
     */
    virtual std::string getTypeName() const = 0;
    
    /**
     * @brief 获取写入统计信息
     */
    struct WriteStats {
        size_t files_written = 0;
        size_t total_bytes = 0;
        size_t streaming_files = 0;
        size_t batch_files = 0;
    };
    
    virtual WriteStats getStats() const = 0;
};

}} // namespace fastexcel::core