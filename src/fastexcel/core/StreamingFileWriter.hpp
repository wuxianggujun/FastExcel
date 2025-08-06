#pragma once

#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 流式文件写入器 - 策略模式实现
 * 
 * 直接将文件内容流式写入ZIP文件，不在内存中缓存。
 * 适用于大规模Excel文件，可以保持常量内存使用。
 */
class StreamingFileWriter : public IFileWriter {
private:
    archive::FileManager* file_manager_;
    bool streaming_file_open_ = false;
    std::string current_streaming_path_;
    
    // 统计信息
    mutable WriteStats stats_;
    
public:
    /**
     * @brief 构造函数
     * @param file_manager 文件管理器指针
     */
    explicit StreamingFileWriter(archive::FileManager* file_manager);
    
    /**
     * @brief 析构函数
     */
    ~StreamingFileWriter() override;
    
    // IFileWriter 接口实现
    bool writeFile(const std::string& path, const std::string& content) override;
    bool openStreamingFile(const std::string& path) override;
    bool writeStreamingChunk(const char* data, size_t size) override;
    bool closeStreamingFile() override;
    std::string getTypeName() const override { return "StreamingFileWriter"; }
    WriteStats getStats() const override { return stats_; }
    
    /**
     * @brief 检查是否有流式文件正在写入
     * @return 是否有流式文件打开
     */
    bool hasOpenStreamingFile() const { return streaming_file_open_; }
    
    /**
     * @brief 获取当前流式文件路径
     * @return 文件路径（如果没有打开的文件则返回空字符串）
     */
    const std::string& getCurrentStreamingPath() const { return current_streaming_path_; }
    
    /**
     * @brief 强制关闭当前流式文件（用于错误恢复）
     * @return 是否成功
     */
    bool forceCloseStreamingFile();
};

}} // namespace fastexcel::core