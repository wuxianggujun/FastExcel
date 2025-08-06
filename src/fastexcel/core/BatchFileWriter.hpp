#pragma once

#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include <vector>
#include <string>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 批量文件写入器 - 策略模式实现
 * 
 * 将所有文件内容收集到内存中，最后一次性批量写入ZIP文件。
 * 适用于小到中等规模的Excel文件，可以获得更好的压缩比。
 */
class BatchFileWriter : public IFileWriter {
private:
    std::vector<std::pair<std::string, std::string>> files_;
    archive::FileManager* file_manager_;
    
    // 流式写入的临时状态
    std::string current_path_;
    std::string current_content_;
    bool streaming_file_open_ = false;
    
    // 统计信息
    mutable WriteStats stats_;
    
public:
    /**
     * @brief 构造函数
     * @param file_manager 文件管理器指针
     */
    explicit BatchFileWriter(archive::FileManager* file_manager);
    
    /**
     * @brief 析构函数
     */
    ~BatchFileWriter() override;
    
    // IFileWriter 接口实现
    bool writeFile(const std::string& path, const std::string& content) override;
    bool openStreamingFile(const std::string& path) override;
    bool writeStreamingChunk(const char* data, size_t size) override;
    bool closeStreamingFile() override;
    std::string getTypeName() const override { return "BatchFileWriter"; }
    WriteStats getStats() const override { return stats_; }
    
    /**
     * @brief 批量写入所有收集的文件
     * @return 是否成功
     */
    bool flush();
    
    /**
     * @brief 获取当前收集的文件数量
     * @return 文件数量
     */
    size_t getFileCount() const { return files_.size(); }
    
    /**
     * @brief 获取预估的内存使用量
     * @return 内存使用量（字节）
     */
    size_t getEstimatedMemoryUsage() const;
    
    /**
     * @brief 清空所有收集的文件
     */
    void clear();
    
    /**
     * @brief 预分配文件容器空间
     * @param expected_files 预期文件数量
     */
    void reserve(size_t expected_files);
};

}} // namespace fastexcel::core