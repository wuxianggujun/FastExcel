#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/core/Path.hpp"
#include <string>
#include <vector>
#include <memory>
#include <cstdint>
#include <string_view>
#include <mutex>
#include <unordered_set>

namespace fastexcel {
namespace archive {

// 错误码枚举（与ZipArchive共享）
enum class ZipError;

/**
 * @brief 高性能ZIP写入器 - 专注于写入操作
 * 
 * 特性：
 * - 线程安全
 * - 批量写入优化
 * - 流式写入支持
 * - 防重复写入
 * - 支持大文件
 * - 原始数据写入（避免重压缩）
 */
class ZipWriter {
public:
    // 文件条目结构
    struct FileEntry {
        std::string internal_path;
        std::string content;
        
        FileEntry() = default;
        
        // 移动构造函数
        FileEntry(std::string&& path, std::string&& data)
            : internal_path(std::move(path)), content(std::move(data)) {}
            
        // 拷贝构造函数
        FileEntry(const std::string& path, const std::string& data)
            : internal_path(path), content(data) {}
            
        // 混合构造函数
        FileEntry(std::string&& path, const std::string& data)
            : internal_path(std::move(path)), content(data) {}
            
        FileEntry(const std::string& path, std::string&& data)
            : internal_path(path), content(std::move(data)) {}
    };
    
    // 构造/析构
    explicit ZipWriter(const core::Path& path);
    ~ZipWriter();
    
    // 禁止拷贝
    ZipWriter(const ZipWriter&) = delete;
    ZipWriter& operator=(const ZipWriter&) = delete;
    
    // 支持移动
    ZipWriter(ZipWriter&& other) noexcept;
    ZipWriter& operator=(ZipWriter&& other) noexcept;
    
    // 文件操作
    
    /**
     * 创建/打开ZIP文件进行写入
     * @param create 是否创建新文件（true=创建新文件，false=追加到现有文件）
     * @return 是否成功
     */
    bool open(bool create = true);
    
    /**
     * 关闭ZIP文件
     * @return 是否成功
     */
    bool close();
    
    /**
     * 检查是否已打开
     */
    bool isOpen() const { return is_open_; }
    
    // 基本写入操作
    
    /**
     * 添加文件（字符串内容）
     * @param internal_path ZIP内部路径
     * @param content 文件内容
     * @return 错误码
     */
    ZipError addFile(std::string_view internal_path, std::string_view content);
    
    /**
     * 添加文件（二进制数据）
     * @param internal_path ZIP内部路径
     * @param data 数据指针
     * @param size 数据大小
     * @return 错误码
     */
    ZipError addFile(std::string_view internal_path, const void* data, size_t size);
    
    /**
     * 添加文件（字节数组）
     * @param internal_path ZIP内部路径
     * @param data 数据
     * @return 错误码
     */
    ZipError addFile(std::string_view internal_path, const uint8_t* data, size_t size);
    
    // 批量写入（性能优化）
    
    /**
     * 批量添加文件
     * @param files 文件条目列表
     * @return 错误码
     */
    ZipError addFiles(const std::vector<FileEntry>& files);
    
    /**
     * 批量添加文件（移动语义版本）
     * @param files 文件条目列表
     * @return 错误码
     */
    ZipError addFiles(std::vector<FileEntry>&& files);
    
    // 流式写入（大文件）
    
    /**
     * 开始流式写入条目
     * @param internal_path ZIP内部路径
     * @return 错误码
     */
    ZipError openEntry(std::string_view internal_path);
    
    /**
     * 写入数据块
     * @param data 数据指针
     * @param size 数据大小
     * @return 错误码
     */
    ZipError writeChunk(const void* data, size_t size);
    
    /**
     * 结束流式写入条目
     * @return 错误码
     */
    ZipError closeEntry();
    
    // 高级功能
    
    /**
     * 写入原始压缩数据（用于高效复制）
     * @param internal_path ZIP内部路径
     * @param raw_data 原始压缩数据
     * @param size 数据大小
     * @param uncompressed_size 未压缩大小
     * @param crc32 CRC32校验值
     * @param compression_method 压缩方法
     * @return 错误码
     * 
     * 注意：这个方法直接写入压缩数据，避免重复压缩
     */
    ZipError writeRawCompressedData(std::string_view internal_path,
                                    const void* raw_data, size_t size,
                                    size_t uncompressed_size,
                                    uint32_t crc32,
                                    int compression_method);
    
    /**
     * 设置压缩级别
     * @param level 压缩级别（0-9，0=无压缩，9=最高压缩）
     * @return 错误码
     */
    ZipError setCompressionLevel(int level);
    
    /**
     * 获取当前压缩级别
     */
    int getCompressionLevel() const { return compression_level_; }
    
    // 状态查询
    
    /**
     * 检查文件是否已写入（防重复）
     * @param internal_path ZIP内部路径
     * @return 是否已写入
     */
    bool hasEntry(const std::string& internal_path) const {
        return written_paths_.count(internal_path) > 0;
    }
    
    /**
     * 获取已写入的文件列表
     */
    std::vector<std::string> getWrittenPaths() const {
        return std::vector<std::string>(written_paths_.begin(), written_paths_.end());
    }
    
    /**
     * 获取统计信息
     */
    struct Stats {
        size_t entries_written = 0;  // 写入的条目数
        size_t bytes_written = 0;    // 写入的字节数
    };
    Stats getStats() const { return stats_; }
    
    /**
     * 获取文件路径
     */
    const core::Path& getPath() const { return filepath_; }
    
private:
    // 内部实现细节
    void* zip_handle_ = nullptr;
    core::Path filepath_;
    std::string filename_;  // UTF-8 filename for logging
    bool is_open_ = false;
    bool stream_entry_open_ = false;  // 流式写入条目是否已打开
    int compression_level_ = 6;  // 压缩级别，默认为6
    mutable std::mutex mutex_;  // 线程安全
    std::unordered_set<std::string> written_paths_;  // 跟踪已写入的路径，防止重复
    Stats stats_;
    
    // 内部辅助方法
    bool initializeWriter();
    void cleanup();
    void initializeFileInfo(void* file_info, const std::string& path, size_t size);
    ZipError writeFileEntry(const std::string& internal_path, const void* data, size_t size);
};

}} // namespace fastexcel::archive
