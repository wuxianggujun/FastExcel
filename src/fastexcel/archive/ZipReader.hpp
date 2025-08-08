#pragma once

#include "fastexcel/core/Path.hpp"
#include <string>
#include <vector>
#include <memory>
#include <string_view>
#include <mutex>
#include <ostream>
#include <functional>
#include <unordered_map>
#include <ostream>

namespace fastexcel {
namespace archive {

// 错误码枚举（与ZipArchive共享）
enum class ZipError;

/**
 * @brief 高性能ZIP读取器 - 专注于读取操作
 * 
 * 特性：
 * - 线程安全
 * - 条目信息缓存
 * - 流式读取支持
 * - 支持大文件
 * - 原始数据访问（用于高效复制）
 */
class ZipReader {
public:
    // ========== 条目信息结构 ==========
    struct EntryInfo {
        std::string path;
        uint64_t compressed_size;
        uint64_t uncompressed_size;
        uint32_t crc32;
        int compression_method;
        time_t modified_date;
        time_t creation_date;
        uint16_t flag;
        bool is_directory;
    };
    
    // ========== 构造/析构 ==========
    explicit ZipReader(const core::Path& path);
    ~ZipReader();
    
    // 禁止拷贝
    ZipReader(const ZipReader&) = delete;
    ZipReader& operator=(const ZipReader&) = delete;
    
    // 支持移动
    ZipReader(ZipReader&& other) noexcept;
    ZipReader& operator=(ZipReader&& other) noexcept;
    
    // ========== 文件操作 ==========
    
    /**
     * 打开ZIP文件进行读取
     * @return 是否成功
     */
    bool open();
    
    /**
     * 关闭ZIP文件
     * @return 是否成功
     */
    bool close();
    
    /**
     * 检查是否已打开
     */
    bool isOpen() const { return is_open_; }
    
    // ========== 条目查询 ==========
    
    /**
     * 获取所有文件列表
     * @return 文件路径列表
     */
    std::vector<std::string> listFiles() const;
    
    /**
     * 获取所有条目的详细信息
     * @return 条目信息列表
     */
    std::vector<EntryInfo> listEntriesInfo() const;
    
    /**
     * 检查文件是否存在
     * @param internal_path ZIP内部路径
     * @return 错误码
     */
    ZipError fileExists(std::string_view internal_path) const;
    
    /**
     * 获取条目信息
     * @param internal_path ZIP内部路径
     * @param info 输出的条目信息
     * @return 是否成功
     */
    bool getEntryInfo(std::string_view internal_path, EntryInfo& info) const;
    
    // ========== 读取操作 ==========
    
    /**
     * 提取文件到字符串
     * @param internal_path ZIP内部路径
     * @param content 输出内容
     * @return 错误码
     */
    ZipError extractFile(std::string_view internal_path, std::string& content);
    
    /**
     * 提取文件到字节数组
     * @param internal_path ZIP内部路径
     * @param data 输出数据
     * @return 错误码
     */
    ZipError extractFile(std::string_view internal_path, std::vector<uint8_t>& data);
    
    /**
     * 流式提取文件到输出流
     * @param internal_path ZIP内部路径
     * @param output 输出流
     * @return 错误码
     */
    ZipError extractFileToStream(std::string_view internal_path, std::ostream& output);
    
    // ========== 高级功能 ==========
    
    /**
     * 获取原始压缩数据（用于高效复制）
     * @param internal_path ZIP内部路径
     * @param raw_data 输出的原始压缩数据
     * @param info 输出的条目信息
     * @return 是否成功
     * 
     * 注意：这个方法获取的是ZIP文件中的原始压缩数据，
     * 可以直接写入到另一个ZIP文件中，避免解压再压缩的开销
     */
    bool getRawCompressedData(std::string_view internal_path, 
                             std::vector<uint8_t>& raw_data,
                             EntryInfo& info) const;
    
    /**
     * 流式读取回调
     * @param internal_path ZIP内部路径
     * @param callback 数据回调函数(data, size) -> continue
     * @param buffer_size 缓冲区大小
     * @return 错误码
     */
    ZipError streamFile(std::string_view internal_path,
                        std::function<bool(const uint8_t*, size_t)> callback,
                        size_t buffer_size = 65536) const;
    
    // ========== 统计信息 ==========
    
    /**
     * 获取ZIP文件统计信息
     */
    struct Stats {
        size_t total_entries = 0;
        size_t total_compressed = 0;
        size_t total_uncompressed = 0;
        double compression_ratio = 0.0;
    };
    Stats getStats() const;
    
    /**
     * 获取文件路径
     */
    const core::Path& getPath() const { return filepath_; }
    
private:
    // 内部实现细节
    void* unzip_handle_ = nullptr;
    core::Path filepath_;
    std::string filename_;  // UTF-8 filename for logging
    bool is_open_ = false;
    mutable std::mutex mutex_;  // 线程安全
    
    // 条目信息缓存
    mutable std::unordered_map<std::string, EntryInfo> entry_cache_;
    mutable bool cache_initialized_ = false;
    
    // 内部辅助方法
    bool initializeReader();
    void cleanup();
    void buildEntryCache() const;
    bool locateEntry(std::string_view path) const;
    ZipError extractFileInternal(std::string_view internal_path, 
                                 std::vector<uint8_t>& data) const;
};

}} // namespace fastexcel::archive
