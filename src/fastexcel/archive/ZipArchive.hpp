#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipWriter.hpp"
#include "fastexcel/archive/ZipError.hpp"
#include "fastexcel/core/ThreadPool.hpp"
#include "fastexcel/parallel/ParallelZipReader.hpp"
#include <string>
#include <vector>
#include <memory>
#include <cstdint>
#include <string_view>
#include <future>
#include <functional>
#include <unordered_map>

namespace fastexcel {
namespace archive {

/**
 * @brief ZIP归档类 - 组合ZipReader和ZipWriter提供完整功能
 * 
 * 这个类组合了ZipReader和ZipWriter，提供了同时读写ZIP文件的能力。
 * 如果只需要读或写，建议直接使用ZipReader或ZipWriter以获得更好的性能。
 */
class ZipArchive {
public:
    // 使用ZipWriter的FileEntry
    using FileEntry = ZipWriter::FileEntry;
    
    // 构造/析构
    explicit ZipArchive(const core::Path& path);
    ~ZipArchive();
    
    // 文件操作
    
    /**
     * 打开ZIP文件
     * @param create true=创建新文件（写模式），false=打开现有文件（读模式）
     * @return 是否成功
     */
    bool open(bool create = true);
    
    /**
     * 关闭ZIP文件
     * @return 是否成功
     */
    bool close();
    
    // 写入操作
    
    ZipError addFile(std::string_view internal_path, std::string_view content);
    ZipError addFile(std::string_view internal_path, const uint8_t* data, size_t size);
    ZipError addFile(std::string_view internal_path, const void* data, size_t size);
    
    // 批量写入
    ZipError addFiles(const std::vector<FileEntry>& files);
    ZipError addFiles(std::vector<FileEntry>&& files);
    
    // 流式写入
    ZipError openEntry(std::string_view internal_path);
    ZipError writeChunk(const void* data, size_t size);
    ZipError closeEntry();
    
    // 读取操作
    
    ZipError extractFile(std::string_view internal_path, std::string& content);
    ZipError extractFile(std::string_view internal_path, std::vector<uint8_t>& data);
    ZipError extractFileToStream(std::string_view internal_path, std::ostream& output);
    ZipError fileExists(std::string_view internal_path) const;
    
    // 并行读取操作
    
    /**
     * @brief 并行提取多个文件
     * @param paths 文件路径列表
     * @return 文件内容映射 (路径 -> 数据)
     */
    std::future<std::unordered_map<std::string, std::vector<uint8_t>>>
    extractFilesParallel(const std::vector<std::string>& paths);
    
    /**
     * @brief 异步提取单个文件
     * @param internal_path 文件内部路径
     * @return 异步结果 future
     */
    std::future<std::vector<uint8_t>> extractFileAsync(std::string_view internal_path);
    
    /**
     * @brief 并行流式处理多个文件
     * @param paths 文件路径列表
     * @param processor 处理函数 (路径, 数据) -> void
     * @param chunk_size 每批处理的文件数量
     * @return 处理结果的future
     */
    std::future<void> processFilesParallel(
        const std::vector<std::string>& paths,
        std::function<void(const std::string&, const std::vector<uint8_t>&)> processor,
        size_t chunk_size = 4
    );
    
    /**
     * @brief 流式并行读取 - 边读取边处理，内存使用更高效
     * @param paths 文件路径列表
     * @param processor 流处理函数 (路径, 输入流) -> void
     * @param max_concurrent 最大并发数
     * @return 处理结果的future
     */
    std::future<void> streamProcessFilesParallel(
        const std::vector<std::string>& paths,
        std::function<void(const std::string&, std::istream&)> processor,
        size_t max_concurrent = 4
    );
    
    /**
     * @brief 预取文件到缓存（用于后续快速访问）
     * @param paths 要预取的文件路径
     */
    ZipError prefetchFiles(const std::vector<std::string>& paths);
    
    // 文件列表
    std::vector<std::string> listFiles() const;
    
    // 状态查询
    
    bool isOpen() const { return is_open_; }
    bool isWritable() const { return mode_ == Mode::Write || mode_ == Mode::ReadWrite; }
    bool isReadable() const { return mode_ == Mode::Read || mode_ == Mode::ReadWrite; }
    
    // 配置
    
    ZipError setCompressionLevel(int level);
    
    /**
     * @brief 并行配置结构
     */
    struct ParallelConfig {
        size_t thread_count = std::thread::hardware_concurrency();
        size_t cache_size_limit = 100 * 1024 * 1024;  // 100MB
        bool enable_cache = true;
        size_t prefetch_size = 10 * 1024 * 1024;      // 10MB
        size_t max_concurrent_streams = 8;             // 最大并发流数
    };
    
    /**
     * @brief 设置并行读取配置
     * @param config 并行配置
     */
    ZipError setParallelConfig(const ParallelConfig& config);
    
    /**
     * @brief 获取并行读取配置
     */
    const ParallelConfig& getParallelConfig() const { return parallel_config_; }
    
    // 直接访问底层对象
    
    /**
     * 获取ZipReader对象（如果可用）
     * @return ZipReader指针，如果不在读模式则返回nullptr
     */
    ZipReader* getReader() { return reader_.get(); }
    const ZipReader* getReader() const { return reader_.get(); }
    
    /**
     * 获取ZipWriter对象（如果可用）
     * @return ZipWriter指针，如果不在写模式则返回nullptr
     */
    ZipWriter* getWriter() { return writer_.get(); }
    const ZipWriter* getWriter() const { return writer_.get(); }
    
private:
    core::Path filepath_;
    std::unique_ptr<ZipReader> reader_;
    std::unique_ptr<ZipWriter> writer_;
    std::unique_ptr<parallel::ParallelZipReader> parallel_reader_;
    std::unique_ptr<core::ThreadPool> thread_pool_;
    ParallelConfig parallel_config_;
    bool is_open_ = false;
    enum class Mode { None, Read, Write, ReadWrite } mode_ = Mode::None;
    
    // 私有辅助方法
    
    /**
     * @brief 初始化并行读取器
     */
    ZipError initializeParallelReader();
    
    /**
     * @brief 检查并行读取是否可用
     */
    bool isParallelReadingAvailable() const {
        return parallel_reader_ != nullptr && isReadable();
    }
};

}} // namespace fastexcel::archive
