#pragma once

#include "fastexcel/utils/ThreadPool.hpp"
#include "mz.h"
#include "mz_zip.h"
#include "mz_strm_mem.h"
#include "zlib.h"  // 添加 zlib 头文件以支持 z_stream
#include <string>
#include <vector>
#include <memory>
#include <future>
#include <map>
#include <mutex>
#include <chrono>

namespace fastexcel {
namespace archive {

/**
 * @brief 压缩任务结构
 */
struct CompressionTask {
    std::string filename;
    std::string content;
    
    CompressionTask(std::string name, std::string data)
        : filename(std::move(name)), content(std::move(data)) {}
};

/**
 * @brief 基于minizip-ng的并行压缩文件结构
 */
struct CompressedFile {
    std::string filename;               // ZIP内的文件名
    std::vector<uint8_t> compressed_data; // 压缩后的数据
    uint32_t crc32;                    // CRC32校验和
    size_t uncompressed_size;          // 原始大小
    size_t compressed_size;            // 压缩后大小
    uint16_t compression_method;       // 压缩方法
    bool success;                      // 压缩是否成功
    std::string error_message;         // 错误信息
    
    CompressedFile() : crc32(0), uncompressed_size(0), compressed_size(0),
                      compression_method(MZ_COMPRESS_METHOD_DEFLATE), success(false) {}
};

/**
 * @brief 基于minizip-ng的并行ZIP写入器
 * 
 * 使用文件级并行策略：
 * 1. 每个线程独立压缩一个文件
 * 2. 使用minizip-ng的内存流进行压缩
 * 3. 主线程收集所有压缩结果
 * 4. 使用单个minizip-ng实例写入最终ZIP文件
 */
class MinizipParallelWriter {
public:
    /**
     * @brief 构造函数
     * @param thread_count 线程数量，0表示使用硬件并发数
     */
    explicit MinizipParallelWriter(size_t thread_count = 0);
    
    /**
     * @brief 析构函数
     */
    ~MinizipParallelWriter();
    
    /**
     * @brief 并行压缩文件并写入ZIP
     * @param zip_filename ZIP文件路径
     * @param files 文件列表 (filename, content)
     * @param compression_level 压缩级别 (0-9)
     * @return 是否成功
     */
    bool compressAndWrite(const std::string& zip_filename,
                         const std::vector<std::pair<std::string, std::string>>& files,
                         int compression_level = MZ_COMPRESS_LEVEL_FAST);
    
    /**
     * @brief 异步压缩单个文件
     * @param filename 文件名
     * @param content 文件内容
     * @param compression_level 压缩级别
     * @return 压缩结果的future
     */
    std::future<CompressedFile> compressFileAsync(const std::string& filename,
                                                 const std::string& content,
                                                 int compression_level = MZ_COMPRESS_LEVEL_FAST);
    
    /**
     * @brief 批量异步压缩文件
     * @param files 文件列表
     * @param compression_level 压缩级别
     * @return 所有压缩结果的future列表
     */
    std::vector<std::future<CompressedFile>> compressFilesAsync(
        const std::vector<std::pair<std::string, std::string>>& files,
        int compression_level = MZ_COMPRESS_LEVEL_FAST);
    
    /**
     * @brief 将压缩结果写入ZIP文件
     * @param zip_filename ZIP文件路径
     * @param compressed_files 压缩文件列表
     * @return 是否成功
     */
    bool writeCompressedFilesToZip(const std::string& zip_filename,
                                  const std::vector<CompressedFile>& compressed_files);
    
    /**
     * @brief 等待所有压缩任务完成
     */
    void waitForAllTasks();
    
    /**
     * @brief 获取统计信息
     */
    struct Statistics {
        size_t thread_count;
        size_t completed_tasks;
        size_t failed_tasks;
        double total_compression_time_ms;
        double total_uncompressed_size_mb;
        double total_compressed_size_mb;
        double compression_ratio;
        double parallel_efficiency; // 并行效率百分比
    };
    
    Statistics getStatistics() const;
    void resetStatistics();

private:
    std::unique_ptr<utils::ThreadPool> thread_pool_;
    
    // 统计信息
    mutable std::mutex stats_mutex_;
    Statistics stats_;
    size_t completed_tasks_;
    size_t failed_tasks_;
    double total_compression_time_ms_;
    size_t total_uncompressed_size_;
    size_t total_compressed_size_;
    std::chrono::high_resolution_clock::time_point start_time_;
    
    // 压缩流重用
    thread_local static std::unique_ptr<z_stream> compression_stream_;
    thread_local static int current_compression_level_;
    
    // 任务分块配置
    static constexpr size_t LARGE_FILE_THRESHOLD = 2 * 1024 * 1024; // 2MB
    static constexpr size_t CHUNK_SIZE = 512 * 1024; // 512KB per chunk
    
    /**
     * @brief 使用minizip-ng压缩单个文件
     * @param filename 文件名
     * @param content 文件内容
     * @param compression_level 压缩级别
     * @return 压缩结果
     */
    CompressedFile compressFile(const std::string& filename,
                               const std::string& content,
                               int compression_level);
    
    CompressedFile compressFileWithMinizip(const std::string& filename,
                                          const std::string& content,
                                          int compression_level);
    
    /**
     * @brief 创建minizip文件信息结构
     * @param filename 文件名
     * @param uncompressed_size 原始大小
     * @param compression_method 压缩方法
     * @return minizip文件信息
     */
    mz_zip_file createFileInfo(const std::string& filename,
                              size_t uncompressed_size,
                              uint16_t compression_method);
    
    /**
     * @brief 创建压缩任务列表，包括大文件分块
     * @param files 原始文件列表
     * @return 任务列表
     */
    std::vector<CompressionTask> createCompressionTasks(
        const std::vector<std::pair<std::string, std::string>>& files);
    
    /**
     * @brief 获取或创建线程本地的压缩流
     * @param compression_level 压缩级别
     * @return 压缩流引用
     */
    z_stream& getOrCreateCompressionStream(int compression_level);
    
    /**
     * @brief 计算统计信息（基于任务）
     * @param tasks 任务列表
     * @param compressed_files 压缩文件列表
     */
    void calculateStatisticsFromTasks(const std::vector<CompressionTask>& tasks,
                                    const std::vector<CompressedFile>& compressed_files);
    
    /**
     * @brief 计算统计信息（原版本兼容）
     * @param original_files 原始文件列表
     * @param compressed_files 压缩文件列表
     */
    void calculateStatistics(const std::vector<std::pair<std::string, std::string>>& original_files,
                           const std::vector<CompressedFile>& compressed_files);
    
    /**
     * @brief 更新统计信息
     * @param result 压缩结果
     * @param compression_time_ms 压缩耗时
     */
    void updateStatistics(const CompressedFile& result, double compression_time_ms);
    
    /**
     * @brief 验证压缩结果
     * @param result 压缩结果
     * @return 是否有效
     */
    bool validateCompressedFile(const CompressedFile& result) const;
};

}} // namespace fastexcel::archive