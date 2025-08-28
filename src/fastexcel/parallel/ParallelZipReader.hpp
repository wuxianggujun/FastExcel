#pragma once

#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/core/ThreadPool.hpp"
#include <unordered_map>
#include <atomic>
#include <future>
#include <shared_mutex>
#include <queue>
#include <condition_variable>
#include <thread>
#include <chrono>

namespace fastexcel {
namespace parallel {

/**
 * @brief 并行ZIP读取器 - 优化大文件读取性能
 * 
 * 特性：
 * - 多线程并行解压
 * - 智能预取缓存
 * - 批量处理优化
 * - 内存映射支持
 */
class ParallelZipReader {
public:
    struct Config {
        size_t thread_count = 4;
        size_t prefetch_size = 10 * 1024 * 1024;  // 10MB
        bool enable_cache = true;
        size_t cache_size_limit = 100 * 1024 * 1024; // 100MB
    };
    
    /**
     * @brief 文件提取任务
     */
    struct ExtractionTask {
        std::string path;
        std::promise<std::vector<uint8_t>> promise;
        std::future<std::vector<uint8_t>> future;
        
        ExtractionTask(const std::string& p) : path(p) {
            future = promise.get_future();
        }
    };
    
    ParallelZipReader(const core::Path& zip_path, const Config& config = Config());
    ~ParallelZipReader();
    

    
    /**
     * @brief 异步提取文件
     * @param path 文件路径
     * @return future对象
     */
    std::future<std::vector<uint8_t>> extractFileAsync(const std::string& path);
    
    /**
     * @brief 预取文件到缓存
     * @param paths 要预取的文件路径
     */
    void prefetchFiles(const std::vector<std::string>& paths);
    
    /**
     * @brief 流式并行处理
     * @param paths 文件路径列表
     * @param processor 处理函数
     */
    void processFilesInParallel(
        const std::vector<std::string>& paths,
        std::function<void(const std::string&, const std::vector<uint8_t>&)> processor
    );
    
    /**
     * @brief 获取缓存统计信息
     */
    struct CacheStats {
        size_t hit_count;
        size_t miss_count;
        size_t cache_size;
        double hit_rate;
    };
    CacheStats getCacheStats() const;
    
    /**
     * @brief 清空缓存
     */
    void clearCache();
    
private:
    // ZIP读取器池（每个线程一个）
    class ReaderPool {
    public:
        ReaderPool(const core::Path& path, size_t pool_size);
        ~ReaderPool();
        
        std::shared_ptr<archive::ZipReader> acquire();
        void release(std::shared_ptr<archive::ZipReader> reader);
        
    private:
        core::Path zip_path_;
        std::vector<std::shared_ptr<archive::ZipReader>> readers_;
        std::queue<std::shared_ptr<archive::ZipReader>> available_;
        std::mutex mutex_;
        std::condition_variable cv_;
    };
    
    // 缓存管理
    class Cache {
    public:
        Cache(size_t size_limit);
        
        bool get(const std::string& key, std::vector<uint8_t>& value);
        void put(const std::string& key, const std::vector<uint8_t>& value);
        void clear();
        size_t size() const;
        
    private:
        struct CacheEntry {
            std::vector<uint8_t> data;
            std::chrono::steady_clock::time_point last_access;
        };
        
        std::unordered_map<std::string, CacheEntry> cache_;
        size_t size_limit_;
        size_t current_size_;
        mutable std::shared_mutex mutex_;
        
        void evictLRU();
    };
    
    // 工作线程函数
    void workerThread();
    
    // 提取单个文件（内部使用）
    std::vector<uint8_t> extractFileInternal(const std::string& path);
    
private:
    Config config_;
    core::Path zip_path_;
    std::unique_ptr<core::ThreadPool> thread_pool_;
    std::unique_ptr<ReaderPool> reader_pool_;
    std::unique_ptr<Cache> cache_;
    
    // 任务队列
    std::queue<std::shared_ptr<ExtractionTask>> task_queue_;
    std::mutex queue_mutex_;
    std::condition_variable queue_cv_;
    
    // 统计信息
    mutable std::atomic<size_t> cache_hits_{0};
    mutable std::atomic<size_t> cache_misses_{0};
    
    // 控制标志
    std::atomic<bool> stop_flag_{false};
    std::vector<std::thread> worker_threads_;
};

/**
 * @brief 工作表并行加载器
 */
class ParallelWorksheetLoader {
public:
    struct WorksheetData {
        std::string name;
        std::string path;
        std::vector<uint8_t> content;
        size_t row_count;
        size_t col_count;
    };
    
    /**
     * @brief 并行加载所有工作表
     * @param zip_reader ZIP读取器
     * @param worksheet_paths 工作表路径列表
     * @return 工作表数据列表
     */
    static std::vector<WorksheetData> loadWorksheetsParallel(
        ParallelZipReader& zip_reader,
        const std::vector<std::pair<std::string, std::string>>& worksheet_paths
    );
    
    /**
     * @brief 使用管道模式处理工作表
     * 
     * 管道阶段：
     * 1. 解压 -> 2. 预解析 -> 3. 数据提取 -> 4. 存储
     */
    static void processWorksheetsPipeline(
        ParallelZipReader& zip_reader,
        const std::vector<std::string>& worksheet_paths,
        std::function<void(const WorksheetData&)> processor
    );
};

}} // namespace fastexcel::parallel
