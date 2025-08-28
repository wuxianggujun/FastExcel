#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <stdexcept>
#include <iostream>
#include <sstream>
#include <thread>

namespace fastexcel {
namespace archive {

// 构造/析构

ZipArchive::ZipArchive(const core::Path& path) 
    : filepath_(path)
    , reader_(nullptr)
    , writer_(nullptr)
    , parallel_reader_(nullptr)
    , thread_pool_(nullptr)
    , parallel_config_()
    , is_open_(false)
    , mode_(Mode::None) {
}

ZipArchive::~ZipArchive() {
    // 确保关闭文件
    if (is_open_) {
        close();
    }
}

// 文件操作

bool ZipArchive::open(bool create) {
    // 如果已经打开，先关闭
    if (is_open_) {
        close();
    }
    
    try {
        if (create) {
            // 创建新文件 - 写模式
            writer_ = std::make_unique<ZipWriter>(filepath_);
            if (writer_->open()) {
                mode_ = Mode::Write;
                is_open_ = true;
                return true;
            }
        } else {
            // 打开现有文件 - 读模式
            reader_ = std::make_unique<ZipReader>(filepath_);
            if (reader_->open()) {
                mode_ = Mode::Read;
                is_open_ = true;
                
                // 初始化并行读取器
                if (initializeParallelReader() != ZipError::Ok) {
                    FASTEXCEL_LOG_WARN("[ARCH] Failed to initialize parallel reader, continuing with sequential reading");
                }
                
                return true;
            }
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to open ZIP archive: {}", e.what());
    }
    
    // 打开失败，清理
    reader_.reset();
    writer_.reset();
    parallel_reader_.reset();
    thread_pool_.reset();
    mode_ = Mode::None;
    is_open_ = false;
    return false;
}

bool ZipArchive::close() {
    if (!is_open_) {
        return true;  // 已经关闭
    }
    
    bool success = true;
    
    // 关闭并行组件
    parallel_reader_.reset();
    thread_pool_.reset();
    
    // 关闭reader
    if (reader_) {
        success = reader_->close() && success;
        reader_.reset();
    }
    
    // 关闭writer
    if (writer_) {
        success = writer_->close() && success;
        writer_.reset();
    }
    
    mode_ = Mode::None;
    is_open_ = false;
    
    return success;
}

// 写入操作

ZipError ZipArchive::addFile(std::string_view internal_path, std::string_view content) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFile 返回 ZipError，直接返回
    return writer_->addFile(internal_path, content);
}

ZipError ZipArchive::addFile(std::string_view internal_path, const uint8_t* data, size_t size) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFile 返回 ZipError，直接返回
    return writer_->addFile(internal_path, data, size);
}

ZipError ZipArchive::addFile(std::string_view internal_path, const void* data, size_t size) {
    return addFile(internal_path, static_cast<const uint8_t*>(data), size);
}

ZipError ZipArchive::addFiles(const std::vector<FileEntry>& files) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFiles 返回 ZipError，直接返回
    return writer_->addFiles(files);
}

ZipError ZipArchive::addFiles(std::vector<FileEntry>&& files) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFiles 返回 ZipError，直接返回
    return writer_->addFiles(std::move(files));
}

ZipError ZipArchive::openEntry(std::string_view internal_path) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::openEntry 返回 ZipError，直接返回
    return writer_->openEntry(internal_path);
}

ZipError ZipArchive::writeChunk(const void* data, size_t size) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::writeChunk 返回 ZipError，直接返回
    return writer_->writeChunk(data, size);
}

ZipError ZipArchive::closeEntry() {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::closeEntry 返回 ZipError，直接返回
    return writer_->closeEntry();
}

// 读取操作

ZipError ZipArchive::extractFile(std::string_view internal_path, std::string& content) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFile 返回 ZipError，直接返回
    return reader_->extractFile(internal_path, content);
}

ZipError ZipArchive::extractFile(std::string_view internal_path, std::vector<uint8_t>& data) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFile 返回 ZipError，直接返回
    return reader_->extractFile(internal_path, data);
}

ZipError ZipArchive::extractFileToStream(std::string_view internal_path, std::ostream& output) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFileToStream 返回 ZipError，直接返回
    return reader_->extractFileToStream(internal_path, output);
}

ZipError ZipArchive::fileExists(std::string_view internal_path) const {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::fileExists 返回 ZipError，直接返回  
    return reader_->fileExists(internal_path);
}

std::vector<std::string> ZipArchive::listFiles() const {
    if (!isReadable()) {
        return {};
    }
    
    return reader_->listFiles();
}

// 配置

ZipError ZipArchive::setCompressionLevel(int level) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // 验证压缩级别范围
    if (level < 0 || level > 9) {
        return ZipError::InvalidParameter;
    }
    
    writer_->setCompressionLevel(level);
    return ZipError::Ok;
}

// 并行配置方法

ZipError ZipArchive::setParallelConfig(const ParallelConfig& config) {
    parallel_config_ = config;
    
    // 如果并行读取器已经初始化，重新创建它
    if (parallel_reader_ && isReadable()) {
        return initializeParallelReader();
    }
    
    return ZipError::Ok;
}

// 并行读取方法实现



std::future<std::vector<uint8_t>> ZipArchive::extractFileAsync(std::string_view internal_path) {
    if (!isParallelReadingAvailable()) {
        // 回退到单线程异步模式
        return std::async(std::launch::async, [this, path = std::string(internal_path)]() {
            std::vector<uint8_t> data;
            extractFile(path, data);
            return data;
        });
    }
    
    return parallel_reader_->extractFileAsync(std::string(internal_path));
}

std::future<void> ZipArchive::processFilesParallel(
    const std::vector<std::string>& paths,
    std::function<void(const std::string&, const std::vector<uint8_t>&)> processor,
    size_t chunk_size) {
    
    if (!isParallelReadingAvailable()) {
        // 回退到单线程处理
        return std::async(std::launch::async, [this, paths, processor]() {
            for (const auto& path : paths) {
                std::vector<uint8_t> data;
                if (extractFile(path, data) == ZipError::Ok) {
                    processor(path, data);
                }
            }
        });
    }
    
    return std::async(std::launch::async, [this, paths, processor, chunk_size]() {
        parallel_reader_->processFilesInParallel(paths, processor);
    });
}

std::future<void> ZipArchive::streamProcessFilesParallel(
    const std::vector<std::string>& paths,
    std::function<void(const std::string&, std::istream&)> processor,
    size_t max_concurrent) {
    
    return std::async(std::launch::async, [this, paths, processor, max_concurrent]() {
        // 创建工作线程池
        if (!thread_pool_) {
            thread_pool_ = std::make_unique<core::ThreadPool>(max_concurrent);
        }
        
        std::vector<std::future<void>> futures;
        
        for (const auto& path : paths) {
            auto future = thread_pool_->enqueue([this, path, processor]() {
                std::vector<uint8_t> data;
                if (extractFile(path, data) == ZipError::Ok) {
                    std::stringstream ss(std::string(data.begin(), data.end()));
                    processor(path, ss);
                }
            });
            
            futures.push_back(std::move(future));
        }
        
        // 等待所有任务完成
        for (auto& future : futures) {
            future.wait();
        }
    });
}

ZipError ZipArchive::prefetchFiles(const std::vector<std::string>& paths) {
    if (!isParallelReadingAvailable()) {
        return ZipError::NotOpen;
    }
    
    parallel_reader_->prefetchFiles(paths);
    return ZipError::Ok;
}

// 私有辅助方法

ZipError ZipArchive::initializeParallelReader() {
    try {
        // 创建并行配置
        parallel::ParallelZipReader::Config config;
        config.thread_count = parallel_config_.thread_count;
        config.prefetch_size = parallel_config_.prefetch_size;
        config.enable_cache = parallel_config_.enable_cache;
        config.cache_size_limit = parallel_config_.cache_size_limit;
        
        // 创建并行读取器
        parallel_reader_ = std::make_unique<parallel::ParallelZipReader>(filepath_, config);
        
        // 创建线程池
        if (!thread_pool_) {
            thread_pool_ = std::make_unique<core::ThreadPool>(parallel_config_.thread_count);
        }
        
        return ZipError::Ok;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to initialize parallel reader: {}", e.what());
        parallel_reader_.reset();
        return ZipError::InternalError;
    }
}

}} // namespace fastexcel::archive
