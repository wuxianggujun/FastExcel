#include "fastexcel/utils/Logger.hpp"
#include "MinizipParallelWriter.hpp"
#include <iostream>
#include <fmt/format.h>
#include <chrono>
#include <algorithm>
#include <cstring>

// Minizip-ng headers
#include "mz.h"
#include "mz_zip.h"
#include "mz_strm.h"
#include "mz_strm_mem.h"
#include "mz_zip_rw.h"
#include "mz_os.h"
#include "mz_strm_os.h"

// zlib-ng headers for direct compression
#include "mz_crypt.h"

#include <zlib.h>

namespace fastexcel {
namespace archive {

// 线程本地存储的压缩流和缓冲区
thread_local std::unique_ptr<z_stream> MinizipParallelWriter::compression_stream_;
thread_local int MinizipParallelWriter::current_compression_level_ = -1;
thread_local std::vector<uint8_t> MinizipParallelWriter::compression_buffer_;

MinizipParallelWriter::MinizipParallelWriter(size_t thread_count)
    : thread_pool_(std::make_unique<core::ThreadPool>(thread_count == 0 ? std::thread::hardware_concurrency() : thread_count)),
      stats_{},
      completed_tasks_(0),
      failed_tasks_(0),
      total_compression_time_ms_(0.0),
      total_uncompressed_size_(0),
      total_compressed_size_(0),
      start_time_(std::chrono::high_resolution_clock::now()) {
    stats_.thread_count = thread_count == 0 ? std::thread::hardware_concurrency() : thread_count;
    
    // 注意：硬件CRC32支持需要检查minizip-ng版本
    // 在某些版本中可能需要不同的API调用
    // mz_crypt_crc32_update_source(MZ_CRC32_AUTO); // 暂时注释掉，避免编译错误
}

MinizipParallelWriter::~MinizipParallelWriter() = default;

bool MinizipParallelWriter::compressAndWrite(
    const std::string& zip_filename,
    const std::vector<std::pair<std::string, std::string>>& files,
    int compression_level) {
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    try {
        // 重置统计信息
        stats_ = Statistics{};
        stats_.thread_count = thread_pool_->size();
        
        // 创建任务列表，包括文件分块
        std::vector<CompressionTask> tasks = createCompressionTasks(files);
        
        // 并行压缩所有任务
        std::vector<std::future<CompressedFile>> futures;
        futures.reserve(tasks.size());
        
        for (const auto& task : tasks) {
            futures.emplace_back(
                thread_pool_->enqueue([this, task, compression_level]() {
                    return compressFile(task.filename, task.content, compression_level);
                })
            );
        }
        
        // 收集压缩结果
        std::vector<CompressedFile> compressed_files;
        compressed_files.reserve(tasks.size());
        
        for (auto& future : futures) {
            auto result = future.get();
            if (result.success) {
                compressed_files.push_back(std::move(result));
                stats_.completed_tasks++;
            } else {
                FASTEXCEL_LOG_ERROR("Failed to compress file: {}", result.filename);
                return false;
            }
        }
        
        // 写入最终ZIP文件
        bool success = writeCompressedFilesToZip(zip_filename, compressed_files);
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        stats_.total_compression_time_ms = static_cast<double>(duration.count());
        
        // 计算压缩比和并行效率
        calculateStatisticsFromTasks(tasks, compressed_files);
        
        return success;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception in compressAndWrite: {}", e.what());
        return false;
    }
}

CompressedFile MinizipParallelWriter::compressFile(
    const std::string& filename,
    const std::string& content,
    int compression_level) {
    
    CompressedFile result;
    result.filename = filename;
    result.uncompressed_size = content.size();
    result.success = false;
    
    try {
        // 计算 CRC32 - 使用 minizip-ng 的 CRC32 函数
        result.crc32 = mz_crypt_crc32_update(0, reinterpret_cast<const uint8_t*>(content.data()), static_cast<int32_t>(content.size()));
        
        // 获取或创建重用的压缩流
        z_stream* strm = getOrCreateCompressionStream(compression_level);
        if (!strm) {
            result.error_message = "Failed to create compression stream";
            return result;
        }
        
        // 使用线程本地缓冲区池化 - 避免频繁的内存分配
        size_t max_compressed_size = content.size() + (content.size() >> 8) + 64;
        
        // 重用线程本地缓冲区，只在需要时扩容
        if (compression_buffer_.size() < max_compressed_size) {
            compression_buffer_.resize(max_compressed_size);
        }
        
        // 设置输入和输出
        strm->next_in = reinterpret_cast<Bytef*>(const_cast<char*>(content.data()));
        strm->avail_in = static_cast<uInt>(content.size());
        strm->next_out = compression_buffer_.data();
        strm->avail_out = static_cast<uInt>(max_compressed_size);
        
        // 执行压缩
        int ret = deflate(strm, Z_FINISH);
        if (ret != Z_STREAM_END) {
            result.error_message = "Failed to compress data";
            return result;
        }
        
        // 获取实际压缩大小并复制到结果
        result.compressed_size = strm->total_out;
        result.compressed_data.assign(compression_buffer_.begin(),
                                    compression_buffer_.begin() + result.compressed_size);
        
        result.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
        result.success = true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception in compressFile for {}: {}", filename, e.what());
        result.error_message = e.what();
    }
    
    return result;
}

bool MinizipParallelWriter::writeCompressedFilesToZip(
    const std::string& zip_filename,
    const std::vector<CompressedFile>& compressed_files) {
    
    try {
        // 创建ZIP句柄（不是writer）
        void* zip_handle = mz_zip_create();
        if (!zip_handle) {
            return false;
        }
        
        // 创建文件流
        void* file_stream = mz_stream_os_create();
        if (!file_stream) {
            mz_zip_delete(&zip_handle);
            return false;
        }
        
        // 打开文件流
        int32_t err = mz_stream_os_open(file_stream, zip_filename.c_str(), MZ_OPEN_MODE_WRITE | MZ_OPEN_MODE_CREATE);
        if (err != MZ_OK) {
            mz_stream_os_delete(&file_stream);
            mz_zip_delete(&zip_handle);
            return false;
        }
        
        // 打开ZIP进行写入
        err = mz_zip_open(zip_handle, file_stream, MZ_OPEN_MODE_WRITE);
        if (err != MZ_OK) {
            mz_stream_os_close(file_stream);
            mz_stream_os_delete(&file_stream);
            mz_zip_delete(&zip_handle);
            return false;
        }
        
        // 写入所有文件
        for (const auto& file : compressed_files) {
            // 设置文件信息
            mz_zip_file file_info = {};
            
            // 初始化必要的字段
            file_info.version_madeby = MZ_VERSION_MADEBY;
            file_info.version_needed = 20; // ZIP 2.0 specification
            file_info.flag = 0; // 不使用 DATA_DESCRIPTOR
            file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE; // 标记为DEFLATE压缩
            file_info.modified_date = time(nullptr);
            file_info.accessed_date = file_info.modified_date;
            file_info.creation_date = file_info.modified_date;
            
            // 关键：填入已经计算好的 CRC、压缩大小和原始大小
            file_info.crc = file.crc32;
            file_info.compressed_size = file.compressed_size;
            file_info.uncompressed_size = file.uncompressed_size;
            
            file_info.filename_size = static_cast<uint16_t>(file.filename.length());
            file_info.extrafield_size = 0;
            file_info.comment_size = 0;
            file_info.disk_number = 0;
            file_info.disk_offset = 0;
            file_info.internal_fa = 0;
            file_info.external_fa = 0;
            file_info.filename = file.filename.c_str();
            file_info.extrafield = nullptr;
            file_info.comment = nullptr;
            file_info.linkname = nullptr;
            file_info.zip64 = 0;
            file_info.aes_version = 0;
            file_info.aes_strength = 0;
            file_info.pk_verify = 0;
            
            // 使用 mz_zip_entry_* API 进行 raw 写入
            err = mz_zip_entry_write_open(zip_handle, &file_info, 0, 1, nullptr); // raw=1
            if (err != MZ_OK) {
                FASTEXCEL_LOG_ERROR("Failed to open entry for {} (error: {})", file.filename, err);
                mz_zip_close(zip_handle);
                mz_stream_os_close(file_stream);
                mz_stream_os_delete(&file_stream);
                mz_zip_delete(&zip_handle);
                return false;
            }
            
            // 直接写入已压缩的 deflate 流
            err = mz_zip_entry_write(zip_handle,
                                   file.compressed_data.data(),
                                   static_cast<int32_t>(file.compressed_size));
            if (err < 0) {
                FASTEXCEL_LOG_ERROR("Failed to write data for {} (error: {})", file.filename, err);
                mz_zip_entry_close(zip_handle);
                mz_zip_close(zip_handle);
                mz_stream_os_close(file_stream);
                mz_stream_os_delete(&file_stream);
                mz_zip_delete(&zip_handle);
                return false;
            }
            
            // 关闭 raw 条目
            err = mz_zip_entry_close_raw(zip_handle, file.uncompressed_size, file.crc32);
            if (err != MZ_OK) {
                FASTEXCEL_LOG_ERROR("Failed to close raw entry for {} (error: {})", file.filename, err);
                mz_zip_close(zip_handle);
                mz_stream_os_close(file_stream);
                mz_stream_os_delete(&file_stream);
                mz_zip_delete(&zip_handle);
                return false;
            }
        }
        
        // 关闭ZIP文件
        err = mz_zip_close(zip_handle);
        mz_stream_os_close(file_stream);
        mz_stream_os_delete(&file_stream);
        mz_zip_delete(&zip_handle);
        
        return err == MZ_OK;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception in writeCompressedFilesToZip: {}", e.what());
        return false;
    }
}

void MinizipParallelWriter::calculateStatistics(
    const std::vector<std::pair<std::string, std::string>>& original_files,
    const std::vector<CompressedFile>& compressed_files) {
    
    size_t total_uncompressed = 0;
    size_t total_compressed = 0;
    
    for (const auto& file : original_files) {
        total_uncompressed += file.second.size();
    }
    
    for (const auto& file : compressed_files) {
        total_compressed += file.compressed_size;
    }
    
    if (total_uncompressed > 0) {
        stats_.compression_ratio = static_cast<double>(total_compressed) / total_uncompressed;
    }
    
    // 计算并行效率 - 基于实际任务分布
    if (stats_.thread_count > 1) {
        // 计算任务分布效率
        size_t tasks_per_thread = compressed_files.size() / stats_.thread_count;
        size_t remaining_tasks = compressed_files.size() % stats_.thread_count;
        
        // 如果任务数少于线程数，效率会下降
        if (compressed_files.size() < stats_.thread_count) {
            stats_.parallel_efficiency = static_cast<double>(compressed_files.size()) / static_cast<double>(stats_.thread_count);
        } else {
            // 考虑负载均衡：如果有剩余任务，部分线程会多处理一个任务
            double load_balance_factor = 1.0;
            if (remaining_tasks > 0) {
                // 计算负载不均衡的影响
                double max_load = static_cast<double>(tasks_per_thread + 1);
                double avg_load = static_cast<double>(compressed_files.size()) / static_cast<double>(stats_.thread_count);
                load_balance_factor = avg_load / max_load;
            }
            
            // 基础并行效率（假设理想情况下 85% 效率）
            stats_.parallel_efficiency = 0.85 * load_balance_factor;
        }
    } else {
        stats_.parallel_efficiency = 1.0;
    }
    
    // 转换为百分比
    stats_.parallel_efficiency *= 100.0;
}

MinizipParallelWriter::Statistics MinizipParallelWriter::getStatistics() const {
    return stats_;
}

void MinizipParallelWriter::resetStatistics() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = Statistics{};
    completed_tasks_ = 0;
    failed_tasks_ = 0;
    total_compression_time_ms_ = 0.0;
    total_uncompressed_size_ = 0;
    total_compressed_size_ = 0;
    start_time_ = std::chrono::high_resolution_clock::now();
}

std::future<CompressedFile> MinizipParallelWriter::compressFileAsync(
    const std::string& filename,
    const std::string& content,
    int compression_level) {
    
    return thread_pool_->enqueue([this, filename, content, compression_level]() {
        return compressFile(filename, content, compression_level);
    });
}

std::vector<std::future<CompressedFile>> MinizipParallelWriter::compressFilesAsync(
    const std::vector<std::pair<std::string, std::string>>& files,
    int compression_level) {
    
    std::vector<std::future<CompressedFile>> futures;
    futures.reserve(files.size());
    
    for (const auto& file : files) {
        futures.emplace_back(compressFileAsync(file.first, file.second, compression_level));
    }
    
    return futures;
}

std::vector<CompressionTask> MinizipParallelWriter::createCompressionTasks(
    const std::vector<std::pair<std::string, std::string>>& files) {
    
    std::vector<CompressionTask> tasks;
    
    for (const auto& [filename, content] : files) {
        // 如果文件大于阈值，进行分块
        if (content.size() > LARGE_FILE_THRESHOLD) {
            size_t chunks = (content.size() + CHUNK_SIZE - 1) / CHUNK_SIZE;
            
            for (size_t i = 0; i < chunks; ++i) {
                size_t start = i * CHUNK_SIZE;
                size_t size = std::min(CHUNK_SIZE, content.size() - start);
                
                std::string chunk_filename = filename;
                if (chunks > 1) {
                    // 为分块文件添加后缀
                    size_t dot_pos = filename.find_last_of('.');
                    if (dot_pos != std::string::npos) {
                        chunk_filename = filename.substr(0, dot_pos) +
                                       "_part" + std::to_string(i) +
                                       filename.substr(dot_pos);
                    } else {
                        chunk_filename = filename + "_part" + std::to_string(i);
                    }
                }
                
                tasks.emplace_back(chunk_filename, content.substr(start, size));
            }
        } else {
            // 小文件直接添加
            tasks.emplace_back(filename, content);
        }
    }
    
    return tasks;
}

z_stream* MinizipParallelWriter::getOrCreateCompressionStream(int compression_level) {
    if (!compression_stream_ || current_compression_level_ != compression_level) {
        // 清理旧流
        if (compression_stream_) {
            deflateEnd(compression_stream_.get());
        }
        
        compression_stream_ = std::make_unique<z_stream>();
        memset(compression_stream_.get(), 0, sizeof(z_stream));
        
        int ret = deflateInit2(compression_stream_.get(), compression_level,
                              Z_DEFLATED, -15, 9, Z_DEFAULT_STRATEGY);
        if (ret != Z_OK) {
            compression_stream_.reset();
            return nullptr;
        }
        
        current_compression_level_ = compression_level;
    } else {
        // 重用现有流，只需重置 - 这是关键优化！
        // 避免重复的 deflateInit2/deflateEnd 调用
        int ret = deflateReset(compression_stream_.get());
        if (ret != Z_OK) {
            // 重置失败，尝试重新初始化
            deflateEnd(compression_stream_.get());
            memset(compression_stream_.get(), 0, sizeof(z_stream));
            
            ret = deflateInit2(compression_stream_.get(), compression_level,
                              Z_DEFLATED, -15, 9, Z_DEFAULT_STRATEGY);
            if (ret != Z_OK) {
                compression_stream_.reset();
                return nullptr;
            }
        }
    }
    
    return compression_stream_.get();
}

void MinizipParallelWriter::calculateStatisticsFromTasks(
    const std::vector<CompressionTask>& tasks,
    const std::vector<CompressedFile>& compressed_files) {
    
    size_t total_uncompressed = 0;
    size_t total_compressed = 0;
    
    for (const auto& task : tasks) {
        total_uncompressed += task.content.size();
    }
    
    for (const auto& file : compressed_files) {
        total_compressed += file.compressed_size;
    }
    
    if (total_uncompressed > 0) {
        stats_.compression_ratio = static_cast<double>(total_compressed) / total_uncompressed;
    }
    
    // 计算并行效率 - 按数据大小加权而非简单任务数量
    if (stats_.thread_count > 1) {
        if (tasks.size() < stats_.thread_count) {
            // 任务数少于线程数，效率受限
            stats_.parallel_efficiency = static_cast<double>(tasks.size()) / static_cast<double>(stats_.thread_count);
        } else {
            // 按数据大小加权计算负载均衡
            std::vector<size_t> task_sizes;
            task_sizes.reserve(tasks.size());
            for (const auto& task : tasks) {
                task_sizes.push_back(task.content.size());
            }
            
            // 模拟任务分配到线程的负载均衡
            std::vector<size_t> thread_loads(stats_.thread_count, 0);
            
            // 简单的贪心分配：每个任务分配给当前负载最小的线程
            for (size_t task_size : task_sizes) {
                auto min_it = std::min_element(thread_loads.begin(), thread_loads.end());
                *min_it += task_size;
            }
            
            // 计算负载均衡系数
            size_t max_load = *std::max_element(thread_loads.begin(), thread_loads.end());
            size_t min_load = *std::min_element(thread_loads.begin(), thread_loads.end());
            double avg_load = static_cast<double>(total_uncompressed) / stats_.thread_count;
            
            double load_balance_factor = 1.0;
            if (max_load > 0) {
                load_balance_factor = avg_load / max_load;
            }
            
            // 考虑任务粒度和负载均衡的综合效率
            double task_granularity_factor = std::min(1.0, static_cast<double>(tasks.size()) / (stats_.thread_count * 2.0));
            stats_.parallel_efficiency = 0.85 * task_granularity_factor * load_balance_factor;
        }
    } else {
        stats_.parallel_efficiency = 1.0;
    }
    
    // 转换为百分比
    stats_.parallel_efficiency *= 100.0;
}

void MinizipParallelWriter::waitForAllTasks() {
    // ThreadPool 的析构函数会等待所有任务完成
    // 这里可以添加额外的等待逻辑如果需要
}

} // namespace archive
} // namespace fastexcel
