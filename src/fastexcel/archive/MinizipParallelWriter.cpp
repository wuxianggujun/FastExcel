#include "MinizipParallelWriter.hpp"
#include <iostream>
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

namespace fastexcel {
namespace archive {

MinizipParallelWriter::MinizipParallelWriter(size_t thread_count)
    : thread_pool_(std::make_unique<utils::ThreadPool>(thread_count == 0 ? std::thread::hardware_concurrency() : thread_count)),
      stats_{},
      completed_tasks_(0),
      failed_tasks_(0),
      total_compression_time_ms_(0.0),
      total_uncompressed_size_(0),
      total_compressed_size_(0),
      start_time_(std::chrono::high_resolution_clock::now()) {
    stats_.thread_count = thread_count == 0 ? std::thread::hardware_concurrency() : thread_count;
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
        
        // 并行压缩所有文件
        std::vector<std::future<CompressedFile>> futures;
        futures.reserve(files.size());
        
        for (const auto& file : files) {
            futures.emplace_back(
                thread_pool_->enqueue([this, file, compression_level]() {
                    return compressFile(file.first, file.second, compression_level);
                })
            );
        }
        
        // 收集压缩结果
        std::vector<CompressedFile> compressed_files;
        compressed_files.reserve(files.size());
        
        for (auto& future : futures) {
            auto result = future.get();
            if (result.success) {
                compressed_files.push_back(std::move(result));
                stats_.completed_tasks++;
            } else {
                std::cerr << "Failed to compress file: " << result.filename << std::endl;
                return false;
            }
        }
        
        // 写入最终ZIP文件
        bool success = writeCompressedFilesToZip(zip_filename, compressed_files);
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        stats_.total_compression_time_ms = static_cast<double>(duration.count());
        
        // 计算压缩比和并行效率
        calculateStatistics(files, compressed_files);
        
        return success;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception in compressAndWrite: " << e.what() << std::endl;
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
        // 简化实现：直接存储未压缩的数据
        // 实际压缩将在 writeCompressedFilesToZip 中进行
        result.compressed_data.assign(content.begin(), content.end());
        result.compressed_size = content.size();
        result.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
        result.success = true;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception in compressFile for " << filename << ": " << e.what() << std::endl;
    }
    
    return result;
}

bool MinizipParallelWriter::writeCompressedFilesToZip(
    const std::string& zip_filename,
    const std::vector<CompressedFile>& compressed_files) {
    
    try {
        // 创建ZIP写入器
        void* zip_writer = mz_zip_writer_create();
        if (!zip_writer) {
            return false;
        }
        
        // 打开ZIP文件进行写入
        int32_t err = mz_zip_writer_open_file(zip_writer, zip_filename.c_str(), 0, 0);
        if (err != MZ_OK) {
            mz_zip_writer_delete(&zip_writer);
            return false;
        }
        
        // 写入所有文件
        for (const auto& file : compressed_files) {
            // 设置文件信息
            mz_zip_file file_info = {};
            
            // 初始化必要的字段
            file_info.version_madeby = MZ_VERSION_MADEBY;
            file_info.version_needed = 20; // ZIP 2.0 specification
            file_info.flag = MZ_ZIP_FLAG_DATA_DESCRIPTOR; // 允许事后写 CRC
            file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
            file_info.modified_date = time(nullptr);
            file_info.accessed_date = file_info.modified_date;
            file_info.creation_date = file_info.modified_date;
            file_info.crc = 0; // 将由minizip-ng计算
            file_info.compressed_size = 0; // 将由minizip-ng计算
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
            
            // 打开文件条目
            err = mz_zip_writer_entry_open(zip_writer, &file_info);
            if (err != MZ_OK) {
                std::cerr << "Failed to open entry for " << file.filename << " (error: " << err << ")" << std::endl;
                mz_zip_writer_close(zip_writer);
                mz_zip_writer_delete(&zip_writer);
                return false;
            }
            
            // 写入文件内容（让minizip-ng进行压缩）
            err = mz_zip_writer_entry_write(zip_writer,
                                          file.compressed_data.data(),
                                          static_cast<int32_t>(file.compressed_data.size()));
            if (err < 0) {
                std::cerr << "Failed to write data for " << file.filename << " (error: " << err << ")" << std::endl;
                mz_zip_writer_entry_close(zip_writer);
                mz_zip_writer_close(zip_writer);
                mz_zip_writer_delete(&zip_writer);
                return false;
            }
            
            // 关闭文件条目
            err = mz_zip_writer_entry_close(zip_writer);
            if (err != MZ_OK) {
                std::cerr << "Failed to close entry for " << file.filename << " (error: " << err << ")" << std::endl;
                mz_zip_writer_close(zip_writer);
                mz_zip_writer_delete(&zip_writer);
                return false;
            }
        }
        
        // 关闭ZIP文件
        err = mz_zip_writer_close(zip_writer);
        mz_zip_writer_delete(&zip_writer);
        
        return err == MZ_OK;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception in writeCompressedFilesToZip: " << e.what() << std::endl;
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
    
    // 计算理论上的并行效率（简化计算）
    // 实际应用中可能需要更复杂的计算
    if (stats_.thread_count > 1) {
        stats_.parallel_efficiency = std::min(1.0, 
            static_cast<double>(compressed_files.size()) / stats_.thread_count);
    } else {
        stats_.parallel_efficiency = 1.0;
    }
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

void MinizipParallelWriter::waitForAllTasks() {
    // ThreadPool 的析构函数会等待所有任务完成
    // 这里可以添加额外的等待逻辑如果需要
}

} // namespace archive
} // namespace fastexcel