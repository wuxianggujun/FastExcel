#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace core {

BatchFileWriter::BatchFileWriter(archive::FileManager* file_manager)
    : file_manager_(file_manager) {
    if (!file_manager_) {
        throw std::invalid_argument("FileManager cannot be null");
    }
}

BatchFileWriter::~BatchFileWriter() {
    if (streaming_file_open_) {
        LOG_WARN("BatchFileWriter destroyed with open streaming file: {}", current_path_);
        closeStreamingFile();
    }
}

bool BatchFileWriter::writeFile(const std::string& path, const std::string& content) {
    if (streaming_file_open_) {
        LOG_ERROR("Cannot write file while streaming file is open: {}", current_path_);
        return false;
    }
    
    files_.emplace_back(path, content);
    stats_.batch_files++;
    stats_.total_bytes += content.size();
    
    LOG_DEBUG("Collected file for batch write: {} ({} bytes)", path, content.size());
    return true;
}

bool BatchFileWriter::openStreamingFile(const std::string& path) {
    if (streaming_file_open_) {
        LOG_ERROR("Streaming file already open: {}", current_path_);
        return false;
    }
    
    current_path_ = path;
    current_content_.clear();
    streaming_file_open_ = true;
    
    LOG_DEBUG("Opened streaming file for batch collection: {}", path);
    return true;
}

bool BatchFileWriter::writeStreamingChunk(const char* data, size_t size) {
    if (!streaming_file_open_) {
        LOG_ERROR("No streaming file is open");
        return false;
    }
    
    if (!data || size == 0) {
        return true; // 空数据块，忽略
    }
    
    current_content_.append(data, size);
    return true;
}

bool BatchFileWriter::closeStreamingFile() {
    if (!streaming_file_open_) {
        LOG_ERROR("No streaming file is open");
        return false;
    }
    
    // 将流式收集的内容添加到批量文件列表
    bool success = writeFile(current_path_, current_content_);
    
    if (success) {
        stats_.streaming_files++;
        LOG_DEBUG("Closed streaming file and added to batch: {} ({} bytes)", 
                 current_path_, current_content_.size());
    }
    
    // 清理状态
    current_path_.clear();
    current_content_.clear();
    streaming_file_open_ = false;
    
    return success;
}

bool BatchFileWriter::flush() {
    if (streaming_file_open_) {
        LOG_WARN("Flushing with open streaming file, closing it first: {}", current_path_);
        closeStreamingFile();
    }
    
    if (files_.empty()) {
        LOG_DEBUG("No files to flush in batch mode");
        return true;
    }
    
    LOG_INFO("Flushing {} files in batch mode (total: {} bytes)", 
             files_.size(), stats_.total_bytes);
    
    // 使用移动语义提高性能
    bool success = file_manager_->writeFiles(std::move(files_));
    
    if (success) {
        stats_.files_written = files_.size();
        LOG_INFO("Successfully flushed {} files in batch mode", stats_.files_written);
    } else {
        LOG_ERROR("Failed to flush files in batch mode");
    }
    
    // 清空文件列表（无论成功与否）
    files_.clear();
    
    return success;
}

size_t BatchFileWriter::getEstimatedMemoryUsage() const {
    size_t usage = sizeof(BatchFileWriter);
    
    // 文件列表内存
    usage += files_.capacity() * sizeof(std::pair<std::string, std::string>);
    for (const auto& [path, content] : files_) {
        usage += path.capacity() + content.capacity();
    }
    
    // 当前流式文件内存
    usage += current_path_.capacity() + current_content_.capacity();
    
    return usage;
}

void BatchFileWriter::clear() {
    if (streaming_file_open_) {
        LOG_WARN("Clearing with open streaming file: {}", current_path_);
        streaming_file_open_ = false;
    }
    
    files_.clear();
    current_path_.clear();
    current_content_.clear();
    stats_ = WriteStats{};
    
    LOG_DEBUG("Cleared all collected files in batch writer");
}

void BatchFileWriter::reserve(size_t expected_files) {
    files_.reserve(expected_files);
    LOG_DEBUG("Reserved space for {} files in batch writer", expected_files);
}

}} // namespace fastexcel::core