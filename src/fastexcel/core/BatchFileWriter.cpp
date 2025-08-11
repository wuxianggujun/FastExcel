#include "fastexcel/utils/ModuleLoggers.hpp"
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
        CORE_WARN("BatchFileWriter destroyed with open streaming file: {}", current_path_);
        closeStreamingFile();
    }
}

bool BatchFileWriter::writeFile(const std::string& path, const std::string& content) {
    // 🔧 关键修复：如果有流式文件打开，先关闭它
    if (streaming_file_open_) {
        CORE_WARN("Auto-closing streaming file {} to write batch file {}", current_path_, path);
        if (!closeStreamingFile()) {
            CORE_ERROR("Failed to close streaming file before writing batch file");
            return false;
        }
    }
    
    files_.emplace_back(path, content);
    stats_.batch_files++;
    stats_.total_bytes += content.size();
    
    CORE_DEBUG("Collected file for batch write: {} ({} bytes)", path, content.size());
    return true;
}

bool BatchFileWriter::openStreamingFile(const std::string& path) {
    if (streaming_file_open_) {
        CORE_ERROR("Streaming file already open: {}", current_path_);
        return false;
    }
    
    current_path_ = path;
    current_content_.clear();
    streaming_file_open_ = true;
    
    CORE_DEBUG("Opened streaming file for batch collection: {}", path);
    return true;
}

bool BatchFileWriter::writeStreamingChunk(const char* data, size_t size) {
    if (!streaming_file_open_) {
        CORE_ERROR("No streaming file is open");
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
        CORE_ERROR("No streaming file is open");
        return false;
    }
    
    // 🔧 关键修复：避免递归调用，直接添加到文件列表而不调用writeFile
    files_.emplace_back(current_path_, current_content_);
    stats_.streaming_files++;
    stats_.batch_files++; // 统计中也算作批量文件
    stats_.total_bytes += current_content_.size();
    
    CORE_DEBUG("Closed streaming file and added to batch: {} ({} bytes)",
             current_path_, current_content_.size());
    
    // 清理状态
    current_path_.clear();
    current_content_.clear();
    streaming_file_open_ = false;
    
    return true;
}

bool BatchFileWriter::flush() {
    if (streaming_file_open_) {
        CORE_WARN("Flushing with open streaming file, closing it first: {}", current_path_);
        closeStreamingFile();
    }
    
    if (files_.empty()) {
        CORE_DEBUG("No files to flush in batch mode");
        return true;
    }
    
    CORE_INFO("Flushing {} files in batch mode (total: {} bytes)", 
             files_.size(), stats_.total_bytes);
    
    // 使用移动语义提高性能
    bool success = file_manager_->writeFiles(std::move(files_));
    
    if (success) {
        stats_.files_written = files_.size();
        CORE_INFO("Successfully flushed {} files in batch mode", stats_.files_written);
    } else {
        CORE_ERROR("Failed to flush files in batch mode");
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
        CORE_WARN("Clearing with open streaming file: {}", current_path_);
        streaming_file_open_ = false;
    }
    
    files_.clear();
    current_path_.clear();
    current_content_.clear();
    stats_ = WriteStats{};
    
    CORE_DEBUG("Cleared all collected files in batch writer");
}

void BatchFileWriter::reserve(size_t expected_files) {
    files_.reserve(expected_files);
    CORE_DEBUG("Reserved space for {} files in batch writer", expected_files);
}

}} // namespace fastexcel::core