
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace core {

StreamingFileWriter::StreamingFileWriter(archive::FileManager* file_manager)
    : file_manager_(file_manager) {
    if (!file_manager_) {
        throw std::invalid_argument("FileManager cannot be null");
    }
}

StreamingFileWriter::~StreamingFileWriter() {
    if (streaming_file_open_) {
        FASTEXCEL_LOG_WARN("StreamingFileWriter destroyed with open streaming file: {}", current_streaming_path_);
        forceCloseStreamingFile();
    }
}

bool StreamingFileWriter::writeFile(const std::string& path, const std::string& content) {
    if (streaming_file_open_) {
        FASTEXCEL_LOG_ERROR("Cannot write file while streaming file is open: {}", current_streaming_path_);
        return false;
    }
    
    FASTEXCEL_LOG_DEBUG("Writing file directly in streaming mode: {} ({} bytes)", path, content.size());
    
    bool success = file_manager_->writeFile(path, content);
    
    if (success) {
        stats_.batch_files++;
        stats_.files_written++;
        stats_.total_bytes += content.size();
        FASTEXCEL_LOG_DEBUG("Successfully wrote file: {}", path);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to write file: {}", path);
    }
    
    return success;
}

bool StreamingFileWriter::openStreamingFile(const std::string& path) {
    if (streaming_file_open_) {
        FASTEXCEL_LOG_ERROR("Streaming file already open: {}", current_streaming_path_);
        return false;
    }
    
    bool success = file_manager_->openStreamingFile(path);
    
    if (success) {
        streaming_file_open_ = true;
        current_streaming_path_ = path;
        FASTEXCEL_LOG_DEBUG("Opened streaming file: {}", path);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to open streaming file: {}", path);
    }
    
    return success;
}

bool StreamingFileWriter::writeStreamingChunk(const char* data, size_t size) {
    if (!streaming_file_open_) {
        FASTEXCEL_LOG_ERROR("No streaming file is open");
        return false;
    }
    
    if (!data || size == 0) {
        return true; // 空数据块，忽略
    }
    
    bool success = file_manager_->writeStreamingChunk(data, size);
    
    if (success) {
        stats_.total_bytes += size;
    } else {
        FASTEXCEL_LOG_ERROR("Failed to write streaming chunk to file: {} ({} bytes)", 
                 current_streaming_path_, size);
    }
    
    return success;
}

bool StreamingFileWriter::closeStreamingFile() {
    if (!streaming_file_open_) {
        FASTEXCEL_LOG_ERROR("No streaming file is open");
        return false;
    }
    
    bool success = file_manager_->closeStreamingFile();
    
    if (success) {
        stats_.streaming_files++;
        stats_.files_written++;
        FASTEXCEL_LOG_DEBUG("Successfully closed streaming file: {}", current_streaming_path_);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to close streaming file: {}", current_streaming_path_);
    }
    
    // 清理状态（无论成功与否）
    streaming_file_open_ = false;
    current_streaming_path_.clear();
    
    return success;
}

bool StreamingFileWriter::forceCloseStreamingFile() {
    if (!streaming_file_open_) {
        return true; // 没有打开的文件
    }
    
    FASTEXCEL_LOG_WARN("Force closing streaming file: {}", current_streaming_path_);
    
    // 尝试正常关闭
    bool success = file_manager_->closeStreamingFile();
    
    // 无论是否成功，都清理状态
    streaming_file_open_ = false;
    current_streaming_path_.clear();
    
    if (success) {
        stats_.streaming_files++;
        stats_.files_written++;
    }
    
    return success;
}

}} // namespace fastexcel::core
