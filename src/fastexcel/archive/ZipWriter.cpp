#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipWriter.hpp"
#include "fastexcel/archive/ZipArchive.hpp"  // 为了使用ZipError枚举
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include <mz.h>
#include <mz_strm.h>  // Add this for stream callback types
#include <mz_zip.h>
#include <mz_zip_rw.h>
#include <cstring>

namespace fastexcel {
namespace archive {

// 构造/析构

ZipWriter::ZipWriter(const core::Path& path)
    : filepath_(path), filename_(path.string()) {
    zip_handle_ = nullptr;
    is_open_ = false;
    stream_entry_open_ = false;
    compression_level_ = 6;
}

ZipWriter::~ZipWriter() {
    cleanup();
}

ZipWriter::ZipWriter(ZipWriter&& other) noexcept
    : zip_handle_(other.zip_handle_),
      filepath_(std::move(other.filepath_)),
      filename_(std::move(other.filename_)),
      is_open_(other.is_open_),
      stream_entry_open_(other.stream_entry_open_),
      compression_level_(other.compression_level_),
      written_paths_(std::move(other.written_paths_)),
      stats_(other.stats_) {
    other.zip_handle_ = nullptr;
    other.is_open_ = false;
    other.stream_entry_open_ = false;
}

ZipWriter& ZipWriter::operator=(ZipWriter&& other) noexcept {
    if (this != &other) {
        cleanup();
        zip_handle_ = other.zip_handle_;
        filepath_ = std::move(other.filepath_);
        filename_ = std::move(other.filename_);
        is_open_ = other.is_open_;
        stream_entry_open_ = other.stream_entry_open_;
        compression_level_ = other.compression_level_;
        written_paths_ = std::move(other.written_paths_);
        stats_ = other.stats_;
        
        other.zip_handle_ = nullptr;
        other.is_open_ = false;
        other.stream_entry_open_ = false;
    }
    return *this;
}

// 文件操作

bool ZipWriter::open(bool create) {
    std::lock_guard<std::mutex> lock(mutex_);
    cleanup();  // 清理之前的状态
    
    if (create) {
        return initializeWriter();
    } else {
        // TODO: 实现追加模式
        FASTEXCEL_LOG_ERROR("[ARCH] Append mode not yet implemented");
        return false;
    }
}

bool ZipWriter::close() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 幂等性检查：如果已经关闭，直接返回成功
    if (!is_open_ || !zip_handle_) {
        return true;  // 已经关闭，避免重复关闭
    }
    
    bool success = true;
    
    // 严格检查 mz_zip_writer_close() 的返回值
    int32_t result = mz_zip_writer_close(zip_handle_);
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to finalize ZIP file: {}, error code: {}", filename_, result);
        FASTEXCEL_LOG_ERROR("[ARCH] This usually means the ZIP central directory was not written properly");
        success = false;
    } else {
        FASTEXCEL_LOG_DEBUG("[ARCH] ZIP file finalized successfully: {}", filename_);
    }
    
    // 删除writer句柄
    mz_zip_writer_delete(&zip_handle_);
    zip_handle_ = nullptr;
    
    is_open_ = false;
    stream_entry_open_ = false;
    written_paths_.clear();
    
    return success;
}

// 基本写入操作

ZipError ZipWriter::addFile(std::string_view internal_path, std::string_view content) {
    return addFile(internal_path, content.data(), content.size());
}

ZipError ZipWriter::addFile(std::string_view internal_path, const uint8_t* data, size_t size) {
    return addFile(internal_path, reinterpret_cast<const void*>(data), size);
}

ZipError ZipWriter::addFile(std::string_view internal_path, const void* data, size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    std::string path_str(internal_path);
    return writeFileEntry(path_str, data, size);
}

// 批量写入

ZipError ZipWriter::addFiles(const std::vector<FileEntry>& files) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (files.empty()) {
        return ZipError::Ok;
    }
    
    FASTEXCEL_LOG_DEBUG("[ARCH] Starting batch write of {} files", files.size());
    
    // 批量写入所有文件
    for (const auto& file : files) {
        ZipError result = writeFileEntry(file.internal_path, file.content.data(), file.content.size());
        if (result != ZipError::Ok) {
            return result;
        }
    }
    
    FASTEXCEL_LOG_INFO("[ARCH] Batch write completed successfully, {} files added", files.size());
    return ZipError::Ok;
}

ZipError ZipWriter::addFiles(std::vector<FileEntry>&& files) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (files.empty()) {
        return ZipError::Ok;
    }
    
    FASTEXCEL_LOG_DEBUG("[ARCH] Starting batch write of {} files (move semantics)", files.size());
    
    // 批量写入所有文件（移动语义版本）
    for (auto& file : files) {
        // 检查是否已经写入过该路径
        if (written_paths_.find(file.internal_path) != written_paths_.end()) {
            FASTEXCEL_LOG_WARN("[ARCH] File {} already exists in zip, skipping duplicate entry", file.internal_path);
            continue;  // 跳过重复条目
        }
        
        // 检查文件大小
        if (file.content.size() > INT32_MAX) {
            FASTEXCEL_LOG_ERROR("[ARCH] File {} is too large ({} bytes), maximum size is {} bytes",
                     file.internal_path, file.content.size(), INT32_MAX);
            return ZipError::TooLarge;
        }
        
        // 初始化文件信息结构
        mz_zip_file file_info = {};
        initializeFileInfo(&file_info, file.internal_path, file.content.size());
        
        // 打开条目
        int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
        if (result != MZ_OK) {
            FASTEXCEL_LOG_ERROR("[ARCH] Failed to open entry for file {} in zip, error: {}", file.internal_path, result);
            return ZipError::IoFail;
        }
        
        // 写入数据
        if (!file.content.empty()) {
            int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_,
                                                            file.content.data(),
                                                            static_cast<int32_t>(file.content.size()));
            
            if (bytes_written != static_cast<int32_t>(file.content.size())) {
                FASTEXCEL_LOG_ERROR("[ARCH] Failed to write complete data for file {} to zip", file.internal_path);
                mz_zip_writer_entry_close(zip_handle_);
                return ZipError::IoFail;
            }
        }
        
        // 关闭条目
        result = mz_zip_writer_entry_close(zip_handle_);
        if (result != MZ_OK) {
            FASTEXCEL_LOG_ERROR("[ARCH] Failed to close entry for file {} in zip, error: {}", file.internal_path, result);
            return ZipError::IoFail;
        }
        
        // 记录已写入的路径
        written_paths_.insert(file.internal_path);
        stats_.entries_written++;
        stats_.bytes_written += file.content.size();
        
        FASTEXCEL_LOG_DEBUG("[ARCH] Successfully added file {} to zip, size: {} bytes", file.internal_path, file.content.size());
        
        // 清空内容以释放内存（移动语义优化）
        file.content.clear();
        file.content.shrink_to_fit();
    }
    
    FASTEXCEL_LOG_INFO("[ARCH] Batch write completed successfully, {} files added (move semantics)", files.size());
    return ZipError::Ok;
}

// 流式写入

ZipError ZipWriter::openEntry(std::string_view internal_path) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (stream_entry_open_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Another entry is already open for streaming");
        return ZipError::InvalidParameter;
    }
    
    std::string path_str(internal_path);
    
    // 检查是否已经写入过该路径
    if (written_paths_.find(path_str) != written_paths_.end()) {
        FASTEXCEL_LOG_WARN("File {} already exists in zip, skipping duplicate entry", path_str);
        return ZipError::Ok;
    }
    
    // 初始化文件信息结构
    mz_zip_file file_info = {};
    file_info.filename = path_str.c_str();
    file_info.uncompressed_size = 0;  // 流式写入时未知大小
    file_info.compressed_size = 0;
    file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
    
    // 使用TimeUtils获取本地时间
    std::tm current_time = utils::TimeUtils::getCurrentTime();
    std::time_t now = utils::TimeUtils::tmToTimeT(current_time);
    file_info.modified_date = now;
    file_info.creation_date = now;
    
    file_info.flag = 0;  // 不使用Data Descriptor
    
#ifdef _WIN32
    file_info.version_madeby = (MZ_HOST_SYSTEM_WINDOWS_NTFS << 8) | 20;
#else
    file_info.version_madeby = (MZ_HOST_SYSTEM_UNIX << 8) | 20;
#endif
    
    // 打开条目
    int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to open entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    stream_entry_open_ = true;
    written_paths_.insert(path_str);
    stats_.entries_written++;
    
    FASTEXCEL_LOG_DEBUG("[ARCH] Successfully opened entry for streaming: {}", internal_path);
    return ZipError::Ok;
}

ZipError ZipWriter::writeChunk(const void* data, size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (!stream_entry_open_) {
        FASTEXCEL_LOG_ERROR("[ARCH] No entry is open for streaming");
        return ZipError::InvalidParameter;
    }
    
    if (size == 0) {
        return ZipError::Ok;  // 空块是合法的
    }
    
    if (size > INT32_MAX) {
        FASTEXCEL_LOG_ERROR("[ARCH] Chunk size {} is too large", size);
        return ZipError::TooLarge;
    }
    
    // 写入数据块
    int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_, data, static_cast<int32_t>(size));
    
    if (bytes_written != static_cast<int32_t>(size)) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to write complete chunk to zip");
        return ZipError::IoFail;
    }
    
    stats_.bytes_written += bytes_written;
    FASTEXCEL_LOG_DEBUG("[ARCH] Successfully wrote chunk of {} bytes", bytes_written);
    return ZipError::Ok;
}

ZipError ZipWriter::closeEntry() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (!stream_entry_open_) {
        FASTEXCEL_LOG_ERROR("[ARCH] No entry is open for streaming");
        return ZipError::InvalidParameter;
    }
    
    int32_t result = mz_zip_writer_entry_close(zip_handle_);
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to close streaming entry, error: {}", result);
        stream_entry_open_ = false;
        return ZipError::IoFail;
    }
    
    stream_entry_open_ = false;
    FASTEXCEL_LOG_DEBUG("[ARCH] Successfully closed streaming entry");
    return ZipError::Ok;
}

// 高级功能

ZipError ZipWriter::writeRawCompressedData(std::string_view internal_path,
                                          const void* raw_data, size_t size,
                                          size_t uncompressed_size,
                                          uint32_t crc32,
                                          int compression_method) {
    // TODO: 实现原始压缩数据写入
    // 目前简化实现：作为普通数据写入
    return addFile(internal_path, raw_data, size);
}

ZipError ZipWriter::setCompressionLevel(int level) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (level < 0 || level > 9) {
        FASTEXCEL_LOG_ERROR("[ARCH] Invalid compression level: {}. Valid range: 0 to 9", level);
        return ZipError::InvalidParameter;
    }
    
    compression_level_ = level;
    
    // 如果已经打开写入模式，立即应用新的压缩级别
    if (is_open_ && zip_handle_) {
        mz_zip_writer_set_compress_level(zip_handle_, static_cast<int16_t>(compression_level_));
    }
    
    FASTEXCEL_LOG_DEBUG("[ARCH] Set compression level to {}", level);
    return ZipError::Ok;
}

// 内部辅助方法

bool ZipWriter::initializeWriter() {
    FASTEXCEL_LOG_DEBUG("[ARCH] Initializing ZIP writer for file: {}", filename_);
    
    zip_handle_ = mz_zip_writer_create();
    if (!zip_handle_) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to create zip writer");
        return false;
    }
    
    // 设置压缩方法和级别
    mz_zip_writer_set_compress_method(zip_handle_, MZ_COMPRESS_METHOD_DEFLATE);
    mz_zip_writer_set_compress_level(zip_handle_, compression_level_);
    FASTEXCEL_LOG_DEBUG("[ARCH] Set compression to DEFLATE with level {}", compression_level_);
    
    // 设置覆盖回调函数
    mz_zip_writer_set_overwrite_cb(zip_handle_, this, [](void* handle, void* userdata, const char* filename) -> int32_t {
        return MZ_OK;  // 总是覆盖
    });
    
    // 如果文件已存在，先删除
    if (filepath_.exists()) {
        filepath_.remove();
        FASTEXCEL_LOG_DEBUG("[ARCH] Removed existing zip file: {}", filename_);
    }
    
    // 打开文件进行写入
    int32_t result = mz_zip_writer_open_file(zip_handle_, filepath_.c_str(), 0, 0);
    
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("[ARCH] Failed to open zip file for writing: {}, error: {}", filename_, result);
        mz_zip_writer_delete(&zip_handle_);
        zip_handle_ = nullptr;
        return false;
    }
    
    // 禁用Data Descriptor以避免兼容性问题
    void* zip_handle = nullptr;
    if (mz_zip_writer_get_zip_handle(zip_handle_, &zip_handle) == MZ_OK && zip_handle) {
        mz_zip_set_data_descriptor(zip_handle, 0);
        FASTEXCEL_LOG_DEBUG("[ARCH] Disabled Data Descriptor for compatibility");
    }
    
    is_open_ = true;
    FASTEXCEL_LOG_DEBUG("[ARCH] ZIP archive successfully opened for writing: {}", filename_);
    return true;
}

void ZipWriter::cleanup() {
    // 如果still open，先正常关闭
    if (is_open_ && zip_handle_) {
        close();  // 使用幂等的close方法
    } else if (zip_handle_) {
        // 如果句柄存在但状态异常，强制清理句柄
        mz_zip_writer_delete(&zip_handle_);
        zip_handle_ = nullptr;
    }
    
    is_open_ = false;
    stream_entry_open_ = false;
    written_paths_.clear();
}

void ZipWriter::initializeFileInfo(void* file_info_ptr, const std::string& path, size_t size) {
    mz_zip_file& file_info = *static_cast<mz_zip_file*>(file_info_ptr);
    file_info = {}; // 清零结构体
    file_info.filename = path.c_str();
    file_info.uncompressed_size = static_cast<uint64_t>(size);
    file_info.compressed_size = 0; // 让minizip自动计算
    
    // 根据压缩级别选择压缩方法
    if (compression_level_ == 0) {
        file_info.compression_method = MZ_COMPRESS_METHOD_STORE;
        FASTEXCEL_LOG_DEBUG("[ARCH] Using STORE compression method");
    } else {
        file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
        FASTEXCEL_LOG_DEBUG("[ARCH] Using DEFLATE compression method");
    }
    
    // 使用TimeUtils获取当前时间
    std::tm current_time = utils::TimeUtils::getCurrentTime();
    std::time_t now = utils::TimeUtils::tmToTimeT(current_time);
    file_info.modified_date = now;
    file_info.creation_date = now;
    
    file_info.flag = 0;  // 不使用Data Descriptor
    
#ifdef _WIN32
    file_info.version_madeby = (MZ_HOST_SYSTEM_WINDOWS_NTFS << 8) | 20;
#else
    file_info.version_madeby = (MZ_HOST_SYSTEM_UNIX << 8) | 20;
#endif
}

ZipError ZipWriter::writeFileEntry(const std::string& internal_path, const void* data, size_t size) {
    // 检查是否已经写入过该路径
    if (written_paths_.find(internal_path) != written_paths_.end()) {
        FASTEXCEL_LOG_WARN("File {} already exists in zip, skipping duplicate entry", internal_path);
        return ZipError::Ok;
    }
    
    // 检查文件大小
    if (size > INT32_MAX) {
        FASTEXCEL_LOG_ERROR("File {} is too large ({} bytes)", internal_path, size);
        return ZipError::TooLarge;
    }
    
    // 初始化文件信息
    mz_zip_file file_info;
    initializeFileInfo(&file_info, internal_path, size);
    
    // 打开条目
    int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("Failed to open entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    // 写入数据
    if (size > 0) {
        int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_, data, static_cast<int32_t>(size));
        
        if (bytes_written != static_cast<int32_t>(size)) {
            FASTEXCEL_LOG_ERROR("Failed to write complete data for file {} to zip", internal_path);
            mz_zip_writer_entry_close(zip_handle_);
            return ZipError::IoFail;
        }
    }
    
    // 关闭条目
    result = mz_zip_writer_entry_close(zip_handle_);
    if (result != MZ_OK) {
        FASTEXCEL_LOG_ERROR("Failed to close entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    // 记录已写入的路径和统计
    written_paths_.insert(internal_path);
    stats_.entries_written++;
    stats_.bytes_written += size;
    
    FASTEXCEL_LOG_DEBUG("Successfully added file {} to zip, size: {} bytes", internal_path, size);
    return ZipError::Ok;
}

}} // namespace fastexcel::archive
