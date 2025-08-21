#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"  // 为了使用ZipError枚举
#include "fastexcel/utils/Logger.hpp"
#include <mz.h>
#include <mz_strm.h>  // Add this for stream callback types
#include <mz_zip.h>
#include <mz_zip_rw.h>
#include <array>
#include <cstring>

namespace fastexcel {
namespace archive {

// 构造/析构

ZipReader::ZipReader(const core::Path& path) 
    : filepath_(path), filename_(path.string()) {
    unzip_handle_ = nullptr;
    is_open_ = false;
}

ZipReader::~ZipReader() {
    cleanup();
}

ZipReader::ZipReader(ZipReader&& other) noexcept
    : unzip_handle_(other.unzip_handle_),
      filepath_(std::move(other.filepath_)),
      filename_(std::move(other.filename_)),
      is_open_(other.is_open_),
      entry_cache_(std::move(other.entry_cache_)),
      cache_initialized_(other.cache_initialized_) {
    other.unzip_handle_ = nullptr;
    other.is_open_ = false;
    other.cache_initialized_ = false;
}

ZipReader& ZipReader::operator=(ZipReader&& other) noexcept {
    if (this != &other) {
        cleanup();
        unzip_handle_ = other.unzip_handle_;
        filepath_ = std::move(other.filepath_);
        filename_ = std::move(other.filename_);
        is_open_ = other.is_open_;
        entry_cache_ = std::move(other.entry_cache_);
        cache_initialized_ = other.cache_initialized_;
        
        other.unzip_handle_ = nullptr;
        other.is_open_ = false;
        other.cache_initialized_ = false;
    }
    return *this;
}

// 文件操作

bool ZipReader::open() {
    std::lock_guard<std::mutex> lock(mutex_);
    cleanup();  // 清理之前的状态
    
    return initializeReader();
}

bool ZipReader::close() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    bool success = true;
    
    if (unzip_handle_) {
        mz_zip_reader_close(unzip_handle_);
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
    }
    
    is_open_ = false;
    entry_cache_.clear();
    cache_initialized_ = false;
    
    return success;
}

// 条目查询

std::vector<std::string> ZipReader::listFiles() const {
    std::lock_guard<std::mutex> lock(mutex_);
    std::vector<std::string> files;
    
    if (!is_open_ || !unzip_handle_) {
        return files;
    }
    
    // 确保缓存已初始化
    if (!cache_initialized_) {
        buildEntryCache();
    }
    
    // 从缓存返回文件列表
    files.reserve(entry_cache_.size());
    for (const auto& [path, info] : entry_cache_) {
        files.push_back(path);
    }
    
    return files;
}

std::vector<ZipReader::EntryInfo> ZipReader::listEntriesInfo() const {
    std::lock_guard<std::mutex> lock(mutex_);
    std::vector<EntryInfo> entries;
    
    if (!is_open_ || !unzip_handle_) {
        return entries;
    }
    
    // 确保缓存已初始化
    if (!cache_initialized_) {
        buildEntryCache();
    }
    
    // 从缓存返回条目信息
    entries.reserve(entry_cache_.size());
    for (const auto& [path, info] : entry_cache_) {
        entries.push_back(info);
    }
    
    return entries;
}

ZipError ZipReader::fileExists(std::string_view internal_path) const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !unzip_handle_) {
        return ZipError::NotOpen;
    }
    
    // 确保缓存已初始化
    if (!cache_initialized_) {
        buildEntryCache();
    }
    
    // 在缓存中查找
    std::string path_str(internal_path);
    if (entry_cache_.find(path_str) != entry_cache_.end()) {
        return ZipError::Ok;
    }
    
    return ZipError::FileNotFound;
}

bool ZipReader::getEntryInfo(std::string_view internal_path, EntryInfo& info) const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !unzip_handle_) {
        return false;
    }
    
    // 确保缓存已初始化
    if (!cache_initialized_) {
        buildEntryCache();
    }
    
    // 在缓存中查找
    std::string path_str(internal_path);
    auto it = entry_cache_.find(path_str);
    if (it != entry_cache_.end()) {
        info = it->second;
        return true;
    }
    
    return false;
}

// 读取操作

ZipError ZipReader::extractFile(std::string_view internal_path, std::string& content) {
    std::vector<uint8_t> data;
    ZipError result = extractFile(internal_path, data);
    if (result != ZipError::Ok) {
        return result;
    }
    
    content.assign(reinterpret_cast<const char*>(data.data()), data.size());
    return ZipError::Ok;
}

ZipError ZipReader::extractFile(std::string_view internal_path, std::vector<uint8_t>& data) {
    std::lock_guard<std::mutex> lock(mutex_);
    return extractFileInternal(internal_path, data);
}

ZipError ZipReader::extractFileInternal(std::string_view internal_path, 
                                        std::vector<uint8_t>& data) const {
    if (!is_open_ || !unzip_handle_) {
        ARCHIVE_ERROR("Zip archive not opened for reading");
        return ZipError::NotOpen;
    }
    
    // 先跳到第一个条目
    if (mz_zip_reader_goto_first_entry(unzip_handle_) != MZ_OK)
        return ZipError::BadFormat;
    
    std::vector<uint8_t> latest;   // 保存最后一次匹配到的内容
    bool found = false;              // 标记是否找到匹配的文件
    
    do {
        mz_zip_file* info = nullptr;
        if (mz_zip_reader_entry_get_info(unzip_handle_, &info) != MZ_OK || !info)
            break;
        
        if (info->filename && internal_path == info->filename) {
            // 读取当前条目
            if (mz_zip_reader_entry_open(unzip_handle_) == MZ_OK) {
                std::vector<uint8_t> buf(info->uncompressed_size);
                int32_t read = info->uncompressed_size > 0
                                ? mz_zip_reader_entry_read(unzip_handle_,
                                                          buf.data(),
                                                          static_cast<int32_t>(info->uncompressed_size))
                                : 0;
                mz_zip_reader_entry_close(unzip_handle_);
                if (read == static_cast<int32_t>(info->uncompressed_size)) {
                    latest.swap(buf);               // 覆盖为"最后一次"
                    found = true;                   // 标记已找到
                }
            }
        }
    } while (mz_zip_reader_goto_next_entry(unzip_handle_) == MZ_OK);
    
    if (!found) {
        ARCHIVE_ERROR("File {} not found in zip archive", internal_path);
        return ZipError::FileNotFound;
    }
    
    data.swap(latest);
    ARCHIVE_DEBUG("Extracted file {} from zip, size: {} bytes", internal_path, data.size());
    return ZipError::Ok;
}

ZipError ZipReader::extractFileToStream(std::string_view internal_path, std::ostream& output) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !unzip_handle_) {
        ARCHIVE_ERROR("Zip archive not opened for reading");
        return ZipError::NotOpen;
    }
    
    // 先跳到第一个条目
    if (mz_zip_reader_goto_first_entry(unzip_handle_) != MZ_OK)
        return ZipError::BadFormat;
    
    bool found = false;
    do {
        mz_zip_file* info = nullptr;
        if (mz_zip_reader_entry_get_info(unzip_handle_, &info) != MZ_OK || !info)
            break;
        
        if (info->filename && internal_path == info->filename) {
            // 找到目标文件，开始流式读取
            if (mz_zip_reader_entry_open(unzip_handle_) == MZ_OK) {
                found = true;
                
                // 使用固定大小的缓冲区进行分块读取
                constexpr size_t BUFFER_SIZE = 8192;  // 8KB 缓冲区
                std::array<uint8_t, BUFFER_SIZE> buffer;
                
                int64_t total_read = 0;
                int32_t bytes_read = 0;
                
                do {
                    bytes_read = mz_zip_reader_entry_read(unzip_handle_, buffer.data(),
                                                         static_cast<int32_t>(buffer.size()));
                    if (bytes_read > 0) {
                        output.write(reinterpret_cast<const char*>(buffer.data()), bytes_read);
                        if (!output) {
                            ARCHIVE_ERROR("Failed to write to output stream for file {}", internal_path);
                            mz_zip_reader_entry_close(unzip_handle_);
                            return ZipError::IoFail;
                        }
                        total_read += bytes_read;
                    }
                } while (bytes_read > 0);
                
                mz_zip_reader_entry_close(unzip_handle_);
                
                // 验证是否读取了完整的数据
                if (total_read != static_cast<int64_t>(info->uncompressed_size)) {
                    ARCHIVE_ERROR("Incomplete read for file {}, expected: {} bytes, read: {} bytes",
                             internal_path, info->uncompressed_size, total_read);
                    return ZipError::IoFail;
                }
                
                ARCHIVE_DEBUG("Extracted file {} to stream, size: {} bytes", internal_path, total_read);
                break;
            }
        }
    } while (mz_zip_reader_goto_next_entry(unzip_handle_) == MZ_OK);
    
    if (!found) {
        ARCHIVE_ERROR("File {} not found in zip archive", internal_path);
        return ZipError::FileNotFound;
    }
    
    return ZipError::Ok;
}

// 高级功能

bool ZipReader::getRawCompressedData(std::string_view internal_path, 
                                     std::vector<uint8_t>& raw_data,
                                     EntryInfo& info) const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !unzip_handle_) {
        return false;
    }
    
    // 获取条目信息
    if (!getEntryInfo(internal_path, info)) {
        return false;
    }
    
    // TODO: 实现真正的原始压缩数据获取
    // 目前简化实现：返回解压后的数据
    std::vector<uint8_t> data;
    if (extractFileInternal(internal_path, data) == ZipError::Ok) {
        raw_data = std::move(data);
        return true;
    }
    
    return false;
}

ZipError ZipReader::streamFile(std::string_view internal_path,
                               std::function<bool(const uint8_t*, size_t)> callback,
                               size_t buffer_size) const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (!is_open_ || !unzip_handle_ || !callback) {
        return ZipError::NotOpen;
    }
    
    if (!locateEntry(internal_path)) {
        return ZipError::FileNotFound;
    }
    
    // 打开条目
    if (mz_zip_reader_entry_open(unzip_handle_) != MZ_OK) {
        ARCHIVE_ERROR("Failed to open entry: {}", internal_path);
        return ZipError::IoFail;
    }
    
    // 流式读取
    std::vector<uint8_t> buffer(buffer_size);
    bool continue_reading = true;
    
    while (continue_reading) {
        int32_t bytes_read = mz_zip_reader_entry_read(unzip_handle_,
                                                      buffer.data(),
                                                      static_cast<int32_t>(buffer.size()));
        
        if (bytes_read <= 0) {
            break;  // 读取完成或出错
        }
        
        // 调用回调
        continue_reading = callback(buffer.data(), static_cast<size_t>(bytes_read));
    }
    
    mz_zip_reader_entry_close(unzip_handle_);
    
    return continue_reading ? ZipError::Ok : ZipError::Ok;  // 用户取消也算成功
}

// 统计信息

ZipReader::Stats ZipReader::getStats() const {
    std::lock_guard<std::mutex> lock(mutex_);
    Stats stats;
    
    if (!is_open_ || !unzip_handle_) {
        return stats;
    }
    
    // 确保缓存已初始化
    if (!cache_initialized_) {
        buildEntryCache();
    }
    
    stats.total_entries = entry_cache_.size();
    
    for (const auto& [path, info] : entry_cache_) {
        stats.total_compressed += info.compressed_size;
        stats.total_uncompressed += info.uncompressed_size;
    }
    
    if (stats.total_uncompressed > 0) {
        stats.compression_ratio = static_cast<double>(stats.total_compressed) / 
                                 static_cast<double>(stats.total_uncompressed);
    }
    
    return stats;
}

// 内部辅助方法

bool ZipReader::initializeReader() {
    unzip_handle_ = mz_zip_reader_create();
    if (!unzip_handle_) {
        ARCHIVE_ERROR("Failed to create zip reader");
        return false;
    }
    
    int32_t result = mz_zip_reader_open_file(unzip_handle_, filepath_.c_str());
    if (result != MZ_OK) {
        ARCHIVE_ERROR("Failed to open zip file for reading: {}, error: {}", filename_, result);
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
        return false;
    }
    
    is_open_ = true;
    ARCHIVE_DEBUG("Zip archive opened for reading: {}", filename_);
    
    // 初始化条目缓存
    buildEntryCache();
    
    return true;
}

void ZipReader::cleanup() {
    if (unzip_handle_) {
        mz_zip_reader_close(unzip_handle_);
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
    }
    
    is_open_ = false;
    entry_cache_.clear();
    cache_initialized_ = false;
}

void ZipReader::buildEntryCache() const {
    if (!unzip_handle_ || cache_initialized_) {
        return;
    }
    
    entry_cache_.clear();
    
    // 跳到第一个条目
    if (mz_zip_reader_goto_first_entry(unzip_handle_) != MZ_OK) {
        return;
    }
    
    do {
        mz_zip_file* file_info = nullptr;
        if (mz_zip_reader_entry_get_info(unzip_handle_, &file_info) == MZ_OK && file_info) {
            if (file_info->filename && file_info->filename[0] != '\0') {
                EntryInfo info;
                info.path = file_info->filename;
                info.compressed_size = file_info->compressed_size;
                info.uncompressed_size = file_info->uncompressed_size;
                info.crc32 = file_info->crc;
                info.compression_method = file_info->compression_method;
                info.modified_date = file_info->modified_date;
                info.creation_date = file_info->creation_date;
                info.flag = file_info->flag;
                info.is_directory = (info.path.back() == '/');
                
                entry_cache_[info.path] = info;
            }
        }
    } while (mz_zip_reader_goto_next_entry(unzip_handle_) == MZ_OK);
    
    cache_initialized_ = true;
    ARCHIVE_DEBUG("Built entry cache with {} entries", entry_cache_.size());
}

bool ZipReader::locateEntry(std::string_view path) const {
    if (!unzip_handle_) {
        return false;
    }
    
    std::string path_str(path);
    int32_t result = mz_zip_reader_locate_entry(unzip_handle_, path_str.c_str(), 1);
    return result == MZ_OK;
}

}} // namespace fastexcel::archive
