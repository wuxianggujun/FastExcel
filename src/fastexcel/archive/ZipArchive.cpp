#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <mz.h>
#include <mz_zip.h>
#include <mz_strm.h>
#include <mz_zip_rw.h>
#include <mz_crypt.h>
#include <mz_os.h>
#include <cstring>
#include <cstdio>
#include <filesystem>
#include <ctime>
#include <ostream>
#include <array>
#include <cstddef>

namespace fastexcel {
namespace archive {

ZipArchive::ZipArchive(const std::string& filename) : filename_(filename) {
    zip_handle_ = nullptr;
    unzip_handle_ = nullptr;
    is_writable_ = false;
    is_readable_ = false;
    stream_entry_open_ = false;
    compression_level_ = 6;  // 使用中等压缩级别，平衡压缩率和速度
}

ZipArchive::~ZipArchive() {
    // 在析构函数中直接调用 cleanup()，避免获取锁
    // 因为析构函数通常在单线程环境中调用
    cleanup();
}

bool ZipArchive::open(bool create) {
    std::lock_guard<std::mutex> lock(mutex_);
    cleanup();  // 直接调用 cleanup() 而不是 close()，避免重复获取锁
    
    if (create) {
        return initForWriting();
    } else {
        return initForReading();
    }
}

bool ZipArchive::close() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    bool success = true;
    
    // 严格检查 mz_zip_writer_close() 的返回值
    if (zip_handle_) {
        int32_t result = mz_zip_writer_close(zip_handle_);
        if (result != MZ_OK) {
            LOG_ERROR("Failed to finalize ZIP file: {}, error code: {}", filename_, result);
            LOG_ERROR("This usually means the ZIP central directory was not written properly");
            success = false;
        } else {
            LOG_DEBUG("ZIP file finalized successfully: {}", filename_);
        }
        
        // 删除writer句柄
        mz_zip_writer_delete(&zip_handle_);
        zip_handle_ = nullptr;
    }
    
    // 清理其他资源
    if (unzip_handle_) {
        mz_zip_reader_close(unzip_handle_);
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
    }
    
    is_writable_ = false;
    is_readable_ = false;
    stream_entry_open_ = false;
    written_paths_.clear();
    
    return success;
}

ZipError ZipArchive::addFile(std::string_view internal_path, std::string_view content) {
    return addFile(internal_path, content.data(), content.size());
}

ZipError ZipArchive::addFile(std::string_view internal_path, const uint8_t* data, size_t size) {
    return addFile(internal_path, reinterpret_cast<const void*>(data), size);
}

// 私有辅助方法：初始化文件信息结构
void ZipArchive::initializeFileInfo(void* file_info_ptr, const std::string& path, size_t size) {
    mz_zip_file& file_info = *static_cast<mz_zip_file*>(file_info_ptr);
    file_info = {}; // 清零结构体
    file_info.filename = path.c_str();
    file_info.uncompressed_size = static_cast<uint64_t>(size);
    file_info.compressed_size = 0; // 让minizip自动计算
    // 对于较大的文件使用Deflate压缩，小文件使用STORE
    file_info.compression_method = (size > 1024) ? MZ_COMPRESS_METHOD_DEFLATE : MZ_COMPRESS_METHOD_STORE;
    
    // 严格按照libxlsxwriter：使用当前时间而不是固定时间戳
    std::time_t now = std::time(nullptr);
    file_info.modified_date = static_cast<time_t>(now);
    file_info.creation_date = static_cast<time_t>(now);
    
    // 不强制设置flag为0，让minizip自己决定需要的位
    // 对于已知大小的文件，minizip会自动处理正确的标志
    // file_info.flag = 0; // 删除这行，让minizip自动处理
}

// 私有辅助方法：写入单个文件条目
ZipError ZipArchive::writeFileEntry(const std::string& internal_path, const void* data, size_t size) {
    // 检查是否已经写入过该路径
    if (written_paths_.find(internal_path) != written_paths_.end()) {
        LOG_WARN("File {} already exists in zip, skipping duplicate entry", internal_path);
        return ZipError::Ok;  // 跳过重复条目，但不报错
    }
    
    // 检查文件大小
    if (size > INT32_MAX) {
        LOG_ERROR("File {} is too large ({} bytes), maximum size is {} bytes",
                 internal_path, size, INT32_MAX);
        return ZipError::TooLarge;
    }
    
    // 初始化文件信息
    mz_zip_file file_info;
    initializeFileInfo(&file_info, internal_path, size);
    
    // 打开条目
    int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    // 写入数据
    if (size > 0) {
        int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_, data, static_cast<int32_t>(size));
        if (bytes_written != static_cast<int32_t>(size)) {
            LOG_ERROR("Failed to write complete data for file {} to zip, written: {} bytes, expected: {} bytes",
                     internal_path, bytes_written, size);
            mz_zip_writer_entry_close(zip_handle_);
            return ZipError::IoFail;
        }
    }
    
    // 关闭条目
    result = mz_zip_writer_entry_close(zip_handle_);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to close entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    // 记录已写入的路径
    written_paths_.insert(internal_path);
    
    LOG_DEBUG("Added file {} to zip, size: {} bytes", internal_path, size);
    return ZipError::Ok;
}

ZipError ZipArchive::addFile(std::string_view internal_path, const void* data, size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    std::string path_str(internal_path);
    return writeFileEntry(path_str, data, size);
}

ZipError ZipArchive::addFiles(const std::vector<FileEntry>& files) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (files.empty()) {
        return ZipError::Ok;
    }
    
    LOG_DEBUG("Starting batch write of {} files", files.size());
    
    // 批量写入所有文件
    for (const auto& file : files) {
        ZipError result = writeFileEntry(file.internal_path, file.content.data(), file.content.size());
        if (result != ZipError::Ok) {
            return result;
        }
    }
    
    LOG_INFO("Batch write completed successfully, {} files added", files.size());
    return ZipError::Ok;
}

ZipError ZipArchive::addFiles(std::vector<FileEntry>&& files) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (files.empty()) {
        return ZipError::Ok;
    }
    
    LOG_DEBUG("Starting batch write of {} files (move semantics)", files.size());
    
    // 批量写入所有文件（移动语义版本）
    for (auto& file : files) {
        // 检查是否已经写入过该路径
        if (written_paths_.find(file.internal_path) != written_paths_.end()) {
            LOG_WARN("File {} already exists in zip, skipping duplicate entry", file.internal_path);
            continue;  // 跳过重复条目
        }
        
        // 检查文件大小
        if (file.content.size() > INT32_MAX) {
            LOG_ERROR("File {} is too large ({} bytes), maximum size is {} bytes",
                     file.internal_path, file.content.size(), INT32_MAX);
            return ZipError::TooLarge;
        }
        
        // 初始化文件信息结构
        mz_zip_file file_info = {};
        file_info.filename = file.internal_path.c_str();
        
        // 设置文件大小信息
        file_info.uncompressed_size = static_cast<uint64_t>(file.content.size());
        file_info.compressed_size = 0; // 让minizip自动计算
        // 对于较大的文件使用Deflate压缩，小文件使用STORE
        file_info.compression_method = (file.content.size() > 1024) ? MZ_COMPRESS_METHOD_DEFLATE : MZ_COMPRESS_METHOD_STORE;
        
        // 严格按照libxlsxwriter：使用当前时间而不是固定时间戳
        std::time_t now = std::time(nullptr);
        file_info.modified_date = static_cast<time_t>(now);
        file_info.creation_date = static_cast<time_t>(now);
        
        // 不强制设置flag为0，让minizip自己决定需要的位
        // 对于已知大小的文件，minizip会自动处理正确的标志
        // file_info.flag = 0; // 删除这行，让minizip自动处理
        
        // 打开条目
        int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
        if (result != MZ_OK) {
            LOG_ERROR("Failed to open entry for file {} in zip, error: {}", file.internal_path, result);
            return ZipError::IoFail;
        }
        
        // 写入数据
        if (!file.content.empty()) {
            int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_,
                                                            file.content.data(),
                                                            static_cast<int32_t>(file.content.size()));
            if (bytes_written != static_cast<int32_t>(file.content.size())) {
                LOG_ERROR("Failed to write complete data for file {} to zip, written: {} bytes, expected: {} bytes",
                         file.internal_path, bytes_written, file.content.size());
                mz_zip_writer_entry_close(zip_handle_);
                return ZipError::IoFail;
            }
        }
        
        // 关闭条目
        result = mz_zip_writer_entry_close(zip_handle_);
        if (result != MZ_OK) {
            LOG_ERROR("Failed to close entry for file {} in zip, error: {}", file.internal_path, result);
            return ZipError::IoFail;
        }
        
        // 记录已写入的路径
        written_paths_.insert(file.internal_path);
        
        LOG_DEBUG("Added file {} to zip, size: {} bytes", file.internal_path, file.content.size());
        
        // 清空内容以释放内存（移动语义优化）
        file.content.clear();
        file.content.shrink_to_fit();
    }
    
    LOG_INFO("Batch write completed successfully, {} files added (move semantics)", files.size());
    return ZipError::Ok;
}

ZipError ZipArchive::extractFile(std::string_view internal_path, std::string& content) {
    std::vector<uint8_t> data;
    ZipError result = extractFile(internal_path, data);
    if (result != ZipError::Ok) {
        return result;
    }
    
    content.assign(reinterpret_cast<const char*>(data.data()), data.size());
    return ZipError::Ok;
}

ZipError ZipArchive::extractFile(std::string_view internal_path, std::vector<uint8_t>& data) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_readable_ || !unzip_handle_) {
        LOG_ERROR("Zip archive not opened for reading");
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
            // 读取当前条目，结果放进 latest
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

    if (!found) {                                    // 没找到
        LOG_ERROR("File {} not found in zip archive", internal_path);
        return ZipError::FileNotFound;
    }
    data.swap(latest);
    LOG_DEBUG("Extracted file {} from zip, size: {} bytes", internal_path, data.size());
    return ZipError::Ok;
}

ZipError ZipArchive::fileExists(std::string_view internal_path) const {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_readable_ || !unzip_handle_) {
        return ZipError::NotOpen;
    }
    
    // 将 string_view 转换为 null 终止的字符串以供 minizip 使用
    std::string path_str(internal_path);
    int32_t result = mz_zip_reader_locate_entry(unzip_handle_, path_str.c_str(), true);
    return result == MZ_OK ? ZipError::Ok : ZipError::FileNotFound;
}

std::vector<std::string> ZipArchive::listFiles() const {
    std::lock_guard<std::mutex> lock(mutex_);
    std::vector<std::string> files;
    
    if (!is_readable_ || !unzip_handle_) {
        return files;
    }
    
    // 获取文件总数，预先分配空间
    int32_t total_entries = 0;
    mz_zip_reader_goto_first_entry(unzip_handle_);
    while (mz_zip_reader_goto_next_entry(unzip_handle_) == MZ_OK) {
        total_entries++;
    }
    mz_zip_reader_goto_first_entry(unzip_handle_);
    
    if (total_entries > 0) {
        files.reserve(static_cast<size_t>(total_entries));
    }
    
    // 回到第一个文件
    if (mz_zip_reader_goto_first_entry(unzip_handle_) != MZ_OK) {
        return files;
    }
    
    mz_zip_file* file_info = nullptr;
    while (mz_zip_reader_entry_get_info(unzip_handle_, &file_info) == MZ_OK && file_info) {
        if (file_info->filename && file_info->filename[0] != '\0') {
            files.push_back(file_info->filename);
        }
        
        if (mz_zip_reader_goto_next_entry(unzip_handle_) != MZ_OK) {
            break;
        }
    }
    
    return files;
}

void ZipArchive::cleanup() {
    // 注意：这个函数主要用于析构函数和异常情况
    // 正常关闭应该使用 close() 函数以检查返回值
    if (zip_handle_) {
        // 尝试关闭，但不检查返回值（因为可能在异常情况下调用）
        mz_zip_writer_close(zip_handle_);
        mz_zip_writer_delete(&zip_handle_);
        zip_handle_ = nullptr;
    }
    
    if (unzip_handle_) {
        mz_zip_reader_close(unzip_handle_);
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
    }
    
    is_writable_ = false;
    is_readable_ = false;
    stream_entry_open_ = false;
    written_paths_.clear();
}

bool ZipArchive::initForWriting() {
    zip_handle_ = mz_zip_writer_create();
    if (!zip_handle_) {
        LOG_ERROR("Failed to create zip writer");
        return false;
    }
    
    // 设置压缩级别和方法 - 使用Deflate压缩
    mz_zip_writer_set_compress_level(zip_handle_, static_cast<int16_t>(compression_level_));
    mz_zip_writer_set_compress_method(zip_handle_, MZ_COMPRESS_METHOD_DEFLATE);
    
    // 关键修复：设置Data Descriptor标志
    // 根据文档，参数0表示在本地文件头中写入CRC和大小（推荐）
    // 参数1表示使用data descriptor（用于流式写入）
    // 注意：需要获取底层的zip handle
    void* zip_handle = nullptr;
    if (mz_zip_writer_get_zip_handle(zip_handle_, &zip_handle) == MZ_OK) {
        // 设置为0，让本地文件头包含正确的CRC和大小信息
        mz_zip_set_data_descriptor(zip_handle, 0);
        LOG_DEBUG("Set data descriptor to 0 (write CRC and sizes in local header)");
    } else {
        LOG_ERROR("Failed to get zip handle for setting data descriptor");
    }
    
    // 设置覆盖回调函数，总是覆盖已存在的文件
    mz_zip_writer_set_overwrite_cb(zip_handle_, this, [](void* handle, void* userdata, const char* filename) -> int32_t {
        // 总是覆盖已存在的文件
        (void)handle;     // 避免未使用参数警告
        (void)userdata;   // 避免未使用参数警告
        (void)filename;   // 避免未使用参数警告
        return MZ_OK;
    });
    
    // 检查文件是否存在
    bool file_exists = std::filesystem::exists(filename_);
    
    // 如果文件已存在，先删除它以确保完全覆盖
    if (file_exists) {
        std::filesystem::remove(filename_);
        LOG_DEBUG("Removed existing zip file: {}", filename_);
    }
    
    // 打开文件进行写入，总是创建新文件（覆盖模式）
    int32_t result = mz_zip_writer_open_file(zip_handle_, filename_.c_str(), 0, 0);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open zip file for writing: {}, error: {}", filename_, result);
        // 清理失败的句柄
        mz_zip_writer_delete(&zip_handle_);
        zip_handle_ = nullptr;
        return false;
    }
    
    is_writable_ = true;
    LOG_DEBUG("Zip archive opened for writing: {}", filename_);
    return true;
}

bool ZipArchive::initForReading() {
    unzip_handle_ = mz_zip_reader_create();
    if (!unzip_handle_) {
        LOG_ERROR("Failed to create zip reader");
        return false;
    }
    
    int32_t result = mz_zip_reader_open_file(unzip_handle_, filename_.c_str());
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open zip file for reading: {}, error: {}", filename_, result);
        // 清理失败的句柄
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
        return false;
    }
    
    is_readable_ = true;
    LOG_DEBUG("Zip archive opened for reading: {}", filename_);
    return true;
}

ZipError ZipArchive::openEntry(std::string_view internal_path) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (stream_entry_open_) {
        LOG_ERROR("Another entry is already open for streaming");
        return ZipError::InvalidParameter;
    }
    
    // 将 string_view 转换为 null 终止的字符串以供 minizip 使用
    std::string path_str(internal_path);
    
    // 检查是否已经写入过该路径
    if (written_paths_.find(path_str) != written_paths_.end()) {
        LOG_WARN("File {} already exists in zip, skipping duplicate entry", path_str);
        return ZipError::Ok;  // 跳过重复条目，但不报错
    }
    
    // 初始化文件信息结构
    mz_zip_file file_info = {};
    file_info.filename = path_str.c_str();
    
    // 对于流式写入，我们不知道最终大小，使用Deflate压缩
    file_info.uncompressed_size = 0;
    file_info.compressed_size = 0;
    file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
    
    // 严格按照libxlsxwriter：使用当前时间而不是固定时间戳
    std::time_t now = std::time(nullptr);
    file_info.modified_date = static_cast<time_t>(now);
    file_info.creation_date = static_cast<time_t>(now);
    
    // 由于已经设置了全局Data Descriptor，不需要单独设置标志
    // minizip会自动处理Data Descriptor的添加
    // 让minizip自动决定需要的标志位
    
    // 打开条目
    int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open entry for file {} in zip, error: {}", internal_path, result);
        return ZipError::IoFail;
    }
    
    stream_entry_open_ = true;
    // 记录已写入的路径（对于流式写入，在打开时就记录）
    written_paths_.insert(path_str);
    LOG_DEBUG("Opened entry for streaming: {}", internal_path);
    return ZipError::Ok;
}

ZipError ZipArchive::writeChunk(const void* data, size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (!stream_entry_open_) {
        LOG_ERROR("No entry is open for streaming");
        return ZipError::InvalidParameter;
    }
    
    if (size == 0) {
        return ZipError::Ok;  // 空块是合法的
    }
    
    // 检查块大小是否超过 int32_t 最大值
    if (size > INT32_MAX) {
        LOG_ERROR("Chunk size {} is too large, maximum size is {} bytes", size, INT32_MAX);
        return ZipError::TooLarge;
    }
    
    // 写入数据块
    int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_, data, static_cast<int32_t>(size));
    if (bytes_written != static_cast<int32_t>(size)) {
        LOG_ERROR("Failed to write complete chunk to zip, written: {} bytes, expected: {} bytes",
                 bytes_written, size);
        return ZipError::IoFail;
    }
    
    return ZipError::Ok;
}

ZipError ZipArchive::closeEntry() {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return ZipError::NotOpen;
    }
    
    if (!stream_entry_open_) {
        LOG_ERROR("No entry is open for streaming");
        return ZipError::InvalidParameter;
    }
    
    // 关闭条目 - 这将让 minizip 自动计算并写入 CRC 和大小信息
    int32_t result = mz_zip_writer_entry_close(zip_handle_);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to close streaming entry in zip, error: {}", result);
        return ZipError::IoFail;
    }
    
    stream_entry_open_ = false;
    LOG_DEBUG("Closed streaming entry");
    return ZipError::Ok;
}

ZipError ZipArchive::extractFileToStream(std::string_view internal_path, std::ostream& output) {
    std::lock_guard<std::mutex> lock(mutex_);
    if (!is_readable_ || !unzip_handle_) {
        LOG_ERROR("Zip archive not opened for reading");
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
                            LOG_ERROR("Failed to write to output stream for file {}", internal_path);
                            mz_zip_reader_entry_close(unzip_handle_);
                            return ZipError::IoFail;
                        }
                        total_read += bytes_read;
                    }
                } while (bytes_read > 0);
                
                mz_zip_reader_entry_close(unzip_handle_);
                
                // 验证是否读取了完整的数据
                if (total_read != static_cast<int64_t>(info->uncompressed_size)) {
                    LOG_ERROR("Incomplete read for file {}, expected: {} bytes, read: {} bytes",
                             internal_path, info->uncompressed_size, total_read);
                    return ZipError::IoFail;
                }
                
                LOG_DEBUG("Extracted file {} to stream, size: {} bytes", internal_path, total_read);
                break;
            }
        }
    } while (mz_zip_reader_goto_next_entry(unzip_handle_) == MZ_OK);

    if (!found) {
        LOG_ERROR("File {} not found in zip archive", internal_path);
        return ZipError::FileNotFound;
    }

    return ZipError::Ok;
}

ZipError ZipArchive::setCompressionLevel(int level) {
    std::lock_guard<std::mutex> lock(mutex_);
    // 验证压缩级别范围 (0-9, 0为无压缩，9为最高压缩)
    if (level < 0 || level > 9) {
        LOG_ERROR("Invalid compression level: {}. Valid range: 0 to 9", level);
        return ZipError::InvalidParameter;
    }
    
    compression_level_ = level;
    
    // 如果已经打开写入模式，立即应用新的压缩级别
    if (is_writable_ && zip_handle_) {
        mz_zip_writer_set_compress_level(zip_handle_, static_cast<int16_t>(compression_level_));
    }
    
    LOG_DEBUG("Set compression level to {}", level);
    return ZipError::Ok;
}

}} // namespace fastexcel::archive
