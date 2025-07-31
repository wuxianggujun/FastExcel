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

namespace fastexcel {
namespace archive {

ZipArchive::ZipArchive(const std::string& filename) : filename_(filename) {
    zip_handle_ = nullptr;
    unzip_handle_ = nullptr;
    is_writable_ = false;
    is_readable_ = false;
}

ZipArchive::~ZipArchive() {
    close();
}

bool ZipArchive::open(bool create) {
    close();
    
    if (create) {
        return initForWriting();
    } else {
        return initForReading();
    }
}

bool ZipArchive::close() {
    cleanup();
    return true;
}

bool ZipArchive::addFile(const std::string& internal_path, const std::string& content) {
    return addFile(internal_path, content.data(), content.size());
}

bool ZipArchive::addFile(const std::string& internal_path, const std::vector<uint8_t>& data) {
    return addFile(internal_path, data.data(), data.size());
}

bool ZipArchive::addFile(const std::string& internal_path, const void* data, size_t size) {
    if (!is_writable_ || !zip_handle_) {
        LOG_ERROR("Zip archive not opened for writing");
        return false;
    }
    
    // 初始化文件信息结构 - 只设置必要的字段，让 minizip 自动计算 CRC 和大小
    mz_zip_file file_info = {};
    file_info.filename = internal_path.c_str();
    file_info.modified_date = time(nullptr);
    file_info.flag = MZ_ZIP_FLAG_UTF8 | MZ_ZIP_FLAG_DATA_DESCRIPTOR;
    
    // 打开条目
    int32_t result = mz_zip_writer_entry_open(zip_handle_, &file_info);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open entry for file {} in zip, error: {}", internal_path, result);
        return false;
    }
    
    // 写入数据
    if (size > 0) {
        int32_t bytes_written = mz_zip_writer_entry_write(zip_handle_, data, static_cast<int32_t>(size));
        if (bytes_written != static_cast<int32_t>(size)) {
            LOG_ERROR("Failed to write complete data for file {} to zip, written: {} bytes, expected: {} bytes",
                     internal_path, bytes_written, size);
            mz_zip_writer_entry_close(zip_handle_);
            return false;
        }
    }
    
    // 关闭条目 - 这将让 minizip 自动计算并写入 CRC 和大小信息
    result = mz_zip_writer_entry_close(zip_handle_);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to close entry for file {} in zip, error: {}", internal_path, result);
        return false;
    }
    
    LOG_DEBUG("Added file {} to zip, size: {} bytes", internal_path, size);
    return true;
}

bool ZipArchive::extractFile(const std::string& internal_path, std::string& content) {
    std::vector<uint8_t> data;
    if (!extractFile(internal_path, data)) {
        return false;
    }
    
    content.assign(data.begin(), data.end());
    return true;
}

bool ZipArchive::extractFile(const std::string& internal_path, std::vector<uint8_t>& data) {
    if (!is_readable_ || !unzip_handle_) {
        LOG_ERROR("Zip archive not opened for reading");
        return false;
    }

    // 先跳到第一个条目
    if (mz_zip_reader_goto_first_entry(unzip_handle_) != MZ_OK)
        return false;

    std::vector<uint8_t> latest;   // 保存最后一次匹配到的内容
    bool found = false;            // 标记是否找到匹配的文件
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
        return false;
    }
    data.swap(latest);
    LOG_DEBUG("Extracted file {} from zip, size: {} bytes", internal_path, data.size());
    return true;
}

bool ZipArchive::fileExists(const std::string& internal_path) const {
    if (!is_readable_ || !unzip_handle_) {
        return false;
    }
    
    int32_t result = mz_zip_reader_locate_entry(unzip_handle_, internal_path.c_str(), true);
    return result == MZ_OK;
}

std::vector<std::string> ZipArchive::listFiles() const {
    std::vector<std::string> files;
    
    if (!is_readable_ || !unzip_handle_) {
        return files;
    }
    
    // 回到第一个文件
    mz_zip_reader_goto_first_entry(unzip_handle_);
    
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
    if (zip_handle_) {
        // 确保所有数据都被写入磁盘
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
}

bool ZipArchive::initForWriting() {
    zip_handle_ = mz_zip_writer_create();
    if (!zip_handle_) {
        LOG_ERROR("Failed to create zip writer");
        return false;
    }
    
    // 设置压缩级别和方法
    mz_zip_writer_set_compress_level(zip_handle_, MZ_COMPRESS_LEVEL_DEFAULT);
    mz_zip_writer_set_compress_method(zip_handle_, MZ_COMPRESS_METHOD_DEFLATE);
    
    // 设置覆盖回调函数，总是覆盖已存在的文件
    mz_zip_writer_set_overwrite_cb(zip_handle_, this, [](void* handle, void* userdata, const char* filename) -> int32_t {
        // 总是覆盖已存在的文件
        return MZ_OK;
    });
    
    // 检查文件是否存在
    bool file_exists = std::filesystem::exists(filename_);
    
    // 打开文件进行写入，如果文件已存在则追加，否则创建
    int32_t result = mz_zip_writer_open_file(zip_handle_, filename_.c_str(), 0, file_exists ? 1 : 0);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open zip file for writing: {}, error: {}", filename_, result);
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
        mz_zip_reader_delete(&unzip_handle_);
        unzip_handle_ = nullptr;
        return false;
    }
    
    is_readable_ = true;
    LOG_DEBUG("Zip archive opened for reading: {}", filename_);
    return true;
}

}} // namespace fastexcel::archive
