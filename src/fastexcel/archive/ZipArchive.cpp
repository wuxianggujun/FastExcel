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
    
    int32_t result = mz_zip_reader_locate_entry(unzip_handle_, internal_path.c_str(), true);
    if (result != MZ_OK) {
        LOG_ERROR("File {} not found in zip archive", internal_path);
        return false;
    }
    
    result = mz_zip_reader_entry_open(unzip_handle_);
    if (result != MZ_OK) {
        LOG_ERROR("Failed to open file {} in zip archive", internal_path);
        return false;
    }
    
    // 获取文件大小
    mz_zip_file* file_info = nullptr;
    result = mz_zip_reader_entry_get_info(unzip_handle_, &file_info);
    if (result != MZ_OK || !file_info) {
        LOG_ERROR("Failed to get file info for {}", internal_path);
        mz_zip_reader_entry_close(unzip_handle_);
        return false;
    }
    
    // 读取文件内容
    data.resize(file_info->uncompressed_size);
    int32_t bytes_read = mz_zip_reader_entry_read(unzip_handle_, data.data(), static_cast<int32_t>(file_info->uncompressed_size));
    
    mz_zip_reader_entry_close(unzip_handle_);
    
    if (bytes_read != static_cast<int32_t>(file_info->uncompressed_size)) {
        LOG_ERROR("Failed to read complete file {} from zip", internal_path);
        return false;
    }
    
    LOG_DEBUG("Extracted file {} from zip, size: {} bytes", internal_path, bytes_read);
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
    
    // 打开文件进行写入，如果文件已存在则覆盖
    int32_t result = mz_zip_writer_open_file(zip_handle_, filename_.c_str(), 0, 0);
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
