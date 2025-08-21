#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <stdexcept>
#include <iostream>

namespace fastexcel {
namespace archive {

// 构造/析构

ZipArchive::ZipArchive(const core::Path& path) 
    : filepath_(path)
    , reader_(nullptr)
    , writer_(nullptr)
    , is_open_(false)
    , mode_(Mode::None) {
}

ZipArchive::~ZipArchive() {
    // 确保关闭文件
    if (is_open_) {
        close();
    }
}

// 文件操作

bool ZipArchive::open(bool create) {
    // 如果已经打开，先关闭
    if (is_open_) {
        close();
    }
    
    try {
        if (create) {
            // 创建新文件 - 写模式
            writer_ = std::make_unique<ZipWriter>(filepath_);
            if (writer_->open()) {
                mode_ = Mode::Write;
                is_open_ = true;
                return true;
            }
        } else {
            // 打开现有文件 - 读模式
            reader_ = std::make_unique<ZipReader>(filepath_);
            if (reader_->open()) {
                mode_ = Mode::Read;
                is_open_ = true;
                return true;
            }
        }
    } catch (const std::exception& e) {
        ARCHIVE_ERROR("Failed to open ZIP archive: {}", e.what());
    }
    
    // 打开失败，清理
    reader_.reset();
    writer_.reset();
    mode_ = Mode::None;
    is_open_ = false;
    return false;
}

bool ZipArchive::close() {
    if (!is_open_) {
        return true;  // 已经关闭
    }
    
    bool success = true;
    
    // 关闭reader
    if (reader_) {
        success = reader_->close() && success;
        reader_.reset();
    }
    
    // 关闭writer
    if (writer_) {
        success = writer_->close() && success;
        writer_.reset();
    }
    
    mode_ = Mode::None;
    is_open_ = false;
    
    return success;
}

// 写入操作

ZipError ZipArchive::addFile(std::string_view internal_path, std::string_view content) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFile 返回 ZipError，直接返回
    return writer_->addFile(internal_path, content);
}

ZipError ZipArchive::addFile(std::string_view internal_path, const uint8_t* data, size_t size) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFile 返回 ZipError，直接返回
    return writer_->addFile(internal_path, data, size);
}

ZipError ZipArchive::addFile(std::string_view internal_path, const void* data, size_t size) {
    return addFile(internal_path, static_cast<const uint8_t*>(data), size);
}

ZipError ZipArchive::addFiles(const std::vector<FileEntry>& files) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFiles 返回 ZipError，直接返回
    return writer_->addFiles(files);
}

ZipError ZipArchive::addFiles(std::vector<FileEntry>&& files) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::addFiles 返回 ZipError，直接返回
    return writer_->addFiles(std::move(files));
}

ZipError ZipArchive::openEntry(std::string_view internal_path) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::openEntry 返回 ZipError，直接返回
    return writer_->openEntry(internal_path);
}

ZipError ZipArchive::writeChunk(const void* data, size_t size) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::writeChunk 返回 ZipError，直接返回
    return writer_->writeChunk(data, size);
}

ZipError ZipArchive::closeEntry() {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // ZipWriter::closeEntry 返回 ZipError，直接返回
    return writer_->closeEntry();
}

// 读取操作

ZipError ZipArchive::extractFile(std::string_view internal_path, std::string& content) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFile 返回 ZipError，直接返回
    return reader_->extractFile(internal_path, content);
}

ZipError ZipArchive::extractFile(std::string_view internal_path, std::vector<uint8_t>& data) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFile 返回 ZipError，直接返回
    return reader_->extractFile(internal_path, data);
}

ZipError ZipArchive::extractFileToStream(std::string_view internal_path, std::ostream& output) {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::extractFileToStream 返回 ZipError，直接返回
    return reader_->extractFileToStream(internal_path, output);
}

ZipError ZipArchive::fileExists(std::string_view internal_path) const {
    if (!isReadable()) {
        return ZipError::NotOpen;
    }
    
    // ZipReader::fileExists 返回 ZipError，直接返回  
    return reader_->fileExists(internal_path);
}

std::vector<std::string> ZipArchive::listFiles() const {
    if (!isReadable()) {
        return {};
    }
    
    return reader_->listFiles();
}

// 配置

ZipError ZipArchive::setCompressionLevel(int level) {
    if (!isWritable()) {
        return ZipError::NotOpen;
    }
    
    // 验证压缩级别范围
    if (level < 0 || level > 9) {
        return ZipError::InvalidParameter;
    }
    
    writer_->setCompressionLevel(level);
    return ZipError::Ok;
}

}} // namespace fastexcel::archive
