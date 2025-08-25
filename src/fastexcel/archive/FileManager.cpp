#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/xml/ContentTypes.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include <fmt/format.h>

#include <chrono>
#include <iomanip>
#include <sstream>
#include <algorithm>
#include <fstream>
#include "fastexcel/utils/TimeUtils.hpp"

namespace fastexcel {
namespace archive {

FileManager::FileManager(const core::Path& path) : filename_(path.string()), filepath_(path) {
}

FileManager::~FileManager() {
    close();
}

bool FileManager::open(bool create) {
    if (isOpen()) {
        close();
    }
    
    archive_ = std::make_unique<ZipArchive>(filepath_);
    if (!archive_->open(create)) {
        FASTEXCEL_LOG_ERROR("Failed to open archive: {}", filename_);
        return false;
    }
    return true;
}

bool FileManager::close() {
    if (archive_) {
        archive_->close();
        archive_.reset();
    }
    return true;
}

bool FileManager::writeFile(const std::string& internal_path, const std::string& content) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::writeFile(const std::string& internal_path, const std::vector<uint8_t>& data) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, data.data(), data.size()) == ZipError::Ok;
}

bool FileManager::writeFiles(const std::vector<std::pair<std::string, std::string>>& files) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (files.empty()) {
        return true;
    }
    
    // 转换为ZipArchive::FileEntry格式
    std::vector<ZipArchive::FileEntry> zip_files;
    zip_files.reserve(files.size());
    
    for (const auto& [path, content] : files) {
        zip_files.emplace_back(path, content);
    }
    
    FASTEXCEL_LOG_INFO("Writing {} files in batch mode", files.size());
    return archive_->addFiles(zip_files) == ZipError::Ok;
}

bool FileManager::writeFiles(std::vector<std::pair<std::string, std::string>>&& files) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (files.empty()) {
        return true;
    }
    
    // 转换为ZipArchive::FileEntry格式（移动语义）
    std::vector<ZipArchive::FileEntry> zip_files;
    zip_files.reserve(files.size());
    
    for (auto& [path, content] : files) {
        zip_files.emplace_back(std::move(path), std::move(content));
    }
    
    FASTEXCEL_LOG_INFO("Writing {} files in batch mode (move semantics)", zip_files.size());
    return archive_->addFiles(std::move(zip_files)) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::string& content) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->extractFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::vector<uint8_t>& data) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->extractFile(internal_path, data) == ZipError::Ok;
}

bool FileManager::fileExists(const std::string& internal_path) const {
    if (!isOpen()) {
        return false;
    }
    
    return archive_->fileExists(internal_path) == ZipError::Ok;
}

std::vector<std::string> FileManager::listFiles() const {
    if (!isOpen()) {
        return {};
    }
    
    return archive_->listFiles();
}

bool FileManager::setCompressionLevel(int level) {
    if (!archive_) {
        FASTEXCEL_LOG_ERROR("Archive not initialized");
        return false;
    }
    
    return archive_->setCompressionLevel(level) == ZipError::Ok;
}

// 流式写入方法实现 - 极致性能模式
bool FileManager::openStreamingFile(const std::string& internal_path) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->openEntry(internal_path);
    if (result != ZipError::Ok) {
        FASTEXCEL_LOG_ERROR("Failed to open streaming entry: {}", internal_path);
        return false;
    }
    
    FASTEXCEL_LOG_DEBUG("Opened streaming file: {}", internal_path);
    return true;
}

bool FileManager::writeStreamingChunk(const void* data, size_t size) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->writeChunk(data, size);
    if (result != ZipError::Ok) {
        FASTEXCEL_LOG_ERROR("Failed to write streaming chunk of size {}", size);
        return false;
    }
    
    return true;
}

bool FileManager::writeStreamingChunk(const std::string& data) {
    return writeStreamingChunk(data.c_str(), data.size());
}

bool FileManager::closeStreamingFile() {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->closeEntry();
    if (result != ZipError::Ok) {
        FASTEXCEL_LOG_ERROR("Failed to close streaming entry");
        return false;
    }
    
    FASTEXCEL_LOG_DEBUG("Closed streaming file");
    return true;
}

bool FileManager::copyFromExistingPackage(const core::Path& source_package,
                                          const std::vector<std::string>& skip_prefixes) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open for writing when copying from existing package");
        return false;
    }

    // 打开源包用于读取
    ZipArchive src(source_package);
    if (!src.open(false)) {
        FASTEXCEL_LOG_ERROR("Failed to open source package for copy: {}", source_package.string());
        return false;
    }

    auto paths = src.listFiles();
    FASTEXCEL_LOG_INFO("Copy-through existing entries: {} files to scan", paths.size());

    for (const auto& p : paths) {
        bool skip = false;
        for (const auto& prefix : skip_prefixes) {
            if (!prefix.empty() && p.rfind(prefix, 0) == 0) { // starts with
                skip = true;
                break;
            }
        }
        if (skip) {
            FASTEXCEL_LOG_DEBUG("Skip passthrough: {}", p);
            continue;
        }

        std::vector<uint8_t> data;
        if (src.extractFile(p, data) == ZipError::Ok) {
            if (archive_->fileExists(p) == ZipError::Ok) {
                FASTEXCEL_LOG_DEBUG("Target already has {}, skipping overwrite", p);
                continue;
            }
            if (archive_->addFile(p, data.data(), data.size()) != ZipError::Ok) {
                FASTEXCEL_LOG_ERROR("Failed to write passthrough entry: {}", p);
                src.close();
                return false;
            }
            FASTEXCEL_LOG_DEBUG("Pass-through copied: {} ({} bytes)", p, data.size());
        } else {
            FASTEXCEL_LOG_WARN("Failed to extract entry from source for passthrough: {}", p);
        }
    }

    src.close();
    return true;
}

// 图片文件管理

bool FileManager::addImageFile(const std::string& image_id,
                               const std::vector<uint8_t>& image_data,
                               core::ImageFormat format) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (image_data.empty()) {
        FASTEXCEL_LOG_ERROR("Image data is empty for image: {}", image_id);
        return false;
    }
    
    std::string internal_path = getImagePath(image_id, format);
    
    bool success = writeFile(internal_path, image_data);
    if (success) {
        FASTEXCEL_LOG_INFO("Added image file: {} ({} bytes)", internal_path, image_data.size());
    } else {
        FASTEXCEL_LOG_ERROR("Failed to add image file: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addImageFile(const std::string& image_id,
                               std::vector<uint8_t>&& image_data,
                               core::ImageFormat format) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (image_data.empty()) {
        FASTEXCEL_LOG_ERROR("Image data is empty for image: {}", image_id);
        return false;
    }
    
    std::string internal_path = getImagePath(image_id, format);
    size_t data_size = image_data.size();
    
    bool success = writeFile(internal_path, image_data);
    if (success) {
        FASTEXCEL_LOG_INFO("Added image file: {} ({} bytes)", internal_path, data_size);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to add image file: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addImageFile(const core::Image& image) {
    if (!image.isValid()) {
        FASTEXCEL_LOG_ERROR("Invalid image object");
        return false;
    }
    
    return addImageFile(image.getId(), image.getData(), image.getFormat());
}

int FileManager::addImageFiles(const std::vector<std::unique_ptr<core::Image>>& images) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return 0;
    }
    
    int success_count = 0;
    
    for (const auto& image : images) {
        if (image && image->isValid()) {
            if (addImageFile(*image)) {
                success_count++;
            }
        }
    }
    
    FASTEXCEL_LOG_INFO("Added {} out of {} image files", success_count, images.size());
    return success_count;
}

bool FileManager::addDrawingXML(int drawing_id, const std::string& xml_content) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (xml_content.empty()) {
        FASTEXCEL_LOG_ERROR("Drawing XML content is empty for drawing: {}", drawing_id);
        return false;
    }
    
    std::string internal_path = getDrawingPath(drawing_id);
    
    bool success = writeFile(internal_path, xml_content);
    if (success) {
        FASTEXCEL_LOG_INFO("Added drawing XML: {} ({} bytes)", internal_path, xml_content.size());
    } else {
        FASTEXCEL_LOG_ERROR("Failed to add drawing XML: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addDrawingRelsXML(int drawing_id, const std::string& xml_content) {
    if (!isOpen()) {
        FASTEXCEL_LOG_ERROR("Archive not open");
        return false;
    }
    
    if (xml_content.empty()) {
        FASTEXCEL_LOG_ERROR("Drawing relationships XML content is empty for drawing: {}", drawing_id);
        return false;
    }
    
    std::string internal_path = getDrawingRelsPath(drawing_id);
    
    bool success = writeFile(internal_path, xml_content);
    if (success) {
        FASTEXCEL_LOG_INFO("Added drawing relationships XML: {} ({} bytes)", internal_path, xml_content.size());
    } else {
        FASTEXCEL_LOG_ERROR("Failed to add drawing relationships XML: {}", internal_path);
    }
    
    return success;
}

bool FileManager::imageExists(const std::string& image_id, core::ImageFormat format) const {
    std::string internal_path = getImagePath(image_id, format);
    return fileExists(internal_path);
}

std::string FileManager::getImagePath(const std::string& image_id, core::ImageFormat format) {
    std::string extension;
    switch (format) {
        case core::ImageFormat::PNG:  extension = "png"; break;
        case core::ImageFormat::JPEG: extension = "jpg"; break;
        case core::ImageFormat::GIF:  extension = "gif"; break;
        case core::ImageFormat::BMP:  extension = "bmp"; break;
        default: extension = "bin"; break;
    }
    
    return fmt::format("xl/media/{}.{}", image_id, extension);
}

std::string FileManager::getDrawingPath(int drawing_id) {
    return fmt::format("xl/drawings/drawing{}.xml", drawing_id);
}

std::string FileManager::getDrawingRelsPath(int drawing_id) {
    return fmt::format("xl/drawings/_rels/drawing{}.xml.rels", drawing_id);
}

}} // namespace fastexcel::archive
