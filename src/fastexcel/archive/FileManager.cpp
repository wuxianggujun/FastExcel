#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/xml/ContentTypes.hpp"
#include "fastexcel/xml/Relationships.hpp"

#include <chrono>
#include <iomanip>
#include <sstream>
#include <algorithm>
#include <fstream>

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
        ARCHIVE_ERROR("Failed to open archive: {}", filename_);
        return false;
    }
    
    // 不在这里创建Excel结构，而是在Workbook::save()时创建
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
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::writeFile(const std::string& internal_path, const std::vector<uint8_t>& data) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, data.data(), data.size()) == ZipError::Ok;
}

bool FileManager::writeFiles(const std::vector<std::pair<std::string, std::string>>& files) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
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
    
    ARCHIVE_INFO("Writing {} files in batch mode", files.size());
    return archive_->addFiles(zip_files) == ZipError::Ok;
}

bool FileManager::writeFiles(std::vector<std::pair<std::string, std::string>>&& files) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
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
    
    ARCHIVE_INFO("Writing {} files in batch mode (move semantics)", zip_files.size());
    return archive_->addFiles(std::move(zip_files)) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::string& content) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    return archive_->extractFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::vector<uint8_t>& data) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
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

bool FileManager::createExcelStructure() {
    // 创建Excel文件所需的基本结构
    if (!addContentTypes()) {
        ARCHIVE_ERROR("Failed to add content types");
        return false;
    }
    
    if (!addRootRels()) {
        ARCHIVE_ERROR("Failed to add root relationships");
        return false;
    }
    
    if (!addDocProps()) {
        ARCHIVE_ERROR("Failed to add document properties");
        return false;
    }
    
    if (!addWorkbookRels()) {
        ARCHIVE_ERROR("Failed to add workbook relationships");
        return false;
    }
    
    ARCHIVE_INFO("Excel file structure created successfully");
    return true;
}

bool FileManager::addContentTypes() {
    // 不使用addExcelDefaults()，而是让Workbook类动态生成正确的Content_Types.xml
    // 这个方法现在只是一个占位符，实际内容由Workbook类的generateContentTypesXML生成
    return true;
}

bool FileManager::addRootRels() {
    // 不在这里生成，让Workbook类动态生成正确的_rels/.rels
    // 这个方法现在只是一个占位符，实际内容由Workbook类的generateRelsXML生成
    return true;
}

bool FileManager::addWorkbookRels() {
    // 不在这里生成，让Workbook类动态生成正确的xl/_rels/workbook.xml.rels
    // 这个方法现在只是一个占位符，实际内容由Workbook类的generateWorkbookRelsXML生成
    return true;
}

bool FileManager::addDocProps() {
    // 创建核心属性
    std::ostringstream core_props;
    core_props << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
    core_props << "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ";
    core_props << "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" ";
    core_props << "xmlns:dcterms=\"http://purl.org/dc/terms/\" ";
    core_props << "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ";
    core_props << "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n";
    core_props << "  <dc:creator>FastExcel</dc:creator>\r\n";
    core_props << "  <cp:lastModifiedBy>FastExcel</cp:lastModifiedBy>\r\n";
    core_props << "  <dcterms:created xsi:type=\"dcterms:W3CDTF\">";
    
    // 添加当前UTC时间（统一封装自 TimeUtils）
    auto tm_utc = ::fastexcel::utils::TimeUtils::getCurrentUTCTime();
    core_props << ::fastexcel::utils::TimeUtils::formatTimeISO8601(tm_utc);
    
    core_props << "</dcterms:created>\r\n";
    core_props << "  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">";
    core_props << ::fastexcel::utils::TimeUtils::formatTimeISO8601(tm_utc);
    core_props << "</dcterms:modified>\r\n";
    core_props << "</cp:coreProperties>\r\n";
    
    if (!writeFile("docProps/core.xml", core_props.str())) {
        return false;
    }
    
    // 创建扩展属性
    std::ostringstream app_props;
    app_props << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
    app_props << "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ";
    app_props << "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\r\n";
    app_props << "  <Application>Microsoft Excel</Application>\r\n";
    app_props << "  <DocSecurity>0</DocSecurity>\r\n";
    app_props << "  <ScaleCrop>false</ScaleCrop>\r\n";
    app_props << "  <LinksUpToDate>false</LinksUpToDate>\r\n";
    app_props << "  <SharedDoc>false</SharedDoc>\r\n";
    app_props << "  <HyperlinksChanged>false</HyperlinksChanged>\r\n";
    app_props << "  <AppVersion>1.0</AppVersion>\r\n";
    app_props << "</Properties>\r\n";
    
    return writeFile("docProps/app.xml", app_props.str());
}

bool FileManager::setCompressionLevel(int level) {
    if (!archive_) {
        ARCHIVE_ERROR("Archive not initialized");
        return false;
    }
    
    return archive_->setCompressionLevel(level) == ZipError::Ok;
}

// 流式写入方法实现 - 极致性能模式
bool FileManager::openStreamingFile(const std::string& internal_path) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->openEntry(internal_path);
    if (result != ZipError::Ok) {
        ARCHIVE_ERROR("Failed to open streaming entry: {}", internal_path);
        return false;
    }
    
    ARCHIVE_DEBUG("Opened streaming file: {}", internal_path);
    return true;
}

bool FileManager::writeStreamingChunk(const void* data, size_t size) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->writeChunk(data, size);
    if (result != ZipError::Ok) {
        ARCHIVE_ERROR("Failed to write streaming chunk of size {}", size);
        return false;
    }
    
    return true;
}

bool FileManager::writeStreamingChunk(const std::string& data) {
    return writeStreamingChunk(data.c_str(), data.size());
}

bool FileManager::closeStreamingFile() {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->closeEntry();
    if (result != ZipError::Ok) {
        ARCHIVE_ERROR("Failed to close streaming entry");
        return false;
    }
    
    ARCHIVE_DEBUG("Closed streaming file");
    return true;
}

bool FileManager::copyFromExistingPackage(const core::Path& source_package,
                                          const std::vector<std::string>& skip_prefixes) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open for writing when copying from existing package");
        return false;
    }

    // 打开源包用于读取
    ZipArchive src(source_package);
    if (!src.open(false)) {
        ARCHIVE_ERROR("Failed to open source package for copy: {}", source_package.string());
        return false;
    }

    auto paths = src.listFiles();
    ARCHIVE_INFO("Copy-through existing entries: {} files to scan", paths.size());

    for (const auto& p : paths) {
        bool skip = false;
        for (const auto& prefix : skip_prefixes) {
            if (!prefix.empty() && p.rfind(prefix, 0) == 0) { // starts with
                skip = true;
                break;
            }
        }
        if (skip) {
            ARCHIVE_DEBUG("Skip passthrough: {}", p);
            continue;
        }

        std::vector<uint8_t> data;
        if (src.extractFile(p, data) == ZipError::Ok) {
            if (archive_->fileExists(p) == ZipError::Ok) {
                ARCHIVE_DEBUG("Target already has {}, skipping overwrite", p);
                continue;
            }
            if (archive_->addFile(p, data.data(), data.size()) != ZipError::Ok) {
                ARCHIVE_ERROR("Failed to write passthrough entry: {}", p);
                src.close();
                return false;
            }
            ARCHIVE_DEBUG("Pass-through copied: {} ({} bytes)", p, data.size());
        } else {
            ARCHIVE_WARN("Failed to extract entry from source for passthrough: {}", p);
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
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    if (image_data.empty()) {
        ARCHIVE_ERROR("Image data is empty for image: {}", image_id);
        return false;
    }
    
    std::string internal_path = getImagePath(image_id, format);
    
    bool success = writeFile(internal_path, image_data);
    if (success) {
        ARCHIVE_INFO("Added image file: {} ({} bytes)", internal_path, image_data.size());
    } else {
        ARCHIVE_ERROR("Failed to add image file: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addImageFile(const std::string& image_id,
                               std::vector<uint8_t>&& image_data,
                               core::ImageFormat format) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    if (image_data.empty()) {
        ARCHIVE_ERROR("Image data is empty for image: {}", image_id);
        return false;
    }
    
    std::string internal_path = getImagePath(image_id, format);
    size_t data_size = image_data.size();
    
    bool success = writeFile(internal_path, image_data);
    if (success) {
        ARCHIVE_INFO("Added image file: {} ({} bytes)", internal_path, data_size);
    } else {
        ARCHIVE_ERROR("Failed to add image file: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addImageFile(const core::Image& image) {
    if (!image.isValid()) {
        ARCHIVE_ERROR("Invalid image object");
        return false;
    }
    
    return addImageFile(image.getId(), image.getData(), image.getFormat());
}

int FileManager::addImageFiles(const std::vector<std::unique_ptr<core::Image>>& images) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
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
    
    ARCHIVE_INFO("Added {} out of {} image files", success_count, images.size());
    return success_count;
}

bool FileManager::addDrawingXML(int drawing_id, const std::string& xml_content) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    if (xml_content.empty()) {
        ARCHIVE_ERROR("Drawing XML content is empty for drawing: {}", drawing_id);
        return false;
    }
    
    std::string internal_path = getDrawingPath(drawing_id);
    
    bool success = writeFile(internal_path, xml_content);
    if (success) {
        ARCHIVE_INFO("Added drawing XML: {} ({} bytes)", internal_path, xml_content.size());
    } else {
        ARCHIVE_ERROR("Failed to add drawing XML: {}", internal_path);
    }
    
    return success;
}

bool FileManager::addDrawingRelsXML(int drawing_id, const std::string& xml_content) {
    if (!isOpen()) {
        ARCHIVE_ERROR("Archive not open");
        return false;
    }
    
    if (xml_content.empty()) {
        ARCHIVE_ERROR("Drawing relationships XML content is empty for drawing: {}", drawing_id);
        return false;
    }
    
    std::string internal_path = getDrawingRelsPath(drawing_id);
    
    bool success = writeFile(internal_path, xml_content);
    if (success) {
        ARCHIVE_INFO("Added drawing relationships XML: {} ({} bytes)", internal_path, xml_content.size());
    } else {
        ARCHIVE_ERROR("Failed to add drawing relationships XML: {}", internal_path);
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
    
    return "xl/media/" + image_id + "." + extension;
}

std::string FileManager::getDrawingPath(int drawing_id) {
    return "xl/drawings/drawing" + std::to_string(drawing_id) + ".xml";
}

std::string FileManager::getDrawingRelsPath(int drawing_id) {
    return "xl/drawings/_rels/drawing" + std::to_string(drawing_id) + ".xml.rels";
}

}} // namespace fastexcel::archive
