#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/xml/ContentTypes.hpp"
#include "fastexcel/xml/Relationships.hpp"

#include <chrono>
#include <iomanip>
#include <sstream>

namespace fastexcel {
namespace archive {

FileManager::FileManager(const std::string& filename) : filename_(filename) {
}

FileManager::~FileManager() {
    close();
}

bool FileManager::open(bool create) {
    if (isOpen()) {
        close();
    }
    
    archive_ = std::make_unique<ZipArchive>(filename_);
    if (!archive_->open(create)) {
        LOG_ERROR("Failed to open archive: {}", filename_);
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
        LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::writeFile(const std::string& internal_path, const std::vector<uint8_t>& data) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->addFile(internal_path, data.data(), data.size()) == ZipError::Ok;
}

bool FileManager::writeFiles(const std::vector<std::pair<std::string, std::string>>& files) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
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
    
    LOG_INFO("Writing {} files in batch mode", files.size());
    return archive_->addFiles(zip_files) == ZipError::Ok;
}

bool FileManager::writeFiles(std::vector<std::pair<std::string, std::string>>&& files) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
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
    
    LOG_INFO("Writing {} files in batch mode (move semantics)", zip_files.size());
    return archive_->addFiles(std::move(zip_files)) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::string& content) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
        return false;
    }
    
    return archive_->extractFile(internal_path, content) == ZipError::Ok;
}

bool FileManager::readFile(const std::string& internal_path, std::vector<uint8_t>& data) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
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
        LOG_ERROR("Failed to add content types");
        return false;
    }
    
    if (!addRootRels()) {
        LOG_ERROR("Failed to add root relationships");
        return false;
    }
    
    if (!addDocProps()) {
        LOG_ERROR("Failed to add document properties");
        return false;
    }
    
    if (!addWorkbookRels()) {
        LOG_ERROR("Failed to add workbook relationships");
        return false;
    }
    
    LOG_INFO("Excel file structure created successfully");
    return true;
}

bool FileManager::addContentTypes() {
    xml::ContentTypes content_types;
    content_types.addExcelDefaults();
    
    // 关键修复：先生成内容到字符串，确保文件大小正确
    std::string xml_content;
    content_types.generate([&xml_content](const char* data, size_t size) {
        xml_content.append(data, size);
    });
    
    // 使用标准writeFile方法，确保ZIP结构正确
    return writeFile("[Content_Types].xml", xml_content);
}

bool FileManager::addRootRels() {
    xml::Relationships rels;
    // 修改顺序以匹配修复后的文件：rId3, rId2, rId1
    rels.addRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
    rels.addRelationship("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
    rels.addRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
    
    // 关键修复：先生成内容到字符串，确保文件大小正确
    std::string xml_content;
    rels.generate([&xml_content](const char* data, size_t size) {
        xml_content.append(data, size);
    });
    
    // 使用标准writeFile方法，确保ZIP结构正确
    return writeFile("_rels/.rels", xml_content);
}

bool FileManager::addWorkbookRels() {
    xml::Relationships rels;
    // 修复：添加完整的关系链，包括theme
    // 按照Workbook类中generateWorkbookRelsXML的顺序：rId3(theme), rId2(sheet2), rId1(sheet1), rId5(sharedStrings), rId4(styles)
    rels.addRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml");
    rels.addRelationship("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet2.xml");
    rels.addRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml");
    rels.addRelationship("rId5", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
    rels.addRelationship("rId4", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
    
    // 关键修复：先生成内容到字符串，确保文件大小正确
    std::string xml_content;
    rels.generate([&xml_content](const char* data, size_t size) {
        xml_content.append(data, size);
    });
    
    // 使用标准writeFile方法，确保ZIP结构正确
    return writeFile("xl/_rels/workbook.xml.rels", xml_content);
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
    
    // 添加当前时间
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    std::tm tm;
    localtime_s(&tm, &time_t);
    core_props << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
    
    core_props << "</dcterms:created>\r\n";
    core_props << "  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">";
    core_props << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
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
        LOG_ERROR("Archive not initialized");
        return false;
    }
    
    return archive_->setCompressionLevel(level) == ZipError::Ok;
}

// 流式写入方法实现 - 极致性能模式
bool FileManager::openStreamingFile(const std::string& internal_path) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->openEntry(internal_path);
    if (result != ZipError::Ok) {
        LOG_ERROR("Failed to open streaming entry: {}", internal_path);
        return false;
    }
    
    LOG_DEBUG("Opened streaming file: {}", internal_path);
    return true;
}

bool FileManager::writeStreamingChunk(const void* data, size_t size) {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->writeChunk(data, size);
    if (result != ZipError::Ok) {
        LOG_ERROR("Failed to write streaming chunk of size {}", size);
        return false;
    }
    
    return true;
}

bool FileManager::writeStreamingChunk(const std::string& data) {
    return writeStreamingChunk(data.c_str(), data.size());
}

bool FileManager::closeStreamingFile() {
    if (!isOpen()) {
        LOG_ERROR("Archive not open");
        return false;
    }
    
    ZipError result = archive_->closeEntry();
    if (result != ZipError::Ok) {
        LOG_ERROR("Failed to close streaming entry");
        return false;
    }
    
    LOG_DEBUG("Closed streaming file");
    return true;
}

}} // namespace fastexcel::archive
