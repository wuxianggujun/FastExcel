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
    
    std::string content = content_types.generate();
    return writeFile("[Content_Types].xml", content);
}

bool FileManager::addWorkbookRels() {
    xml::Relationships rels;
    rels.addRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
    rels.addRelationship("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
    rels.addRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
    
    std::string content = rels.generate();
    return writeFile("_rels/.rels", content);
}

bool FileManager::addRootRels() {
    xml::Relationships rels;
    rels.addRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
    rels.addRelationship("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
    rels.addRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml");
    
    std::string content = rels.generate();
    return writeFile("xl/_rels/workbook.xml.rels", content);
}

bool FileManager::addDocProps() {
    // 创建核心属性
    std::ostringstream core_props;
    core_props << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    core_props << "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ";
    core_props << "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" ";
    core_props << "xmlns:dcterms=\"http://purl.org/dc/terms/\" ";
    core_props << "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ";
    core_props << "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n";
    core_props << "  <dc:creator>FastExcel</dc:creator>\n";
    core_props << "  <cp:lastModifiedBy>FastExcel</cp:lastModifiedBy>\n";
    core_props << "  <dcterms:created xsi:type=\"dcterms:W3CDTF\">";
    
    // 添加当前时间
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    std::tm tm;
    localtime_s(&tm, &time_t);
    core_props << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
    
    core_props << "</dcterms:created>\n";
    core_props << "  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">";
    core_props << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
    core_props << "</dcterms:modified>\n";
    core_props << "</cp:coreProperties>\n";
    
    if (!writeFile("docProps/core.xml", core_props.str())) {
        return false;
    }
    
    // 创建扩展属性
    std::ostringstream app_props;
    app_props << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    app_props << "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ";
    app_props << "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n";
    app_props << "  <Application>FastExcel</Application>\n";
    app_props << "  <DocSecurity>0</DocSecurity>\n";
    app_props << "  <ScaleCrop>false</ScaleCrop>\n";
    app_props << "  <LinksUpToDate>false</LinksUpToDate>\n";
    app_props << "  <SharedDoc>false</SharedDoc>\n";
    app_props << "  <HyperlinksChanged>false</HyperlinksChanged>\n";
    app_props << "  <AppVersion>1.0</AppVersion>\n";
    app_props << "</Properties>\n";
    
    return writeFile("docProps/app.xml", app_props.str());
}

}} // namespace fastexcel::archive
