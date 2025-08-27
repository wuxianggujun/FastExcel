/**
 * @file ReadOnlyXLSXReader.cpp  
 * @brief 只读模式专用XLSX解析器实现
 */

#include "fastexcel/reader/ReadOnlyXLSXReader.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/reader/SharedStringsParser.hpp"
#include "fastexcel/reader/RelationshipsParser.hpp"
#include "fastexcel/reader/ReadOnlyWorksheetParser.hpp"
#include "fastexcel/reader/ReadOnlyWorkbookParser.hpp"
#include <sstream>

namespace fastexcel {
namespace reader {

ReadOnlyXLSXReader::ReadOnlyXLSXReader(const core::Path& file_path, 
                                      const core::WorkbookOptions* options)
    : path_(file_path), options_(options) {
    if (!options_) {
        static core::WorkbookOptions default_options;
        options_ = &default_options;
    }
}

core::ErrorCode ReadOnlyXLSXReader::parse() {
    try {
        if (!path_.exists()) {
            FASTEXCEL_LOG_ERROR("XLSX file not found: {}", path_.string());
            return core::ErrorCode::FileNotFound;
        }
        
        FASTEXCEL_LOG_INFO("Starting read-only XLSX parsing: {}", path_.string());
        
        // 打开ZIP文件
        zip_archive_ = std::make_unique<archive::ZipArchive>(path_);
        if (!zip_archive_->open(false)) { // false = 读模式
            FASTEXCEL_LOG_ERROR("Failed to open XLSX file as ZIP archive: {}", path_.string());
            return core::ErrorCode::FileReadError;
        }
        
        // 1. 解析共享字符串表
        auto result = parseSharedStrings();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to parse shared strings");
            return result;
        }
        
        // 2. 解析工作簿结构和工作表
        result = parseWorkbook();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to parse workbook structure");
            return result;
        }
        
        FASTEXCEL_LOG_INFO("Successfully parsed {} worksheets in read-only mode", 
                          worksheet_infos_.size());
        
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during read-only XLSX parsing: {}", e.what());
        return core::ErrorCode::InternalError;
    }
}

std::string ReadOnlyXLSXReader::extractXMLFromZip(const std::string& path) {
    if (!zip_archive_) {
        return "";
    }
    
    std::string content;
    auto error = zip_archive_->extractFile(path, content);
    if (archive::isError(error)) {
        FASTEXCEL_LOG_DEBUG("Failed to extract {}: {}", path, static_cast<int>(error));
        return "";
    }
    
    return content;
}

core::ErrorCode ReadOnlyXLSXReader::parseSharedStrings() {
    try {
        // 检查共享字符串文件是否存在
        auto error = zip_archive_->fileExists("xl/sharedStrings.xml");
        if (archive::isError(error)) {
            // 共享字符串文件不存在，这是正常的
            FASTEXCEL_LOG_INFO("No shared string table found, continuing without SST");
            return core::ErrorCode::Ok;
        }
        
        if (!zip_archive_ || !zip_archive_->getReader()) {
            FASTEXCEL_LOG_ERROR("ZIP存档或读取器不可用");
            return core::ErrorCode::InternalError;
        }
        
        SharedStringsParser parser;
        auto* zr = zip_archive_->getReader();
        
        if (!parser.parseStream(zr, "xl/sharedStrings.xml")) {
            FASTEXCEL_LOG_ERROR("共享字符串流式解析失败");
            return core::ErrorCode::XmlParseError;
        }
        
        shared_strings_ = parser.getStrings();
        FASTEXCEL_LOG_INFO("Loaded {} shared strings", shared_strings_.size());
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception parsing shared strings: {}", e.what());
        return core::ErrorCode::InternalError;
    }
}

core::ErrorCode ReadOnlyXLSXReader::parseWorkbook() {
    try {
        // 1. 首先使用RelationshipsParser解析关系文件
        std::unordered_map<std::string, std::string> relationships;
        std::string rels_content = extractXMLFromZip("xl/_rels/workbook.xml.rels");
        if (!rels_content.empty()) {
            RelationshipsParser rels_parser;
            if (rels_parser.parse(rels_content)) {
                // 提取工作表相关的关系
                for (const auto& rel : rels_parser.getRelationships()) {
                    if (rel.type.find("worksheet") != std::string::npos) {
                        relationships[rel.id] = rel.target;
                    }
                }
                FASTEXCEL_LOG_DEBUG("解析到 {} 个工作表关系", relationships.size());
            } else {
                FASTEXCEL_LOG_WARN("关系文件解析失败，使用默认路径");
            }
        }
        
        // 2. 使用ReadOnlyWorkbookParser解析工作簿
        std::string workbook_content = extractXMLFromZip("xl/workbook.xml");
        if (workbook_content.empty()) {
            FASTEXCEL_LOG_ERROR("Failed to extract workbook.xml");
            return core::ErrorCode::FileReadError;
        }
        
        ReadOnlyWorkbookParser workbook_parser;
        workbook_parser.setRelationships(std::move(relationships));
        
        if (!workbook_parser.parseXML(workbook_content)) {
            FASTEXCEL_LOG_ERROR("Failed to parse workbook.xml");
            return core::ErrorCode::XmlParseError;
        }
        
        auto sheets = workbook_parser.takeSheets();
        FASTEXCEL_LOG_INFO("发现 {} 个工作表", sheets.size());
        
        // 3. 为每个工作表创建ColumnarStorageManager并解析数据
        for (auto& sheet : sheets) {
            auto storage_manager = std::make_shared<core::ColumnarStorageManager>();
            storage_manager->enableColumnarStorage(options_);
            
            // 使用ReadOnlyWorksheetParser解析工作表
            auto result = parseWorksheet(sheet.worksheet_path, sheet.name, storage_manager);
            if (result != core::ErrorCode::Ok) {
                FASTEXCEL_LOG_WARN("Failed to parse worksheet '{}', skipping", sheet.name);
                continue;
            }
            
            // 解析完成后获取实际的数据范围
            uint32_t first_row = 1, last_row = 1, first_col = 1, last_col = 1;
            if (storage_manager->hasData()) {
                first_row = storage_manager->getFirstRow();
                last_row = storage_manager->getLastRow();
                first_col = storage_manager->getFirstColumn();
                last_col = storage_manager->getLastColumn();
            }
            
            worksheet_infos_.emplace_back(
                std::move(sheet.name),  // 正确使用 move 避免字符串复制
                storage_manager, 
                static_cast<int>(first_row), static_cast<int>(first_col),
                static_cast<int>(last_row), static_cast<int>(last_col)
            );
        }
        
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception parsing workbook: {}", e.what());
        return core::ErrorCode::InternalError;
    }
}

core::ErrorCode ReadOnlyXLSXReader::parseWorksheet(const std::string& worksheet_path,
                                                   const std::string& worksheet_name,
                                                   std::shared_ptr<core::ColumnarStorageManager> storage) {
    try {
        std::string worksheet_content = extractXMLFromZip(worksheet_path);
        
        if (worksheet_content.empty()) {
            FASTEXCEL_LOG_ERROR("Failed to extract worksheet: {}", worksheet_path);
            return core::ErrorCode::FileReadError;
        }
        
        // 使用专用的只读工作表流式解析器
        ReadOnlyWorksheetParser worksheet_parser;
        worksheet_parser.configure(storage, &shared_strings_, options_);
        
        if (!worksheet_parser.parseXML(worksheet_content)) {
            FASTEXCEL_LOG_ERROR("Failed to parse worksheet '{}' with streaming parser", worksheet_name);
            return core::ErrorCode::XmlParseError;
        }
        
        size_t cells_processed = worksheet_parser.getCellsProcessed();
        FASTEXCEL_LOG_INFO("Processed {} cells from worksheet '{}' using streaming parser", 
                          cells_processed, worksheet_name);
        
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception parsing worksheet '{}': {}", worksheet_name, e.what());
        return core::ErrorCode::InternalError;
    }
}

}} // namespace fastexcel::reader