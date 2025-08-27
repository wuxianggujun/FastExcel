#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#include "XLSXReader.hpp"
#include "SharedStringsParser.hpp"
#include "RelationshipsParser.hpp"
#include "ContentTypesParser.hpp"
#include "WorkbookParser.hpp"
#include "StylesParser.hpp"
#include "WorksheetParser.hpp"
#include "DocPropsParser.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/theme/ThemeParser.hpp"
#include <iostream>
#include <sstream>

namespace fastexcel {
namespace reader {

// 构造函数
XLSXReader::XLSXReader(const std::string& filename)
    : filepath_(filename)
    , filename_(filename)
    , zip_archive_(std::make_unique<archive::ZipArchive>(filepath_))
    , is_open_(false) {
}

// Path构造函数
XLSXReader::XLSXReader(const core::Path& path)
    : filepath_(path)
    , filename_(path.string())
    , zip_archive_(std::make_unique<archive::ZipArchive>(filepath_))
    , is_open_(false) {
}

// 析构函数
XLSXReader::~XLSXReader() {
    if (is_open_) {
        close();
    }
}

// 打开XLSX文件 - 系统层高性能API
core::ErrorCode XLSXReader::open() {
    if (is_open_) {
        return core::ErrorCode::Ok;
    }
    
    try {
        // 打开ZIP文件进行读取
        if (!zip_archive_->open(false)) {  // false表示不创建新文件，只读取
            FASTEXCEL_LOG_ERROR("无法打开XLSX文件: {}", filename_);
            return core::ErrorCode::FileAccessDenied;
        }
        
        // 验证XLSX文件结构
        if (!validateXLSXStructure()) {
            FASTEXCEL_LOG_ERROR("无效的XLSX文件格式: {}", filename_);
            zip_archive_->close();
            return core::ErrorCode::XmlInvalidFormat;
        }
        
        is_open_ = true;
        FASTEXCEL_LOG_INFO("成功打开XLSX文件: {}", filename_);
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("打开XLSX文件时发生异常: {}", e.what());
        return core::ErrorCode::InternalError;
    }
}

// 关闭文件 - 系统层高性能API
core::ErrorCode XLSXReader::close() {
    if (!is_open_) {
        return core::ErrorCode::Ok;
    }
    
    try {
        zip_archive_->close();
        is_open_ = false;
        
        // 清理缓存数据
        worksheet_names_.clear();
        defined_names_.clear();
        worksheet_paths_.clear();
        shared_strings_.clear();
        styles_.clear();
        theme_xml_.clear();
        parsed_theme_.reset();
        
        FASTEXCEL_LOG_INFO("成功关闭XLSX文件: {}", filename_);
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("关闭XLSX文件时发生异常: {}", e.what());
        return core::ErrorCode::InternalError;
    }
}

// 加载整个工作簿 - 系统层高性能API
core::ErrorCode XLSXReader::loadWorkbook(std::unique_ptr<core::Workbook>& workbook) {
    if (!is_open_) {
        FASTEXCEL_LOG_ERROR("文件未打开，无法加载工作簿");
        return core::ErrorCode::InvalidArgument;
    }
    
    try {
        // 创建内存工作簿容器（绝不创建文件）
        core::Path memory_path("::memory::reader_" + std::to_string(reinterpret_cast<uintptr_t>(this)));
        workbook = std::make_unique<core::Workbook>(memory_path);
        
        // 确保内存工作簿正确初始化
        if (!workbook->open()) {
            FASTEXCEL_LOG_ERROR("无法初始化内存工作簿");
            return core::ErrorCode::InternalError;
        }
        
        FASTEXCEL_LOG_INFO("开始解析XLSX文件结构: {}", filename_);
        
        // 解析各种XML文件，使用错误码检查
        auto result = parseSharedStringsXML();
        if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
            FASTEXCEL_LOG_WARN("解析共享字符串失败，错误码: {}", static_cast<int>(result));
            // 共享字符串不是必需的，继续执行
        }
        
        result = parseStylesXML();
        if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
            FASTEXCEL_LOG_WARN("解析样式失败，错误码: {}", static_cast<int>(result));
            // 样式解析失败不影响主要功能，继续执行
        }
        
        // 解析主题文件以保持原始样式
        result = parseThemeXML();
        if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
            FASTEXCEL_LOG_WARN("解析主题失败，错误码: {}", static_cast<int>(result));
            // 主题解析失败不影响主要功能，继续执行
        }
        
        result = parseWorkbookXML();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("解析工作簿结构失败，错误码: {}", static_cast<int>(result));
            return result;
        }
        
        result = parseDocPropsXML();
        if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
            FASTEXCEL_LOG_WARN("解析文档属性失败，错误码: {}", static_cast<int>(result));
            // 文档属性解析失败不影响主要功能，继续执行
        }
        
        // 利用现有的XMLStreamReader解析工作表
        // 直接在内存中构建数据结构，无需临时文件
        for (const auto& sheet_name : worksheet_names_) {
            try {
                auto worksheet = workbook->addSheet(sheet_name);
                if (worksheet) {
                    auto it = worksheet_paths_.find(sheet_name);
                    if (it != worksheet_paths_.end()) {
                        // 使用流式XML解析器处理工作表数据
                        auto parse_result = parseWorksheetXML(it->second, worksheet.get());
                        if (parse_result != core::ErrorCode::Ok) {
                            FASTEXCEL_HANDLE_WARNING(
                                "解析工作表失败: " + sheet_name, "loadWorkbook");
                            // 继续处理其他工作表
                        }
                    } else {
                        FASTEXCEL_HANDLE_WARNING(
                            "工作表路径未找到: " + sheet_name, "loadWorkbook");
                    }
                } else {
                    FASTEXCEL_HANDLE_WARNING(
                        "无法创建工作表: " + sheet_name, "loadWorkbook");
                }
            } catch (const core::FastExcelException& e) {
                FASTEXCEL_HANDLE_WARNING(
                    "处理工作表时发生异常: " + sheet_name + " - " + e.what(), 
                    "loadWorkbook");
            }
        }
        
        // 将解析的 FormatDescriptor 样式导入到工作簿的 FormatRepository 中（保持原始 ID）
        if (!styles_.empty()) {
            FASTEXCEL_LOG_DEBUG("开始导入 {} 个FormatDescriptor样式到工作簿格式仓储", styles_.size());
            
            // 获取工作簿的格式仓储
            auto& format_repo = const_cast<core::FormatRepository&>(workbook->getStyles());
            
            // 清空之前的映射
            style_id_mapping_.clear();
            
            // 尽量保持原始 ID 映射，避免样式 ID 重新分配
            int imported_count = 0;
            for (const auto& style_pair : styles_) {
                int original_style_id = style_pair.first;
                const auto& format_desc = style_pair.second;
                
                if (format_desc) {
                    // 尝试直接使用原始ID添加格式（如果可能的话）
                    int new_id = format_repo.addFormat(*format_desc);
                    imported_count++;
                    
                    // 优先使用原始 ID 作为映射
                    // 如果新分配的ID与原始ID不同，我们仍然需要记录映射
                    // 但对于列样式，我们尝试保持一致性
                    if (original_style_id != new_id) {
                        style_id_mapping_[original_style_id] = new_id;
                        FASTEXCEL_LOG_DEBUG("样式ID重映射: {} -> {} (可能影响兼容性)", original_style_id, new_id);
                    } else {
                        // ID一致时不需要映射
                        FASTEXCEL_LOG_TRACE("样式ID保持不变: {}", original_style_id);
                    }
                }
            }
            
            FASTEXCEL_LOG_INFO("成功导入 {} 个FormatDescriptor样式到工作簿格式仓储", imported_count);
            FASTEXCEL_LOG_INFO("样式ID映射数量: {} (映射数越少表示兼容性越好)", style_id_mapping_.size());
        } else {
            FASTEXCEL_LOG_DEBUG("未检测到自定义样式，使用默认样式");
        }
        
        // 将解析的主题 XML 设置到工作簿，以保持原始主题
        if (!theme_xml_.empty()) {
            // 保真：保存原始主题XML到工作簿（不触发脏标记）
            workbook->setOriginalThemeXML(theme_xml_);
            // 同时，如果我们有解析对象，也顺带注入，便于后续编辑
            // setOriginalThemeXML 内部会解析为对象并缓存到工作簿，不需要再次 setTheme 以避免置脏
            FASTEXCEL_LOG_DEBUG("已注入原始主题XML到工作簿 ({} 字节)", theme_xml_.size());
        } else {
            FASTEXCEL_LOG_DEBUG("未检测到主题文件，使用默认主题");
        }
        
        FASTEXCEL_LOG_INFO("成功加载工作簿，包含 {} 个工作表", worksheet_names_.size());
        return core::ErrorCode::Ok;
        
    } catch (const core::FastExcelException& e) {
        FASTEXCEL_LOG_ERROR("加载工作簿时发生FastExcel异常: {}", e.getDetailedMessage());
        FASTEXCEL_HANDLE_ERROR(e);
        return core::ErrorCode::InternalError;
    } catch (const std::exception& e) {
        core::OperationException oe("加载工作簿时发生未知错误: " + std::string(e.what()), 
                                   "loadWorkbook");
        FASTEXCEL_LOG_ERROR("加载工作簿时发生标准异常: {}", e.what());
        FASTEXCEL_HANDLE_ERROR(oe);
        return core::ErrorCode::InternalError;
    }
}

// 加载单个工作表 - 系统层高性能API
core::ErrorCode XLSXReader::loadWorksheet(const std::string& name, std::shared_ptr<core::Worksheet>& worksheet) {
    if (!is_open_) {
        FASTEXCEL_LOG_ERROR("文件未打开，无法加载工作表");
        return core::ErrorCode::InvalidArgument;
    }
    
    try {
        // 如果还没有解析工作簿结构，先解析
        if (worksheet_names_.empty()) {
            auto result = parseWorkbookXML();
            if (result != core::ErrorCode::Ok) {
                FASTEXCEL_LOG_ERROR("解析工作簿结构失败，错误码: {}", static_cast<int>(result));
                return result;
            }
        }
        
        // 检查工作表是否存在
        auto it = worksheet_paths_.find(name);
        if (it == worksheet_paths_.end()) {
            FASTEXCEL_LOG_ERROR("工作表不存在: {}", name);
            return core::ErrorCode::InvalidWorksheet;
        }
        
        // 解析共享字符串（如果还没有解析）
        if (shared_strings_.empty()) {
            auto result = parseSharedStringsXML();
            if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
                FASTEXCEL_HANDLE_WARNING("解析共享字符串失败", "loadWorksheet");
            }
        }
        
        // 解析样式（如果还没有解析）
        if (styles_.empty()) {
            auto result = parseStylesXML();
            if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
                FASTEXCEL_HANDLE_WARNING("解析样式失败", "loadWorksheet");
            }
        }
        
        // 创建轻量级内存工作簿来容纳单个工作表
        core::Path memory_path("::memory::" + std::to_string(reinterpret_cast<uintptr_t>(this)) + "_" + name);
        auto temp_workbook = std::make_shared<core::Workbook>(memory_path);
        
        // 确保内存工作簿处于打开状态
        if (!temp_workbook->open()) {
            FASTEXCEL_LOG_ERROR("无法打开内存工作簿用于工作表: {}", name);
            return core::ErrorCode::InternalError;
        }
        
        worksheet = std::make_shared<core::Worksheet>(name, temp_workbook);
        
        // 解析工作表数据
        auto result = parseWorksheetXML(it->second, worksheet.get());
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("解析工作表失败: {}，错误码: {}", name, static_cast<int>(result));
            return result;
        }
        
        FASTEXCEL_LOG_INFO("成功加载工作表: {}", name);
        return core::ErrorCode::Ok;
        
    } catch (const core::FastExcelException& e) {
        FASTEXCEL_LOG_ERROR("加载工作表时发生FastExcel异常: {}", e.getDetailedMessage());
        FASTEXCEL_HANDLE_ERROR(e);
        return core::ErrorCode::InternalError;
    } catch (const std::exception& e) {
        core::WorksheetException we("加载工作表时发生未知错误: " + std::string(e.what()), name);
        FASTEXCEL_LOG_ERROR("加载工作表时发生标准异常: {}", e.what());
        FASTEXCEL_HANDLE_ERROR(we);
        return core::ErrorCode::InternalError;
    }
}

// 获取工作表名称列表 - 系统层高性能API
core::ErrorCode XLSXReader::getSheetNames(std::vector<std::string>& names) {
    if (!is_open_) {
        FASTEXCEL_HANDLE_WARNING("文件未打开，无法获取工作表名称", "getSheetNames");
        return core::ErrorCode::InvalidArgument;
    }
    
    // 如果还没有解析工作簿结构，先解析
    if (worksheet_names_.empty()) {
        auto result = parseWorkbookXML();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_HANDLE_WARNING("解析工作簿结构失败", "getSheetNames");
            return result;
        }
    }
    
    names = worksheet_names_;
    return core::ErrorCode::Ok;
}

// 获取元数据 - 系统层高性能API
core::ErrorCode XLSXReader::getMetadata(WorkbookMetadata& metadata) {
    if (!is_open_) {
        FASTEXCEL_HANDLE_WARNING("文件未打开，无法获取元数据", "getMetadata");
        return core::ErrorCode::InvalidArgument;
    }
    
    // 如果还没有解析文档属性，先解析
    if (metadata_.title.empty() && metadata_.author.empty()) {
        auto result = parseDocPropsXML();
        if (result != core::ErrorCode::Ok && result != core::ErrorCode::FileNotFound) {
            FASTEXCEL_HANDLE_WARNING("解析文档属性失败", "getMetadata");
            return result;
        }
    }
    
    metadata = metadata_;
    return core::ErrorCode::Ok;
}

// 获取定义名称列表 - 系统层高性能API
core::ErrorCode XLSXReader::getDefinedNames(std::vector<std::string>& names) {
    if (!is_open_) {
        FASTEXCEL_HANDLE_WARNING("文件未打开，无法获取定义名称", "getDefinedNames");
        return core::ErrorCode::InvalidArgument;
    }
    
    // 如果还没有解析工作簿结构，先解析
    if (defined_names_.empty()) {
        auto result = parseWorkbookXML();
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_HANDLE_WARNING("解析工作簿结构失败", "getDefinedNames");
            return result;
        }
    }
    
    names = defined_names_;
    return core::ErrorCode::Ok;
}

// 从ZIP中提取XML文件内容
std::string XLSXReader::extractXMLFromZip(const std::string& path) {
    std::string content;
    auto error = zip_archive_->extractFile(path, content);
    
    if (archive::isError(error)) {
        FASTEXCEL_LOG_ERROR("提取文件失败: {}", path);
        return "";
    }
    
    FASTEXCEL_LOG_DEBUG("成功提取XML文件: {} ({} bytes)", path, content.size());
    return content;
}

// 验证XLSX文件结构 - 使用ContentTypesParser
bool XLSXReader::validateXLSXStructure() {
    // 检查必需的文件是否存在
    std::vector<std::string> required_files = {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml"
    };
    
    for (const auto& file : required_files) {
        auto error = zip_archive_->fileExists(file);
        if (archive::isError(error)) {
            FASTEXCEL_LOG_ERROR("缺少必需文件: {}", file);
            return false;
        }
    }
    
    // 使用ContentTypesParser验证内容类型定义
    std::string content_types_xml = extractXMLFromZip("[Content_Types].xml");
    if (!content_types_xml.empty()) {
        ContentTypesParser parser;
        if (parser.parse(content_types_xml)) {
            // 验证必需的内容类型是否存在
            std::string workbook_type = parser.getContentType("/xl/workbook.xml");
            bool has_workbook = !workbook_type.empty();
            if (!has_workbook) {
                FASTEXCEL_LOG_WARN("Content_Types.xml中缺少工作簿内容类型定义");
            }
            FASTEXCEL_LOG_DEBUG("成功解析Content_Types.xml，包含{}个默认类型和{}个重写类型", 
                              parser.getDefaultCount(), parser.getOverrideCount());
        } else {
            FASTEXCEL_LOG_WARN("Content_Types.xml解析失败，但文件验证将继续");
        }
    }
    
    return true;
}

// 解析工作簿XML - 使用WorkbookParser
core::ErrorCode XLSXReader::parseWorkbookXML() {
    std::string xml_content = extractXMLFromZip("xl/workbook.xml");
    if (xml_content.empty()) {
        return core::ErrorCode::FileNotFound;
    }
    
    try {
        // 清理之前的数据
        worksheet_names_.clear();
        worksheet_paths_.clear();
        defined_names_.clear();
        
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
                FASTEXCEL_LOG_WARN("无法解析工作簿关系文件，使用默认路径");
            }
        }
        
        // 2. 使用WorkbookParser解析工作簿XML
        WorkbookParser workbook_parser;
        workbook_parser.setRelationships(std::move(relationships));
        
        if (!workbook_parser.parseXML(xml_content)) {
            FASTEXCEL_LOG_ERROR("工作簿XML解析失败");
            return core::ErrorCode::XmlParseError;
        }
        
        // 3. 提取解析结果
        auto worksheets = workbook_parser.takeWorksheets();
        for (const auto& ws : worksheets) {
            worksheet_names_.push_back(ws.name);
            worksheet_paths_[ws.name] = ws.worksheet_path;
        }
        
        defined_names_ = workbook_parser.getDefinedNameStrings();
        
        FASTEXCEL_LOG_INFO("成功解析工作簿: {} 个工作表, {} 个定义名称", 
                          worksheet_names_.size(), defined_names_.size());
        
        return worksheet_names_.empty() ? core::ErrorCode::XmlMissingElement : core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析工作簿XML时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

// 解析工作表XML - 系统层ErrorCode版本
core::ErrorCode XLSXReader::parseWorksheetXML(const std::string& path, core::Worksheet* worksheet) {
    if (!worksheet) {
        FASTEXCEL_LOG_ERROR("工作表对象为空");
        return core::ErrorCode::InvalidArgument;
    }
    
    if (!zip_archive_ || !zip_archive_->getReader()) {
        FASTEXCEL_LOG_ERROR("ZIP存档或读取器不可用");
        return core::ErrorCode::InternalError;
    }
    
    try {
        WorksheetParser parser;
        auto* zr = zip_archive_->getReader();
        
        if (!parser.parseStream(zr, path, worksheet, shared_strings_, styles_, style_id_mapping_, &options_)) {
            FASTEXCEL_LOG_ERROR("流式解析工作表失败: {}", path);
            return core::ErrorCode::XmlParseError;
        }
        
        FASTEXCEL_LOG_DEBUG("成功以流式解析工作表: {}", worksheet->getName());
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析工作表时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

core::ErrorCode XLSXReader::parseStylesXML() {
    // 检查样式文件是否存在
    auto error = zip_archive_->fileExists("xl/styles.xml");
    if (archive::isError(error)) {
        // 样式文件不存在，这是正常的（某些Excel文件可能没有自定义样式）
        return core::ErrorCode::FileNotFound;
    }
    
    std::string xml_content = extractXMLFromZip("xl/styles.xml");
    if (xml_content.empty()) {
        return core::ErrorCode::Ok; // 文件为空也是正常的
    }
    
    try {
        StylesParser parser;
        if (!parser.parse(xml_content)) {
            FASTEXCEL_LOG_ERROR("解析样式XML失败");
            return core::ErrorCode::XmlParseError;
        }
        
        // 将解析结果转换为styles_映射
        styles_.clear();
        for (size_t i = 0; i < parser.getFormatCount(); ++i) {
            auto format = parser.getFormat(static_cast<int>(i));
            if (format) {
                styles_[static_cast<int>(i)] = format;
            }
        }
        
        FASTEXCEL_LOG_DEBUG("成功解析 {} 个样式", styles_.size());
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析样式时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

core::ErrorCode XLSXReader::parseSharedStringsXML() {
    // 检查共享字符串文件是否存在
    auto error = zip_archive_->fileExists("xl/sharedStrings.xml");
    if (archive::isError(error)) {
        // 共享字符串文件不存在，这是正常的（某些Excel文件可能没有共享字符串）
        return core::ErrorCode::FileNotFound;
    }
    
    if (!zip_archive_ || !zip_archive_->getReader()) {
        FASTEXCEL_LOG_ERROR("ZIP存档或读取器不可用");
        return core::ErrorCode::InternalError;
    }
    
    try {
        SharedStringsParser parser;
        auto* zr = zip_archive_->getReader();
        
        if (!parser.parseStream(zr, "xl/sharedStrings.xml")) {
            FASTEXCEL_LOG_ERROR("共享字符串流式解析失败");
            return core::ErrorCode::XmlParseError;
        }
        
        shared_strings_ = parser.getStrings();
        FASTEXCEL_LOG_DEBUG("成功流式解析 {} 个共享字符串", shared_strings_.size());
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析共享字符串时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

core::ErrorCode XLSXReader::parseDocPropsXML() {
    try {
        DocPropsParser parser;
        
        // 尝试解析核心文档属性 (docProps/core.xml)
        auto error = zip_archive_->fileExists("docProps/core.xml");
        if (archive::isSuccess(error)) {
            std::string xml_content = extractXMLFromZip("docProps/core.xml");
            if (!xml_content.empty()) {
                if (!parser.parseCoreProps(xml_content)) {
                    FASTEXCEL_LOG_WARN("解析核心文档属性失败");
                } else {
                    FASTEXCEL_LOG_DEBUG("成功解析核心文档属性");
                }
            }
        }
        
        // 尝试解析应用程序属性 (docProps/app.xml)
        error = zip_archive_->fileExists("docProps/app.xml");
        if (archive::isSuccess(error)) {
            std::string xml_content = extractXMLFromZip("docProps/app.xml");
            if (!xml_content.empty()) {
                if (!parser.parseAppProps(xml_content)) {
                    FASTEXCEL_LOG_WARN("解析应用程序属性失败");
                } else {
                    FASTEXCEL_LOG_DEBUG("成功解析应用程序属性");
                }
            }
        }
        
        // 获取解析结果并设置到元数据中
        const auto& doc_props = parser.getDocProps();
        metadata_.title = doc_props.title;
        metadata_.subject = doc_props.subject;
        metadata_.author = doc_props.author;
        metadata_.company = doc_props.company;
        metadata_.application = doc_props.application;
        metadata_.app_version = doc_props.app_version;
        metadata_.manager = doc_props.manager;
        metadata_.category = doc_props.category;
        metadata_.keywords = doc_props.keywords;
        metadata_.created_time = doc_props.created;
        metadata_.modified_time = doc_props.modified;
        
        FASTEXCEL_LOG_DEBUG("文档属性解析完成");
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析文档属性时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

core::ErrorCode XLSXReader::parseThemeXML() {
    // 检查主题文件是否存在
    auto error = zip_archive_->fileExists("xl/theme/theme1.xml");
    if (archive::isError(error)) {
        // 主题文件不存在，这是正常的（某些Excel文件可能没有自定义主题）
        FASTEXCEL_LOG_DEBUG("主题文件不存在，将使用默认主题");
        return core::ErrorCode::FileNotFound;
    }
    
    std::string xml_content = extractXMLFromZip("xl/theme/theme1.xml");
    if (xml_content.empty()) {
        FASTEXCEL_LOG_DEBUG("主题文件为空，将使用默认主题");
        return core::ErrorCode::Ok; // 文件为空也是正常的
    }
    
    try {
        // 保存原始主题XML内容
        theme_xml_ = xml_content;

        // 使用 ThemeParser 解析为结构化对象，便于颜色/字体主题映射
        parsed_theme_ = theme::ThemeParser::parseFromXML(xml_content);
        if (!parsed_theme_) {
            FASTEXCEL_LOG_WARN("主题XML解析为对象失败，继续使用原始XML");
        } else {
            FASTEXCEL_LOG_DEBUG("成功解析主题对象: name={}", parsed_theme_->getName());
        }

        FASTEXCEL_LOG_DEBUG("成功解析主题文件 ({} 字节)", theme_xml_.size());
        return core::ErrorCode::Ok;

    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析主题文件时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

} // namespace reader
} // namespace fastexcel
