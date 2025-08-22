#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#include "XLSXReader.hpp"
#include "SharedStringsParser.hpp"
#include "StylesParser.hpp"
#include "WorksheetParser.hpp"
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

// 验证XLSX文件结构
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
    
    return true;
}

// 解析工作簿XML - 系统层ErrorCode版本
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
        
        // 首先解析关系文件来获取工作表的实际路径
        std::unordered_map<std::string, std::string> relationships;
        if (!parseWorkbookRelationships(relationships)) {
            FASTEXCEL_LOG_WARN("无法解析工作簿关系文件，使用默认路径");
        }
        
        // 解析工作表信息
        size_t sheets_start = xml_content.find("<sheets");
        if (sheets_start == std::string::npos) {
            FASTEXCEL_LOG_ERROR("未找到工作表定义");
            return core::ErrorCode::XmlMissingElement;
        }
        
        size_t sheets_end = xml_content.find("</sheets>", sheets_start);
        if (sheets_end == std::string::npos) {
            FASTEXCEL_LOG_ERROR("工作表定义格式错误");
            return core::ErrorCode::XmlInvalidFormat;
        }
        
        std::string sheets_content = xml_content.substr(sheets_start, sheets_end - sheets_start);
        
        // 解析每个工作表
        size_t pos = 0;
        while ((pos = sheets_content.find("<sheet ", pos)) != std::string::npos) {
            size_t sheet_end = sheets_content.find("/>", pos);
            if (sheet_end == std::string::npos) {
                sheet_end = sheets_content.find("</sheet>", pos);
                if (sheet_end == std::string::npos) {
                    break;
                }
                sheet_end += 8; // 包含 </sheet>
            } else {
                sheet_end += 2; // 包含 />
            }
            
            std::string sheet_xml = sheets_content.substr(pos, sheet_end - pos);
            
            // 提取工作表属性
            std::string sheet_name = extractAttribute(sheet_xml, "name");
            std::string sheet_id = extractAttribute(sheet_xml, "sheetId");
            std::string rel_id = extractAttribute(sheet_xml, "r:id");
            
            if (!sheet_name.empty()) {
                worksheet_names_.push_back(sheet_name);
                
                // 确定工作表文件路径
                std::string sheet_path;
                if (!rel_id.empty() && relationships.find(rel_id) != relationships.end()) {
                    sheet_path = "xl/" + relationships[rel_id];
                } else if (!sheet_id.empty()) {
                    // 回退到默认路径
                    sheet_path = "xl/worksheets/sheet" + sheet_id + ".xml";
                } else {
                    FASTEXCEL_LOG_ERROR("无法确定工作表 {} 的路径", sheet_name);
                    continue;
                }
                
                worksheet_paths_[sheet_name] = sheet_path;
            }
            
            pos = sheet_end;
        }
        
        // 解析定义名称
        parseDefinedNames(xml_content);
        
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
    
    std::string xml_content = extractXMLFromZip(path);
    if (xml_content.empty()) {
        FASTEXCEL_LOG_ERROR("无法提取工作表XML: {}", path);
        return core::ErrorCode::FileNotFound;
    }
    
    try {
        WorksheetParser parser;
        if (!parser.parse(xml_content, worksheet, shared_strings_, styles_, style_id_mapping_)) {
            FASTEXCEL_LOG_ERROR("解析工作表XML失败: {}", path);
            return core::ErrorCode::XmlParseError;
        }
        
        FASTEXCEL_LOG_DEBUG("成功解析工作表: {}", worksheet->getName());
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
    
    std::string xml_content = extractXMLFromZip("xl/sharedStrings.xml");
    if (xml_content.empty()) {
        return core::ErrorCode::Ok; // 文件为空也是正常的
    }
    
    try {
        SharedStringsParser parser;
        if (!parser.parse(xml_content)) {
            FASTEXCEL_LOG_ERROR("解析共享字符串XML失败");
            return core::ErrorCode::XmlParseError;
        }
        
        // 将解析结果复制到成员变量
        shared_strings_ = parser.getStrings();

        FASTEXCEL_LOG_DEBUG("成功解析 {} 个共享字符串", shared_strings_.size());
        return core::ErrorCode::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析共享字符串时发生异常: {}", e.what());
        return core::ErrorCode::XmlParseError;
    }
}

core::ErrorCode XLSXReader::parseContentTypesXML() {
    // TODO: 实现内容类型XML解析
    FASTEXCEL_LOG_WARN("parseContentTypesXML 尚未实现");
    return core::ErrorCode::NotImplemented;
}

core::ErrorCode XLSXReader::parseRelationshipsXML() {
    // TODO: 实现关系XML解析
    FASTEXCEL_LOG_WARN("parseRelationshipsXML 尚未实现");
    return core::ErrorCode::NotImplemented;
}

core::ErrorCode XLSXReader::parseDocPropsXML() {
    // 尝试解析核心文档属性
    auto error = zip_archive_->fileExists("docProps/core.xml");
    if (archive::isSuccess(error)) {
        std::string xml_content = extractXMLFromZip("docProps/core.xml");
        if (!xml_content.empty()) {
            // 简单解析标题和作者信息
            size_t title_start = xml_content.find("<dc:title>");
            if (title_start != std::string::npos) {
                title_start += 10; // 跳过 <dc:title>
                size_t title_end = xml_content.find("</dc:title>", title_start);
                if (title_end != std::string::npos) {
                    metadata_.title = xml_content.substr(title_start, title_end - title_start);
                }
            }
            
            size_t author_start = xml_content.find("<dc:creator>");
            if (author_start != std::string::npos) {
                author_start += 12; // 跳过 <dc:creator>
                size_t author_end = xml_content.find("</dc:creator>", author_start);
                if (author_end != std::string::npos) {
                    metadata_.author = xml_content.substr(author_start, author_end - author_start);
                }
            }
            
            size_t subject_start = xml_content.find("<dc:subject>");
            if (subject_start != std::string::npos) {
                subject_start += 12; // 跳过 <dc:subject>
                size_t subject_end = xml_content.find("</dc:subject>", subject_start);
                if (subject_end != std::string::npos) {
                    metadata_.subject = xml_content.substr(subject_start, subject_end - subject_start);
                }
            }
        }
    }
    
    // 尝试解析应用程序属性
    error = zip_archive_->fileExists("docProps/app.xml");
    if (archive::isSuccess(error)) {
        std::string xml_content = extractXMLFromZip("docProps/app.xml");
        if (!xml_content.empty()) {
            size_t company_start = xml_content.find("<Company>");
            if (company_start != std::string::npos) {
                company_start += 9; // 跳过 <Company>
                size_t company_end = xml_content.find("</Company>", company_start);
                if (company_end != std::string::npos) {
                    metadata_.company = xml_content.substr(company_start, company_end - company_start);
                }
            }
            
            size_t app_start = xml_content.find("<Application>");
            if (app_start != std::string::npos) {
                app_start += 13; // 跳过 <Application>
                size_t app_end = xml_content.find("</Application>", app_start);
                if (app_end != std::string::npos) {
                    metadata_.application = xml_content.substr(app_start, app_end - app_start);
                }
            }
        }
    }
    
    return core::ErrorCode::Ok;
}

std::string XLSXReader::getCellValue(const std::string& cell_xml, core::CellType& type) {
    // 提取单元格类型
    std::string cell_type = extractAttribute(cell_xml, "t");
    
    // 查找值标签
    size_t v_start = cell_xml.find("<v>");
    if (v_start != std::string::npos) {
        v_start += 3; // 跳过 <v>
        size_t v_end = cell_xml.find("</v>", v_start);
        if (v_end != std::string::npos) {
            std::string value = cell_xml.substr(v_start, v_end - v_start);
            
            if (cell_type == "s") {
                type = core::CellType::String;
                // 共享字符串索引
                try {
                    int index = std::stoi(value);
                    auto it = shared_strings_.find(index);
                    return (it != shared_strings_.end()) ? it->second : "";
                } catch (...) {
                    return "";
                }
            } else if (cell_type == "b") {
                type = core::CellType::Boolean;
                return value;
            } else if (cell_type == "str") {
                type = core::CellType::String;
                return value;
            } else {
                type = core::CellType::Number;
                return value;
            }
        }
    }
    
    // 查找内联字符串
    size_t is_start = cell_xml.find("<is>");
    if (is_start != std::string::npos) {
        size_t t_start = cell_xml.find("<t>", is_start);
        if (t_start != std::string::npos) {
            t_start += 3;
            size_t t_end = cell_xml.find("</t>", t_start);
            if (t_end != std::string::npos) {
                type = core::CellType::String;
                return cell_xml.substr(t_start, t_end - t_start);
            }
        }
    }
    
    type = core::CellType::Empty;
    return "";
}

std::shared_ptr<core::FormatDescriptor> XLSXReader::getStyleByIndex(int index) {
    auto it = styles_.find(index);
    if (it != styles_.end()) {
        return it->second;
    }
    return nullptr;
}

// 新增的辅助方法
std::string XLSXReader::extractAttribute(const std::string& xml, const std::string& attr_name) {
    std::string search_pattern = attr_name + "=\"";
    size_t attr_start = xml.find(search_pattern);
    if (attr_start == std::string::npos) {
        return "";
    }
    
    attr_start += search_pattern.length();
    size_t attr_end = xml.find("\"", attr_start);
    if (attr_end == std::string::npos) {
        return "";
    }
    
    return xml.substr(attr_start, attr_end - attr_start);
}

bool XLSXReader::parseWorkbookRelationships(std::unordered_map<std::string, std::string>& relationships) {
    std::string xml_content = extractXMLFromZip("xl/_rels/workbook.xml.rels");
    if (xml_content.empty()) {
        return false;
    }
    
    try {
        size_t pos = 0;
        while ((pos = xml_content.find("<Relationship ", pos)) != std::string::npos) {
            size_t rel_end = xml_content.find("/>", pos);
            if (rel_end == std::string::npos) {
                rel_end = xml_content.find("</Relationship>", pos);
                if (rel_end == std::string::npos) {
                    break;
                }
                rel_end += 15;
            } else {
                rel_end += 2;
            }
            
            std::string rel_xml = xml_content.substr(pos, rel_end - pos);
            
            std::string id = extractAttribute(rel_xml, "Id");
            std::string target = extractAttribute(rel_xml, "Target");
            std::string type = extractAttribute(rel_xml, "Type");
            
            // 只处理工作表关系
            if (!id.empty() && !target.empty() &&
                type.find("worksheet") != std::string::npos) {
                relationships[id] = target;
            }
            
            pos = rel_end;
        }
        
        return !relationships.empty();
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析关系文件时发生异常: {}", e.what());
        return false;
    }
}

bool XLSXReader::parseDefinedNames(const std::string& xml_content) {
    size_t names_start = xml_content.find("<definedNames");
    if (names_start == std::string::npos) {
        return true; // 没有定义名称是正常的
    }
    
    size_t names_end = xml_content.find("</definedNames>", names_start);
    if (names_end == std::string::npos) {
        return false;
    }
    
    std::string names_content = xml_content.substr(names_start, names_end - names_start);
    
    size_t pos = 0;
    while ((pos = names_content.find("<definedName ", pos)) != std::string::npos) {
        size_t name_end = names_content.find("</definedName>", pos);
        if (name_end == std::string::npos) {
            break;
        }
        
        std::string name_xml = names_content.substr(pos, name_end - pos);
        std::string name = extractAttribute(name_xml, "name");
        
        if (!name.empty()) {
            defined_names_.push_back(name);
        }
        
        pos = name_end + 14; // 跳过 </definedName>
    }
    
    return true;
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
