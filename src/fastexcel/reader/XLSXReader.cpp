//
// Created by wuxianggujun on 25-8-4.
//

#include "XLSXReader.hpp"
#include "SharedStringsParser.hpp"
#include "WorksheetParser.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <iostream>
#include <sstream>
#include <stdexcept>

namespace fastexcel {
namespace reader {

// 构造函数
XLSXReader::XLSXReader(const std::string& filename)
    : filename_(filename)
    , zip_archive_(std::make_unique<archive::ZipArchive>(filename))
    , is_open_(false) {
}

// 析构函数
XLSXReader::~XLSXReader() {
    if (is_open_) {
        close();
    }
}

// 打开XLSX文件
bool XLSXReader::open() {
    if (is_open_) {
        return true;
    }
    
    try {
        // 打开ZIP文件进行读取
        if (!zip_archive_->open(false)) {  // false表示不创建新文件，只读取
            std::cerr << "无法打开XLSX文件: " << filename_ << std::endl;
            return false;
        }
        
        // 验证XLSX文件结构
        if (!validateXLSXStructure()) {
            std::cerr << "无效的XLSX文件格式: " << filename_ << std::endl;
            zip_archive_->close();
            return false;
        }
        
        is_open_ = true;
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "打开XLSX文件时发生错误: " << e.what() << std::endl;
        return false;
    }
}

// 关闭文件
bool XLSXReader::close() {
    if (!is_open_) {
        return true;
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
        
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "关闭XLSX文件时发生错误: " << e.what() << std::endl;
        return false;
    }
}

// 加载整个工作簿
std::unique_ptr<core::Workbook> XLSXReader::loadWorkbook() {
    if (!is_open_) {
        std::cerr << "文件未打开，无法加载工作簿" << std::endl;
        return nullptr;
    }
    
    try {
        // 创建新的工作簿，使用不同的临时文件名避免冲突
        std::string temp_filename = filename_ + ".temp_workbook";
        auto workbook = std::make_unique<core::Workbook>(temp_filename);
        
        // 打开工作簿以便添加工作表
        if (!workbook->open()) {
            std::cerr << "无法打开工作簿进行写入" << std::endl;
            return nullptr;
        }
        
        // 解析各种XML文件
        if (!parseSharedStringsXML()) {
            std::cerr << "解析共享字符串失败" << std::endl;
            // 共享字符串不是必需的，继续执行
        }
        
        if (!parseStylesXML()) {
            // 样式解析失败不影响主要功能，继续执行
        }
        
        if (!parseWorkbookXML()) {
            std::cerr << "解析工作簿失败" << std::endl;
            return nullptr;
        }
        
        if (!parseDocPropsXML()) {
            // 文档属性解析失败不影响主要功能，继续执行
        }
        
        // 为每个工作表创建Worksheet对象
        for (const auto& sheet_name : worksheet_names_) {
            auto worksheet = workbook->addWorksheet(sheet_name);
            if (worksheet) {
                auto it = worksheet_paths_.find(sheet_name);
                if (it != worksheet_paths_.end()) {
                    if (!parseWorksheetXML(it->second, worksheet.get())) {
                        std::cerr << "解析工作表 " << sheet_name << " 失败" << std::endl;
                        // 继续处理其他工作表
                    }
                }
            }
        }
        
        return workbook;
        
    } catch (const std::exception& e) {
        std::cerr << "加载工作簿时发生错误: " << e.what() << std::endl;
        return nullptr;
    }
}

// 加载单个工作表
std::shared_ptr<core::Worksheet> XLSXReader::loadWorksheet(const std::string& name) {
    if (!is_open_) {
        std::cerr << "文件未打开，无法加载工作表" << std::endl;
        return nullptr;
    }
    
    try {
        // 如果还没有解析工作簿结构，先解析
        if (worksheet_names_.empty()) {
            if (!parseWorkbookXML()) {
                std::cerr << "解析工作簿结构失败" << std::endl;
                return nullptr;
            }
        }
        
        // 检查工作表是否存在
        auto it = worksheet_paths_.find(name);
        if (it == worksheet_paths_.end()) {
            std::cerr << "工作表 " << name << " 不存在" << std::endl;
            return nullptr;
        }
        
        // 解析共享字符串（如果还没有解析）
        if (shared_strings_.empty()) {
            parseSharedStringsXML();  // 忽略错误，某些文件可能没有共享字符串
        }
        
        // 解析样式（如果还没有解析）
        if (styles_.empty()) {
            parseStylesXML();  // 忽略错误，某些文件可能没有自定义样式
        }
        
        // 创建临时工作簿来容纳工作表
        auto temp_workbook = std::make_shared<core::Workbook>("temp");
        auto worksheet = std::make_shared<core::Worksheet>(name, temp_workbook);
        
        // 解析工作表数据
        if (!parseWorksheetXML(it->second, worksheet.get())) {
            std::cerr << "解析工作表 " << name << " 失败" << std::endl;
            return nullptr;
        }
        
        return worksheet;
        
    } catch (const std::exception& e) {
        std::cerr << "加载工作表时发生错误: " << e.what() << std::endl;
        return nullptr;
    }
}

// 获取工作表名称列表
std::vector<std::string> XLSXReader::getWorksheetNames() {
    if (!is_open_) {
        std::cerr << "文件未打开，无法获取工作表名称" << std::endl;
        return {};
    }
    
    // 如果还没有解析工作簿结构，先解析
    if (worksheet_names_.empty()) {
        if (!parseWorkbookXML()) {
            std::cerr << "解析工作簿结构失败" << std::endl;
            return {};
        }
    }
    
    return worksheet_names_;
}

// 获取元数据
WorkbookMetadata XLSXReader::getMetadata() {
    if (!is_open_) {
        std::cerr << "文件未打开，无法获取元数据" << std::endl;
        return WorkbookMetadata{};
    }
    
    // 如果还没有解析文档属性，先解析
    if (metadata_.title.empty() && metadata_.author.empty()) {
        parseDocPropsXML();  // 忽略错误
    }
    
    return metadata_;
}

// 获取定义名称列表
std::vector<std::string> XLSXReader::getDefinedNames() {
    if (!is_open_) {
        std::cerr << "文件未打开，无法获取定义名称" << std::endl;
        return {};
    }
    
    // 如果还没有解析工作簿结构，先解析
    if (defined_names_.empty()) {
        if (!parseWorkbookXML()) {
            std::cerr << "解析工作簿结构失败" << std::endl;
            return {};
        }
    }
    
    return defined_names_;
}

// 从ZIP中提取XML文件内容
std::string XLSXReader::extractXMLFromZip(const std::string& path) {
    std::string content;
    auto error = zip_archive_->extractFile(path, content);
    
    if (archive::isError(error)) {
        std::cerr << "提取文件 " << path << " 失败" << std::endl;
        return "";
    }
    
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
            std::cerr << "缺少必需文件: " << file << std::endl;
            return false;
        }
    }
    
    return true;
}

// 解析工作簿XML - 基本实现
bool XLSXReader::parseWorkbookXML() {
    std::string xml_content = extractXMLFromZip("xl/workbook.xml");
    if (xml_content.empty()) {
        return false;
    }
    
    // TODO: 实现完整的XML解析
    // 这里先提供一个简单的实现来提取工作表名称
    
    // 简单的字符串搜索来找到工作表信息
    size_t pos = 0;
    while ((pos = xml_content.find("<sheet ", pos)) != std::string::npos) {
        size_t name_start = xml_content.find("name=\"", pos);
        if (name_start != std::string::npos) {
            name_start += 6; // 跳过 name="
            size_t name_end = xml_content.find("\"", name_start);
            if (name_end != std::string::npos) {
                std::string sheet_name = xml_content.substr(name_start, name_end - name_start);
                worksheet_names_.push_back(sheet_name);
                
                // 构造工作表文件路径
                size_t id_start = xml_content.find("sheetId=\"", pos);
                if (id_start != std::string::npos) {
                    id_start += 9; // 跳过 sheetId="
                    size_t id_end = xml_content.find("\"", id_start);
                    if (id_end != std::string::npos) {
                        std::string sheet_id = xml_content.substr(id_start, id_end - id_start);
                        std::string sheet_path = "xl/worksheets/sheet" + sheet_id + ".xml";
                        worksheet_paths_[sheet_name] = sheet_path;
                    }
                }
            }
        }
        pos++;
    }
    
    return !worksheet_names_.empty();
}

// 解析工作表XML
bool XLSXReader::parseWorksheetXML(const std::string& path, core::Worksheet* worksheet) {
    if (!worksheet) {
        std::cerr << "工作表对象为空" << std::endl;
        return false;
    }
    
    std::string xml_content = extractXMLFromZip(path);
    if (xml_content.empty()) {
        std::cerr << "无法提取工作表XML: " << path << std::endl;
        return false;
    }
    
    try {
        WorksheetParser parser;
        if (!parser.parse(xml_content, worksheet, shared_strings_, styles_)) {
            std::cerr << "解析工作表XML失败: " << path << std::endl;
            return false;
        }
        
        std::cout << "成功解析工作表: " << worksheet->getName() << std::endl;
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "解析工作表时发生异常: " << e.what() << std::endl;
        return false;
    }
}

bool XLSXReader::parseStylesXML() {
    // 检查样式文件是否存在
    auto error = zip_archive_->fileExists("xl/styles.xml");
    if (archive::isError(error)) {
        // 样式文件不存在，这是正常的（某些Excel文件可能没有自定义样式）
        return true;
    }
    
    std::string xml_content = extractXMLFromZip("xl/styles.xml");
    if (xml_content.empty()) {
        return true; // 文件为空也是正常的
    }
    
    try {
        // TODO: 实现完整的样式解析
        // 目前只是简单地标记样式文件存在，不输出错误信息
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "解析样式时发生异常: " << e.what() << std::endl;
        return false;
    }
}

bool XLSXReader::parseSharedStringsXML() {
    // 检查共享字符串文件是否存在
    auto error = zip_archive_->fileExists("xl/sharedStrings.xml");
    if (archive::isError(error)) {
        // 共享字符串文件不存在，这是正常的（某些Excel文件可能没有共享字符串）
        return true;
    }
    
    std::string xml_content = extractXMLFromZip("xl/sharedStrings.xml");
    if (xml_content.empty()) {
        return true; // 文件为空也是正常的
    }
    
    try {
        SharedStringsParser parser;
        if (!parser.parse(xml_content)) {
            std::cerr << "解析共享字符串XML失败" << std::endl;
            return false;
        }
        
        // 将解析结果复制到成员变量
        shared_strings_ = parser.getStrings();
        
        std::cout << "成功解析 " << shared_strings_.size() << " 个共享字符串" << std::endl;
        return true;
        
    } catch (const std::exception& e) {
        std::cerr << "解析共享字符串时发生异常: " << e.what() << std::endl;
        return false;
    }
}

bool XLSXReader::parseContentTypesXML() {
    // TODO: 实现内容类型XML解析
    std::cerr << "parseContentTypesXML 尚未实现" << std::endl;
    return true;
}

bool XLSXReader::parseRelationshipsXML() {
    // TODO: 实现关系XML解析
    std::cerr << "parseRelationshipsXML 尚未实现" << std::endl;
    return true;
}

bool XLSXReader::parseDocPropsXML() {
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
    
    return true;
}

std::string XLSXReader::getCellValue(const std::string& cell_xml, core::CellType& type) {
    // TODO: 实现单元格值解析
    type = core::CellType::Empty;
    return "";
}

std::shared_ptr<core::Format> XLSXReader::getStyleByIndex(int index) {
    auto it = styles_.find(index);
    if (it != styles_.end()) {
        return it->second;
    }
    return nullptr;
}

} // namespace reader
} // namespace fastexcel
