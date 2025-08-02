#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <sstream>
#include <algorithm>
#include <iomanip>
#include <functional>
#include <ctime>
#include <fstream>

namespace fastexcel {
namespace core {

// ========== DocumentProperties 实现 ==========

DocumentProperties::DocumentProperties() {
    // 设置默认时间为当前时间
    std::time_t now = std::time(nullptr);
#ifdef _WIN32
    localtime_s(&created_time, &now);
#else
    created_time = *std::localtime(&now);
#endif
    modified_time = created_time;
}

// ========== Workbook 实现 ==========

std::unique_ptr<Workbook> Workbook::create(const std::string& filename) {
    return std::make_unique<Workbook>(filename);
}

Workbook::Workbook(const std::string& filename) : filename_(filename) {
    file_manager_ = std::make_unique<archive::FileManager>(filename);
    
    // 设置默认文档属性
    doc_properties_.author = "FastExcel";
    doc_properties_.company = "FastExcel Library";
}

Workbook::~Workbook() {
    close();
}

// ========== 文件操作 ==========

bool Workbook::open() {
    if (is_open_) {
        return true;
    }
    
    is_open_ = file_manager_->open(true);
    if (is_open_) {
        LOG_INFO("Workbook opened: {}", filename_);
    }
    
    return is_open_;
}

bool Workbook::save() {
    if (!is_open_) {
        throw std::runtime_error("Workbook is not open");
    }
    
    try {
        // 更新修改时间
        std::time_t now = std::time(nullptr);
#ifdef _WIN32
        localtime_s(&doc_properties_.modified_time, &now);
#else
        doc_properties_.modified_time = *std::localtime(&now);
#endif
        
        // 设置ZIP压缩级别
        if (file_manager_->isOpen()) {
            if (!file_manager_->setCompressionLevel(options_.compression_level)) {
                LOG_WARN("Failed to set compression level to {}", options_.compression_level);
            } else {
                LOG_DEBUG("Set ZIP compression level to {}", options_.compression_level);
            }
        }
        
        // 根据选项决定是否收集共享字符串
        if (options_.use_shared_strings) {
            LOG_DEBUG("Collecting shared strings (enabled)");
            collectSharedStrings();
        } else {
            LOG_DEBUG("Skipping shared strings collection (disabled for performance)");
            shared_strings_.clear();
            shared_strings_list_.clear();
        }
        
        // 更新格式索引
        updateFormatIndices();
        
        // 生成Excel文件结构
        if (!generateExcelStructure()) {
            LOG_ERROR("Failed to generate Excel structure");
            return false;
        }
        
        LOG_INFO("Workbook saved successfully: {}", filename_);
        return true;
    } catch (const std::exception& e) {
        LOG_ERROR("Failed to save workbook: {}", e.what());
        return false;
    }
}

bool Workbook::saveAs(const std::string& filename) {
    std::string old_filename = filename_;
    filename_ = filename;
    
    // 重新创建文件管理器
    file_manager_ = std::make_unique<archive::FileManager>(filename);
    
    if (!file_manager_->open(true)) {
        // 恢复原文件名
        filename_ = old_filename;
        file_manager_ = std::make_unique<archive::FileManager>(old_filename);
        return false;
    }
    
    return save();
}

bool Workbook::close() {
    if (is_open_) {
        file_manager_->close();
        is_open_ = false;
        LOG_INFO("Workbook closed: {}", filename_);
    }
    return true;
}

// ========== 工作表管理 ==========

std::shared_ptr<Worksheet> Workbook::addWorksheet(const std::string& name) {
    if (!is_open_) {
        throw std::runtime_error("Workbook is not open");
    }
    
    std::string sheet_name;
    if (name.empty()) {
        sheet_name = generateUniqueSheetName("Sheet1");
    } else {
        // 检查名称是否已存在，如果存在则生成唯一名称
        if (getWorksheet(name) != nullptr) {
            sheet_name = generateUniqueSheetName(name);
        } else {
            sheet_name = name;
        }
    }
    
    if (!validateSheetName(sheet_name)) {
        LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.push_back(worksheet);
    
    LOG_DEBUG("Added worksheet: {}", sheet_name);
    return worksheet;
}

std::shared_ptr<Worksheet> Workbook::insertWorksheet(size_t index, const std::string& name) {
    if (!is_open_) {
        LOG_ERROR("Workbook is not open");
        return nullptr;
    }
    
    if (index > worksheets_.size()) {
        index = worksheets_.size();
    }
    
    std::string sheet_name = name.empty() ? generateUniqueSheetName("Sheet1") : name;
    
    if (!validateSheetName(sheet_name)) {
        LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.insert(worksheets_.begin() + index, worksheet);
    
    LOG_DEBUG("Inserted worksheet: {} at index {}", sheet_name, index);
    return worksheet;
}

bool Workbook::removeWorksheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        worksheets_.erase(it);
        LOG_DEBUG("Removed worksheet: {}", name);
        return true;
    }
    
    return false;
}

bool Workbook::removeWorksheet(size_t index) {
    if (index < worksheets_.size()) {
        std::string name = worksheets_[index]->getName();
        worksheets_.erase(worksheets_.begin() + index);
        LOG_DEBUG("Removed worksheet: {} at index {}", name, index);
        return true;
    }
    
    return false;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(), 
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(size_t index) {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(size_t index) const {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::vector<std::string> Workbook::getWorksheetNames() const {
    std::vector<std::string> names;
    names.reserve(worksheets_.size());
    
    for (const auto& worksheet : worksheets_) {
        names.push_back(worksheet->getName());
    }
    
    return names;
}

bool Workbook::renameWorksheet(const std::string& old_name, const std::string& new_name) {
    auto worksheet = getWorksheet(old_name);
    if (!worksheet) {
        return false;
    }
    
    if (!validateSheetName(new_name)) {
        return false;
    }
    
    worksheet->setName(new_name);
    LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

bool Workbook::moveWorksheet(size_t from_index, size_t to_index) {
    if (from_index >= worksheets_.size() || to_index >= worksheets_.size()) {
        return false;
    }
    
    if (from_index == to_index) {
        return true;
    }
    
    auto worksheet = worksheets_[from_index];
    worksheets_.erase(worksheets_.begin() + from_index);
    
    if (to_index > from_index) {
        to_index--;
    }
    
    worksheets_.insert(worksheets_.begin() + to_index, worksheet);
    
    LOG_DEBUG("Moved worksheet from index {} to {}", from_index, to_index);
    return true;
}

std::shared_ptr<Worksheet> Workbook::copyWorksheet(const std::string& source_name, const std::string& new_name) {
    auto source_worksheet = getWorksheet(source_name);
    if (!source_worksheet) {
        return nullptr;
    }
    
    if (!validateSheetName(new_name)) {
        return nullptr;
    }
    
    // 创建新工作表
    auto new_worksheet = std::make_shared<Worksheet>(new_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 这里应该实现深拷贝逻辑
    // 简化版本，实际需要复制所有单元格、格式、设置等
    
    worksheets_.push_back(new_worksheet);
    
    LOG_DEBUG("Copied worksheet: {} -> {}", source_name, new_name);
    return new_worksheet;
}

void Workbook::setActiveWorksheet(size_t index) {
    // 取消所有工作表的选中状态
    for (auto& worksheet : worksheets_) {
        worksheet->setTabSelected(false);
    }
    
    // 设置指定工作表为活动状态
    if (index < worksheets_.size()) {
        worksheets_[index]->setTabSelected(true);
    }
}

// ========== 格式管理 ==========

std::shared_ptr<Format> Workbook::createFormat() {
    if (!is_open_) {
        throw std::runtime_error("Workbook is not open");
    }
    
    auto format = std::make_shared<Format>();
    int xf_index = static_cast<int>(formats_.size()); // 从0开始索引
    format->setXfIndex(xf_index);
    
    // 存储格式
    formats_[xf_index] = format;
    
    return format;
}

std::shared_ptr<Format> Workbook::getFormat(int format_id) const {
    auto it = formats_.find(format_id);
    if (it != formats_.end()) {
        return it->second;
    }
    return nullptr;
}

// ========== 自定义属性 ==========

void Workbook::setCustomProperty(const std::string& name, const std::string& value) {
    // 查找是否已存在
    auto it = std::find_if(custom_properties_.begin(), custom_properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    
    if (it != custom_properties_.end()) {
        it->value = value;
        it->type = CustomProperty::String;
    } else {
        custom_properties_.emplace_back(name, value);
    }
}

void Workbook::setCustomProperty(const std::string& name, double value) {
    auto it = std::find_if(custom_properties_.begin(), custom_properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    
    if (it != custom_properties_.end()) {
        it->value = std::to_string(value);
        it->type = CustomProperty::Number;
    } else {
        custom_properties_.emplace_back(name, value);
    }
}

void Workbook::setCustomProperty(const std::string& name, bool value) {
    auto it = std::find_if(custom_properties_.begin(), custom_properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    
    if (it != custom_properties_.end()) {
        it->value = value ? "true" : "false";
        it->type = CustomProperty::Boolean;
    } else {
        custom_properties_.emplace_back(name, value);
    }
}

std::string Workbook::getCustomProperty(const std::string& name) const {
    auto it = std::find_if(custom_properties_.begin(), custom_properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    
    if (it != custom_properties_.end()) {
        return it->value;
    }
    
    return "";
}

bool Workbook::removeCustomProperty(const std::string& name) {
    auto it = std::find_if(custom_properties_.begin(), custom_properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    
    if (it != custom_properties_.end()) {
        custom_properties_.erase(it);
        return true;
    }
    
    return false;
}

// ========== 定义名称 ==========

void Workbook::defineName(const std::string& name, const std::string& formula, const std::string& scope) {
    // 查找是否已存在
    auto it = std::find_if(defined_names_.begin(), defined_names_.end(),
                          [&name, &scope](const DefinedName& dn) {
                              return dn.name == name && dn.scope == scope;
                          });
    
    if (it != defined_names_.end()) {
        it->formula = formula;
    } else {
        defined_names_.emplace_back(name, formula, scope);
    }
}

std::string Workbook::getDefinedName(const std::string& name, const std::string& scope) const {
    auto it = std::find_if(defined_names_.begin(), defined_names_.end(),
                          [&name, &scope](const DefinedName& dn) {
                              return dn.name == name && dn.scope == scope;
                          });
    
    if (it != defined_names_.end()) {
        return it->formula;
    }
    
    return "";
}

bool Workbook::removeDefinedName(const std::string& name, const std::string& scope) {
    auto it = std::find_if(defined_names_.begin(), defined_names_.end(),
                          [&name, &scope](const DefinedName& dn) {
                              return dn.name == name && dn.scope == scope;
                          });
    
    if (it != defined_names_.end()) {
        defined_names_.erase(it);
        return true;
    }
    
    return false;
}

// ========== VBA项目 ==========

bool Workbook::addVbaProject(const std::string& vba_project_path) {
    // 检查文件是否存在
    std::ifstream file(vba_project_path, std::ios::binary);
    if (!file.is_open()) {
        LOG_ERROR("VBA project file not found: {}", vba_project_path);
        return false;
    }
    
    vba_project_path_ = vba_project_path;
    has_vba_ = true;
    
    LOG_INFO("Added VBA project: {}", vba_project_path);
    return true;
}

// ========== 工作簿保护 ==========

void Workbook::protect(const std::string& password, bool lock_structure, bool lock_windows) {
    protected_ = true;
    protection_password_ = password;
    lock_structure_ = lock_structure;
    lock_windows_ = lock_windows;
}

void Workbook::unprotect() {
    protected_ = false;
    protection_password_.clear();
    lock_structure_ = false;
    lock_windows_ = false;
}

// ========== 工作簿选项 ==========

void Workbook::setCalcOptions(bool calc_on_load, bool full_calc_on_load) {
    options_.calc_on_load = calc_on_load;
    options_.full_calc_on_load = full_calc_on_load;
}

// ========== 共享字符串管理 ==========

int Workbook::addSharedString(const std::string& str) {
    auto it = shared_strings_.find(str);
    if (it != shared_strings_.end()) {
        return it->second;
    }
    
    int index = static_cast<int>(shared_strings_list_.size());
    shared_strings_[str] = index;
    shared_strings_list_.push_back(str);
    
    return index;
}

int Workbook::getSharedStringIndex(const std::string& str) const {
    auto it = shared_strings_.find(str);
    if (it != shared_strings_.end()) {
        return it->second;
    }
    
    return -1;
}

// ========== 内部方法 ==========

bool Workbook::generateExcelStructure() {
    if (options_.streaming_xml) {
        return generateExcelStructureStreaming();
    } else {
        return generateExcelStructureBatch();
    }
}

bool Workbook::generateExcelStructureBatch() {
    LOG_INFO("Starting Excel structure generation with batch mode");
    
    // 预估文件数量：基础文件8个 + 工作表文件 + 工作表关系文件 + 自定义属性
    size_t estimated_files = 8 + worksheets_.size() * 2;
    if (!custom_properties_.empty()) {
        estimated_files++;
    }
    
    // 预分配空间以提高性能
    std::vector<std::pair<std::string, std::string>> files;
    files.reserve(estimated_files);
    
    // 收集所有文件内容（批量模式）
    LOG_DEBUG("Generating XML content for {} estimated files", estimated_files);
    
    // 基础文件
    files.emplace_back("[Content_Types].xml", generateContentTypesXML());
    files.emplace_back("_rels/.rels", generateRelsXML());
    files.emplace_back("docProps/app.xml", generateDocPropsAppXML());
    files.emplace_back("docProps/core.xml", generateDocPropsCoreXML());
    
    // 自定义属性（如果有）
    if (!custom_properties_.empty()) {
        files.emplace_back("docProps/custom.xml", generateDocPropsCustomXML());
    }
    
    // Excel核心文件
    files.emplace_back("xl/_rels/workbook.xml.rels", generateWorkbookRelsXML());
    files.emplace_back("xl/workbook.xml", generateWorkbookXML());
    files.emplace_back("xl/styles.xml", generateStylesXML());
    files.emplace_back("xl/sharedStrings.xml", generateSharedStringsXML());
    
    // 工作表文件
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        std::string worksheet_xml = generateWorksheetXML(worksheets_[i]);
        std::string worksheet_path = getWorksheetPath(static_cast<int>(i + 1));
        files.emplace_back(std::move(worksheet_path), std::move(worksheet_xml));
        
        // 工作表关系文件（如果有超链接等）
        std::string worksheet_rels_xml = worksheets_[i]->generateRelsXML();
        if (!worksheet_rels_xml.empty()) {
            std::string rels_path = "xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels";
            files.emplace_back(std::move(rels_path), std::move(worksheet_rels_xml));
        }
    }
    
    LOG_INFO("Generated {} files, starting batch write to ZIP", files.size());
    
    // 批量写入所有文件（使用移动语义提高性能）
    bool success = file_manager_->writeFiles(std::move(files));
    
    if (success) {
        LOG_INFO("Excel structure generation completed successfully in batch mode");
    } else {
        LOG_ERROR("Failed to write files in batch mode");
    }
    
    return success;
}

bool Workbook::generateExcelStructureStreaming() {
    LOG_INFO("Starting Excel structure generation with streaming mode");
    
    try {
        // 基础文件（这些通常较小，直接生成）
        if (!file_manager_->writeFile("[Content_Types].xml", generateContentTypesXML())) {
            LOG_ERROR("Failed to write Content_Types.xml");
            return false;
        }
        
        if (!file_manager_->writeFile("_rels/.rels", generateRelsXML())) {
            LOG_ERROR("Failed to write _rels/.rels");
            return false;
        }
        
        if (!file_manager_->writeFile("docProps/app.xml", generateDocPropsAppXML())) {
            LOG_ERROR("Failed to write docProps/app.xml");
            return false;
        }
        
        if (!file_manager_->writeFile("docProps/core.xml", generateDocPropsCoreXML())) {
            LOG_ERROR("Failed to write docProps/core.xml");
            return false;
        }
        
        // 自定义属性（如果有）
        if (!custom_properties_.empty()) {
            if (!file_manager_->writeFile("docProps/custom.xml", generateDocPropsCustomXML())) {
                LOG_ERROR("Failed to write docProps/custom.xml");
                return false;
            }
        }
        
        // Excel核心文件
        if (!file_manager_->writeFile("xl/_rels/workbook.xml.rels", generateWorkbookRelsXML())) {
            LOG_ERROR("Failed to write xl/_rels/workbook.xml.rels");
            return false;
        }
        
        if (!file_manager_->writeFile("xl/workbook.xml", generateWorkbookXML())) {
            LOG_ERROR("Failed to write xl/workbook.xml");
            return false;
        }
        
        if (!file_manager_->writeFile("xl/styles.xml", generateStylesXML())) {
            LOG_ERROR("Failed to write xl/styles.xml");
            return false;
        }
        
        if (!file_manager_->writeFile("xl/sharedStrings.xml", generateSharedStringsXML())) {
            LOG_ERROR("Failed to write xl/sharedStrings.xml");
            return false;
        }
        
        // 工作表文件（流式写入）
        for (size_t i = 0; i < worksheets_.size(); ++i) {
            std::string worksheet_path = getWorksheetPath(static_cast<int>(i + 1));
            
            // 使用流式写入生成工作表XML
            if (!generateWorksheetXMLStreaming(worksheets_[i], worksheet_path)) {
                LOG_ERROR("Failed to write worksheet {}", worksheet_path);
                return false;
            }
            
            // 工作表关系文件（如果有超链接等）
            std::string worksheet_rels_xml = worksheets_[i]->generateRelsXML();
            if (!worksheet_rels_xml.empty()) {
                std::string rels_path = "xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels";
                if (!file_manager_->writeFile(rels_path, worksheet_rels_xml)) {
                    LOG_ERROR("Failed to write worksheet relations {}", rels_path);
                    return false;
                }
            }
        }
        
        LOG_INFO("Excel structure generation completed successfully in streaming mode");
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception in streaming Excel structure generation: {}", e.what());
        return false;
    }
}

std::string Workbook::generateWorkbookXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 工作簿属性
    if (options_.read_only_recommended) {
        writer.startElement("fileVersion");
        writer.writeAttribute("readOnlyRecommended", "1");
        writer.endElement(); // fileVersion
    }
    
    // 工作簿保护
    if (protected_) {
        writer.startElement("workbookProtection");
        if (lock_structure_) {
            writer.writeAttribute("lockStructure", "1");
        }
        if (lock_windows_) {
            writer.writeAttribute("lockWindows", "1");
        }
        if (!protection_password_.empty()) {
            writer.writeAttribute("workbookPassword", hashPassword(protection_password_).c_str());
        }
        writer.endElement(); // workbookProtection
    }
    
    // 工作簿视图
    writer.startElement("bookViews");
    writer.startElement("workbookView");
    writer.writeAttribute("xWindow", "240");
    writer.writeAttribute("yWindow", "15");
    writer.writeAttribute("windowWidth", "16095");
    writer.writeAttribute("windowHeight", "9660");
    writer.endElement(); // workbookView
    writer.endElement(); // bookViews
    
    // 工作表
    writer.startElement("sheets");
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        writer.startElement("sheet");
        writer.writeAttribute("name", worksheets_[i]->getName().c_str());
        writer.writeAttribute("sheetId", std::to_string(worksheets_[i]->getSheetId()).c_str());
        writer.writeAttribute("r:id", ("rId" + std::to_string(i + 1)).c_str());
        writer.endElement(); // sheet
    }
    writer.endElement(); // sheets
    
    // 定义名称
    if (!defined_names_.empty()) {
        writer.startElement("definedNames");
        for (const auto& defined_name : defined_names_) {
            writer.startElement("definedName");
            writer.writeAttribute("name", defined_name.name.c_str());
            if (!defined_name.scope.empty()) {
                writer.writeAttribute("localSheetId", "0"); // 简化处理
            }
            if (defined_name.hidden) {
                writer.writeAttribute("hidden", "1");
            }
            writer.writeText(defined_name.formula.c_str());
            writer.endElement(); // definedName
        }
        writer.endElement(); // definedNames
    }
    
    // 计算属性
    writer.startElement("calcPr");
    writer.writeAttribute("calcId", "124519");
    if (!options_.calc_on_load) {
        writer.writeAttribute("calcMode", "manual");
    }
    if (options_.full_calc_on_load) {
        writer.writeAttribute("fullCalcOnLoad", "1");
    }
    writer.endElement(); // calcPr
    
    writer.endElement(); // workbook
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateStylesXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("styleSheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    // 数字格式
    writer.startElement("numFmts");
    writer.writeAttribute("count", "0");
    writer.endElement(); // numFmts
    
    // 字体
    writer.startElement("fonts");
    writer.writeAttribute("count", std::to_string(next_font_id_).c_str());
    
    // 默认字体
    writer.startElement("font");
    writer.startElement("sz");
    writer.writeAttribute("val", "11");
    writer.endElement(); // sz
    writer.startElement("name");
    writer.writeAttribute("val", "Calibri");
    writer.endElement(); // name
    writer.startElement("family");
    writer.writeAttribute("val", "2");
    writer.endElement(); // family
    writer.startElement("scheme");
    writer.writeAttribute("val", "minor");
    writer.endElement(); // scheme
    writer.endElement(); // font
    
    // 生成其他字体
    for (const auto& [id, format] : formats_) {
        if (format->hasFont()) {
            writer.writeRaw(format->generateFontXML().c_str());
        }
    }
    
    writer.endElement(); // fonts
    
    // 填充
    writer.startElement("fills");
    writer.writeAttribute("count", std::to_string(next_fill_id_).c_str());
    
    // 默认填充
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "none");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    writer.startElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "gray125");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    
    // 生成其他填充
    for (const auto& [id, format] : formats_) {
        if (format->hasFill()) {
            writer.writeRaw(format->generateFillXML().c_str());
        }
    }
    
    writer.endElement(); // fills
    
    // 边框
    writer.startElement("borders");
    writer.writeAttribute("count", std::to_string(next_border_id_).c_str());
    
    // 默认边框
    writer.startElement("border");
    writer.startElement("left");
    writer.endElement(); // left
    writer.startElement("right");
    writer.endElement(); // right
    writer.startElement("top");
    writer.endElement(); // top
    writer.startElement("bottom");
    writer.endElement(); // bottom
    writer.startElement("diagonal");
    writer.endElement(); // diagonal
    writer.endElement(); // border
    
    // 生成其他边框
    for (const auto& [id, format] : formats_) {
        if (format->hasBorder()) {
            writer.writeRaw(format->generateBorderXML().c_str());
        }
    }
    
    writer.endElement(); // borders
    
    // 单元格样式XF
    writer.startElement("cellStyleXfs");
    writer.writeAttribute("count", "1");
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.endElement(); // xf
    writer.endElement(); // cellStyleXfs
    
    // 单元格XF
    writer.startElement("cellXfs");
    writer.writeAttribute("count", std::to_string(next_xf_id_).c_str());
    
    // 默认XF
    writer.startElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.writeAttribute("xfId", "0");
    writer.endElement(); // xf
    
    // 生成其他XF
    for (const auto& [id, format] : formats_) {
        writer.startElement("xf");
        writer.writeAttribute("numFmtId", std::to_string(format->getNumberFormatIndex()).c_str());
        writer.writeAttribute("fontId", std::to_string(format->getFontIndex()).c_str());
        writer.writeAttribute("fillId", std::to_string(format->getFillIndex()).c_str());
        writer.writeAttribute("borderId", std::to_string(format->getBorderIndex()).c_str());
        writer.writeAttribute("xfId", "0");
        
        if (format->hasAlignment()) {
            writer.writeRaw(format->generateAlignmentXML().c_str());
        }
        
        if (format->hasProtection()) {
            writer.writeRaw(format->generateProtectionXML().c_str());
        }
        
        writer.endElement(); // xf
    }
    
    writer.endElement(); // cellXfs
    
    // 单元格样式
    writer.startElement("cellStyles");
    writer.writeAttribute("count", "1");
    writer.startElement("cellStyle");
    writer.writeAttribute("name", "Normal");
    writer.writeAttribute("xfId", "0");
    writer.writeAttribute("builtinId", "0");
    writer.endElement(); // cellStyle
    writer.endElement(); // cellStyles
    
    writer.endElement(); // styleSheet
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateSharedStringsXML() const {
    // 如果禁用了共享字符串，生成空的共享字符串文件
    if (!options_.use_shared_strings || shared_strings_list_.empty()) {
        xml::XMLStreamWriter writer;
        writer.startDocument();
        writer.startElement("sst");
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.writeAttribute("count", "0");
        writer.writeAttribute("uniqueCount", "0");
        writer.endElement(); // sst
        writer.endDocument();
        return writer.toString();
    }
    
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("sst");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("count", std::to_string(shared_strings_list_.size()).c_str());
    writer.writeAttribute("uniqueCount", std::to_string(shared_strings_list_.size()).c_str());
    
    for (const auto& str : shared_strings_list_) {
        writer.startElement("si");
        writer.startElement("t");
        writer.writeText(str.c_str());
        writer.endElement(); // t
        writer.endElement(); // si
    }
    
    writer.endElement(); // sst
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateWorksheetXML(const std::shared_ptr<Worksheet>& worksheet) const {
    return worksheet->generateXML();
}

std::string Workbook::generateDocPropsAppXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Properties");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    
    writer.startElement("Application");
    writer.writeText("FastExcel");
    writer.endElement(); // Application
    
    writer.startElement("DocSecurity");
    writer.writeText("0");
    writer.endElement(); // DocSecurity
    
    writer.startElement("Company");
    writer.writeText(doc_properties_.company.c_str());
    writer.endElement(); // Company
    
    writer.startElement("LinksUpToDate");
    writer.writeText("false");
    writer.endElement(); // LinksUpToDate
    
    writer.startElement("SharedDoc");
    writer.writeText("false");
    writer.endElement(); // SharedDoc
    
    writer.startElement("HyperlinksChanged");
    writer.writeText("false");
    writer.endElement(); // HyperlinksChanged
    
    writer.startElement("AppVersion");
    writer.writeText("16.0300");
    writer.endElement(); // AppVersion
    
    writer.endElement(); // Properties
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateDocPropsCoreXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("cp:coreProperties");
    writer.writeAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
    writer.writeAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
    writer.writeAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
    writer.writeAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
    writer.writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    
    if (!doc_properties_.title.empty()) {
        writer.startElement("dc:title");
        writer.writeText(doc_properties_.title.c_str());
        writer.endElement(); // dc:title
    }
    
    if (!doc_properties_.subject.empty()) {
        writer.startElement("dc:subject");
        writer.writeText(doc_properties_.subject.c_str());
        writer.endElement(); // dc:subject
    }
    
    if (!doc_properties_.author.empty()) {
        writer.startElement("dc:creator");
        writer.writeText(doc_properties_.author.c_str());
        writer.endElement(); // dc:creator
    }
    
    if (!doc_properties_.keywords.empty()) {
        writer.startElement("cp:keywords");
        writer.writeText(doc_properties_.keywords.c_str());
        writer.endElement(); // cp:keywords
    }
    
    if (!doc_properties_.comments.empty()) {
        writer.startElement("dc:description");
        writer.writeText(doc_properties_.comments.c_str());
        writer.endElement(); // dc:description
    }
    
    writer.startElement("cp:lastModifiedBy");
    writer.writeText(doc_properties_.author.c_str());
    writer.endElement(); // cp:lastModifiedBy
    
    writer.startElement("dcterms:created");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeText(formatTime(doc_properties_.created_time).c_str());
    writer.endElement(); // dcterms:created
    
    writer.startElement("dcterms:modified");
    writer.writeAttribute("xsi:type", "dcterms:W3CDTF");
    writer.writeText(formatTime(doc_properties_.modified_time).c_str());
    writer.endElement(); // dcterms:modified
    
    if (!doc_properties_.category.empty()) {
        writer.startElement("cp:category");
        writer.writeText(doc_properties_.category.c_str());
        writer.endElement(); // cp:category
    }
    
    if (!doc_properties_.status.empty()) {
        writer.startElement("cp:contentStatus");
        writer.writeText(doc_properties_.status.c_str());
        writer.endElement(); // cp:contentStatus
    }
    
    writer.endElement(); // cp:coreProperties
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateDocPropsCustomXML() const {
    if (custom_properties_.empty()) {
        return "";
    }
    
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Properties");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
    writer.writeAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    
    int pid = 2;
    for (const auto& prop : custom_properties_) {
        writer.startElement("property");
        writer.writeAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        writer.writeAttribute("pid", std::to_string(pid++).c_str());
        writer.writeAttribute("name", prop.name.c_str());
        
        writer.startElement("vt:lpwstr");
        writer.writeText(prop.value.c_str());
        writer.endElement(); // vt:lpwstr
        
        writer.endElement(); // property
    }
    
    writer.endElement(); // Properties
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateContentTypesXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Types");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
    
    // 默认类型
    writer.startElement("Default");
    writer.writeAttribute("Extension", "rels");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    writer.endElement(); // Default
    
    writer.startElement("Default");
    writer.writeAttribute("Extension", "xml");
    writer.writeAttribute("ContentType", "application/xml");
    writer.endElement(); // Default
    
    // 覆盖类型
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/xl/workbook.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    writer.endElement(); // Override
    
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", ("/xl/worksheets/sheet" + std::to_string(i + 1) + ".xml").c_str());
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        writer.endElement(); // Override
    }
    
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/xl/styles.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    writer.endElement(); // Override
    
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/xl/sharedStrings.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    writer.endElement(); // Override
    
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/docProps/core.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
    writer.endElement(); // Override
    
    writer.startElement("Override");
    writer.writeAttribute("PartName", "/docProps/app.xml");
    writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
    writer.endElement(); // Override
    
    if (!custom_properties_.empty()) {
        writer.startElement("Override");
        writer.writeAttribute("PartName", "/docProps/custom.xml");
        writer.writeAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
        writer.endElement(); // Override
    }
    
    writer.endElement(); // Types
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateRelsXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    writer.startElement("Relationship");
    writer.writeAttribute("Id", "rId1");
    writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    writer.writeAttribute("Target", "xl/workbook.xml");
    writer.endElement(); // Relationship
    
    writer.startElement("Relationship");
    writer.writeAttribute("Id", "rId2");
    writer.writeAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
    writer.writeAttribute("Target", "docProps/core.xml");
    writer.endElement(); // Relationship
    
    writer.startElement("Relationship");
    writer.writeAttribute("Id", "rId3");
    writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
    writer.writeAttribute("Target", "docProps/app.xml");
    writer.endElement(); // Relationship
    
    if (!custom_properties_.empty()) {
        writer.startElement("Relationship");
        writer.writeAttribute("Id", "rId4");
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
        writer.writeAttribute("Target", "docProps/custom.xml");
        writer.endElement(); // Relationship
    }
    
    writer.endElement(); // Relationships
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateWorkbookRelsXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("Relationships");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    
    // 工作表关系
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        writer.startElement("Relationship");
        writer.writeAttribute("Id", ("rId" + std::to_string(i + 1)).c_str());
        writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
        writer.writeAttribute("Target", getWorksheetRelPath(static_cast<int>(i + 1)).c_str());
        writer.endElement(); // Relationship
    }
    
    // 样式关系
    writer.startElement("Relationship");
    writer.writeAttribute("Id", ("rId" + std::to_string(worksheets_.size() + 1)).c_str());
    writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    writer.writeAttribute("Target", "styles.xml");
    writer.endElement(); // Relationship
    
    // 共享字符串关系
    writer.startElement("Relationship");
    writer.writeAttribute("Id", ("rId" + std::to_string(worksheets_.size() + 2)).c_str());
    writer.writeAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
    writer.writeAttribute("Target", "sharedStrings.xml");
    writer.endElement(); // Relationship
    
    writer.endElement(); // Relationships
    writer.endDocument();
    
    return writer.toString();
}

// ========== 格式管理内部方法 ==========

int Workbook::getFontId(const Format& format) {
    if (!format.hasFont()) {
        return 0; // 默认字体
    }
    
    size_t hash = format.hash();
    auto it = font_hash_to_id_.find(hash);
    if (it != font_hash_to_id_.end()) {
        return it->second;
    }
    
    int id = next_font_id_++;
    font_hash_to_id_[hash] = id;
    return id;
}

int Workbook::getFillId(const Format& format) {
    if (!format.hasFill()) {
        return 0; // 默认填充
    }
    
    size_t hash = format.hash();
    auto it = fill_hash_to_id_.find(hash);
    if (it != fill_hash_to_id_.end()) {
        return it->second;
    }
    
    int id = next_fill_id_++;
    fill_hash_to_id_[hash] = id;
    return id;
}

int Workbook::getBorderId(const Format& format) {
    if (!format.hasBorder()) {
        return 0; // 默认边框
    }
    
    size_t hash = format.hash();
    auto it = border_hash_to_id_.find(hash);
    if (it != border_hash_to_id_.end()) {
        return it->second;
    }
    
    int id = next_border_id_++;
    border_hash_to_id_[hash] = id;
    return id;
}

int Workbook::getXfId(const Format& format) {
    size_t hash = format.hash();
    auto it = xf_hash_to_id_.find(hash);
    if (it != xf_hash_to_id_.end()) {
        return it->second;
    }
    
    int id = next_xf_id_++;
    xf_hash_to_id_[hash] = id;
    return id;
}

// ========== 辅助函数 ==========

std::string Workbook::generateUniqueSheetName(const std::string& base_name) const {
    // 如果base_name不存在，直接返回
    if (getWorksheet(base_name) == nullptr) {
        return base_name;
    }
    
    // 如果base_name是"Sheet1"，从"Sheet2"开始尝试
    if (base_name == "Sheet1") {
        int counter = 2;
        std::string name = "Sheet" + std::to_string(counter);
        while (getWorksheet(name) != nullptr) {
            name = "Sheet" + std::to_string(++counter);
        }
        return name;
    }
    
    // 对于其他base_name，添加数字后缀
    int suffix_counter = 1;
    std::string name = base_name + std::to_string(suffix_counter);
    while (getWorksheet(name) != nullptr) {
        name = base_name + std::to_string(++suffix_counter);
    }
    
    return name;
}

bool Workbook::validateSheetName(const std::string& name) const {
    // 检查长度
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查非法字符
    const std::string invalid_chars = ":\\/?*[]";
    if (name.find_first_of(invalid_chars) != std::string::npos) {
        return false;
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 不检查是否已存在，因为这个方法也被用于验证新名称
    // 重复名称的检查应该在调用方处理
    
    return true;
}

void Workbook::collectSharedStrings() {
    shared_strings_.clear();
    shared_strings_list_.clear();
    
    for (const auto& worksheet : worksheets_) {
        // 这里需要访问工作表的单元格来收集字符串
        // 简化版本，实际实现需要遍历所有字符串单元格
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    if (cell.isString()) {
                        addSharedString(cell.getStringValue());
                    }
                }
            }
        }
    }
}

void Workbook::updateFormatIndices() {
    for (auto& [id, format] : formats_) {
        format->setFontIndex(getFontId(*format));
        format->setFillIndex(getFillId(*format));
        format->setBorderIndex(getBorderId(*format));
        format->setXfIndex(getXfId(*format));
    }
}

std::string Workbook::getWorksheetPath(int sheet_id) const {
    return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::getWorksheetRelPath(int sheet_id) const {
    return "worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::formatTime(const std::tm& time) const {
    std::ostringstream oss;
    oss << std::put_time(&time, "%Y-%m-%dT%H:%M:%SZ");
    return oss.str();
}

std::string Workbook::hashPassword(const std::string& password) const {
    // 简化的密码哈希实现
    // 实际应该使用Excel的密码哈希算法
    std::hash<std::string> hasher;
    size_t hash = hasher(password);
    
    std::ostringstream oss;
    oss << std::hex << std::uppercase << hash;
    return oss.str();
}

void Workbook::setHighPerformanceMode(bool enable) {
    if (enable) {
        LOG_INFO("Enabling high performance mode");
        
        // 禁用共享字符串以避免哈希和排序开销
        options_.use_shared_strings = false;
        
        // 启用流式XML写入以减少内存占用
        options_.streaming_xml = true;
        
        // 设置较大的行缓冲以减少I/O次数
        options_.row_buffer_size = 5000;
        
        // 使用较低的压缩级别以提高速度
        options_.compression_level = 1;  // 快速压缩
        
        // 增大XML缓冲区以减少内存分配
        options_.xml_buffer_size = 4 * 1024 * 1024;  // 4MB
        
        LOG_INFO("High performance mode configured: SharedStrings=OFF, StreamingXML=ON, RowBuffer={}, Compression={}",
                options_.row_buffer_size, options_.compression_level);
    } else {
        LOG_INFO("Disabling high performance mode, using balanced settings");
        
        // 恢复平衡设置
        options_.use_shared_strings = true;
        options_.streaming_xml = false;
        options_.row_buffer_size = 1000;
        options_.compression_level = 6;  // 平衡压缩
        options_.xml_buffer_size = 1024 * 1024;  // 1MB
    }
}

bool Workbook::generateWorksheetXMLStreaming(const std::shared_ptr<Worksheet>& worksheet, const std::string& path) {
    LOG_DEBUG("Generating worksheet XML using streaming mode: {}", path);
    
    try {
        // 创建XMLStreamWriter并设置为回调模式，直接写入文件
        xml::XMLStreamWriter writer;
        
        // 设置回调模式，直接将数据块写入ZIP文件
        writer.setCallbackMode([this, &path](const std::string& chunk) {
            // 这里需要实现直接写入ZIP文件的逻辑
            // 暂时先累积到字符串中，后续可以优化为直接写入ZIP流
            static std::string accumulated_xml;
            accumulated_xml += chunk;
            
            // 当累积到一定大小时，可以考虑分批写入
            if (accumulated_xml.size() > options_.xml_buffer_size) {
                // 这里可以实现分批写入逻辑
                // 目前先简单累积
            }
        }, true); // 启用自动刷新
        
        // 写入XML声明和工作表开始标签
        writer.startDocument();
        writer.startElement("worksheet");
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        
        // 工作表属性（如果有）
        if (worksheet->hasFrozenPanes()) {
            auto freeze_info = worksheet->getFreezeInfo();
            writer.startElement("sheetViews");
            writer.startElement("sheetView");
            writer.writeAttribute("workbookViewId", "0");
            writer.startElement("pane");
            writer.writeAttribute("ySplit", std::to_string(freeze_info.row));
            writer.writeAttribute("xSplit", std::to_string(freeze_info.col));
            writer.writeAttribute("topLeftCell", "A1"); // 简化处理
            writer.writeAttribute("activePane", "bottomRight");
            writer.writeAttribute("state", "frozen");
            writer.endElement(); // pane
            writer.endElement(); // sheetView
            writer.endElement(); // sheetViews
        }
        
        // 列信息（如果有）
        // 这里简化处理，实际需要从worksheet获取列信息
        
        // 工作表数据 - 这是最重要的部分，需要流式处理
        writer.startElement("sheetData");
        
        // 获取使用范围
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        // 按行处理数据，实现真正的流式写入
        size_t rows_processed = 0;
        for (int row = 0; row <= max_row; ++row) {
            bool has_data_in_row = false;
            
            // 检查这一行是否有数据
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    has_data_in_row = true;
                    break;
                }
            }
            
            if (has_data_in_row) {
                // 写入行开始标签
                writer.startElement("row");
                writer.writeAttribute("r", std::to_string(row + 1));
                
                for (int col = 0; col <= max_col; ++col) {
                    if (worksheet->hasCellAt(row, col)) {
                        const auto& cell = worksheet->getCell(row, col);
                        std::string cell_ref = columnToLetter(col) + std::to_string(row + 1);
                        
                        writer.startElement("c");
                        writer.writeAttribute("r", cell_ref);
                        
                        // 根据单元格类型添加属性和内容
                        if (cell.isString()) {
                            if (options_.use_shared_strings) {
                                int sst_index = getSharedStringIndex(cell.getStringValue());
                                if (sst_index >= 0) {
                                    writer.writeAttribute("t", "s");
                                    writer.startElement("v");
                                    writer.writeText(std::to_string(sst_index));
                                    writer.endElement(); // v
                                } else {
                                    writer.writeAttribute("t", "inlineStr");
                                    writer.startElement("is");
                                    writer.startElement("t");
                                    writer.writeText(cell.getStringValue());
                                    writer.endElement(); // t
                                    writer.endElement(); // is
                                }
                            } else {
                                writer.writeAttribute("t", "inlineStr");
                                writer.startElement("is");
                                writer.startElement("t");
                                writer.writeText(cell.getStringValue());
                                writer.endElement(); // t
                                writer.endElement(); // is
                            }
                        } else if (cell.isNumber()) {
                            writer.startElement("v");
                            writer.writeText(std::to_string(cell.getNumberValue()));
                            writer.endElement(); // v
                        } else if (cell.isBoolean()) {
                            writer.writeAttribute("t", "b");
                            writer.startElement("v");
                            writer.writeText(cell.getBooleanValue() ? "1" : "0");
                            writer.endElement(); // v
                        } else if (cell.isFormula()) {
                            writer.startElement("f");
                            writer.writeText(cell.getStringValue());
                            writer.endElement(); // f
                        }
                        
                        writer.endElement(); // c
                    }
                }
                
                writer.endElement(); // row
                rows_processed++;
                
                // 每处理一定数量的行就刷新缓冲区
                if (rows_processed % options_.row_buffer_size == 0) {
                    writer.flushBuffer();
                    LOG_DEBUG("Processed {} rows, flushed buffer", rows_processed);
                }
            }
        }
        
        writer.endElement(); // sheetData
        
        // 合并单元格（如果有）
        const auto& merge_ranges = worksheet->getMergeRanges();
        if (!merge_ranges.empty()) {
            writer.startElement("mergeCells");
            writer.writeAttribute("count", std::to_string(merge_ranges.size()));
            
            for (const auto& range : merge_ranges) {
                writer.startElement("mergeCell");
                std::string range_ref = columnToLetter(range.first_col) + std::to_string(range.first_row + 1) +
                                      ":" + columnToLetter(range.last_col) + std::to_string(range.last_row + 1);
                writer.writeAttribute("ref", range_ref);
                writer.endElement(); // mergeCell
            }
            
            writer.endElement(); // mergeCells
        }
        
        // 自动筛选（如果有）
        if (worksheet->hasAutoFilter()) {
            auto filter_range = worksheet->getAutoFilterRange();
            writer.startElement("autoFilter");
            std::string range_ref = columnToLetter(filter_range.first_col) + std::to_string(filter_range.first_row + 1) +
                                  ":" + columnToLetter(filter_range.last_col) + std::to_string(filter_range.last_row + 1);
            writer.writeAttribute("ref", range_ref);
            writer.endElement(); // autoFilter
        }
        
        writer.endElement(); // worksheet
        writer.endDocument();
        
        // 最终刷新并获取完整的XML内容
        writer.flushBuffer();
        std::string final_xml = writer.toString();
        
        // 将XML写入文件
        bool success = file_manager_->writeFile(path, final_xml);
        
        if (success) {
            LOG_INFO("Successfully generated streaming worksheet XML: {}, {} rows processed", path, rows_processed);
        } else {
            LOG_ERROR("Failed to write streaming worksheet XML: {}", path);
        }
        
        return success;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception in streaming worksheet XML generation: {}", e.what());
        return false;
    }
}

// 辅助方法：列号转字母
std::string Workbook::columnToLetter(int col) const {
    std::string result;
    while (col >= 0) {
        result = static_cast<char>('A' + (col % 26)) + result;
        col = col / 26 - 1;
    }
    return result;
}

// 辅助方法：XML转义
std::string Workbook::escapeXML(const std::string& text) const {
    std::string result;
    result.reserve(text.size() * 1.2);
    
    for (char c : text) {
        switch (c) {
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '&': result += "&amp;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default:
                if (c < 0x20 && c != 0x09 && c != 0x0A && c != 0x0D) {
                    continue; // 跳过无效控制字符
                }
                result.push_back(c);
                break;
        }
    }
    
    return result;
}

}} // namespace fastexcel::core