#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <sstream>
#include <algorithm>

namespace fastexcel {
namespace core {

Workbook::Workbook(const std::string& filename) : filename_(filename) {
    file_manager_ = std::make_unique<archive::FileManager>(filename);
}

Workbook::~Workbook() {
    close();
}

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
        LOG_ERROR("Workbook is not open");
        return false;
    }
    
    try {
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

bool Workbook::close() {
    if (is_open_) {
        file_manager_->close();
        is_open_ = false;
        LOG_INFO("Workbook closed: {}", filename_);
    }
    return true;
}

std::shared_ptr<Worksheet> Workbook::addWorksheet(const std::string& name) {
    if (!is_open_) {
        LOG_ERROR("Workbook is not open");
        return nullptr;
    }
    
    std::string sheet_name = name.empty() ? generateUniqueSheetName("Sheet") : name;
    
    if (!validateSheetName(sheet_name)) {
        LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, shared_from_this(), next_sheet_id_++);
    worksheets_.push_back(worksheet);
    
    LOG_DEBUG("Added worksheet: {}", sheet_name);
    return worksheet;
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

std::shared_ptr<Format> Workbook::createFormat() {
    auto format = std::make_shared<Format>();
    format->setFormatId(next_format_id_++);
    
    // 存储格式以便后续使用
    std::string key = std::to_string(format->getFormatId());
    formats_[key] = format;
    
    return format;
}

std::shared_ptr<Format> Workbook::getFormat(int format_id) const {
    std::string key = std::to_string(format_id);
    auto it = formats_.find(key);
    if (it != formats_.end()) {
        return it->second;
    }
    return nullptr;
}

bool Workbook::generateExcelStructure() {
    // 生成工作簿XML
    std::string workbook_xml = generateWorkbookXML();
    if (!file_manager_->writeFile("xl/workbook.xml", workbook_xml)) {
        return false;
    }
    
    // 生成样式XML
    std::string styles_xml = generateStylesXML();
    if (!file_manager_->writeFile("xl/styles.xml", styles_xml)) {
        return false;
    }
    
    // 生成共享字符串XML
    std::string shared_strings_xml = generateSharedStringsXML();
    if (!file_manager_->writeFile("xl/sharedStrings.xml", shared_strings_xml)) {
        return false;
    }
    
    // 生成工作表XML
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        std::string worksheet_xml = generateWorksheetXML(worksheets_[i]);
        std::string worksheet_path = getWorksheetPath(static_cast<int>(i + 1));
        if (!file_manager_->writeFile(worksheet_path, worksheet_xml)) {
            return false;
        }
    }
    
    return true;
}

std::string Workbook::generateWorkbookXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 添加工作表视图
    writer.startElement("bookViews");
    writer.writeEmptyElement("workbookView");
    writer.writeAttribute("xWindow", "240");
    writer.writeAttribute("yWindow", "15");
    writer.writeAttribute("windowWidth", "16095");
    writer.writeAttribute("windowHeight", "9660");
    writer.endElement(); // bookViews
    
    // 添加工作表
    writer.startElement("sheets");
    for (size_t i = 0; i < worksheets_.size(); ++i) {
        writer.writeEmptyElement("sheet");
        writer.writeAttribute("name", worksheets_[i]->getName().c_str());
        std::string sheet_id = std::to_string(worksheets_[i]->getSheetId());
        writer.writeAttribute("sheetId", sheet_id.c_str());
        std::string r_id = "rId" + std::to_string(i + 1);
        writer.writeAttribute("r:id", r_id.c_str());
    }
    writer.endElement(); // sheets
    
    writer.endElement(); // workbook
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateStylesXML() const {
    xml::XMLStreamWriter writer;
    writer.startDocument();
    writer.startElement("styleSheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    // 添加数字格式
    writer.startElement("numFmts");
    writer.writeAttribute("count", "0");
    writer.endElement(); // numFmts
    
    // 添加字体
    writer.startElement("fonts");
    writer.writeAttribute("count", "1");
    writer.startElement("font");
    writer.writeEmptyElement("sz");
    writer.writeAttribute("val", "11");
    writer.writeEmptyElement("name");
    writer.writeAttribute("val", "Calibri");
    writer.writeEmptyElement("family");
    writer.writeAttribute("val", "2");
    writer.writeEmptyElement("scheme");
    writer.writeAttribute("val", "minor");
    writer.endElement(); // font
    writer.endElement(); // fonts
    
    // 添加填充
    writer.startElement("fills");
    writer.writeAttribute("count", "2");
    writer.writeEmptyElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "none");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    writer.writeEmptyElement("fill");
    writer.startElement("patternFill");
    writer.writeAttribute("patternType", "gray125");
    writer.endElement(); // patternFill
    writer.endElement(); // fill
    writer.endElement(); // fills
    
    // 添加边框
    writer.startElement("borders");
    writer.writeAttribute("count", "1");
    writer.startElement("border");
    writer.writeEmptyElement("left");
    writer.writeEmptyElement("right");
    writer.writeEmptyElement("top");
    writer.writeEmptyElement("bottom");
    writer.writeEmptyElement("diagonal");
    writer.endElement(); // border
    writer.endElement(); // borders
    
    // 添加单元格样式
    writer.startElement("cellXfs");
    writer.writeAttribute("count", "1");
    writer.writeEmptyElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.writeAttribute("xfId", "0");
    writer.endElement(); // cellXfs
    
    // 添加单元格样式
    writer.startElement("cellStyles");
    writer.writeAttribute("count", "1");
    writer.writeEmptyElement("cellStyle");
    writer.writeAttribute("name", "Normal");
    writer.writeAttribute("xfId", "0");
    writer.writeAttribute("builtinId", "0");
    writer.endElement(); // cellStyles
    
    writer.endElement(); // styleSheet
    writer.endDocument();
    
    return writer.toString();
}

std::string Workbook::generateSharedStringsXML() const {
    xml::SharedStrings shared_strings;
    
    // 收集所有字符串
    std::map<std::string, int> string_map;
    collectSharedStrings(string_map);
    
    // 添加字符串到共享字符串表
    for (const auto& [str, index] : string_map) {
        shared_strings.addString(str);
    }
    
    return shared_strings.generate();
}

std::string Workbook::generateWorksheetXML(const std::shared_ptr<Worksheet>& worksheet) const {
    return worksheet->generateXML();
}

std::string Workbook::generateUniqueSheetName(const std::string& base_name) const {
    std::string name = base_name;
    int counter = 1;
    
    while (getWorksheet(name) != nullptr) {
        name = base_name + std::to_string(counter++);
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
    
    // 检查是否已存在
    if (getWorksheet(name) != nullptr) {
        return false;
    }
    
    return true;
}

void Workbook::collectSharedStrings(std::map<std::string, int>& shared_strings) const {
    int index = 0;
    for (const auto& worksheet : worksheets_) {
        // 这里需要访问Worksheet的私有数据来收集字符串
        // 简化版本，实际实现可能需要添加友元类或公共方法
        // 暂时留空
    }
}

std::string Workbook::getWorksheetPath(int sheet_id) const {
    return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::getWorksheetRelPath(int sheet_id) const {
    return "worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

}} // namespace fastexcel::core