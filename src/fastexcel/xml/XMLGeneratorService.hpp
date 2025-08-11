#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include <string>
#include <memory>
#include <functional>
#include <sstream>

namespace fastexcel {

// 前向声明
namespace core {
    class Workbook;
    class Worksheet;
}

namespace xml {

/**
 * @brief XML生成器服务接口
 * 遵循接口隔离原则 (ISP)：每个接口只包含客户端需要的方法
 */
class IXMLGenerator {
public:
    virtual ~IXMLGenerator() = default;
    
    virtual std::string generateWorkbookXML() = 0;
    virtual std::string generateWorksheetXML(const std::string& sheet_name) = 0;
    virtual std::string generateStylesXML() = 0;
    virtual std::string generateSharedStringsXML() = 0;
    virtual std::string generateContentTypesXML() = 0;
    virtual std::string generateWorkbookRelsXML() = 0;
};

/**
 * @brief 基于现有Workbook的XML生成器实现
 * 这是一个适配器类，将现有的Workbook XML生成方法适配到新接口
 * 遵循单一职责原则 (SRP)：只负责XML生成
 */
class WorkbookXMLGenerator : public IXMLGenerator {
private:
    core::Workbook* workbook_;  // 不拥有所有权，只是引用
    
    // 通用的回调到字符串转换器
    std::string callbackToString(const std::function<void(const std::function<void(const char*, size_t)>&)>& generator) {
        std::ostringstream buffer;
        auto callback = [&buffer](const char* data, size_t size) {
            buffer.write(data, static_cast<std::streamsize>(size));
        };
        generator(callback);
        return buffer.str();
    }
    
public:
    explicit WorkbookXMLGenerator(core::Workbook* workbook) 
        : workbook_(workbook) {}
    
    std::string generateWorkbookXML() override {
        if (!workbook_) return \"\";
        return callbackToString([this](const auto& callback) {
            workbook_->generateWorkbookXML(callback);
        });
    }
    
    std::string generateWorksheetXML(const std::string& sheet_name) override {
        if (!workbook_) return \"\";
        auto worksheet = workbook_->getSheet(sheet_name);
        if (!worksheet) return \"\";
        
        return callbackToString([this, worksheet](const auto& callback) {
            workbook_->generateWorksheetXML(worksheet, callback);
        });
    }
    
    std::string generateStylesXML() override {
        if (!workbook_) return \"\";
        return callbackToString([this](const auto& callback) {
            workbook_->generateStylesXML(callback);
        });
    }
    
    std::string generateSharedStringsXML() override {
        if (!workbook_) return \"\";
        return callbackToString([this](const auto& callback) {
            workbook_->generateSharedStringsXML(callback);
        });
    }
    
    std::string generateContentTypesXML() override {
        if (!workbook_) return \"\";
        return callbackToString([this](const auto& callback) {
            workbook_->generateContentTypesXML(callback);
        });
    }
    
    std::string generateWorkbookRelsXML() override {
        if (!workbook_) return \"\";
        return callbackToString([this](const auto& callback) {
            workbook_->generateWorkbookRelsXML(callback);
        });
    }
};

/**
 * @brief 轻量级XML生成器
 * 用于简单场景，不依赖完整的Workbook对象
 * 遵循YAGNI原则：只实现当前需要的功能
 */
class LightweightXMLGenerator : public IXMLGenerator {
private:
    struct WorkbookData {
        std::vector<std::string> sheet_names;
        std::string active_sheet;
    } data_;
    
public:
    void addSheet(const std::string& name) {
        data_.sheet_names.push_back(name);
    }
    
    void setActiveSheet(const std::string& name) {
        data_.active_sheet = name;
    }
    
    std::string generateWorkbookXML() override {
        std::ostringstream xml;
        xml << \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\";
        xml << \"<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>\\n\";
        xml << \"  <sheets>\\n\";
        
        for (size_t i = 0; i < data_.sheet_names.size(); ++i) {
            xml << \"    <sheet name='\" << data_.sheet_names[i] 
                << \"' sheetId='\" << (i + 1) 
                << \"' r:id='rId\" << (i + 1) << \"'/>\\n\";
        }
        
        xml << \"  </sheets>\\n\";
        xml << \"</workbook>\";
        return xml.str();
    }
    
    std::string generateWorksheetXML(const std::string& sheet_name) override {
        std::ostringstream xml;
        xml << \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\";
        xml << \"<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>\\n\";
        xml << \"  <sheetData/>\\n\";
        xml << \"</worksheet>\";
        return xml.str();
    }
    
    std::string generateStylesXML() override {
        return \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\"
               \"<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>\\n\"
               \"  <fonts count='1'><font><sz val='11'/><name val='Calibri'/></font></fonts>\\n\"
               \"  <fills count='2'><fill><patternFill patternType='none'/></fill>\"
               \"  <fill><patternFill patternType='gray125'/></fill></fills>\\n\"
               \"  <borders count='1'><border><left/><right/><top/><bottom/><diagonal/></border></borders>\\n\"
               \"  <cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs>\\n\"
               \"  <cellXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/></cellXfs>\\n\"
               \"</styleSheet>\";
    }
    
    std::string generateSharedStringsXML() override {
        return \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\"
               \"<sst xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count='0' uniqueCount='0'/>\\n\";
    }
    
    std::string generateContentTypesXML() override {
        std::ostringstream xml;
        xml << \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\";
        xml << \"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>\\n\";
        xml << \"  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>\\n\";
        xml << \"  <Default Extension='xml' ContentType='application/xml'/>\\n\";
        xml << \"  <Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>\\n\";
        
        for (size_t i = 0; i < data_.sheet_names.size(); ++i) {
            xml << \"  <Override PartName='/xl/worksheets/sheet\" << (i + 1) 
                << \".xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>\\n\";
        }
        
        xml << \"  <Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/>\\n\";
        xml << \"  <Override PartName='/xl/sharedStrings.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'/>\\n\";
        xml << \"</Types>\";
        return xml.str();
    }
    
    std::string generateWorkbookRelsXML() override {
        std::ostringstream xml;
        xml << \"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\\n\";
        xml << \"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>\\n\";
        
        for (size_t i = 0; i < data_.sheet_names.size(); ++i) {
            xml << \"  <Relationship Id='rId\" << (i + 1) 
                << \"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' \"
                << \"Target='worksheets/sheet\" << (i + 1) << \".xml'/>\\n\";
        }
        
        xml << \"  <Relationship Id='rId\" << (data_.sheet_names.size() + 1) 
            << \"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' \"
            << \"Target='styles.xml'/>\\n\";
        xml << \"  <Relationship Id='rId\" << (data_.sheet_names.size() + 2)
            << \"' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings' \"
            << \"Target='sharedStrings.xml'/>\\n\";
        xml << \"</Relationships>\";
        return xml.str();
    }
};

} // namespace xml
} // namespace fastexcel