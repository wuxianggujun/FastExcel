//
// Created by wuxianggujun on 25-8-4.
//

#pragma once
#include "fastexcel/core/Format.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <memory>
#include <unordered_map>
#include <vector>

namespace fastexcel {
namespace reader {

// 工作簿元数据结构
struct WorkbookMetadata {
    std::string title;
    std::string subject;
    std::string author;
    std::string manager;
    std::string company;
    std::string category;
    std::string keywords;
    std::string comments;
    std::string created_time;
    std::string modified_time;
    std::string application;
    std::string app_version;
    
    WorkbookMetadata() = default;
};

class XLSXReader {

public:
  explicit XLSXReader(const std::string& filename);
  explicit XLSXReader(const core::Path& path);
  ~XLSXReader();

  bool open();
  std::unique_ptr<core::Workbook> loadWorkbook();
  bool close();

  std::shared_ptr<core::Worksheet> loadWorksheet(const std::string& name);
  std::vector<std::string> getWorksheetNames();

  WorkbookMetadata getMetadata();
  std::vector<std::string> getDefinedNames();

private:
    // 成员变量
    core::Path filepath_;
    std::string filename_;
    std::unique_ptr<archive::ZipArchive> zip_archive_;
    bool is_open_;
    
    // 解析后的数据
    WorkbookMetadata metadata_;
    std::vector<std::string> worksheet_names_;
    std::vector<std::string> defined_names_;
    std::unordered_map<std::string, std::string> worksheet_paths_;  // name -> path
    std::unordered_map<int, std::string> shared_strings_;           // index -> string
    std::unordered_map<int, std::shared_ptr<core::Format>> styles_; // index -> format
    
    // 解析方法
    bool parseWorkbookXML();
    bool parseWorksheetXML(const std::string& path, core::Worksheet* worksheet);
    bool parseStylesXML();
    bool parseSharedStringsXML();
    bool parseContentTypesXML();
    bool parseRelationshipsXML();
    bool parseDocPropsXML();
    
    // 辅助方法
    std::string extractXMLFromZip(const std::string& path);
    bool validateXLSXStructure();
    std::string getCellValue(const std::string& cell_xml, core::CellType& type);
    std::shared_ptr<core::Format> getStyleByIndex(int index);
    
    // 新增的XML解析辅助方法
    std::string extractAttribute(const std::string& xml, const std::string& attr_name);
    bool parseWorkbookRelationships(std::unordered_map<std::string, std::string>& relationships);
    bool parseDefinedNames(const std::string& xml_content);
};
} // namespace reader
} // namespace fastexcel
