//
// Created by wuxianggujun on 25-8-4.
//

#pragma once
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/ErrorCode.hpp"
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

  // 系统层API：高性能，使用错误码
  core::ErrorCode open();
  core::ErrorCode loadWorkbook(std::unique_ptr<core::Workbook>& workbook);
  core::ErrorCode close();

  core::ErrorCode loadWorksheet(const std::string& name, std::shared_ptr<core::Worksheet>& worksheet);
  core::ErrorCode getWorksheetNames(std::vector<std::string>& names);
  
  core::ErrorCode getMetadata(WorkbookMetadata& metadata);
  core::ErrorCode getDefinedNames(std::vector<std::string>& names);

  // 状态查询
  bool isOpen() const { return is_open_; }
  const std::string& getFilename() const { return filename_; }
  
  // 样式访问
  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& getStyles() const { return styles_; }

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
    std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>> styles_; // index -> format descriptor
    
    // 内部解析方法 - 系统层，使用ErrorCode
    core::ErrorCode parseWorkbookXML();
    core::ErrorCode parseWorksheetXML(const std::string& path, core::Worksheet* worksheet);
    core::ErrorCode parseStylesXML();
    core::ErrorCode parseSharedStringsXML();
    core::ErrorCode parseContentTypesXML();
    core::ErrorCode parseRelationshipsXML();
    core::ErrorCode parseDocPropsXML();
    
    // 辅助方法
    std::string extractXMLFromZip(const std::string& path);
    bool validateXLSXStructure();
    std::string getCellValue(const std::string& cell_xml, core::CellType& type);
    std::shared_ptr<core::FormatDescriptor> getStyleByIndex(int index);
    
    // 新增的XML解析辅助方法
    std::string extractAttribute(const std::string& xml, const std::string& attr_name);
    bool parseWorkbookRelationships(std::unordered_map<std::string, std::string>& relationships);
    bool parseDefinedNames(const std::string& xml_content);
};
} // namespace reader
} // namespace fastexcel
