#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#pragma once
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"  // 添加WorkbookOptions支持
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/theme/Theme.hpp"

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

  // 设置工作簿选项（用于列式存储等优化）
  void setWorkbookOptions(const core::WorkbookOptions& options) { options_ = options; }
  const core::WorkbookOptions& getWorkbookOptions() const { return options_; }

  core::ErrorCode loadWorksheet(const std::string& name, std::shared_ptr<core::Worksheet>& worksheet);
  core::ErrorCode getSheetNames(std::vector<std::string>& names);
  
  core::ErrorCode getMetadata(WorkbookMetadata& metadata);
  core::ErrorCode getDefinedNames(std::vector<std::string>& names);

  // 状态查询
  bool isOpen() const { return is_open_; }
  const std::string& getFilename() const { return filename_; }
  
  // 样式访问
  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& getStyles() const { return styles_; }
  const std::unordered_map<int, int>& getStyleIdMapping() const { return style_id_mapping_; }
  
  // 主题访问
  const std::string& getThemeXML() const { return theme_xml_; }
   // 解析后的主题对象（可能为nullptr，表示未解析或无主题文件）
   const theme::Theme* getParsedTheme() const { return parsed_theme_.get(); }

private:
    // 成员变量
    core::Path filepath_;
    std::string filename_;
    std::unique_ptr<archive::ZipArchive> zip_archive_;
    bool is_open_;
    
    // 工作簿选项（用于列式存储等优化）
    core::WorkbookOptions options_;
    
    // 解析后的数据
    WorkbookMetadata metadata_;
    std::vector<std::string> worksheet_names_;
    std::vector<std::string> defined_names_;
    std::unordered_map<std::string, std::string> worksheet_paths_;  // name -> path
    std::unordered_map<int, std::string> shared_strings_;           // index -> string
    std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>> styles_; // index -> format descriptor
    std::unordered_map<int, int> style_id_mapping_; // 原始样式ID -> FormatRepository中的ID
    std::string theme_xml_; // 原始主题文件内容
    std::unique_ptr<theme::Theme> parsed_theme_; // 解析后的主题对象
    
    // 内部解析方法 - 系统层，使用ErrorCode
    core::ErrorCode parseWorkbookXML();
    core::ErrorCode parseWorksheetXML(const std::string& path, core::Worksheet* worksheet);
    core::ErrorCode parseStylesXML();
    core::ErrorCode parseSharedStringsXML();
    core::ErrorCode parseDocPropsXML();
    core::ErrorCode parseThemeXML();
    
    // 辅助方法
    std::string extractXMLFromZip(const std::string& path);
    bool validateXLSXStructure();
    
    // 字符串解析辅助方法
    std::string extractAttribute(const std::string& xml, const std::string& attr_name);
    bool parseWorkbookRelationships(std::unordered_map<std::string, std::string>& relationships);
    bool parseDefinedNames(const std::string& xml_content);
private:
};
} // namespace reader
} // namespace fastexcel
