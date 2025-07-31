#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Format.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>

namespace fastexcel {
namespace core {

class Workbook : public std::enable_shared_from_this<Workbook> {
private:
    std::string filename_;
    std::vector<std::shared_ptr<Worksheet>> worksheets_;
    std::map<std::string, std::shared_ptr<Format>> formats_;
    std::unique_ptr<archive::FileManager> file_manager_;
    int next_format_id_ = 164; // Excel内置格式之后从164开始
    int next_sheet_id_ = 1;
    bool is_open_ = false;
    
public:
    explicit Workbook(const std::string& filename);
    ~Workbook();
    
    // 文件操作
    bool open();
    bool save();
    bool close();
    
    // 工作表管理
    std::shared_ptr<Worksheet> addWorksheet(const std::string& name = "");
    std::shared_ptr<Worksheet> getWorksheet(const std::string& name);
    std::shared_ptr<Worksheet> getWorksheet(size_t index);
    std::shared_ptr<const Worksheet> getWorksheet(const std::string& name) const;
    std::shared_ptr<const Worksheet> getWorksheet(size_t index) const;
    size_t getWorksheetCount() const { return worksheets_.size(); }
    
    // 格式管理
    std::shared_ptr<Format> createFormat();
    std::shared_ptr<Format> getFormat(int format_id) const;
    
    // 获取状态
    bool isOpen() const { return is_open_; }
    std::string getFilename() const { return filename_; }
    
private:
    // 生成Excel文件结构
    bool generateExcelStructure();
    
    // 生成各种XML文件
    std::string generateWorkbookXML() const;
    std::string generateStylesXML() const;
    std::string generateSharedStringsXML() const;
    std::string generateWorksheetXML(const std::shared_ptr<Worksheet>& worksheet) const;
    
    // 辅助函数
    std::string generateUniqueSheetName(const std::string& base_name) const;
    bool validateSheetName(const std::string& name) const;
    void collectSharedStrings(std::map<std::string, int>& shared_strings) const;
    
    // 文件路径生成
    std::string getWorksheetPath(int sheet_id) const;
    std::string getWorksheetRelPath(int sheet_id) const;
};

}} // namespace fastexcel::core
