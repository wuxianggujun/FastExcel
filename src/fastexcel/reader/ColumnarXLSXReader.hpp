#pragma once

#include "fastexcel/core/columnar/ReadOnlyWorkbook.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include <string>
#include <memory>
#include <vector>

namespace fastexcel {
namespace reader {

// 列式XLSX读取器 - ReadOnlyWorkbook的主要入口
class ColumnarXLSXReader {
private:
    std::unique_ptr<archive::ZipReader> zip_reader_;
    core::columnar::ReadOnlyOptions options_;
    
public:
    explicit ColumnarXLSXReader(const core::columnar::ReadOnlyOptions& options = {});
    ~ColumnarXLSXReader() = default;
    
    // 禁用拷贝
    ColumnarXLSXReader(const ColumnarXLSXReader&) = delete;
    ColumnarXLSXReader& operator=(const ColumnarXLSXReader&) = delete;
    
    // 主要解析方法
    std::unique_ptr<core::columnar::ReadOnlyWorkbook> parse(const std::string& filename);
    
    // 流式解析 - 用于超大文件
    bool parseStream(const std::string& filename, 
                    core::columnar::ReadOnlyWorkbook* workbook);
                    
private:
    // 解析步骤
    bool parseSharedStrings(core::columnar::ReadOnlyWorkbook* workbook);
    bool parseWorkbook(core::columnar::ReadOnlyWorkbook* workbook);
    bool parseWorksheets(core::columnar::ReadOnlyWorkbook* workbook);
    bool parseWorksheet(const std::string& worksheet_path,
                       core::columnar::ReadOnlyWorksheet* worksheet);
                       
    // 辅助方法
    std::vector<std::string> getWorksheetPaths() const;
    std::string getWorksheetName(const std::string& worksheet_path) const;
};

}} // namespace fastexcel::reader