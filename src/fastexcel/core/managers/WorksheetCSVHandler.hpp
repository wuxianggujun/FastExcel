#pragma once

#include "fastexcel/core/CSVProcessor.hpp"
#include <string>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;

class WorksheetCSVHandler {
public:
    explicit WorksheetCSVHandler(Worksheet& worksheet);
    
    // CSV导入
    CSVParseInfo loadFromCSV(const std::string& filepath, const CSVOptions& options = CSVOptions());
    CSVParseInfo loadFromCSVString(const std::string& csv_content, const CSVOptions& options = CSVOptions());
    
    // CSV导出
    bool saveAsCSV(const std::string& filepath, const CSVOptions& options = CSVOptions()) const;
    std::string toCSVString(const CSVOptions& options = CSVOptions()) const;
    std::string rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                                const CSVOptions& options = CSVOptions()) const;
    
    // CSV工具方法
    static CSVParseInfo previewCSV(const std::string& filepath, const CSVOptions& options = CSVOptions());
    static CSVOptions detectCSVOptions(const std::string& filepath);
    static bool isCSVFile(const std::string& filepath);
    
    // 便捷方法
    std::string getCellDisplayValue(int row, int col) const;

private:
    Worksheet& worksheet_;
};

} // namespace core
} // namespace fastexcel