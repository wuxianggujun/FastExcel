//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include <string>
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace reader {

/**
 * @brief 工作表解析器
 *
 * 负责解析Excel文件中的worksheet*.xml文件，
 * 提取单元格数据并填充到Worksheet对象中
 */
class WorksheetParser {
public:
    WorksheetParser() = default;
    ~WorksheetParser() = default;
    
    /**
     * @brief 解析工作表XML内容
     * @param xml_content XML内容
     * @param worksheet 目标工作表对象
     * @param shared_strings 共享字符串映射
     * @param styles 样式映射
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content,
               core::Worksheet* worksheet,
               const std::unordered_map<int, std::string>& shared_strings,
               const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);

private:
    // 辅助方法
    bool parseSheetData(const std::string& xml_content,
                       core::Worksheet* worksheet,
                       const std::unordered_map<int, std::string>& shared_strings,
                       const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);
    
    bool parseRow(const std::string& row_xml,
                  core::Worksheet* worksheet,
                  const std::unordered_map<int, std::string>& shared_strings,
                  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);
    
    bool parseCell(const std::string& cell_xml,
                   core::Worksheet* worksheet,
                   const std::unordered_map<int, std::string>& shared_strings,
                   const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);
    
    // 工具方法
    std::pair<int, int> parseCellReference(const std::string& ref);
    std::string extractCellValue(const std::string& cell_xml);
    std::string extractCellType(const std::string& cell_xml);
    int extractStyleIndex(const std::string& cell_xml);
    std::string decodeXMLEntities(const std::string& text);
    int columnLetterToNumber(const std::string& column);
    
    // 新增的工具方法
    std::string extractFormula(const std::string& cell_xml);
    bool isDateFormat(int style_index, const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);
    std::string convertExcelDateToString(double excel_date);
};

} // namespace reader
} // namespace fastexcel

