#include "fastexcel/utils/Logger.hpp"
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
     * @param style_id_mapping 样式ID映射（原始ID -> FormatRepository ID）
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content,
               core::Worksheet* worksheet,
               const std::unordered_map<int, std::string>& shared_strings,
               const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
               const std::unordered_map<int, int>& style_id_mapping = {});

private:
    // 辅助方法
    bool parseSheetData(const std::string& xml_content,
                       core::Worksheet* worksheet,
                       const std::unordered_map<int, std::string>& shared_strings,
                       const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                       const std::unordered_map<int, int>& style_id_mapping);
    
    bool parseColumns(const std::string& xml_content,
                     core::Worksheet* worksheet,
                     const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                     const std::unordered_map<int, int>& style_id_mapping);

    // 解析合并单元格 (<mergeCells><mergeCell ref="A1:C3"/></mergeCells>)
    bool parseMergeCells(const std::string& xml_content,
                         core::Worksheet* worksheet);
    
    bool parseRow(const std::string& row_xml,
                  core::Worksheet* worksheet,
                  const std::unordered_map<int, std::string>& shared_strings,
                  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                  const std::unordered_map<int, int>& style_id_mapping);
    
    bool parseCell(const std::string& cell_xml,
                   core::Worksheet* worksheet,
                   const std::unordered_map<int, std::string>& shared_strings,
                   const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                   const std::unordered_map<int, int>& style_id_mapping);
    
    // 工具方法
    std::pair<int, int> parseCellReference(const std::string& ref);
    std::string extractCellValue(const std::string& cell_xml);
    std::string extractCellType(const std::string& cell_xml);
    int extractStyleIndex(const std::string& cell_xml);
    std::string decodeXMLEntities(const std::string& text);
    int columnLetterToNumber(const std::string& column);
    
    // XML属性提取工具方法
    int extractIntAttribute(const std::string& xml, const std::string& attr_name);
    double extractDoubleAttribute(const std::string& xml, const std::string& attr_name);
    std::string extractStringAttribute(const std::string& xml, const std::string& attr_name);
    
    // 新增的工具方法
    std::string extractFormula(const std::string& cell_xml);
    
    // 解析共享公式
    void parseSharedFormulas(const std::string& xml_content, core::Worksheet* worksheet);
    bool extractSharedFormulaInfo(const std::string& f_tag, int& si, std::string& ref, std::string& formula);
    
    bool isDateFormat(int style_index, const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles);
    std::string convertExcelDateToString(double excel_date);

    // 解析区间引用（例如 A1:C3）
    bool parseRangeRef(const std::string& ref, int& first_row, int& first_col, int& last_row, int& last_col);
};

} // namespace reader
} // namespace fastexcel

