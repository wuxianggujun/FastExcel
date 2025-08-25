#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#include "WorksheetParser.hpp"

#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/utils/CommonUtils.hpp"

#include <algorithm>
#include "fastexcel/utils/TimeUtils.hpp"
#include <cctype>
#include <ctime>
#include <iostream>
#include <sstream>
#include <string>

namespace fastexcel {
namespace reader {

bool WorksheetParser::parse(const std::string& xml_content, 
                           core::Worksheet* worksheet,
                           const std::unordered_map<int, std::string>& shared_strings,
                           const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                           const std::unordered_map<int, int>& style_id_mapping) {
    if (xml_content.empty() || !worksheet) {
        return false;
    }
    
    try {
        // 先解析列样式定义
        FASTEXCEL_LOG_DEBUG("开始解析列样式定义");
        parseColumns(xml_content, worksheet, styles, style_id_mapping);
        FASTEXCEL_LOG_DEBUG("列样式解析完成");

        // 解析合并单元格（否则编辑保存后会丢失）
        FASTEXCEL_LOG_DEBUG("开始解析合并单元格");
        parseMergeCells(xml_content, worksheet);
        FASTEXCEL_LOG_DEBUG("合并单元格解析完成");
        
        // 解析共享公式（在解析单元格数据之前）
        FASTEXCEL_LOG_DEBUG("开始解析共享公式");
        parseSharedFormulas(xml_content, worksheet);
        FASTEXCEL_LOG_DEBUG("共享公式解析完成");
        
        // 解析工作表数据（行/单元格），并在行级别读取行高
        return parseSheetData(xml_content, worksheet, shared_strings, styles, style_id_mapping);
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析工作表时发生错误: {}", e.what());
        return false;
    }
}

bool WorksheetParser::parseSheetData(const std::string& xml_content, 
                                    core::Worksheet* worksheet,
                                    const std::unordered_map<int, std::string>& shared_strings,
                                    const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                                    const std::unordered_map<int, int>& style_id_mapping) {
    // 查找 <sheetData> 标签
    size_t sheet_data_start = xml_content.find("<sheetData");
    if (sheet_data_start == std::string::npos) {
        // 没有数据，这是正常的（空工作表）
        return true;
    }
    
    // 找到 <sheetData> 标签的结束位置
    size_t content_start = xml_content.find(">", sheet_data_start);
    if (content_start == std::string::npos) {
        return false;
    }
    content_start++;
    
    // 检查是否是自闭合标签
    if (xml_content.substr(sheet_data_start, content_start - sheet_data_start).find("/>") != std::string::npos) {
        // 自闭合标签，没有数据
        return true;
    }
    
    // 找到 </sheetData> 标签
    size_t sheet_data_end = xml_content.find("</sheetData>", content_start);
    if (sheet_data_end == std::string::npos) {
        return false;
    }
    
    // 提取 sheetData 内容
    std::string sheet_data_content = xml_content.substr(content_start, sheet_data_end - content_start);
    
    // 解析所有行
    size_t pos = 0;
    while ((pos = sheet_data_content.find("<row ", pos)) != std::string::npos) {
        // 找到行的结束位置
        size_t row_end = sheet_data_content.find("</row>", pos);
        if (row_end == std::string::npos) {
            // 尝试查找自闭合行标签
            size_t self_close = sheet_data_content.find("/>", pos);
            if (self_close != std::string::npos && self_close < sheet_data_content.find("<row ", pos + 1)) {
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        // 提取行内容
        std::string row_xml = sheet_data_content.substr(pos, row_end - pos + 6); // 包含 </row>
        
        // 解析行
        if (!parseRow(row_xml, worksheet, shared_strings, styles, style_id_mapping)) {
            FASTEXCEL_LOG_WARN("解析行失败");
            // 继续处理其他行
        }
        
        pos = row_end + 6; // 跳过 </row>
    }
    
    return true;
}

bool WorksheetParser::parseRow(const std::string& row_xml, 
                              core::Worksheet* worksheet,
                              const std::unordered_map<int, std::string>& shared_strings,
                              const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                              const std::unordered_map<int, int>& style_id_mapping) {
    // 先解析行级属性：r（行号）、ht（行高）、customHeight、hidden
    int excel_row = extractIntAttribute(row_xml, "r");
    if (excel_row > 0) {
        int row_index = excel_row - 1; // 转为0基
        double ht = extractDoubleAttribute(row_xml, "ht");
        std::string custom_height = extractStringAttribute(row_xml, "customHeight");
        std::string hidden = extractStringAttribute(row_xml, "hidden");
        if (ht > 0 && (custom_height == "1" || custom_height == "true" || !custom_height.empty())) {
            // 只有明确设置自定义行高时才设置；部分文件也会直接提供ht且customHeight缺省，兼容性处理：只要有ht就设置
            worksheet->setRowHeight(row_index, ht);
        } else if (ht > 0) {
            worksheet->setRowHeight(row_index, ht);
        }
        if (hidden == "1" || hidden == "true") {
            worksheet->hideRow(row_index);
        }
    }

    // 解析行中的所有单元格
    size_t pos = 0;
    while ((pos = row_xml.find("<c ", pos)) != std::string::npos) {
        // 首先检查这个单元格是否是自闭合的
        size_t tag_end = row_xml.find(">", pos);
        if (tag_end == std::string::npos) {
            break;
        }
        
        // 检查标签是否以 "/>" 结尾（自闭合标签）
        if (tag_end > 0 && row_xml[tag_end - 1] == '/') {
            // 自闭合标签处理
            std::string cell_xml = row_xml.substr(pos, tag_end - pos + 1);
            parseCell(cell_xml, worksheet, shared_strings, styles, style_id_mapping);
            pos = tag_end + 1;
            continue;
        }
        
        // 非自闭合标签：寻找对应的 </c> 结束标签
        size_t cell_end = row_xml.find("</c>", tag_end);
        if (cell_end == std::string::npos) {
            break;
        }
        
        // 提取单元格内容
        std::string cell_xml = row_xml.substr(pos, cell_end - pos + 4); // 包含 </c>
        
        // 解析单元格
        if (!parseCell(cell_xml, worksheet, shared_strings, styles, style_id_mapping)) {
            FASTEXCEL_LOG_WARN("解析单元格失败: {}", cell_xml);
            // 继续处理其他单元格
        }
        
        pos = cell_end + 4; // 跳过 </c>
    }
    
    return true;
}

bool WorksheetParser::parseCell(const std::string& cell_xml,
                               core::Worksheet* worksheet,
                               const std::unordered_map<int, std::string>& shared_strings,
                               const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                               const std::unordered_map<int, int>& style_id_mapping) {
    // 提取单元格引用 (r="A1")
    size_t r_start = cell_xml.find("r=\"");
    if (r_start == std::string::npos) {
        return false;
    }
    r_start += 3; // 跳过 r="
    
    size_t r_end = cell_xml.find("\"", r_start);
    if (r_end == std::string::npos) {
        return false;
    }
    
    std::string cell_ref = cell_xml.substr(r_start, r_end - r_start);
    auto [row, col] = parseCellReference(cell_ref);
    
    if (row < 0 || col < 0) {
        return false;
    }
    
    // 提取单元格类型
    std::string cell_type = extractCellType(cell_xml);
    
    // 提取单元格值
    std::string cell_value = extractCellValue(cell_xml);
    
    // 提取公式（如果有）
    std::string formula = extractFormula(cell_xml);
    
    // 提取样式索引
    int style_index = extractStyleIndex(cell_xml);
    
    // 根据类型设置单元格值
    if (cell_type == "s") {
        // 共享字符串
        try {
            int string_index = std::stoi(cell_value);
            auto it = shared_strings.find(string_index);
            if (it != shared_strings.end()) {
                // 使用 addSharedStringWithIndex 保持原始索引
                if (auto wb = worksheet->getParentWorkbook()) {
                    wb->addSharedStringWithIndex(it->second, string_index);
                }
                worksheet->setValue(row, col, it->second);
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("解析共享字符串索引失败: {}", e.what());
        }
    } else if (cell_type == "inlineStr") {
        // 内联字符串
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->setValue(row, col, decoded_value);
    } else if (cell_type == "b") {
        // 布尔值
        bool bool_value = (cell_value == "1" || cell_value == "true");
        worksheet->setValue(row, col, bool_value);
    } else if (cell_type == "str") {
        // 公式字符串结果
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->setValue(row, col, decoded_value);
    } else if (cell_type == "e") {
        // 错误值
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->setValue(row, col, "#ERROR: " + decoded_value);
    } else if (cell_type == "d") {
        // 日期值（ISO 8601格式）
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->setValue(row, col, decoded_value);
    } else {
        // 数字或默认类型 - Excel中没有t属性的单元格默认是数字！
        if (!cell_value.empty()) {
            try {
                // 解析为数字 - 这是最常见的情况，应该优先处理
                double number_value = std::stod(cell_value);
                
                // 重要：日期在Excel中也是数字，不要转换为字符串！
                // 日期格式应该由样式来控制显示，而不是改变数据类型
                worksheet->setValue(row, col, number_value);
                
            } catch (const std::exception& /*e*/) {
                // 极少数情况：如果真的不能解析为数字（可能是Excel文件损坏）
                // 记录警告并跳过该单元格，而不是错误地转换为字符串
                FASTEXCEL_LOG_WARN("单元格{}无法解析为数字，值='{}'", 
                           utils::CommonUtils::cellReference(row, col), cell_value);
                // 不设置任何值，保持单元格为空
            }
        } else if (!formula.empty()) {
            // 只有公式没有值的情况
            worksheet->setFormula(row, col, formula);
        }
    }
    
    // 应用样式（如果有）
    if (style_index >= 0) {
        // 使用样式 ID 映射来获取正确的 FormatRepository 中的样式
        int mapped_style_id = style_index;
        if (!style_id_mapping.empty()) {
            auto mapping_it = style_id_mapping.find(style_index);
            if (mapping_it != style_id_mapping.end()) {
                mapped_style_id = mapping_it->second;
            }
        }
        
        auto style_it = styles.find(mapped_style_id);
        if (style_it != styles.end()) {
            auto& cell = worksheet->getCell(row, col);
            cell.setFormat(style_it->second);
        }
    }
    
    return true;
}

std::pair<int, int> WorksheetParser::parseCellReference(const std::string& ref) {
    if (ref.empty()) {
        return {-1, -1};
    }
    
    // 分离列字母和行号
    size_t i = 0;
    std::string column;
    std::string row_str;
    
    // 提取列字母
    while (i < ref.length() && std::isalpha(ref[i])) {
        column += ref[i];
        i++;
    }
    
    // 提取行号
    while (i < ref.length() && std::isdigit(ref[i])) {
        row_str += ref[i];
        i++;
    }
    
    if (column.empty() || row_str.empty()) {
        return {-1, -1};
    }
    
    try {
        int row = std::stoi(row_str) - 1; // Excel行号从1开始，转换为0开始
        int col = columnLetterToNumber(column);
        return {row, col};
    } catch (const std::exception& /*e*/) {
        return {-1, -1};
    }
}

std::string WorksheetParser::extractCellValue(const std::string& cell_xml) {
    // 查找 <v> 标签（值）
    size_t v_start = cell_xml.find("<v>");
    if (v_start == std::string::npos) {
        // 查找内联字符串 <is><t>
        size_t is_start = cell_xml.find("<is>");
        if (is_start != std::string::npos) {
            size_t t_start = cell_xml.find("<t>", is_start);
            if (t_start != std::string::npos) {
                t_start += 3; // 跳过 <t>
                size_t t_end = cell_xml.find("</t>", t_start);
                if (t_end != std::string::npos) {
                    return cell_xml.substr(t_start, t_end - t_start);
                }
            }
        }
        return "";
    }
    
    v_start += 3; // 跳过 <v>
    size_t v_end = cell_xml.find("</v>", v_start);
    if (v_end == std::string::npos) {
        return "";
    }
    
    return cell_xml.substr(v_start, v_end - v_start);
}

std::string WorksheetParser::extractCellType(const std::string& cell_xml) {
    size_t t_start = cell_xml.find("t=\"");
    if (t_start == std::string::npos) {
        return ""; // 默认类型（数字）
    }
    
    t_start += 3; // 跳过 t="
    size_t t_end = cell_xml.find("\"", t_start);
    if (t_end == std::string::npos) {
        return "";
    }
    
    return cell_xml.substr(t_start, t_end - t_start);
}

int WorksheetParser::extractStyleIndex(const std::string& cell_xml) {
    size_t s_start = cell_xml.find("s=\"");
    if (s_start == std::string::npos) {
        return -1; // 没有样式
    }
    
    s_start += 3; // 跳过 s="
    size_t s_end = cell_xml.find("\"", s_start);
    if (s_end == std::string::npos) {
        return -1;
    }
    
    try {
        return std::stoi(cell_xml.substr(s_start, s_end - s_start));
    } catch (const std::exception& /*e*/) {
        return -1;
    }
}

std::string WorksheetParser::decodeXMLEntities(const std::string& text) {
    std::string result = text;
    
    // 替换常见的XML实体
    size_t pos = 0;
    
    // &lt; -> <
    while ((pos = result.find("&lt;", pos)) != std::string::npos) {
        result.replace(pos, 4, "<");
        pos += 1;
    }
    
    pos = 0;
    // &gt; -> >
    while ((pos = result.find("&gt;", pos)) != std::string::npos) {
        result.replace(pos, 4, ">");
        pos += 1;
    }
    
    pos = 0;
    // &amp; -> &
    while ((pos = result.find("&amp;", pos)) != std::string::npos) {
        result.replace(pos, 5, "&");
        pos += 1;
    }
    
    pos = 0;
    // &quot; -> "
    while ((pos = result.find("&quot;", pos)) != std::string::npos) {
        result.replace(pos, 6, "\"");
        pos += 1;
    }
    
    pos = 0;
    // &apos; -> '
    while ((pos = result.find("&apos;", pos)) != std::string::npos) {
        result.replace(pos, 6, "'");
        pos += 1;
    }
    
    return result;
}

int WorksheetParser::columnLetterToNumber(const std::string& column) {
    int result = 0;
    for (char c : column) {
        result = result * 26 + (std::toupper(c) - 'A' + 1);
    }
    return result - 1; // 转换为0开始的索引
}

// 新增的辅助方法
std::string WorksheetParser::extractFormula(const std::string& cell_xml) {
    size_t f_start = cell_xml.find("<f>");
    if (f_start == std::string::npos) {
        return "";
    }
    
    f_start += 3; // 跳过 <f>
    size_t f_end = cell_xml.find("</f>", f_start);
    if (f_end == std::string::npos) {
        return "";
    }
    
    return decodeXMLEntities(cell_xml.substr(f_start, f_end - f_start));
}

bool WorksheetParser::isDateFormat(int style_index,
                                  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles) {
    if (style_index < 0) {
        return false;
    }
    
    auto it = styles.find(style_index);
    if (it == styles.end()) {
        return false;
    }
    
    // 检查格式是否为日期格式
    // 这里可以根据Format对象的属性来判断
    // 简化实现：检查内置的日期格式ID
    return (style_index >= 14 && style_index <= 22) || // 常见日期格式
           (style_index >= 176 && style_index <= 180);  // 更多日期格式
}

std::string WorksheetParser::convertExcelDateToString(double excel_date) {
    // Excel日期从1900年1月1日开始计算
    // 注意：Excel错误地认为1900年是闰年，所以需要调整
    const int EXCEL_EPOCH_OFFSET = 25569; // 1900-01-01到1970-01-01的天数差
    const int SECONDS_PER_DAY = 86400;
    
    // 调整Excel的1900年闰年错误
    if (excel_date >= 60) {
        excel_date -= 1;
    }
    
    // 转换为Unix时间戳
    time_t unix_time = static_cast<time_t>((excel_date - EXCEL_EPOCH_OFFSET) * SECONDS_PER_DAY);
    
    // 使用项目时间工具类格式化
    std::tm time_info_buf{};
#ifdef _WIN32
    gmtime_s(&time_info_buf, &unix_time);
#else
    gmtime_r(&unix_time, &time_info_buf);
#endif
    return fastexcel::utils::TimeUtils::formatTime(time_info_buf, "%Y-%m-%d");
}

bool WorksheetParser::parseColumns(const std::string& xml_content,
                                  core::Worksheet* worksheet,
                                  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                                  const std::unordered_map<int, int>& style_id_mapping) {
    FASTEXCEL_LOG_DEBUG("parseColumns被调用，xml_content长度: {}", xml_content.length());
    
    // 查找 <cols> 标签
    size_t cols_start = xml_content.find("<cols");
    if (cols_start == std::string::npos) {
        // 没有列定义，这是正常的
        FASTEXCEL_LOG_DEBUG("没有找到<cols>标签");
        return true;
    }
    
    FASTEXCEL_LOG_DEBUG("找到<cols>标签在位置: {}", cols_start);
    
    // 找到 <cols> 标签的结束位置
    size_t content_start = xml_content.find(">", cols_start);
    if (content_start == std::string::npos) {
        return false;
    }
    content_start++;
    
    // 检查是否是自闭合标签
    if (xml_content.substr(cols_start, content_start - cols_start).find("/>") != std::string::npos) {
        // 自闭合标签，没有列定义
        return true;
    }
    
    // 找到 </cols> 标签
    size_t cols_end = xml_content.find("</cols>", content_start);
    if (cols_end == std::string::npos) {
        return false;
    }
    
    // 提取 cols 内容
    std::string cols_content = xml_content.substr(content_start, cols_end - content_start);
    
    // 解析所有列定义
    size_t pos = 0;
    while ((pos = cols_content.find("<col ", pos)) != std::string::npos) {
        // 找到列定义的结束位置
        size_t col_end = cols_content.find("/>", pos);
        if (col_end == std::string::npos) {
            // 查找完整的col标签
            col_end = cols_content.find("</col>", pos);
            if (col_end == std::string::npos) {
                break;
            }
            col_end += 6;
        } else {
            col_end += 2;
        }
        
        // 提取列定义内容
        std::string col_xml = cols_content.substr(pos, col_end - pos);
        
        // 解析列属性
        int min_col = extractIntAttribute(col_xml, "min");
        int max_col = extractIntAttribute(col_xml, "max");
        double width = extractDoubleAttribute(col_xml, "width");
        int style_index = extractIntAttribute(col_xml, "style");
        bool custom_width = extractStringAttribute(col_xml, "customWidth") == "1";
        bool hidden = extractStringAttribute(col_xml, "hidden") == "1";
        
        if (min_col > 0 && max_col > 0) {
            // Excel列索引从1开始，转换为0开始
            int first_col = min_col - 1;
            int last_col = max_col - 1;
            
            // 设置列宽（保留Excel默认列宽，即使没有customWidth属性）
            if (width > 0) {
                // 使用新的智能列宽设置，逐列处理
                for (int col = first_col; col <= last_col; ++col) {
                    worksheet->setColumnWidth(col, width);
                }
                FASTEXCEL_LOG_DEBUG("设置列宽：列 {}-{} 宽度 {} custom_width={}", first_col, last_col, width, custom_width);
            }
            
            // 设置列样式
            if (style_index >= 0) {
                // 使用样式 ID 映射来获取正确的格式
                int mapped_style_id = style_index;
                if (!style_id_mapping.empty()) {
                    auto mapping_it = style_id_mapping.find(style_index);
                    if (mapping_it != style_id_mapping.end()) {
                        mapped_style_id = mapping_it->second;
                    }
                }
                
                auto style_it = styles.find(mapped_style_id);
                if (style_it != styles.end()) {
                    // 设置列格式 ID 到工作表
                    worksheet->setColumnFormatId(first_col, last_col, mapped_style_id);
                    FASTEXCEL_LOG_DEBUG("设置列样式：列 {}-{} 原始样式ID {} 映射样式ID {}", first_col, last_col, style_index, mapped_style_id);
                }
            }
            
            // 设置隐藏状态
            if (hidden) {
                worksheet->hideColumn(first_col, last_col);
            }
        }
        
        pos = col_end;
    }
    
    return true;
}

// 辅助方法：提取整数属性
int WorksheetParser::extractIntAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return -1;
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return -1;
    }
    
    try {
        return std::stoi(xml.substr(start, end - start));
    } catch (const std::exception& /*e*/) {
        return -1;
    }
}

// 辅助方法：提取双精度浮点数属性
double WorksheetParser::extractDoubleAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return -1.0;
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return -1.0;
    }
    
    try {
        return std::stod(xml.substr(start, end - start));
    } catch (const std::exception& /*e*/) {
        return -1.0;
    }
}

// 辅助方法：提取字符串属性
std::string WorksheetParser::extractStringAttribute(const std::string& xml, const std::string& attr_name) {
    std::string pattern = attr_name + "=\"";
    size_t start = xml.find(pattern);
    if (start == std::string::npos) {
        return "";
    }
    
    start += pattern.length();
    size_t end = xml.find("\"", start);
    if (end == std::string::npos) {
        return "";
    }
    
    return xml.substr(start, end - start);
}

// 解析合并单元格
bool WorksheetParser::parseMergeCells(const std::string& xml_content,
                                     core::Worksheet* worksheet) {
    size_t merges_start = xml_content.find("<mergeCells");
    if (merges_start == std::string::npos) {
        return true; // 没有合并单元格，正常
    }

    // 找到开始标签结束符
    size_t merges_tag_end = xml_content.find('>', merges_start);
    if (merges_tag_end == std::string::npos) {
        return false;
    }

    // 查找闭合标签位置
    size_t merges_end = xml_content.find("</mergeCells>", merges_tag_end);
    if (merges_end == std::string::npos) {
        return false;
    }

    std::string merges_content = xml_content.substr(merges_tag_end + 1, merges_end - (merges_tag_end + 1));

    size_t pos = 0;
    while ((pos = merges_content.find("<mergeCell", pos)) != std::string::npos) {
        size_t tag_end = merges_content.find("/>", pos);
        if (tag_end == std::string::npos) {
            tag_end = merges_content.find('>', pos);
            if (tag_end == std::string::npos) break;
        }
        std::string mc_xml = merges_content.substr(pos, tag_end - pos + 2);
        std::string ref = extractStringAttribute(mc_xml, "ref");
        if (!ref.empty()) {
            int r1, c1, r2, c2;
            if (parseRangeRef(ref, r1, c1, r2, c2)) {
                worksheet->mergeCells(r1, c1, r2, c2);
            }
        }
        pos = tag_end + 2;
    }

    return true;
}

// 解析区间引用 A1:C3 -> (r1,c1,r2,c2)
bool WorksheetParser::parseRangeRef(const std::string& ref, int& first_row, int& first_col, int& last_row, int& last_col) {
    size_t colon = ref.find(':');
    if (colon == std::string::npos) {
        // 单一单元格区间
        auto [r, c] = parseCellReference(ref);
        if (r < 0 || c < 0) return false;
        first_row = last_row = r;
        first_col = last_col = c;
        return true;
    }
    std::string start_ref = ref.substr(0, colon);
    std::string end_ref = ref.substr(colon + 1);
    auto [r1, c1] = parseCellReference(start_ref);
    auto [r2, c2] = parseCellReference(end_ref);
    if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return false;
    first_row = r1; first_col = c1; last_row = r2; last_col = c2;
    return true;
}

// 解析共享公式
void WorksheetParser::parseSharedFormulas(const std::string& xml_content, core::Worksheet* worksheet) {
    if (!worksheet) {
        FASTEXCEL_LOG_ERROR("Worksheet is null in parseSharedFormulas");
        return;
    }
    
    FASTEXCEL_LOG_DEBUG("正在解析共享公式...");
    
    // 存储共享公式主定义（si -> {formula, range}）
    std::unordered_map<int, std::pair<std::string, std::string>> shared_formulas;
    
    // 第一轮：查找所有共享公式主定义（包含 ref 属性的）
    size_t pos = 0;
    while ((pos = xml_content.find("<f t=\"shared\"", pos)) != std::string::npos) {
        size_t f_end = xml_content.find("</f>", pos);
        size_t f_self_close = xml_content.find("/>", pos);
        
        size_t actual_end = std::string::npos;
        if (f_end != std::string::npos && (f_self_close == std::string::npos || f_end < f_self_close)) {
            actual_end = f_end + 4; // 包含 </f>
        } else if (f_self_close != std::string::npos) {
            actual_end = f_self_close + 2; // 包含 />
        }
        
        if (actual_end == std::string::npos) {
            pos++;
            continue;
        }
        
        std::string f_tag = xml_content.substr(pos, actual_end - pos);
        
        int si;
        std::string ref, formula;
        if (extractSharedFormulaInfo(f_tag, si, ref, formula)) {
            if (!ref.empty() && !formula.empty()) {
                // 这是主公式定义
                shared_formulas[si] = {formula, ref};
                FASTEXCEL_LOG_DEBUG("发现共享公式主定义: si={}, ref={}, formula={}", si, ref, formula);
            }
        }
        
        pos = actual_end;
    }
    
    FASTEXCEL_LOG_DEBUG("找到 {} 个共享公式主定义", shared_formulas.size());
    
    // 为每个共享公式创建 SharedFormulaManager 中的条目
    for (const auto& [si, formula_info] : shared_formulas) {
        const auto& [formula, ref] = formula_info;
        
        // 解析范围
        int first_row, first_col, last_row, last_col;
        if (parseRangeRef(ref, first_row, first_col, last_row, last_col)) {
            // 使用 worksheet 的 createSharedFormula 方法
            int created_si = worksheet->createSharedFormula(first_row, first_col, last_row, last_col, formula);
            if (created_si >= 0) {
                FASTEXCEL_LOG_DEBUG("成功创建共享公式: si={}, 范围={}:{}-{}:{}", 
                         created_si, first_row, first_col, last_row, last_col);
            } else {
                FASTEXCEL_LOG_ERROR("创建共享公式失败: si={}, 范围={}", si, ref);
            }
        } else {
            FASTEXCEL_LOG_ERROR("无法解析共享公式范围: {}", ref);
        }
    }
}

bool WorksheetParser::extractSharedFormulaInfo(const std::string& f_tag, int& si, std::string& ref, std::string& formula) {
    // 提取 si 属性
    size_t si_pos = f_tag.find("si=\"");
    if (si_pos == std::string::npos) return false;
    
    si_pos += 4; // 跳过 si="
    size_t si_end = f_tag.find("\"", si_pos);
    if (si_end == std::string::npos) return false;
    
    try {
        si = std::stoi(f_tag.substr(si_pos, si_end - si_pos));
    } catch (const std::invalid_argument& e) {
        FASTEXCEL_LOG_WARN("Invalid shared formula index format in f tag: {}", e.what());
        return false;
    } catch (const std::out_of_range& e) {
        FASTEXCEL_LOG_WARN("Shared formula index out of range in f tag: {}", e.what());
        return false;
    }
    
    // 提取 ref 属性（可选，只有主公式才有）
    size_t ref_pos = f_tag.find("ref=\"");
    if (ref_pos != std::string::npos) {
        ref_pos += 5; // 跳过 ref="
        size_t ref_end = f_tag.find("\"", ref_pos);
        if (ref_end != std::string::npos) {
            ref = f_tag.substr(ref_pos, ref_end - ref_pos);
        }
    }
    
    // 提取公式内容
    size_t content_start = f_tag.find(">") + 1;
    size_t content_end = f_tag.rfind("</f>");
    
    if (content_end == std::string::npos) {
        // 自闭合标签，没有公式内容
        formula = "";
    } else if (content_start < content_end) {
        formula = f_tag.substr(content_start, content_end - content_start);
        // 解码 XML 实体
        formula = decodeXMLEntities(formula);
    }
    
    return true;
}

} // namespace reader
} // namespace fastexcel
