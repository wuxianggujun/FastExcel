//
// Created by wuxianggujun on 25-8-4.
//

#include "WorksheetParser.hpp"
#include <iostream>
#include <sstream>
#include <algorithm>
#include <cctype>

namespace fastexcel {
namespace reader {

bool WorksheetParser::parse(const std::string& xml_content, 
                           core::Worksheet* worksheet,
                           const std::unordered_map<int, std::string>& shared_strings,
                           const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
    if (xml_content.empty() || !worksheet) {
        return false;
    }
    
    try {
        // 解析工作表数据
        return parseSheetData(xml_content, worksheet, shared_strings, styles);
        
    } catch (const std::exception& e) {
        std::cerr << "解析工作表时发生错误: " << e.what() << std::endl;
        return false;
    }
}

bool WorksheetParser::parseSheetData(const std::string& xml_content, 
                                    core::Worksheet* worksheet,
                                    const std::unordered_map<int, std::string>& shared_strings,
                                    const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
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
        if (!parseRow(row_xml, worksheet, shared_strings, styles)) {
            std::cerr << "解析行失败" << std::endl;
            // 继续处理其他行
        }
        
        pos = row_end + 6; // 跳过 </row>
    }
    
    return true;
}

bool WorksheetParser::parseRow(const std::string& row_xml, 
                              core::Worksheet* worksheet,
                              const std::unordered_map<int, std::string>& shared_strings,
                              const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
    // 解析行中的所有单元格
    size_t pos = 0;
    while ((pos = row_xml.find("<c ", pos)) != std::string::npos) {
        // 找到单元格的结束位置
        size_t cell_end = row_xml.find("</c>", pos);
        if (cell_end == std::string::npos) {
            // 尝试查找自闭合单元格标签
            size_t self_close = row_xml.find("/>", pos);
            if (self_close != std::string::npos && self_close < row_xml.find("<c ", pos + 1)) {
                std::string cell_xml = row_xml.substr(pos, self_close - pos + 2);
                parseCell(cell_xml, worksheet, shared_strings, styles);
                pos = self_close + 2;
                continue;
            }
            break;
        }
        
        // 提取单元格内容
        std::string cell_xml = row_xml.substr(pos, cell_end - pos + 4); // 包含 </c>
        
        // 解析单元格
        if (!parseCell(cell_xml, worksheet, shared_strings, styles)) {
            std::cerr << "解析单元格失败" << std::endl;
            // 继续处理其他单元格
        }
        
        pos = cell_end + 4; // 跳过 </c>
    }
    
    return true;
}

bool WorksheetParser::parseCell(const std::string& cell_xml, 
                               core::Worksheet* worksheet,
                               const std::unordered_map<int, std::string>& shared_strings,
                               const std::unordered_map<int, std::shared_ptr<core::Format>>& styles) {
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
    
    // 提取样式索引
    int style_index = extractStyleIndex(cell_xml);
    
    // 根据类型设置单元格值
    if (cell_type == "s") {
        // 共享字符串
        try {
            int string_index = std::stoi(cell_value);
            auto it = shared_strings.find(string_index);
            if (it != shared_strings.end()) {
                worksheet->writeString(row, col, it->second);
            }
        } catch (const std::exception& e) {
            std::cerr << "解析共享字符串索引失败: " << e.what() << std::endl;
        }
    } else if (cell_type == "inlineStr") {
        // 内联字符串
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->writeString(row, col, decoded_value);
    } else if (cell_type == "b") {
        // 布尔值
        bool bool_value = (cell_value == "1" || cell_value == "true");
        worksheet->writeBoolean(row, col, bool_value);
    } else if (cell_type == "str") {
        // 公式字符串结果
        std::string decoded_value = decodeXMLEntities(cell_value);
        worksheet->writeString(row, col, decoded_value);
    } else {
        // 数字或默认类型
        if (!cell_value.empty()) {
            try {
                double number_value = std::stod(cell_value);
                worksheet->writeNumber(row, col, number_value);
            } catch (const std::exception& e) {
                // 如果不能转换为数字，当作字符串处理
                std::string decoded_value = decodeXMLEntities(cell_value);
                worksheet->writeString(row, col, decoded_value);
            }
        }
    }
    
    // 应用样式（如果有）
    if (style_index >= 0) {
        auto style_it = styles.find(style_index);
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
    } catch (const std::exception& e) {
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
    } catch (const std::exception& e) {
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

} // namespace reader
} // namespace fastexcel
