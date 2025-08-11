#pragma once

#include <string>
#include <tuple>
#include <regex>
#include <stdexcept>

namespace fastexcel {
namespace utils {

/**
 * @brief Excel 地址解析工具类
 * 
 * 提供解析 Excel 地址格式的工具方法，支持：
 * - 单个地址：A1, B2, AA100 等
 * - 带工作表：Sheet1!A1, "工作表"!B2 等  
 * - 范围地址：A1:C3, Sheet1!A1:C3 等
 * - 整行整列：A:A, 1:1 等
 */
class AddressParser {
public:
    /**
     * @brief 解析单个地址 "A1" 或 "Sheet1!A1"
     * @param address Excel 地址字符串
     * @return std::tuple<std::string, int, int> (工作表名, 行索引, 列索引)，索引基于0
     * 
     * @example
     * auto [sheet, row, col] = AddressParser::parseAddress("A1");      // ("", 0, 0)
     * auto [sheet, row, col] = AddressParser::parseAddress("B2");      // ("", 1, 1)  
     * auto [sheet, row, col] = AddressParser::parseAddress("Sheet1!C3"); // ("Sheet1", 2, 2)
     */
    static std::tuple<std::string, int, int> parseAddress(const std::string& address) {
        // 正则表达式匹配 [SheetName!]A1 格式
        // 支持工作表名包含空格和中文，用引号括起来的情况
        static const std::regex addr_regex(R"(^(?:([^!]+)!)?([A-Z]+)(\d+)$)");
        std::smatch matches;
        
        if (!std::regex_match(address, matches, addr_regex)) {
            throw std::invalid_argument("Invalid address format: " + address);
        }
        
        std::string sheet_name = matches[1].str();
        std::string col_str = matches[2].str();
        int row = std::stoi(matches[3].str()) - 1; // 转换为0基索引
        
        int col = columnStringToIndex(col_str);
        
        return std::make_tuple(sheet_name, row, col);
    }
    
    /**
     * @brief 解析范围地址 "A1:C3" 或 "Sheet1!A1:C3"
     * @param range Excel 范围字符串
     * @return std::tuple<std::string, int, int, int, int> (工作表名, 开始行, 开始列, 结束行, 结束列)
     * 
     * @example  
     * auto [sheet, sr, sc, er, ec] = AddressParser::parseRange("A1:C3");        // ("", 0, 0, 2, 2)
     * auto [sheet, sr, sc, er, ec] = AddressParser::parseRange("Sheet1!B2:D4"); // ("Sheet1", 1, 1, 3, 3)
     */
    static std::tuple<std::string, int, int, int, int> parseRange(const std::string& range) {
        // 查找冒号分隔符
        auto colon_pos = range.find(':');
        if (colon_pos == std::string::npos) {
            // 单个单元格当作1x1范围
            auto [sheet, row, col] = parseAddress(range);
            return std::make_tuple(sheet, row, col, row, col);
        }
        
        std::string start_addr = range.substr(0, colon_pos);
        std::string end_addr = range.substr(colon_pos + 1);
        
        auto [start_sheet, start_row, start_col] = parseAddress(start_addr);
        auto [end_sheet, end_row, end_col] = parseAddress(end_addr);
        
        // 如果结束地址没有指定工作表，使用开始地址的工作表
        if (end_sheet.empty()) {
            end_sheet = start_sheet;
        }
        
        // 工作表名必须相同（如果都指定了的话）
        if (!start_sheet.empty() && !end_sheet.empty() && start_sheet != end_sheet) {
            throw std::invalid_argument("Range cannot span multiple sheets: " + range);
        }
        
        // 确保起始位置小于等于结束位置
        if (start_row > end_row || start_col > end_col) {
            std::swap(start_row, end_row);
            std::swap(start_col, end_col);
        }
        
        return std::make_tuple(start_sheet, start_row, start_col, end_row, end_col);
    }
    
    /**
     * @brief 将行列索引转换为 Excel 地址
     * @param row 行索引（基于0）
     * @param col 列索引（基于0）
     * @param sheet_name 工作表名称（可选）
     * @return Excel 地址字符串
     * 
     * @example
     * std::string addr = AddressParser::indexToAddress(0, 0);           // "A1"
     * std::string addr = AddressParser::indexToAddress(1, 2);           // "C2"
     * std::string addr = AddressParser::indexToAddress(0, 0, "Sheet1"); // "Sheet1!A1"
     */
    static std::string indexToAddress(int row, int col, const std::string& sheet_name = "") {
        std::string col_str = indexToColumnString(col);
        std::string addr = col_str + std::to_string(row + 1);  // 转换为1基索引
        
        if (!sheet_name.empty()) {
            // 如果工作表名包含特殊字符，需要用引号括起来
            if (needsQuoting(sheet_name)) {
                addr = "'" + sheet_name + "'!" + addr;
            } else {
                addr = sheet_name + "!" + addr;
            }
        }
        
        return addr;
    }
    
    /**
     * @brief 将范围索引转换为 Excel 范围地址
     * @param start_row 开始行索引（基于0）
     * @param start_col 开始列索引（基于0）
     * @param end_row 结束行索引（基于0）
     * @param end_col 结束列索引（基于0）
     * @param sheet_name 工作表名称（可选）
     * @return Excel 范围地址字符串
     */
    static std::string indexToRange(int start_row, int start_col, int end_row, int end_col, 
                                   const std::string& sheet_name = "") {
        std::string start_addr = indexToAddress(start_row, start_col);
        std::string end_addr = indexToAddress(end_row, end_col);
        std::string range = start_addr + ":" + end_addr;
        
        if (!sheet_name.empty()) {
            if (needsQuoting(sheet_name)) {
                range = "'" + sheet_name + "'!" + range;
            } else {
                range = sheet_name + "!" + range;
            }
        }
        
        return range;
    }
    
    /**
     * @brief 验证地址格式是否有效
     * @param address Excel 地址字符串
     * @return 是否有效
     */
    static bool isValidAddress(const std::string& address) noexcept {
        try {
            parseAddress(address);
            return true;
        } catch (...) {
            return false;
        }
    }
    
    /**
     * @brief 验证范围格式是否有效
     * @param range Excel 范围字符串
     * @return 是否有效
     */
    static bool isValidRange(const std::string& range) noexcept {
        try {
            parseRange(range);
            return true;
        } catch (...) {
            return false;
        }
    }

private:
    /**
     * @brief 将列字符串转换为索引 (A->0, B->1, ..., Z->25, AA->26, ...)
     * @param col_str 列字符串（如 "A", "B", "AA"）
     * @return 列索引（基于0）
     */
    static int columnStringToIndex(const std::string& col_str) {
        int result = 0;
        for (char c : col_str) {
            if (c < 'A' || c > 'Z') {
                throw std::invalid_argument("Invalid column string: " + col_str);
            }
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1; // 转换为0基索引
    }
    
    /**
     * @brief 将索引转换为列字符串 (0->A, 1->B, ..., 25->Z, 26->AA, ...)
     * @param index 列索引（基于0）
     * @return 列字符串
     */
    static std::string indexToColumnString(int index) {
        if (index < 0) {
            throw std::invalid_argument("Column index cannot be negative");
        }
        
        std::string result;
        index++; // 转换为1基索引进行计算
        
        while (index > 0) {
            index--; // 调整为0基索引
            result = char('A' + (index % 26)) + result;
            index /= 26;
        }
        
        return result;
    }
    
    /**
     * @brief 判断工作表名是否需要引号括起来
     * @param sheet_name 工作表名
     * @return 是否需要引号
     */
    static bool needsQuoting(const std::string& sheet_name) {
        // 如果包含空格、特殊字符或非ASCII字符，需要引号
        for (char c : sheet_name) {
            if (c == ' ' || c == '!' || c == '\'' || c < 0) { // c < 0 检查非ASCII字符
                return true;
            }
        }
        return false;
    }
};

}} // namespace fastexcel::utils