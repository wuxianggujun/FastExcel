#include "fastexcel/core/CSVProcessor.hpp"
#include <fstream>
#include <sstream>
#include <fmt/format.h>
#include <regex>
#include <algorithm>
#include <locale>
#include <codecvt>
#include <iomanip>
#include <climits>
#include <unordered_map>

namespace fastexcel {
namespace core {

// 简单的CSV文件读取函数
std::vector<std::vector<std::string>> readCSVFromFile(const std::string& filepath, const CSVOptions& options) {
    std::vector<std::vector<std::string>> result;
    
    std::ifstream file(filepath);
    if (!file.is_open()) {
        return result;
    }
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    // 读取文件内容
    std::string content((std::istreambuf_iterator<char>(file)),
                       std::istreambuf_iterator<char>());
    file.close();
    
    return processor.parseString(content);
}

// 简单的CSV文件写入函数
bool writeCSVToFile(const std::string& filepath, const std::vector<std::vector<std::string>>& data, const CSVOptions& options) {
    std::ofstream file(filepath);
    if (!file.is_open()) {
        return false;
    }
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    for (size_t i = 0; i < data.size(); ++i) {
        if (i > 0) {
            file << "\n";
        }
        file << processor.formatRow(data[i]);
    }
    
    file.close();
    return true;
}

// CSV格式检测函数
char detectDelimiter(const std::string& sample) {
    // 统计各种分隔符出现的频率
    std::unordered_map<char, int> counts;
    const std::vector<char> candidates = {',', ';', '\t', '|'};
    
    for (char delimiter : candidates) {
        counts[delimiter] = static_cast<int>(std::count(sample.begin(), sample.end(), delimiter));
    }
    
    // 找出频率最高的分隔符
    char best_delimiter = ',';
    int max_count = 0;
    
    for (const auto& [delimiter, count] : counts) {
        if (count > max_count) {
            max_count = count;
            best_delimiter = delimiter;
        }
    }
    
    return best_delimiter;
}

// 编码检测函数
std::string detectEncoding(const std::vector<uint8_t>& data) {
    if (data.size() >= 3) {
        // 检测UTF-8 BOM
        if (data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
            return "UTF-8";
        }
    }
    
    if (data.size() >= 2) {
        // 检测UTF-16 BOM
        if ((data[0] == 0xFF && data[1] == 0xFE) || 
            (data[0] == 0xFE && data[1] == 0xFF)) {
            return "UTF-16";
        }
    }
    
    // 默认假设UTF-8
    return "UTF-8";
}

// 解析单行CSV
std::vector<std::string> parseLine(const std::string& line, char delimiter, char quote_char, char escape_char) {
    std::vector<std::string> result;
    
    if (line.empty()) {
        return result;
    }
    
    std::string current_field;
    bool in_quotes = false;
    bool escape_next = false;
    
    for (size_t i = 0; i < line.length(); ++i) {
        char c = line[i];
        
        if (escape_next) {
            current_field += c;
            escape_next = false;
        } else if (c == escape_char && in_quotes) {
            if (i + 1 < line.length() && line[i + 1] == quote_char) {
                escape_next = true;
            } else {
                current_field += c;
            }
        } else if (c == quote_char) {
            in_quotes = !in_quotes;
        } else if (c == delimiter && !in_quotes) {
            result.push_back(current_field);
            current_field.clear();
        } else {
            current_field += c;
        }
    }
    
    result.push_back(current_field);
    return result;
}

// 解析内容到CSV数据
CSVParseInfo parseContent(const std::string& content, const CSVOptions& options) {
    CSVParseInfo info;
    std::vector<std::vector<std::string>> data;
    
    if (content.empty()) {
        info.success = false;
        info.error_message = "Empty content";
        return info;
    }
    
    // 按行分割
    std::istringstream stream(content);
    std::string line;
    int row_count = 0;
    
    while (std::getline(stream, line)) {
        if (options.skip_empty_lines && line.empty()) {
            continue;
        }
        
        std::vector<std::string> row = parseLine(line, options.delimiter, options.quote_char, options.escape_char);
        
        // 去除空白字符（如果启用）
        if (options.trim_whitespace) {
            for (auto& field : row) {
                // 简单的trim实现
                size_t start = field.find_first_not_of(" \t\r\n");
                if (start != std::string::npos) {
                    size_t end = field.find_last_not_of(" \t\r\n");
                    field = field.substr(start, end - start + 1);
                } else {
                    field.clear();
                }
            }
        }
        
        // 数字和日期解析（如果启用）
        if (options.parse_numbers || options.parse_dates) {
            for (auto& field : row) {
                if (options.parse_numbers && !field.empty()) {
                    // 尝试解析为数字（简单实现）
                    try {
                        if (field.find('.') != std::string::npos) {
                            std::stod(field); // 测试是否为有效数字
                        } else {
                            std::stoi(field); // 测试是否为有效整数
                        }
                    } catch (...) {
                        // 不是数字，保持原样
                    }
                }
            }
        }
        
        data.push_back(row);
        row_count++;
        
        if (row.size() > static_cast<size_t>(info.columns_detected)) {
            info.columns_detected = static_cast<int>(row.size());
        }
    }
    
    info.success = true;
    info.rows_parsed = row_count;
    
    // 如果有标题行
    if (options.has_header && !data.empty()) {
        info.has_header_row = true;
        info.column_names = data[0];
    }
    
    return info;
}

// 字段转义函数
std::string escapeField(const std::string& field, const CSVOptions& options) {
    bool needs_quotes = field.find(options.delimiter) != std::string::npos ||
                       field.find(options.quote_char) != std::string::npos ||
                       field.find('\n') != std::string::npos ||
                       field.find('\r') != std::string::npos;
    
    if (needs_quotes) {
        std::string escaped = std::string(1, options.quote_char);
        for (char c : field) {
            if (c == options.quote_char) {
                escaped += std::string(1, options.escape_char) + std::string(1, options.quote_char);
            } else {
                escaped += c;
            }
        }
        escaped += std::string(1, options.quote_char);
        return escaped;
    }
    
    return field;
}

// 检查字段是否需要引号
bool needsQuoting(const std::string& field, const CSVOptions& options) {
    return field.find(options.delimiter) != std::string::npos ||
           field.find(options.quote_char) != std::string::npos ||
           field.find('\n') != std::string::npos ||
           field.find('\r') != std::string::npos;
}

// CSV工具函数
bool isCSVFile(const std::string& filepath) {
    size_t dot_pos = filepath.find_last_of('.');
    if (dot_pos == std::string::npos) {
        return false;
    }
    
    std::string ext = filepath.substr(dot_pos + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    return ext == "csv" || ext == "tsv" || ext == "txt";
}

CSVOptions detectCSVOptions(const std::string& filepath) {
    CSVOptions options;
    
    // 读取文件开头几行进行检测
    std::ifstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        return options;
    }
    
    std::string sample;
    std::string line;
    int lines_read = 0;
    const int max_lines = 5; // 只读前5行进行检测
    
    while (std::getline(file, line) && lines_read < max_lines) {
        sample += fmt::format("{}\n", line);
        lines_read++;
    }
    file.close();
    
    if (!sample.empty()) {
        options.delimiter = detectDelimiter(sample);
        
        // 简单的编码检测
        std::vector<uint8_t> data(sample.begin(), sample.end());
        options.encoding = detectEncoding(data);
    }
    
    return options;
}

}} // namespace fastexcel::core
