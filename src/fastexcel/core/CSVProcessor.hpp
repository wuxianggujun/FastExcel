#pragma once

#include <string>
#include <vector>
#include <sstream>

namespace fastexcel {
namespace core {

struct CSVOptions {
    char delimiter = ',';
    char quote_char = '"';
    char escape_char = '"';
    bool has_header = false;
    bool skip_empty_lines = true;
    std::string encoding = "UTF-8";
    std::string line_terminator = "\n";
    bool trim_whitespace = false;
    bool auto_detect_types = false;
    bool parse_numbers = false;
    bool parse_dates = false;
    
    // 为了兼容性，添加standard别名
    static CSVOptions standard() {
        return CSVOptions{};
    }
    
    CSVOptions() = default;
};

// CSV解析结果信息
struct CSVParseInfo {
    bool success = true;
    std::string error_message;
    int rows_parsed = 0;
    int columns_detected = 0;
    bool has_header_row = false;
    std::vector<std::string> column_names;
    
    CSVParseInfo() = default;
    explicit CSVParseInfo(bool success) : success(success) {}
};

class CSVProcessor {
public:
    CSVProcessor() = default;
    ~CSVProcessor() = default;
    
    void setOptions(const CSVOptions& options) {
        options_ = options;
    }
    
    std::vector<std::vector<std::string>> parseString(const std::string& content) {
        std::vector<std::vector<std::string>> result;
        
        if (content.empty()) {
            return result;
        }
        
        // 简单的CSV解析实现
        std::istringstream stream(content);
        std::string line;
        
        while (std::getline(stream, line)) {
            if (options_.skip_empty_lines && line.empty()) {
                continue;
            }
            
            std::vector<std::string> row = parseLine(line);
            result.push_back(row);
        }
        
        return result;
    }
    
    std::string formatRow(const std::vector<std::string>& row) {
        std::ostringstream oss;
        
        for (size_t i = 0; i < row.size(); ++i) {
            if (i > 0) {
                oss << options_.delimiter;
            }
            
            std::string cell = row[i];
            bool needs_quotes = cell.find(options_.delimiter) != std::string::npos ||
                               cell.find(options_.quote_char) != std::string::npos ||
                               cell.find('\n') != std::string::npos ||
                               cell.find('\r') != std::string::npos;
            
            if (needs_quotes) {
                oss << options_.quote_char;
                // 转义引号
                for (char c : cell) {
                    if (c == options_.quote_char) {
                        oss << options_.escape_char << options_.quote_char;
                    } else {
                        oss << c;
                    }
                }
                oss << options_.quote_char;
            } else {
                oss << cell;
            }
        }
        
        return oss.str();
    }

private:
    CSVOptions options_;
    
    std::vector<std::string> parseLine(const std::string& line) {
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
            } else if (c == options_.escape_char && in_quotes) {
                if (i + 1 < line.length() && line[i + 1] == options_.quote_char) {
                    escape_next = true;
                } else {
                    current_field += c;
                }
            } else if (c == options_.quote_char) {
                in_quotes = !in_quotes;
            } else if (c == options_.delimiter && !in_quotes) {
                result.push_back(current_field);
                current_field.clear();
            } else {
                current_field += c;
            }
        }
        
        result.push_back(current_field);
        return result;
    }
};

}} // namespace fastexcel::core

// Utility functions for CSV processing
namespace fastexcel {
namespace core {

// Function declarations
std::vector<std::vector<std::string>> readCSVFromFile(const std::string& filepath, const CSVOptions& options = CSVOptions{});
bool writeCSVToFile(const std::string& filepath, const std::vector<std::vector<std::string>>& data, const CSVOptions& options = CSVOptions{});
char detectDelimiter(const std::string& sample);
std::string detectEncoding(const std::vector<uint8_t>& data);
std::vector<std::string> parseLine(const std::string& line, char delimiter, char quote_char, char escape_char);
CSVParseInfo parseContent(const std::string& content, const CSVOptions& options);
std::string escapeField(const std::string& field, const CSVOptions& options);
bool needsQuoting(const std::string& field, const CSVOptions& options);
bool isCSVFile(const std::string& filepath);
CSVOptions detectCSVOptions(const std::string& filepath);

}} // namespace fastexcel::core