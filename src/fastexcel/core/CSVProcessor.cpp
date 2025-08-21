#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <fstream>
#include <sstream>
#include <regex>
#include <algorithm>
#include <locale>
#include <codecvt>
#include <iomanip>
#include <climits>

namespace fastexcel {
namespace core {

// CSVReader 实现

CSVParseInfo CSVReader::loadFromFile(const std::string& filepath, 
                                    Worksheet& worksheet, 
                                    const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Loading CSV from file: {}", filepath);
    
    // 读取文件内容（使用UTF-8编码）
    std::ifstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        CSVParseInfo info;
        info.errors.push_back("Failed to open file: " + filepath);
        FASTEXCEL_LOG_ERROR("Failed to open CSV file: {}", filepath);
        return info;
    }
    
    // 设置UTF-8编码（Windows特定处理）
    file.imbue(std::locale(""));
    
    // 读取文件到字符串
    std::string content((std::istreambuf_iterator<char>(file)),
                       std::istreambuf_iterator<char>());
    file.close();
    
    if (content.empty()) {
        CSVParseInfo info;
        info.errors.push_back("File is empty: " + filepath);
        return info;
    }
    
    // 如果启用自动检测，先检测选项
    CSVOptions final_options = options;
    if (options.auto_detect_delimiter || options.auto_detect_encoding) {
        auto detected = detectOptions(filepath);
        if (options.auto_detect_delimiter) {
            final_options.delimiter = detected.delimiter;
        }
        if (options.auto_detect_encoding) {
            final_options.encoding = detected.encoding;
        }
    }
    
    return parseContent(content, worksheet, final_options);
}

CSVParseInfo CSVReader::loadFromString(const std::string& csv_content,
                                      Worksheet& worksheet,
                                      const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Loading CSV from string, length: {}", csv_content.length());
    return parseContent(csv_content, worksheet, options);
}

CSVParseInfo CSVReader::previewFile(const std::string& filepath,
                                   const CSVOptions& options) {
    // 只读取前几行进行预览
    std::ifstream file(filepath);
    if (!file.is_open()) {
        CSVParseInfo info;
        info.errors.push_back("Failed to open file: " + filepath);
        return info;
    }
    
    std::string preview_content;
    std::string line;
    size_t lines_read = 0;
    
    while (std::getline(file, line) && lines_read < options.preview_lines) {
        preview_content += line + "\\n";
        lines_read++;
    }
    file.close();
    
    // 创建临时工作簿和工作表进行解析
    auto temp_workbook = Workbook::create(Path("temp.xlsx"));
    auto temp_worksheet = temp_workbook->addSheet("temp");
    return parseContent(preview_content, *temp_worksheet, options);
}

CSVOptions CSVReader::detectOptions(const std::string& filepath) {
    CSVOptions options;
    
    // 读取文件的前几行用于检测
    std::ifstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("Failed to open file for detection: {}", filepath);
        return options;
    }
    
    // 读取前1KB用于检测
    const size_t sample_size = 1024;
    std::vector<char> buffer(sample_size);
    file.read(buffer.data(), sample_size);
    size_t bytes_read = file.gcount();
    file.close();
    
    std::string sample(buffer.begin(), buffer.begin() + bytes_read);
    
    // 检测分隔符
    options.delimiter = detectDelimiter(sample);
    
    // 检测编码（简化实现）
    options.encoding = detectEncoding(std::vector<uint8_t>(buffer.begin(), buffer.begin() + bytes_read));
    
    // 检测是否有标题行（简单启发式）
    std::istringstream ss(sample);
    std::string first_line;
    if (std::getline(ss, first_line)) {
        auto fields = parseLine(first_line, options);
        bool likely_header = true;
        
        // 如果第一行的字段看起来像数字，可能不是标题
        for (const auto& field : fields) {
            if (!field.empty() && std::all_of(field.begin(), field.end(), 
                [](unsigned char c) { return std::isdigit(c) || c == '.' || c == '-'; })) {
                likely_header = false;
                break;
            }
        }
        options.has_header = likely_header;
    }
    
    FASTEXCEL_LOG_DEBUG("Detected CSV options: delimiter='{}', encoding={}, has_header={}", 
                       options.delimiter, options.encoding, options.has_header);
    
    return options;
}

CSVParseInfo CSVReader::parseContent(const std::string& content,
                                    Worksheet& worksheet,
                                    const CSVOptions& options) {
    CSVParseInfo info;
    info.detected_delimiter = options.delimiter;
    info.detected_encoding = options.encoding;
    
    std::istringstream stream(content);
    std::string line;
    int current_row = 0;
    bool is_first_row = true;
    
    while (std::getline(stream, line)) {
        // 跳过空行
        if (options.skip_empty_lines && line.empty()) {
            continue;
        }
        
        try {
            auto fields = parseLine(line, options);
            
            if (fields.empty()) {
                continue;
            }
            
            // 更新列数统计
            info.columns_detected = std::max(info.columns_detected, fields.size());
            
            // 处理标题行
            if (is_first_row && options.has_header) {
                info.column_names = fields;
                info.has_header_detected = true;
                // 将header写入第0行
                for (size_t col = 0; col < fields.size(); ++col) {
                    worksheet.setValue(current_row, static_cast<int>(col), fields[col]);
                }
                current_row++;
                is_first_row = false;
                continue;
            }
            
            // 设置单元格值
            for (size_t col = 0; col < fields.size(); ++col) {
                setCellValue(worksheet, current_row, static_cast<int>(col), fields[col], options);
            }
            
            current_row++;
            info.rows_processed++;
            is_first_row = false;
            
        } catch (const std::exception& e) {
            std::string error_msg = "Error parsing line " + std::to_string(info.rows_processed + 1) + ": " + e.what();
            info.errors.push_back(error_msg);
            
            if (options.strict_mode || info.errors.size() >= options.max_errors) {
                break;
            }
        }
    }
    
    FASTEXCEL_LOG_INFO("CSV parsing completed: {} rows, {} columns, {} errors", 
                      info.rows_processed, info.columns_detected, info.errors.size());
    
    return info;
}

char CSVReader::detectDelimiter(const std::string& sample) {
    // 常见分隔符的候选列表
    std::vector<char> candidates = {',', ';', '\t', '|'};
    std::map<char, int> scores;
    
    // 分析每个候选分隔符的出现频率和一致性
    for (char delim : candidates) {
        std::istringstream ss(sample);
        std::string line;
        std::vector<size_t> field_counts;
        
        while (std::getline(ss, line) && field_counts.size() < 10) {
            size_t count = 1;
            bool in_quote = false;
            
            for (char c : line) {
                if (c == '"') {
                    in_quote = !in_quote;
                } else if (c == delim && !in_quote) {
                    count++;
                }
            }
            
            if (count > 1) {
                field_counts.push_back(count);
            }
        }
        
        // 计算一致性分数
        if (!field_counts.empty()) {
            auto first_count = field_counts[0];
            int consistency = 0;
            for (auto count : field_counts) {
                if (count == first_count) {
                    consistency++;
                }
            }
            scores[delim] = consistency * 10 + static_cast<int>(field_counts.size());
        }
    }
    
    // 选择得分最高的分隔符
    char best_delimiter = ',';
    int best_score = 0;
    for (const auto& pair : scores) {
        if (pair.second > best_score) {
            best_score = pair.second;
            best_delimiter = pair.first;
        }
    }
    
    return best_delimiter;
}

std::string CSVReader::detectEncoding(const std::vector<uint8_t>& data) {
    // 简化的编码检测，检查BOM和一些启发式规则
    if (data.size() >= 3) {
        // UTF-8 BOM
        if (data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
            return "UTF-8";
        }
    }
    
    if (data.size() >= 2) {
        // UTF-16 BOM
        if ((data[0] == 0xFF && data[1] == 0xFE) || (data[0] == 0xFE && data[1] == 0xFF)) {
            return "UTF-16";
        }
    }
    
    // 检查是否包含非ASCII字符
    bool has_non_ascii = false;
    for (uint8_t byte : data) {
        if (byte > 127) {
            has_non_ascii = true;
            break;
        }
    }
    
    return has_non_ascii ? "UTF-8" : "ASCII";
}

std::vector<std::string> CSVReader::parseLine(const std::string& line, const CSVOptions& options) {
    std::vector<std::string> fields;
    std::string current_field;
    bool in_quote = false;
    bool quote_started = false;
    
    for (size_t i = 0; i < line.length(); ++i) {
        char c = line[i];
        
        if (c == options.quote_char) {
            if (!in_quote) {
                in_quote = true;
                quote_started = true;
            } else {
                // 检查是否是转义的引号
                if (i + 1 < line.length() && line[i + 1] == options.quote_char) {
                    current_field += options.quote_char;
                    i++; // 跳过下一个引号
                } else {
                    in_quote = false;
                }
            }
        } else if (c == options.delimiter && !in_quote) {
            // 字段结束
            if (options.trim_whitespace && !quote_started) {
                // 去除前后空白（只有非引号字段）
                size_t start = current_field.find_first_not_of(" \t");
                size_t end = current_field.find_last_not_of(" \t");
                if (start != std::string::npos) {
                    current_field = current_field.substr(start, end - start + 1);
                } else {
                    current_field.clear();
                }
            }
            fields.push_back(current_field);
            current_field.clear();
            quote_started = false;
        } else {
            current_field += c;
        }
    }
    
    // 添加最后一个字段
    if (options.trim_whitespace && !quote_started) {
        size_t start = current_field.find_first_not_of(" \t");
        size_t end = current_field.find_last_not_of(" \t");
        if (start != std::string::npos) {
            current_field = current_field.substr(start, end - start + 1);
        } else {
            current_field.clear();
        }
    }
    fields.push_back(current_field);
    
    return fields;
}

void CSVReader::setCellValue(Worksheet& worksheet, int row, int col, 
                           const std::string& value, const CSVOptions& options) {
    if (value.empty()) {
        return; // 空值不设置
    }
    
    // 如果不启用类型推断，直接设置为字符串
    if (!options.auto_detect_types) {
        worksheet.setValue(row, col, value);
        return;
    }
    
    // 尝试解析数字
    if (options.parse_numbers) {
        try {
            // 改进的数字检测逻辑
            std::string trimmed_value = value;
            
            // 去除前后空白
            trimmed_value.erase(0, trimmed_value.find_first_not_of(" \t"));
            trimmed_value.erase(trimmed_value.find_last_not_of(" \t") + 1);
            
            // 检查是否为有效数字格式
            if (!trimmed_value.empty()) {
                bool is_negative = false;
                size_t start_pos = 0;
                
                // 处理负号
                if (trimmed_value[0] == '-') {
                    is_negative = true;
                    start_pos = 1;
                }
                
                // 检查是否包含小数点或科学记数法
                bool has_decimal = trimmed_value.find('.', start_pos) != std::string::npos;
                bool has_scientific = trimmed_value.find_first_of("eE", start_pos) != std::string::npos;
                
                if (!has_decimal && !has_scientific) {
                    // 尝试解析为整数
                    long long int_val = std::stoll(trimmed_value);
                    // 检查是否在int范围内
                    if (int_val >= INT_MIN && int_val <= INT_MAX) {
                        worksheet.setValue(row, col, static_cast<int>(int_val));
                        return;
                    }
                }
                
                // 尝试解析为浮点数
                double double_val = std::stod(trimmed_value);
                worksheet.setValue(row, col, double_val);
                return;
            }
        } catch (const std::exception&) {
            // 不是数字，继续其他类型检测
        }
    }
    
    // 尝试解析日期（简化实现）
    if (options.parse_dates) {
        // 修复正则表达式 - 移除多余的转义符
        std::regex date_pattern(R"(\d{4}[-/]\d{1,2}[-/]\d{1,2})");
        if (std::regex_match(value, date_pattern)) {
            // 简单的日期格式，这里可以扩展
            worksheet.setValue(row, col, value); // 暂时作为字符串处理
            return;
        }
    }
    
    // 尝试解析布尔值
    std::string lower_value = value;
    std::transform(lower_value.begin(), lower_value.end(), lower_value.begin(), 
                  [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    
    if (lower_value == "true" || lower_value == "1" || lower_value == "yes" || lower_value == "y") {
        worksheet.setValue(row, col, true);
        return;
    } else if (lower_value == "false" || lower_value == "0" || lower_value == "no" || lower_value == "n") {
        worksheet.setValue(row, col, false);
        return;
    }
    
    // 默认作为字符串处理
    worksheet.setValue(row, col, value);
}

// CSVWriter 实现

bool CSVWriter::saveToFile(const Worksheet& worksheet,
                          const std::string& filepath,
                          const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Saving worksheet to CSV file: {}", filepath);
    
    std::string csv_content = saveToString(worksheet, options);
    if (csv_content.empty()) {
        FASTEXCEL_LOG_ERROR("Failed to generate CSV content");
        return false;
    }
    
    std::ofstream file(filepath, std::ios::binary);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("Failed to open file for writing: {}", filepath);
        return false;
    }
    
    // 设置UTF-8编码
    file.imbue(std::locale(""));
    
    // 写入UTF-8 BOM以确保正确识别编码
    file.write("\xEF\xBB\xBF", 3);
    
    file << csv_content;
    file.close();
    
    FASTEXCEL_LOG_INFO("Successfully saved CSV file: {}", filepath);
    return true;
}

std::string CSVWriter::saveToString(const Worksheet& worksheet, const CSVOptions& options) {
    // 检查工作表是否为空
    if (worksheet.isEmpty()) {
        FASTEXCEL_LOG_DEBUG("Worksheet is empty, returning empty CSV");
        return "";
    }
    
    // 获取工作表的完整使用范围
    auto [min_row, max_row, min_col, max_col] = worksheet.getUsedRangeFull();
    
    if (max_row < 0 || max_col < 0) {
        return "";
    }
    
    return saveRangeToString(worksheet, min_row, min_col, max_row, max_col, options);
}

std::string CSVWriter::saveRangeToString(const Worksheet& worksheet,
                                        int start_row, int start_col,
                                        int end_row, int end_col,
                                        const CSVOptions& options) {
    std::ostringstream csv_stream;
    
    for (int row = start_row; row <= end_row; ++row) {
        std::vector<std::string> row_fields;
        
        for (int col = start_col; col <= end_col; ++col) {
            std::string cell_value = worksheet.getCellDisplayValue(row, col);
            std::string escaped_value = escapeField(cell_value, options);
            row_fields.push_back(escaped_value);
        }
        
        // 连接字段
        for (size_t i = 0; i < row_fields.size(); ++i) {
            if (i > 0) {
                csv_stream << options.delimiter;
            }
            csv_stream << row_fields[i];
        }
        
        csv_stream << options.line_terminator;
    }
    
    return csv_stream.str();
}

std::string CSVWriter::escapeField(const std::string& field, const CSVOptions& options) {
    if (field.empty()) {
        return field;
    }
    
    bool needs_quoting = needsQuoting(field, options);
    
    if (!needs_quoting) {
        return field;
    }
    
    // 转义引号字符
    std::string escaped_field;
    for (char c : field) {
        if (c == options.quote_char) {
            escaped_field += options.escape_char;
            escaped_field += c;
        } else {
            escaped_field += c;
        }
    }
    
    // 添加引号包围
    return std::string(1, options.quote_char) + escaped_field + std::string(1, options.quote_char);
}

bool CSVWriter::needsQuoting(const std::string& field, const CSVOptions& options) {
    // 检查是否包含需要转义的字符
    for (char c : field) {
        if (c == options.delimiter || c == options.quote_char || 
            c == '\\n' || c == '\\r' || c == '\\t') {
            return true;
        }
    }
    
    // 检查是否以空格开头或结尾
    if (!field.empty() && (field.front() == ' ' || field.back() == ' ')) {
        return true;
    }
    
    return false;
}

// CSVUtils 实现

bool CSVUtils::isCSVFile(const std::string& filepath) {
    // 检查文件扩展名
    size_t dot_pos = filepath.find_last_of('.');
    if (dot_pos == std::string::npos) {
        return false;
    }
    
    std::string extension = filepath.substr(dot_pos + 1);
    std::transform(extension.begin(), extension.end(), extension.begin(), 
                  [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    
    return extension == "csv" || extension == "tsv" || extension == "txt";
}

CSVOptions CSVUtils::optionsFromExtension(const std::string& filepath) {
    CSVOptions options;
    
    size_t dot_pos = filepath.find_last_of('.');
    if (dot_pos != std::string::npos) {
        std::string extension = filepath.substr(dot_pos + 1);
        std::transform(extension.begin(), extension.end(), extension.begin(), 
                      [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
        
        if (extension == "tsv") {
            options.delimiter = '\t';
        } else if (extension == "csv") {
            options.delimiter = ',';
        }
    }
    
    return options;
}

bool CSVUtils::validateOptions(const CSVOptions& options, std::string& error_message) {
    // 验证分隔符
    if (options.delimiter == options.quote_char) {
        error_message = "Delimiter and quote character cannot be the same";
        return false;
    }
    
    // 验证字段大小限制
    if (options.max_field_size == 0) {
        error_message = "Max field size must be greater than 0";
        return false;
    }
    
    // 验证预览行数
    if (options.preview_lines == 0) {
        error_message = "Preview lines must be greater than 0";
        return false;
    }
    
    return true;
}

std::string CSVUtils::getDelimiterName(char delimiter) {
    switch (delimiter) {
        case ',': return "Comma";
        case ';': return "Semicolon";
        case '\\t': return "Tab";
        case '|': return "Pipe";
        case ' ': return "Space";
        default: return "Custom (" + std::string(1, delimiter) + ")";
    }
}

size_t CSVUtils::countLines(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        return 0;
    }
    
    size_t line_count = 0;
    std::string line;
    while (std::getline(file, line)) {
        line_count++;
    }
    
    return line_count;
}

}} // namespace fastexcel::core
