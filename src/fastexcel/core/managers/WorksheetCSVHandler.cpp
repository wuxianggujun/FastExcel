#include "WorksheetCSVHandler.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/CSVProcessor.hpp"
#include <iomanip>
#include <sstream>
#include <cmath>
#include <fstream>

// External function declarations from CSVProcessor.cpp
extern "C++" {
    std::vector<std::vector<std::string>> readCSVFromFile(const std::string& filepath, const fastexcel::core::CSVOptions& options);
    bool writeCSVToFile(const std::string& filepath, const std::vector<std::vector<std::string>>& data, const fastexcel::core::CSVOptions& options);
    fastexcel::core::CSVParseInfo parseContent(const std::string& content, const fastexcel::core::CSVOptions& options);
    fastexcel::core::CSVOptions detectCSVOptions(const std::string& filepath);
    bool isCSVFile(const std::string& filepath);
}

namespace fastexcel {
namespace core {

WorksheetCSVHandler::WorksheetCSVHandler(Worksheet& worksheet)
    : worksheet_(worksheet) {
}

CSVParseInfo WorksheetCSVHandler::loadFromCSV(const std::string& filepath, const CSVOptions& options) {
    FASTEXCEL_LOG_INFO("Loading CSV from file: {} into worksheet", filepath);
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    std::ifstream file(filepath);
    if (!file.is_open()) {
        CSVParseInfo info(false);
        info.error_message = fmt::format("Failed to open file: {}", filepath);
        return info;
    }
    
    std::string content((std::istreambuf_iterator<char>(file)),
                       std::istreambuf_iterator<char>());
    file.close();
    
    if (content.empty()) {
        CSVParseInfo info(false);
        info.error_message = fmt::format("File is empty: {}", filepath);
        return info;
    }
    
    auto data = processor.parseString(content);
    
    // 将数据写入worksheet
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            worksheet_.setValue(static_cast<int>(row), static_cast<int>(col), data[row][col]);
        }
    }
    
    CSVParseInfo info(true);
    info.rows_parsed = static_cast<int>(data.size());
    info.columns_detected = data.empty() ? 0 : static_cast<int>(data[0].size());
    
    if (options.has_header && !data.empty()) {
        info.has_header_row = true;
        info.column_names = data[0];
    }
    
    return info;
}

CSVParseInfo WorksheetCSVHandler::loadFromCSVString(const std::string& csv_content, const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Loading CSV from string, content length: {}", csv_content.length());
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    auto data = processor.parseString(csv_content);
    
    // 将数据写入worksheet
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            worksheet_.setValue(static_cast<int>(row), static_cast<int>(col), data[row][col]);
        }
    }
    
    CSVParseInfo info(true);
    info.rows_parsed = static_cast<int>(data.size());
    info.columns_detected = data.empty() ? 0 : static_cast<int>(data[0].size());
    
    if (options.has_header && !data.empty()) {
        info.has_header_row = true;
        info.column_names = data[0];
    }
    
    return info;
}

bool WorksheetCSVHandler::saveAsCSV(const std::string& filepath, const CSVOptions& options) const {
    FASTEXCEL_LOG_INFO("Saving worksheet as CSV to file: {}", filepath);
    
    std::ofstream file(filepath);
    if (!file.is_open()) {
        return false;
    }
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    auto [max_row, max_col] = worksheet_.getUsedRange();
    if (max_row == -1 || max_col == -1) {
        file.close();
        return true; // 空文件
    }
    
    // 获取数据并格式化为CSV
    for (int row = 0; row <= max_row; ++row) {
        if (row > 0) {
            file << "\n";
        }
        
        std::vector<std::string> row_data;
        for (int col = 0; col <= max_col; ++col) {
            row_data.push_back(getCellDisplayValue(row, col));
        }
        
        file << processor.formatRow(row_data);
    }
    
    file.close();
    return true;
}

std::string WorksheetCSVHandler::toCSVString(const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting worksheet to CSV string");
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    auto [max_row, max_col] = worksheet_.getUsedRange();
    if (max_row == -1 || max_col == -1) {
        return "";
    }
    
    std::ostringstream oss;
    
    for (int row = 0; row <= max_row; ++row) {
        if (row > 0) {
            oss << "\n";
        }
        
        std::vector<std::string> row_data;
        for (int col = 0; col <= max_col; ++col) {
            row_data.push_back(getCellDisplayValue(row, col));
        }
        
        oss << processor.formatRow(row_data);
    }
    
    return oss.str();
}

std::string WorksheetCSVHandler::rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                                                 const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting range ({},{}) to ({},{}) to CSV string", 
                       start_row, start_col, end_row, end_col);
    
    CSVProcessor processor;
    processor.setOptions(options);
    
    std::ostringstream oss;
    
    for (int row = start_row; row <= end_row; ++row) {
        if (row > start_row) {
            oss << "\n";
        }
        
        std::vector<std::string> row_data;
        for (int col = start_col; col <= end_col; ++col) {
            row_data.push_back(getCellDisplayValue(row, col));
        }
        
        oss << processor.formatRow(row_data);
    }
    
    return oss.str();
}

CSVParseInfo WorksheetCSVHandler::previewCSV(const std::string& filepath, const CSVOptions& options) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        CSVParseInfo info(false);
        info.error_message = fmt::format("Failed to open file: {}", filepath);
        return info;
    }
    
    // 只读取前几行用于预览
    std::string preview_content;
    std::string line;
    size_t lines_read = 0;
    const size_t max_preview_lines = 10;
    
    while (std::getline(file, line) && lines_read < max_preview_lines) {
        if (lines_read > 0) {
            preview_content += "\n";
        }
        preview_content += line;
        lines_read++;
    }
    file.close();
    
    // 解析预览内容
    return ::fastexcel::core::parseContent(preview_content, options);
}

CSVOptions WorksheetCSVHandler::detectCSVOptions(const std::string& filepath) {
    return ::fastexcel::core::detectCSVOptions(filepath);
}

bool WorksheetCSVHandler::isCSVFile(const std::string& filepath) {
    return ::fastexcel::core::isCSVFile(filepath);
}

std::string WorksheetCSVHandler::getCellDisplayValue(int row, int col) const {
    try {
        if (row < 0 || col < 0) {
            return "";
        }
        
        if (!worksheet_.hasCellAt(row, col)) {
            return "";
        }
        
        const auto& cell = worksheet_.getCell(row, col);
        
        switch (cell.getType()) {
            case CellType::Empty:
                return "";
                
            case CellType::Number:
                {
                    double value = cell.getValue<double>();
                    if (value == std::floor(value) && std::abs(value) < 1e15) {
                        return fmt::format("{}", static_cast<long long>(value));
                    } else {
                        std::ostringstream oss;
                        oss << std::fixed << std::setprecision(10) << value;
                        std::string result = oss.str();
                        // 移除末尾的零
                        result.erase(result.find_last_not_of('0') + 1, std::string::npos);
                        result.erase(result.find_last_not_of('.') + 1, std::string::npos);
                        return result;
                    }
                }
                
            case CellType::String:
                return cell.getValue<std::string>();
                
            case CellType::Boolean:
                return cell.getValue<bool>() ? "TRUE" : "FALSE";
                
            case CellType::Formula:
                try {
                    double result = cell.getFormulaResult();
                    if (result == std::floor(result) && std::abs(result) < 1e15) {
                        return fmt::format("{}", static_cast<long long>(result));
                    } else {
                        std::ostringstream oss;
                        oss << std::fixed << std::setprecision(10) << result;
                        std::string result_str = oss.str();
                        result_str.erase(result_str.find_last_not_of('0') + 1, std::string::npos);
                        result_str.erase(result_str.find_last_not_of('.') + 1, std::string::npos);
                        return result_str;
                    }
                } catch (const std::runtime_error& e) {
                    FASTEXCEL_LOG_DEBUG("Formula execution error: {}", e.what());
                    try {
                        std::string formula = cell.getFormula();
                        return formula.empty() ? "=" : fmt::format("={}", formula);
                    } catch (const fmt::format_error& e) {
                        FASTEXCEL_LOG_DEBUG("Formula format error: {}", e.what());
                        return "#FORMULA_ERROR";
                    } catch (const std::exception& e) {
                        FASTEXCEL_LOG_DEBUG("Exception getting formula: {}", e.what());
                        return "#FORMULA_ERROR";
                    }
                } catch (const std::exception& e) {
                    FASTEXCEL_LOG_DEBUG("Exception getting formula result: {}", e.what());
                    try {
                        std::string formula = cell.getFormula();
                        return formula.empty() ? "=" : fmt::format("={}", formula);
                    } catch (const fmt::format_error& e) {
                        FASTEXCEL_LOG_DEBUG("Formula format error: {}", e.what());
                        return "#FORMULA_ERROR";
                    } catch (const std::exception& e) {
                        FASTEXCEL_LOG_DEBUG("Exception getting formula: {}", e.what());
                        return "#FORMULA_ERROR";
                    }
                }
                
            case CellType::Error:
                return "#ERROR";
                
            default:
                return "";
        }
        
    } catch (const std::exception& /*e*/) {
        return "";
    }
}

} // namespace core
} // namespace fastexcel
