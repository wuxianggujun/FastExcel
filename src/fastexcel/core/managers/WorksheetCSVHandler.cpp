#include "WorksheetCSVHandler.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iomanip>
#include <sstream>
#include <cmath>

namespace fastexcel {
namespace core {

WorksheetCSVHandler::WorksheetCSVHandler(Worksheet& worksheet)
    : worksheet_(worksheet) {
}

CSVParseInfo WorksheetCSVHandler::loadFromCSV(const std::string& filepath, const CSVOptions& options) {
    FASTEXCEL_LOG_INFO("Loading CSV from file: {} into worksheet", filepath);
    return CSVReader::loadFromFile(filepath, worksheet_, options);
}

CSVParseInfo WorksheetCSVHandler::loadFromCSVString(const std::string& csv_content, const CSVOptions& options) {
    FASTEXCEL_LOG_DEBUG("Loading CSV from string, content length: {}", csv_content.length());
    return CSVReader::loadFromString(csv_content, worksheet_, options);
}

bool WorksheetCSVHandler::saveAsCSV(const std::string& filepath, const CSVOptions& options) const {
    FASTEXCEL_LOG_INFO("Saving worksheet as CSV to file: {}", filepath);
    return CSVWriter::saveToFile(worksheet_, filepath, options);
}

std::string WorksheetCSVHandler::toCSVString(const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting worksheet to CSV string");
    return CSVWriter::saveToString(worksheet_, options);
}

std::string WorksheetCSVHandler::rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                                                 const CSVOptions& options) const {
    FASTEXCEL_LOG_DEBUG("Converting range ({},{}) to ({},{}) to CSV string", 
                       start_row, start_col, end_row, end_col);
    return CSVWriter::saveRangeToString(worksheet_, start_row, start_col, end_row, end_col, options);
}

CSVParseInfo WorksheetCSVHandler::previewCSV(const std::string& filepath, const CSVOptions& options) {
    return CSVReader::previewFile(filepath, options);
}

CSVOptions WorksheetCSVHandler::detectCSVOptions(const std::string& filepath) {
    return CSVReader::detectOptions(filepath);
}

bool WorksheetCSVHandler::isCSVFile(const std::string& filepath) {
    return CSVUtils::isCSVFile(filepath);
}

std::string WorksheetCSVHandler::getCellDisplayValue(int row, int col) const {
    try {
        if (row < 0 || col < 0) {
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
                        return std::to_string(static_cast<long long>(value));
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
                        return std::to_string(static_cast<long long>(result));
                    } else {
                        std::ostringstream oss;
                        oss << std::fixed << std::setprecision(10) << result;
                        std::string result_str = oss.str();
                        result_str.erase(result_str.find_last_not_of('0') + 1, std::string::npos);
                        result_str.erase(result_str.find_last_not_of('.') + 1, std::string::npos);
                        return result_str;
                    }
                } catch (...) {
                    try {
                        std::string formula = cell.getFormula();
                        return formula.empty() ? "=" : "=" + formula;
                    } catch (...) {
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