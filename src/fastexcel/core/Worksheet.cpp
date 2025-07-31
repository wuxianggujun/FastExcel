#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/xml/XMLWriter.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <sstream>
#include <stdexcept>

namespace fastexcel {
namespace core {

Worksheet::Worksheet(const std::string& name, std::shared_ptr<Workbook> workbook, int sheet_id)
    : name_(name), parent_workbook_(workbook), sheet_id_(sheet_id) {
}

Cell& Worksheet::getCell(int row, int col) {
    validateCellPosition(row, col);
    return cells_[std::make_pair(row, col)];
}

const Cell& Worksheet::getCell(int row, int col) const {
    validateCellPosition(row, col);
    auto it = cells_.find(std::make_pair(row, col));
    if (it == cells_.end()) {
        static Cell empty_cell;
        return empty_cell;
    }
    return it->second;
}

void Worksheet::writeString(int row, int col, const std::string& value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    cell.setValue(value);
    if (format) {
        cell.setFormat(format);
    }
}

void Worksheet::writeNumber(int row, int col, double value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    cell.setValue(value);
    if (format) {
        cell.setFormat(format);
    }
}

void Worksheet::writeBoolean(int row, int col, bool value, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    cell.setValue(value);
    if (format) {
        cell.setFormat(format);
    }
}

void Worksheet::writeFormula(int row, int col, const std::string& formula, std::shared_ptr<Format> format) {
    validateCellPosition(row, col);
    auto& cell = cells_[std::make_pair(row, col)];
    cell.setFormula(formula);
    if (format) {
        cell.setFormat(format);
    }
}

void Worksheet::writeRange(int start_row, int start_col, const std::vector<std::vector<std::string>>& data) {
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            writeString(start_row + row, start_col + col, data[row][col]);
        }
    }
}

void Worksheet::writeRange(int start_row, int start_col, const std::vector<std::vector<double>>& data) {
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            writeNumber(start_row + row, start_col + col, data[row][col]);
        }
    }
}

std::pair<int, int> Worksheet::getUsedRange() const {
    int max_row = -1;
    int max_col = -1;
    
    for (const auto& [pos, cell] : cells_) {
        if (!cell.isEmpty()) {
            max_row = std::max(max_row, pos.first);
            max_col = std::max(max_col, pos.second);
        }
    }
    
    return {max_row, max_col};
}

std::string Worksheet::generateXML() const {
    xml::XMLWriter writer;
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 获取使用范围
    auto [max_row, max_col] = getUsedRange();
    
    if (max_row >= 0 && max_col >= 0) {
        writer.startElement("sheetData");
        
        // 按行处理数据
        for (int row = 0; row <= max_row; ++row) {
            bool has_data_in_row = false;
            
            // 检查这一行是否有数据
            for (int col = 0; col <= max_col; ++col) {
                auto it = cells_.find(std::make_pair(row, col));
                if (it != cells_.end() && !it->second.isEmpty()) {
                    has_data_in_row = true;
                    break;
                }
            }
            
            if (has_data_in_row) {
                writer.startElement("row");
                writer.writeAttribute("r", std::to_string(row + 1));
                
                for (int col = 0; col <= max_col; ++col) {
                    auto it = cells_.find(std::make_pair(row, col));
                    if (it != cells_.end() && !it->second.isEmpty()) {
                        const auto& cell = it->second;
                        
                        writer.startElement("c");
                        writer.writeAttribute("r", cellReference(row, col));
                        
                        // 如果有格式，添加样式ID
                        if (cell.getFormat()) {
                            writer.writeAttribute("s", std::to_string(cell.getFormat()->getFormatId()));
                        }
                        
                        if (cell.isString()) {
                            writer.writeAttribute("t", "inlineStr");
                            writer.startElement("is");
                            writer.startElement("t");
                            writer.writeText(cell.getStringValue());
                            writer.endElement(); // t
                            writer.endElement(); // is
                        } else if (cell.isNumber()) {
                            writer.startElement("v");
                            writer.writeText(std::to_string(cell.getNumberValue()));
                            writer.endElement(); // v
                        } else if (cell.isBoolean()) {
                            writer.writeAttribute("t", "b");
                            writer.startElement("v");
                            writer.writeText(cell.getBooleanValue() ? "1" : "0");
                            writer.endElement(); // v
                        } else if (cell.isFormula()) {
                            writer.writeAttribute("t", "str");
                            writer.startElement("f");
                            writer.writeText(cell.getFormula());
                            writer.endElement(); // f
                        }
                        
                        writer.endElement(); // c
                    }
                }
                
                writer.endElement(); // row
            }
        }
        
        writer.endElement(); // sheetData
    }
    
    writer.endElement(); // worksheet
    writer.endDocument();
    
    return writer.toString();
}

void Worksheet::clear() {
    cells_.clear();
}

std::string Worksheet::columnToLetter(int col) const {
    std::string result;
    while (col >= 0) {
        result = static_cast<char>('A' + (col % 26)) + result;
        col = col / 26 - 1;
    }
    return result;
}

std::string Worksheet::cellReference(int row, int col) const {
    return columnToLetter(col) + std::to_string(row + 1);
}

void Worksheet::validateCellPosition(int row, int col) const {
    if (row < 0 || col < 0) {
        throw std::invalid_argument("Cell position cannot be negative");
    }
    if (row > 1048575 || col > 16383) { // Excel 2007+ limits
        throw std::invalid_argument("Cell position exceeds Excel limits");
    }
}

}} // namespace fastexcel::core