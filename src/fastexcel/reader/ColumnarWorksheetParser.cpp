#include "ColumnarWorksheetParser.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <stdexcept>

namespace fastexcel {
namespace reader {

bool ColumnarWorksheetParser::parseToColumnar(archive::ZipReader* zip_reader,
                                            const std::string& internal_path,
                                            core::columnar::ReadOnlyWorksheet* worksheet,
                                            const std::unordered_map<int, std::string>& shared_strings,
                                            const core::columnar::ReadOnlyOptions& options) {
    if (!zip_reader || !worksheet) {
        return false;
    }

    // 初始化解析状态
    state_.reset();
    state_.worksheet = worksheet;
    state_.shared_strings = &shared_strings;
    state_.setupProjection(options);

    try {
        // 提取工作表XML内容
        std::string xml_content;
        auto error = zip_reader->extractFile(internal_path, xml_content);
        if (error != archive::ZipError::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to extract worksheet: {}", internal_path);
            return false;
        }

        // 解析XML内容
        return parseToColumnar(xml_content, worksheet, shared_strings, options);
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during columnar worksheet parsing: {}", e.what());
        return false;
    }
}

bool ColumnarWorksheetParser::parseToColumnar(const std::string& xml_content,
                                            core::columnar::ReadOnlyWorksheet* worksheet,
                                            const std::unordered_map<int, std::string>& shared_strings,
                                            const core::columnar::ReadOnlyOptions& options) {
    if (xml_content.empty() || !worksheet) {
        return false;
    }

    try {
        // 初始化解析状态
        state_.reset();
        state_.worksheet = worksheet;
        state_.shared_strings = &shared_strings;
        state_.setupProjection(options);

        // 简化实现：直接解析sheetData部分
        // 这里应该使用XML解析器，但为了快速实现，我们先用简单的字符串查找
        
        // 查找sheetData标签
        size_t sheet_data_start = xml_content.find("<sheetData");
        if (sheet_data_start == std::string::npos) {
            // 没有数据，但这不是错误
            FASTEXCEL_LOG_DEBUG("No sheetData found in worksheet");
            return true;
        }
        
        size_t sheet_data_end = xml_content.find("</sheetData>", sheet_data_start);
        if (sheet_data_end == std::string::npos) {
            FASTEXCEL_LOG_ERROR("Malformed sheetData in worksheet");
            return false;
        }

        // 提取sheetData内容进行解析
        std::string sheet_data = xml_content.substr(sheet_data_start, sheet_data_end - sheet_data_start);
        
        // 简单的行解析 (实际应该使用完整的XML解析器)
        parseSheetDataSimple(sheet_data);
        
        return true;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during columnar parsing: {}", e.what());
        return false;
    }
}

void ColumnarWorksheetParser::parseSheetDataSimple(const std::string& sheet_data) {
    // 简化实现：解析基本的行和单元格数据
    // 在实际实现中应该使用完整的XML解析器
    
    size_t pos = 0;
    while ((pos = sheet_data.find("<row ", pos)) != std::string::npos) {
        // 提取行号
        size_t r_pos = sheet_data.find(" r=\"", pos);
        if (r_pos == std::string::npos) {
            pos++;
            continue;
        }
        
        size_t r_end = sheet_data.find("\"", r_pos + 4);
        if (r_end == std::string::npos) {
            pos++;
            continue;
        }
        
        std::string row_str = sheet_data.substr(r_pos + 4, r_end - r_pos - 4);
        uint32_t row = static_cast<uint32_t>(std::stoi(row_str)) - 1; // Excel行号从1开始，我们从0开始
        
        // 检查行限制
        if (state_.shouldSkipRow(row)) {
            // 跳到下一行
            size_t row_end = sheet_data.find("</row>", pos);
            if (row_end != std::string::npos) {
                pos = row_end + 6;
            } else {
                break;
            }
            continue;
        }
        
        // 解析该行的单元格
        size_t row_end = sheet_data.find("</row>", pos);
        if (row_end == std::string::npos) {
            break;
        }
        
        std::string row_data = sheet_data.substr(pos, row_end - pos);
        parseCellsInRow(row_data, row);
        
        pos = row_end + 6;
    }
}

void ColumnarWorksheetParser::parseCellsInRow(const std::string& row_data, uint32_t row) {
    size_t pos = 0;
    while ((pos = row_data.find("<c ", pos)) != std::string::npos) {
        // 提取单元格引用
        size_t r_pos = row_data.find(" r=\"", pos);
        if (r_pos == std::string::npos) {
            pos++;
            continue;
        }
        
        size_t r_end = row_data.find("\"", r_pos + 4);
        if (r_end == std::string::npos) {
            pos++;
            continue;
        }
        
        std::string cell_ref = row_data.substr(r_pos + 4, r_end - r_pos - 4);
        uint32_t col = parseColumnReference(cell_ref);
        
        // 检查列过滤
        if (state_.shouldSkipColumn(col)) {
            // 跳到下一个单元格
            size_t cell_end = row_data.find("</c>", pos);
            if (cell_end != std::string::npos) {
                pos = cell_end + 4;
            } else {
                pos++;
            }
            continue;
        }
        
        // 提取单元格类型
        std::string cell_type = "n"; // 默认为数值
        size_t t_pos = row_data.find(" t=\"", pos);
        if (t_pos != std::string::npos && t_pos < r_end + 100) { // 确保在当前单元格范围内
            size_t t_end = row_data.find("\"", t_pos + 4);
            if (t_end != std::string::npos) {
                cell_type = row_data.substr(t_pos + 4, t_end - t_pos - 4);
            }
        }
        
        // 提取单元格值
        size_t v_start = row_data.find("<v>", pos);
        if (v_start != std::string::npos) {
            size_t v_end = row_data.find("</v>", v_start);
            if (v_end != std::string::npos) {
                std::string value = row_data.substr(v_start + 3, v_end - v_start - 3);
                processCellValue(value, cell_type, row, col);
            }
        }
        
        // 移动到下一个单元格
        size_t cell_end = row_data.find("</c>", pos);
        if (cell_end != std::string::npos) {
            pos = cell_end + 4;
        } else {
            pos++;
        }
    }
}

void ColumnarWorksheetParser::processCellValue(const std::string& value, 
                                             const std::string& cell_type, 
                                             uint32_t row, uint32_t col) {
    if (!state_.worksheet || value.empty()) {
        return;
    }
    
    try {
        if (cell_type == "s") {
            // 共享字符串
            int sst_index = std::stoi(value);
            if (state_.shared_strings) {
                auto it = state_.shared_strings->find(sst_index);
                if (it != state_.shared_strings->end()) {
                    state_.worksheet->setValue(row, col, std::string_view(it->second));
                }
            }
        } else if (cell_type == "str" || cell_type == "inlineStr") {
            // 内联字符串
            state_.worksheet->setValue(row, col, std::string_view(value));
        } else if (cell_type == "b") {
            // 布尔值
            bool bool_value = (value == "1" || value == "true");
            state_.worksheet->setValue(row, col, bool_value);
        } else {
            // 数值 (默认)
            double num_value = std::stod(value);
            state_.worksheet->setValue(row, col, num_value);
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("Failed to process cell value: {}", e.what());
        // 作为字符串存储
        state_.worksheet->setValue(row, col, std::string_view(value));
    }
}

uint32_t ColumnarWorksheetParser::parseColumnReference(const std::string& cell_ref) const {
    uint32_t col = 0;
    for (char c : cell_ref) {
        if (c >= 'A' && c <= 'Z') {
            col = col * 26 + (c - 'A' + 1);
        } else {
            break; // 遇到数字就停止
        }
    }
    return col > 0 ? col - 1 : 0; // Excel列从1开始，我们从0开始
}

uint32_t ColumnarWorksheetParser::parseColumnReference(std::string_view cell_ref) const {
    uint32_t col = 0;
    for (char c : cell_ref) {
        if (c >= 'A' && c <= 'Z') {
            col = col * 26 + (c - 'A' + 1);
        } else {
            break; // 遇到数字就停止
        }
    }
    return col > 0 ? col - 1 : 0; // Excel列从1开始，我们从0开始
}

uint32_t ColumnarWorksheetParser::parseRowReference(const std::string& cell_ref) const {
    size_t num_start = 0;
    for (size_t i = 0; i < cell_ref.length(); ++i) {
        if (cell_ref[i] >= '0' && cell_ref[i] <= '9') {
            num_start = i;
            break;
        }
    }
    
    if (num_start < cell_ref.length()) {
        std::string row_str = cell_ref.substr(num_start);
        return static_cast<uint32_t>(std::stoi(row_str)) - 1; // Excel行从1开始，我们从0开始
    }
    
    return 0;
}

uint32_t ColumnarWorksheetParser::parseRowReference(std::string_view cell_ref) const {
    size_t num_start = 0;
    for (size_t i = 0; i < cell_ref.length(); ++i) {
        if (cell_ref[i] >= '0' && cell_ref[i] <= '9') {
            num_start = i;
            break;
        }
    }
    
    if (num_start < cell_ref.length()) {
        std::string row_str(cell_ref.substr(num_start));
        return static_cast<uint32_t>(std::stoi(row_str)) - 1; // Excel行从1开始，我们从0开始
    }
    
    return 0;
}

bool ColumnarWorksheetParser::isProjectedColumn(uint32_t col) const {
    return !state_.shouldSkipColumn(col);
}

// 未实现的方法 - 用于完整的XML解析
void ColumnarWorksheetParser::handleStartElement(std::string_view name, 
                                                xml::span<const xml::XMLAttribute> attributes, 
                                                int depth) {
    // TODO: 实现完整的XML解析
}

void ColumnarWorksheetParser::handleEndElement(std::string_view name, int depth) {
    // TODO: 实现完整的XML解析
}

void ColumnarWorksheetParser::handleText(std::string_view text, int depth) {
    // TODO: 实现完整的XML解析
}

void ColumnarWorksheetParser::processCellElement(xml::span<const xml::XMLAttribute> attributes) {
    // TODO: 实现完整的XML解析
}

}} // namespace fastexcel::reader