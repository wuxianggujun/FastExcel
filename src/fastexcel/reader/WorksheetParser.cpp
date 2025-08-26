//
// Created by wuxianggujun on 25-8-4.
//

#include "WorksheetParser.hpp"

#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/core/RangeFormatter.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"

#include <algorithm>
#include "fastexcel/utils/TimeUtils.hpp"
#include <cctype>
#include <ctime>
#include <iostream>
#include <sstream>
#include <string>
#include <optional>

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
    
    // 初始化解析状态
    state_.reset();
    state_.worksheet = worksheet;
    state_.shared_strings = &shared_strings;
    state_.styles = &styles;
    state_.style_id_mapping = &style_id_mapping;
    
    try {
        // 使用SAX流式解析
        xml::XMLStreamReader reader;
        
        // 设置回调函数 - 正确的函数签名包含depth参数
        reader.setStartElementCallback([this](const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
            this->handleStartElement(name, attributes, depth);
        });
        
        reader.setEndElementCallback([this](const std::string& name, int depth) {
            this->handleEndElement(name, depth);
        });
        
        reader.setTextCallback([this](const std::string& text, int depth) {
            this->handleText(text, depth);
        });
        
        // 执行解析并检查结果
        xml::XMLParseError result = reader.parseFromString(xml_content);
        return result == xml::XMLParseError::Ok;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("解析工作表时发生错误: {}", e.what());
        return false;
    }
}

// SAX helper methods - 将这些移到前面以便后面的方法调用
std::optional<std::string> WorksheetParser::findAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) {
    for (const auto& attr : attributes) {
        if (attr.name == name) {
            return attr.value;
        }
    }
    return std::nullopt;
}

std::optional<int> WorksheetParser::findIntAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) {
    auto value_opt = findAttribute(attributes, name);
    if (!value_opt) {
        return std::nullopt;
    }
    try {
        return std::stoi(value_opt.value());
    } catch (const std::exception& /*e*/) {
        return std::nullopt;
    }
}

std::optional<double> WorksheetParser::findDoubleAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) {
    auto value_opt = findAttribute(attributes, name);
    if (!value_opt) {
        return std::nullopt;
    }
    try {
        return std::stod(value_opt.value());
    } catch (const std::exception& /*e*/) {
        return std::nullopt;
    }
}

// Utility methods (using existing project utilities)

bool WorksheetParser::parseRangeRef(const std::string& ref, int& first_row, int& first_col, int& last_row, int& last_col) {
    size_t colon = ref.find(':');
    if (colon == std::string::npos) {
        // 单一单元格区间
        try {
            auto [r, c] = utils::CommonUtils::parseReference(ref);
            first_row = last_row = r;
            first_col = last_col = c;
            return true;
        } catch (const std::exception& /*e*/) {
            return false;
        }
    }
    std::string start_ref = ref.substr(0, colon);
    std::string end_ref = ref.substr(colon + 1);
    try {
        auto [r1, c1] = utils::CommonUtils::parseReference(start_ref);
        auto [r2, c2] = utils::CommonUtils::parseReference(end_ref);
        first_row = r1; first_col = c1; last_row = r2; last_col = c2;
        return true;
    } catch (const std::exception& /*e*/) {
        return false;
    }
}

bool WorksheetParser::isDateFormat(int style_index) const {
    if (style_index < 0) {
        return false;
    }
    
    auto it = state_.styles->find(style_index);
    if (it == state_.styles->end()) {
        return false;
    }
    
    // 检查格式是否为日期格式
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

// SAX事件处理器实现
void WorksheetParser::handleStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
    // 推入元素栈
    if (name == "worksheet") {
        state_.element_stack.push(ParseState::Element::Worksheet);
    } else if (name == "cols") {
        state_.element_stack.push(ParseState::Element::Cols);
    } else if (name == "col" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::Cols) {
        handleColumnElement(attributes);
    } else if (name == "mergeCells") {
        state_.element_stack.push(ParseState::Element::MergeCells);
    } else if (name == "mergeCell" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::MergeCells) {
        handleMergeCellElement(attributes);
    } else if (name == "sheetData") {
        state_.element_stack.push(ParseState::Element::SheetData);
    } else if (name == "row" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::SheetData) {
        handleRowStartElement(attributes);
        state_.element_stack.push(ParseState::Element::Row);
    } else if (name == "c" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::Row) {
        handleCellStartElement(attributes);
        state_.element_stack.push(ParseState::Element::Cell);
    } else if (name == "v" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::Cell) {
        state_.element_stack.push(ParseState::Element::CellValue);
        state_.current_value.clear();
    } else if (name == "f" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::Cell) {
        state_.element_stack.push(ParseState::Element::CellFormula);
        state_.current_formula.clear();
        // 处理共享公式属性 - 保持原有的共享公式支持
        auto type_opt = findAttribute(attributes, "t");
        auto si_opt = findIntAttribute(attributes, "si");
        auto ref_opt = findAttribute(attributes, "ref");
        
        if (type_opt && type_opt.value() == "shared" && si_opt && ref_opt) {
            // 这是共享公式的主定义，稍后在公式内容解析完成后处理
        }
    } else if (name == "is" && !state_.element_stack.empty() && 
               state_.element_stack.top() == ParseState::Element::Cell) {
        // 内联字符串开始 - 需要找到其中的<t>标签
    } else if (name == "t") {
        // 文本内容，可能在内联字符串<is>中
        state_.current_value.clear(); // 准备收集文本
    }
}

void WorksheetParser::handleEndElement(const std::string& name, int depth) {
    if (!state_.element_stack.empty()) {
        ParseState::Element current = state_.element_stack.top();
        state_.element_stack.pop();
        
        if (name == "c" && current == ParseState::Element::Cell) {
            // 单元格结束，处理单元格数据
            processCellData();
        } else if (name == "v" && current == ParseState::Element::CellValue) {
            // 单元格值结束
        } else if (name == "f" && current == ParseState::Element::CellFormula) {
            // 公式结束
        } else if (name == "row" && current == ParseState::Element::Row) {
            // 行结束
            state_.current_row = -1;
        }
    }
}

void WorksheetParser::handleText(const std::string& text, int depth) {
    if (!state_.element_stack.empty()) {
        ParseState::Element current = state_.element_stack.top();
        if (current == ParseState::Element::CellValue) {
            state_.current_value += text;
        } else if (current == ParseState::Element::CellFormula) {
            state_.current_formula += text;
        }
    }
}

// 私有SAX事件处理辅助方法实现
void WorksheetParser::handleColumnElement(const std::vector<xml::XMLAttribute>& attributes) {
    // 提取列属性 - 保持原有的完整列处理逻辑
    auto min_col_opt = findIntAttribute(attributes, "min");
    auto max_col_opt = findIntAttribute(attributes, "max");
    auto width_opt = findDoubleAttribute(attributes, "width");
    auto style_opt = findIntAttribute(attributes, "style");
    auto hidden_opt = findAttribute(attributes, "hidden");
    auto custom_width_opt = findAttribute(attributes, "customWidth");
    
    if (min_col_opt && max_col_opt) {
        int min_col = min_col_opt.value();
        int max_col = max_col_opt.value();
        
        // Excel列索引从1开始，转换为0开始
        int first_col = min_col - 1;
        int last_col = max_col - 1;
        
        // 设置列宽（保留Excel默认列宽，即使没有customWidth属性）
        if (width_opt) {
            double width = width_opt.value();
            // 使用新的智能列宽设置，逐列处理
            for (int col = first_col; col <= last_col; ++col) {
                state_.worksheet->setColumnWidth(col, width);
            }
            FASTEXCEL_LOG_DEBUG("设置列宽：列 {}-{} 宽度 {} custom_width={}", 
                               first_col, last_col, width, 
                               custom_width_opt ? custom_width_opt.value() : "false");
        }
        
        // 设置列样式
        if (style_opt) {
            int style_index = style_opt.value();
            // 使用样式 ID 映射来获取正确的格式
            int mapped_style_id = style_index;
            if (!state_.style_id_mapping->empty()) {
                auto mapping_it = state_.style_id_mapping->find(style_index);
                if (mapping_it != state_.style_id_mapping->end()) {
                    mapped_style_id = mapping_it->second;
                }
            }
            
            auto style_it = state_.styles->find(mapped_style_id);
            if (style_it != state_.styles->end()) {
                // 设置列格式 ID 到工作表
                state_.worksheet->setColumnFormatId(first_col, last_col, mapped_style_id);
                FASTEXCEL_LOG_DEBUG("设置列样式：列 {}-{} 原始样式ID {} 映射样式ID {}", 
                                   first_col, last_col, style_index, mapped_style_id);
            }
        }
        
        // 设置隐藏状态
        if (hidden_opt && (hidden_opt.value() == "1" || hidden_opt.value() == "true")) {
            state_.worksheet->hideColumn(first_col, last_col);
        }
    }
}

void WorksheetParser::handleMergeCellElement(const std::vector<xml::XMLAttribute>& attributes) {
    auto ref_opt = findAttribute(attributes, "ref");
    if (ref_opt) {
        std::string ref = ref_opt.value();
        int r1, c1, r2, c2;
        if (parseRangeRef(ref, r1, c1, r2, c2)) {
            state_.worksheet->mergeCells(r1, c1, r2, c2);
        }
    }
}

void WorksheetParser::handleRowStartElement(const std::vector<xml::XMLAttribute>& attributes) {
    // 先解析行级属性：r（行号）、ht（行高）、customHeight、hidden - 保持原有完整逻辑
    auto r_opt = findIntAttribute(attributes, "r");
    if (r_opt) {
        int excel_row = r_opt.value();
        state_.current_row = excel_row - 1; // 转为0基
        
        auto ht_opt = findDoubleAttribute(attributes, "ht");
        auto custom_height_opt = findAttribute(attributes, "customHeight");
        auto hidden_opt = findAttribute(attributes, "hidden");
        
        if (ht_opt) {
            double ht = ht_opt.value();
            if (ht > 0 && (custom_height_opt && 
                          (custom_height_opt.value() == "1" || custom_height_opt.value() == "true" || !custom_height_opt.value().empty()))) {
                // 只有明确设置自定义行高时才设置；部分文件也会直接提供ht且customHeight缺省，兼容性处理：只要有ht就设置
                state_.worksheet->setRowHeight(state_.current_row, ht);
            } else if (ht > 0) {
                state_.worksheet->setRowHeight(state_.current_row, ht);
            }
        }
        
        if (hidden_opt && (hidden_opt.value() == "1" || hidden_opt.value() == "true")) {
            state_.worksheet->hideRow(state_.current_row);
        }
    }
}

void WorksheetParser::handleCellStartElement(const std::vector<xml::XMLAttribute>& attributes) {
    // 初始化当前单元格状态
    state_.current_col = -1;
    state_.current_cell_ref.clear();
    state_.current_cell_type.clear();
    state_.current_style_id = -1;
    state_.current_value.clear();
    state_.current_formula.clear();
    
    // 提取单元格引用 (r="A1")
    auto ref_opt = findAttribute(attributes, "r");
    if (ref_opt) {
        state_.current_cell_ref = ref_opt.value();
        try {
            auto [row, col] = utils::CommonUtils::parseReference(state_.current_cell_ref);
            state_.current_col = col;
        } catch (const std::exception& /*e*/) {
            // 解析失败，使用列属性
            if (auto col_opt = findIntAttribute(attributes, "c")) {
                state_.current_col = col_opt.value();
            }
        }
    }
    
    // 提取单元格类型
    auto type_opt = findAttribute(attributes, "t");
    if (type_opt) {
        state_.current_cell_type = type_opt.value();
    }
    
    // 提取样式索引
    auto style_opt = findIntAttribute(attributes, "s");
    if (style_opt) {
        state_.current_style_id = style_opt.value();
    }
}

// 单元格处理方法实现
void WorksheetParser::processCellData() {
    if (state_.current_row < 0 || state_.current_col < 0) {
        return;
    }
    
    int row = state_.current_row;
    int col = state_.current_col;
    
    // 根据类型设置单元格值 - 保持原有的完整类型处理逻辑
    if (state_.current_cell_type == "s") {
        // 共享字符串
        try {
            int string_index = std::stoi(state_.current_value);
            auto it = state_.shared_strings->find(string_index);
            if (it != state_.shared_strings->end()) {
                // 使用 addSharedStringWithIndex 保持原始索引
                if (auto wb = state_.worksheet->getParentWorkbook()) {
                    wb->addSharedStringWithIndex(it->second, string_index);
                }
                state_.worksheet->setValue(row, col, it->second);
            }
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("解析共享字符串索引失败: {}", e.what());
        }
    } else if (state_.current_cell_type == "inlineStr") {
        // 内联字符串
        std::string decoded_value = utils::XMLUtils::unescapeXML(state_.current_value);
        state_.worksheet->setValue(row, col, decoded_value);
    } else if (state_.current_cell_type == "b") {
        // 布尔值
        bool bool_value = (state_.current_value == "1" || state_.current_value == "true");
        state_.worksheet->setValue(row, col, bool_value);
    } else if (state_.current_cell_type == "str") {
        // 公式字符串结果
        std::string decoded_value = utils::XMLUtils::unescapeXML(state_.current_value);
        state_.worksheet->setValue(row, col, decoded_value);
    } else if (state_.current_cell_type == "e") {
        // 错误值
        std::string decoded_value = utils::XMLUtils::unescapeXML(state_.current_value);
        state_.worksheet->setValue(row, col, "#ERROR: " + decoded_value);
    } else if (state_.current_cell_type == "d") {
        // 日期值（ISO 8601格式）
        std::string decoded_value = utils::XMLUtils::unescapeXML(state_.current_value);
        state_.worksheet->setValue(row, col, decoded_value);
    } else {
        // 数字或默认类型 - Excel中没有t属性的单元格默认是数字！
        if (!state_.current_value.empty()) {
            try {
                // 解析为数字 - 这是最常见的情况，应该优先处理
                double number_value = std::stod(state_.current_value);
                
                // 重要：日期在Excel中也是数字，不要转换为字符串！
                // 日期格式应该由样式来控制显示，而不是改变数据类型
                state_.worksheet->setValue(row, col, number_value);
                
            } catch (const std::exception& /*e*/) {
                // 极少数情况：如果真的不能解析为数字（可能是Excel文件损坏）
                // 记录警告并跳过该单元格，而不是错误地转换为字符串
                FASTEXCEL_LOG_WARN("单元格{}无法解析为数字，值='{}'", 
                           utils::CommonUtils::cellReference(row, col), state_.current_value);
                // 不设置任何值，保持单元格为空
            }
        } else if (!state_.current_formula.empty()) {
            // 只有公式没有值的情况
            state_.worksheet->setFormula(row, col, state_.current_formula);
        }
    }
    
    // 应用样式（如果有）
    if (state_.current_style_id >= 0) {
        // 使用样式 ID 映射来获取正确的 FormatRepository 中的样式
        int mapped_style_id = state_.current_style_id;
        if (!state_.style_id_mapping->empty()) {
            auto mapping_it = state_.style_id_mapping->find(state_.current_style_id);
            if (mapping_it != state_.style_id_mapping->end()) {
                mapped_style_id = mapping_it->second;
            }
        }
        
        auto style_it = state_.styles->find(mapped_style_id);
        if (style_it != state_.styles->end()) {
            auto& cell = state_.worksheet->getCell(row, col);
            cell.setFormat(style_it->second);
        }
    }
    
    // 处理公式
    if (!state_.current_formula.empty()) {
        state_.worksheet->setFormula(row, col, utils::XMLUtils::unescapeXML(state_.current_formula));
    }
}

void WorksheetParser::processColumnDefinition() {
    // 列定义处理已在handleColumnElement中完成
}

void WorksheetParser::processMergeCell(const std::vector<xml::XMLAttribute>& attributes) {
    auto ref_opt = findAttribute(attributes, "ref");
    if (ref_opt) {
        std::string ref = ref_opt.value();
        int r1, c1, r2, c2;
        if (parseRangeRef(ref, r1, c1, r2, c2)) {
            state_.worksheet->mergeCells(r1, c1, r2, c2);
        }
    }
}

} // namespace reader
} // namespace fastexcel