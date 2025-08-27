//
// Created by wuxianggujun on 25-8-4.
//

#include "WorksheetParser.hpp"

#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <fast_float/fast_float.h>
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/xml/XMLEscapes.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <cstring>
#include <algorithm>

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
    
    try {
        // 初始化解析状态
        state_.reset();
        state_.worksheet = worksheet;
        state_.shared_strings = &shared_strings;
        state_.styles = &styles;
        state_.style_id_mapping = &style_id_mapping;
        
        // 使用混合架构：一次SAX解析处理所有内容
        // 这样避免多次解析同一个XML文档的开销
        return parseXML(xml_content);
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("WorksheetParser解析失败: {}", e.what());
        return false;
    }
}

// 流式解析实现：使用 XMLStreamReader 的回调收集行级 XML，从 ZipReader 逐块喂给解析器
bool WorksheetParser::parseStream(archive::ZipReader* zip_reader,
                                  const std::string& internal_path,
                                  core::Worksheet* worksheet,
                                  const std::unordered_map<int, std::string>& shared_strings,
                                  const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                                  const std::unordered_map<int, int>& style_id_mapping,
                                  const core::WorkbookOptions* options) {
    if (!zip_reader || !worksheet) {
        return false;
    }

    // 初始化解析状态
    state_.reset();
    state_.worksheet = worksheet;
    state_.shared_strings = &shared_strings;
    state_.styles = &styles;
    state_.style_id_mapping = &style_id_mapping;
    state_.options = options;
    
    // 设置列式优化选项
    state_.setupColumnarOptions();

    xml::XMLStreamReader reader;

    reader.setStartElementCallback([this](std::string_view name, span<const xml::XMLAttribute> attributes, int depth) {
        this->handleStartElement(name, attributes, depth);
    });
    reader.setEndElementCallback([this](std::string_view name, int depth) {
        this->handleEndElement(name, depth);
    });
    reader.setTextCallback([this](std::string_view text, int depth) {
        this->handleText(text, depth);
    });
    reader.setErrorCallback([this](xml::XMLParseError /*err*/, const std::string& msg, int line, int col) {
        FASTEXCEL_LOG_ERROR("Worksheet stream parse error at {}:{} -> {}", line, col, msg);
    });

    if (reader.beginParsing() != xml::XMLParseError::Ok) {
        return false;
    }

    auto err = zip_reader->streamFile(internal_path,
        [&reader](const uint8_t* data, size_t size) -> bool {
            if (reader.feedData(reinterpret_cast<const char*>(data), size) != xml::XMLParseError::Ok) {
                return false;
            }
            return true;
        },
        1 << 16 // 64KB
    );

    if (archive::isError(err)) {
        FASTEXCEL_LOG_ERROR("Zip stream failed: {}", internal_path);
        return false;
    }

    if (reader.endParsing() != xml::XMLParseError::Ok) {
        return false;
    }

    return true;
}

// === 完整的混合架构实现：处理所有工作表元素 ===
void WorksheetParser::onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int /*depth*/) {
    // 列定义处理 - 影响整个工作表的列格式和宽度
    if (name == "cols") {
        // 进入列定义区域
    }
    else if (name == "col") {
        handleColumnElement(attributes);
    }
    
    // 合并单元格处理
    else if (name == "mergeCells") {
        // 进入合并单元格区域
    }
    else if (name == "mergeCell") {
        handleMergeCellElement(attributes);
    }
    
    // 工作表数据处理 - 这里使用混合架构优化
    else if (name == "sheetData") {
        state_.in_sheet_data = true;
    }
    else if (name == "row" && state_.in_sheet_data) {
        // 行开始：初始化行处理状态
        state_.in_row = true;
        state_.row_buffer.clear();
        
        // 提取行号和行属性
        handleRowStartElement(attributes);
        
        // 高性能XML重构：使用预分配缓冲区避免重分配
        state_.format_buffer.clear();
        state_.format_buffer.reserve(512); // 预估单行XML大小
        
        state_.format_buffer = "<row";
        for (const auto& attr : attributes) {
            state_.format_buffer += " ";
            state_.format_buffer.append(attr.name.data(), attr.name.size());
            state_.format_buffer += "=\"";
            state_.format_buffer.append(attr.value.data(), attr.value.size());
            state_.format_buffer += "\"";
        }
        state_.format_buffer += ">";
        
        // 移动语义避免拷贝
        state_.row_xml_buffer = std::move(state_.format_buffer);
    }
    else if (state_.in_row) {
        // 在行内的任何元素：优化字符串拼接
        state_.row_xml_buffer += "<";
        state_.row_xml_buffer.append(name.data(), name.size());
        
        for (const auto& attr : attributes) {
            state_.row_xml_buffer += " ";
            state_.row_xml_buffer.append(attr.name.data(), attr.name.size());
            state_.row_xml_buffer += "=\"";
            state_.row_xml_buffer.append(attr.value.data(), attr.value.size());
            state_.row_xml_buffer += "\"";
        }
        state_.row_xml_buffer += ">";
    }
    
    // 其他非关键元素忽略 - 关键优化：减少SAX回调处理
}

void WorksheetParser::onEndElement(std::string_view name, int /*depth*/) {
    if (name == "row" && state_.in_row) {
        // 行结束：完成XML收集，进行指针扫描解析
        state_.row_xml_buffer += "</row>";
        
        if (state_.current_row >= 0) {
            // 使用指针扫描解析整行数据
            parseRowWithPointerScan(state_.row_xml_buffer, state_.row_buffer);
            // 批量处理单元格数据
            processBatchCellData(state_.current_row, state_.row_buffer);
        }
        
        state_.in_row = false;
        state_.current_row = -1;
    }
    else if (state_.in_row) {
        // 在行内的结束标签也要收集：优化拼接
        state_.row_xml_buffer += "</";
        state_.row_xml_buffer.append(name.data(), name.size());
        state_.row_xml_buffer += ">";
    }
    else if (name == "sheetData") {
        state_.in_sheet_data = false;
    }
}

void WorksheetParser::onText(std::string_view text, int /*depth*/) {
    if (state_.in_row && !text.empty()) {
        // 在行内：直接收集文本内容用于指针扫描，避免拷贝
        state_.row_xml_buffer.append(text.data(), text.size());
    }
}

// === 修正的核心优化：零分配指针扫描器 ===
void WorksheetParser::parseRowWithPointerScan(std::string_view row_xml, std::vector<FastCellData>& cells) {
    cells.clear();
    
    const char* p = row_xml.data();
    const char* end = p + row_xml.size();
    
    // 扫描所有 <c 开始标签
    while (p < end) {
        // 查找下一个 <c 开始
        const char* c_tag = std::strstr(p, "<c ");
        if (!c_tag || c_tag >= end) {
            // 也检查自闭合标签 <c .../>
            c_tag = std::strstr(p, "<c>");
            if (!c_tag || c_tag >= end) break;
        }
        
        FastCellData cell;
        if (extractCellInfo(c_tag, end, cell)) {
            cells.push_back(cell);
        }
        
        // 移动到下一个位置
        p = c_tag + 1;
    }
}

// 修正的提取单元格信息的核心函数
bool WorksheetParser::extractCellInfo(const char*& p, const char* end, FastCellData& cell) {
    // 跳过 "<c"
    p += 2;
    
    // 解析属性直到 '>' 或 '/>'
    while (p < end && *p != '>' && !(p[0] == '/' && p[1] == '>')) {
        // 跳过空白字符
        while (p < end && (*p == ' ' || *p == '\t' || *p == '\n' || *p == '\r')) p++;
        
        if (p >= end) break;
        
        // 检查 r= 属性 (单元格引用)
        if (p + 3 < end && p[0] == 'r' && p[1] == '=' && p[2] == '"') {
            p += 3; // 跳过 'r="'
            const char* ref_start = p;
            while (p < end && *p != '"') p++;
            if (p < end) {
                // 使用CommonUtils解析单元格引用
                try {
                    std::string ref(ref_start, p - ref_start);
                    auto [row, col] = utils::CommonUtils::parseReference(ref);
                    cell.col = static_cast<uint32_t>(col);
                } catch (...) {
                    cell.col = UINT32_MAX; // 解析失败
                }
                p++; // 跳过结束引号
            }
        }
        // 检查 t= 属性 (单元格类型)
        else if (p + 3 < end && p[0] == 't' && p[1] == '=' && p[2] == '"') {
            p += 3; // 跳过 't="'
            if (p < end) {
                char type_char = *p;
                if (type_char == 's') {
                    cell.type = FastCellData::SharedString;
                } else if (type_char == 'b') {
                    cell.type = FastCellData::Boolean;
                } else if (type_char == 's' && p + 3 < end && std::strncmp(p, "str", 3) == 0) {
                    cell.type = FastCellData::String;
                }
                // 跳到引号结束
                while (p < end && *p != '"') p++;
                if (p < end) p++; // 跳过结束引号
            }
        }
        // 检查 s= 属性 (样式索引)
        else if (p + 3 < end && p[0] == 's' && p[1] == '=' && p[2] == '"') {
            p += 3; // 跳过 's="'
            const char* style_start = p;
            int style_id = 0;
            while (p < end && *p >= '0' && *p <= '9') {
                style_id = style_id * 10 + (*p - '0');
                p++;
            }
            if (p > style_start && p < end && *p == '"') {
                cell.style_id = style_id;
                p++; // 跳过结束引号
            }
        }
        else {
            p++;
        }
    }
    
    if (p >= end) return false;
    
    // 检查是否为自闭合标签
    if (p[0] == '/' && p[1] == '>') {
        // 自闭合标签，没有值
        return cell.col != UINT32_MAX; // 至少要有有效的列号
    }
    
    p++; // 跳过 '>'
    
    // 查找 <v> 值或 <is><t>值
    const char* v_start = std::strstr(p, "<v>");
    if (v_start && v_start < end) {
        v_start += 3; // 跳过 "<v>"
        const char* v_end = std::strstr(v_start, "</v>");
        if (v_end && v_end < end) {
            cell.value = std::string_view(v_start, v_end - v_start);
            cell.is_empty = false;
        }
    } else {
        // 查找内联字符串 <is><t>...</t></is>
        const char* is_start = std::strstr(p, "<is>");
        if (is_start && is_start < end) {
            const char* t_start = std::strstr(is_start, "<t>");
            if (t_start && t_start < end) {
                t_start += 3; // 跳过 "<t>"
                const char* t_end = std::strstr(t_start, "</t>");
                if (t_end && t_end < end) {
                    cell.value = std::string_view(t_start, t_end - t_start);
                    cell.is_empty = false;
                    cell.type = FastCellData::String;
                }
            }
        }
    }
    
    return !cell.is_empty || cell.col != UINT32_MAX;
}

// 批量处理单元格数据 - 减少worksheet调用开销
void WorksheetParser::processBatchCellData(int row, const std::vector<FastCellData>& cells) {
    // 列式优化：行过滤
    if (state_.shouldSkipRow(row)) {
        FASTEXCEL_LOG_DEBUG("Skipping row {} due to row limit", row);
        return;
    }
    
    for (const auto& cell_data : cells) {
        // 列式优化：列过滤
        if (state_.shouldSkipColumn(cell_data.col)) {
            continue;
        }
        setCellValue(row, cell_data.col, cell_data);
    }
}

// 设置单元格值 - 现在只使用传统Cell对象方式（编辑模式）
void WorksheetParser::setCellValue(int row, uint32_t col, const FastCellData& cell_data) {
    if (cell_data.is_empty || cell_data.col == UINT32_MAX) return;
    
    // **编辑模式：使用setValue方法设置值**
    switch (cell_data.type) {
        case FastCellData::SharedString: {
            try {
                int string_index = 0;
                for (char c : cell_data.value) {
                    if (c >= '0' && c <= '9') {
                        string_index = string_index * 10 + (c - '0');
                    }
                }
                // 使用setValue设置SST索引
                state_.worksheet->setValue(row, col, string_index);
            } catch (...) {
                // 忽略错误
            }
            break;
        }
        
        case FastCellData::Number: {
            try {
                // 使用 fast_float 高性能解析数值
                double number_value;
                auto result = fast_float::from_chars(cell_data.value.data(),
                                                   cell_data.value.data() + cell_data.value.size(),
                                                   number_value);
                if (result.ec == std::errc{}) {
                    // 使用setValue设置数值
                    state_.worksheet->setValue(row, col, number_value);
                } else {
                    // 解析失败，设置为错误值
                    std::string error_value = "#VALUE!";
                    state_.worksheet->setValue(row, col, error_value);
                }
            } catch (...) {
                // 解析失败，设置为字符串
                std::string error_value = "#VALUE!";
                state_.worksheet->setValue(row, col, error_value);
            }
            break;
        }
        
        case FastCellData::Boolean: {
            bool bool_value = (cell_data.value == "1" || cell_data.value == "true");
            // 使用setValue设置布尔值
            state_.worksheet->setValue(row, col, bool_value);
            break;
        }
        
        case FastCellData::String:
        default: {
            // 内联字符串使用setValue设置
            std::string str_value = decodeXMLEntities(std::string(cell_data.value));
            state_.worksheet->setValue(row, col, str_value);
            break;
        }
    }
    
    // 编辑模式下保留样式设置
}

// 工具方法

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
    
    // 使用TimeUtils格式化
    std::tm time_info;
#ifdef _WIN32
    gmtime_s(&time_info, &unix_time);
#else
    time_info = *gmtime(&unix_time);
#endif
    
    return utils::TimeUtils::formatTime(time_info, "%Y-%m-%d");
}

// === 完整的工作表元素处理方法实现 ===

void WorksheetParser::handleColumnElement(span<const xml::XMLAttribute> attributes) {
    // 提取列属性 - 完整的列处理逻辑
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
        
        // 设置列宽
        if (width_opt) {
            double width = width_opt.value();
            for (int col = first_col; col <= last_col; ++col) {
                state_.worksheet->setColumnWidth(col, width);
            }
            FASTEXCEL_LOG_DEBUG("设置列宽：列 {}-{} 宽度 {}", first_col, last_col, width);
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
                FASTEXCEL_LOG_DEBUG("设置列样式：列 {}-{} 样式ID {}", first_col, last_col, mapped_style_id);
            }
        }
        
        // 设置隐藏状态
        if (hidden_opt && (hidden_opt.value() == "1" || hidden_opt.value() == "true")) {
            state_.worksheet->hideColumn(first_col, last_col);
        }
    }
}

void WorksheetParser::handleMergeCellElement(span<const xml::XMLAttribute> attributes) {
    auto ref_opt = findAttribute(attributes, "ref");
    if (ref_opt) {
        std::string ref = ref_opt.value();
        int r1, c1, r2, c2;
        if (parseRangeReference(ref, r1, c1, r2, c2)) {
            state_.worksheet->mergeCells(r1, c1, r2, c2);
            FASTEXCEL_LOG_DEBUG("合并单元格：{} -> ({},{}) - ({},{})", ref, r1, c1, r2, c2);
        }
    }
}

void WorksheetParser::handleRowStartElement(span<const xml::XMLAttribute> attributes) {
    // 解析行级属性：r（行号）、ht（行高）、customHeight、hidden
    auto r_opt = findIntAttribute(attributes, "r");
    if (r_opt) {
        int excel_row = r_opt.value();
        state_.current_row = excel_row - 1; // 转为0基
        
        auto ht_opt = findDoubleAttribute(attributes, "ht");
        auto custom_height_opt = findAttribute(attributes, "customHeight");
        auto hidden_opt = findAttribute(attributes, "hidden");
        
        // 设置行高
        if (ht_opt) {
            double ht = ht_opt.value();
            if (ht > 0) {
                state_.worksheet->setRowHeight(state_.current_row, ht);
                FASTEXCEL_LOG_DEBUG("设置行高：行 {} 高度 {}", state_.current_row, ht);
            }
        }
        
        // 设置隐藏状态
        if (hidden_opt && (hidden_opt.value() == "1" || hidden_opt.value() == "true")) {
            state_.worksheet->hideRow(state_.current_row);
            FASTEXCEL_LOG_DEBUG("隐藏行：{}", state_.current_row);
        }
    }
}

} // namespace reader
} // namespace fastexcel
