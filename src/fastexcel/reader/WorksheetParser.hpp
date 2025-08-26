#include "fastexcel/utils/Logger.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/core/RangeFormatter.hpp"
#include <string>
#include <unordered_map>
#include <memory>
#include <stack>
#include <optional>

namespace fastexcel {
namespace reader {

/**
 * @brief 高性能SAX流式工作表解析器
 *
 * 重写为基于libexpat的SAX流式解析，彻底解决字符串查找性能瓶颈：
 * - 使用XMLStreamReader进行真正的流式解析
 * - 消除所有find/substr字符串操作
 * - 减少临时对象创建，大幅降低内存开销
 * - 支持600万+单元格的高效解析
 */
class WorksheetParser {
public:
    WorksheetParser() = default;
    ~WorksheetParser() = default;
    
    /**
     * @brief 解析工作表XML内容 - SAX流式版本
     * @param xml_content XML内容
     * @param worksheet 目标工作表对象
     * @param shared_strings 共享字符串映射
     * @param styles 样式映射
     * @param style_id_mapping 样式ID映射（原始ID -> FormatRepository ID）
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content,
               core::Worksheet* worksheet,
               const std::unordered_map<int, std::string>& shared_strings,
               const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
               const std::unordered_map<int, int>& style_id_mapping = {});

private:
    // SAX事件处理状态
    struct ParseState {
        core::Worksheet* worksheet = nullptr;
        const std::unordered_map<int, std::string>* shared_strings = nullptr;
        const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>* styles = nullptr;
        const std::unordered_map<int, int>* style_id_mapping = nullptr;
        
        // 当前解析状态
        enum class Element {
            None,
            Worksheet,
            SheetData, 
            Row,
            Cell,
            CellValue,
            CellFormula,
            Cols,
            Col,
            MergeCells,
            MergeCell
        };
        
        std::stack<Element> element_stack;
        
        // 当前行/单元格状态
        int current_row = -1;
        int current_col = -1;
        std::string current_cell_ref;
        std::string current_cell_type;
        int current_style_id = -1;
        
        // 当前单元格数据收集
        std::string current_value;
        std::string current_formula;
        
        // 列定义状态
        int col_min = -1;
        int col_max = -1;
        double col_width = -1.0;
        bool col_hidden = false;
        int col_style = -1;
        
        // 性能优化：预分配容器
        std::vector<core::Cell> row_cells_buffer; // 重用的行单元格缓冲区
        
        void reset() {
            while (!element_stack.empty()) element_stack.pop();
            current_row = -1;
            current_col = -1;
            current_cell_ref.clear();
            current_cell_type.clear();
            current_style_id = -1;
            current_value.clear();
            current_formula.clear();
            row_cells_buffer.clear();
        }
    } state_;
    
    // SAX事件处理器
    void handleStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth);
    void handleEndElement(const std::string& name, int depth);
    void handleText(const std::string& text, int depth);
    
    // 私有SAX事件处理辅助方法
    void handleColumnElement(const std::vector<xml::XMLAttribute>& attributes);
    void handleMergeCellElement(const std::vector<xml::XMLAttribute>& attributes);
    void handleRowStartElement(const std::vector<xml::XMLAttribute>& attributes);
    void handleCellStartElement(const std::vector<xml::XMLAttribute>& attributes);
    
    // 属性解析优化版本 - 直接从vector<XMLAttribute>提取
    std::optional<std::string> findAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name);
    std::optional<int> findIntAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name);
    std::optional<double> findDoubleAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name);
    
    // 单元格处理
    void processCellData();
    void processColumnDefinition();
    void processMergeCell(const std::vector<xml::XMLAttribute>& attributes);
    
    // 工具方法 - 保留必要的转换功能
    bool parseRangeRef(const std::string& ref, int& first_row, int& first_col, int& last_row, int& last_col);
    
    // 字符串转换优化
    bool isDateFormat(int style_index) const;
    std::string convertExcelDateToString(double excel_date);
};

} // namespace reader
} // namespace fastexcel

