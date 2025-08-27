//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"  // 添加WorkbookOptions支持
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/span.hpp"
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>  // 添加unordered_set支持
#include <memory>
#include <vector>
#include <optional>
#include <cstring>

using fastexcel::core::span;  // Import span into this namespace

namespace fastexcel {
namespace reader {

/**
 * @brief 极高性能混合架构工作表解析器
 * 
 * 采用分层处理策略，解决SAX回调风暴问题：
 * 1. Expat SAX 只处理 <row> 级别 (减少回调次数1000倍)
 * 2. 行内使用零分配指针扫描处理 <c> 元素
 * 3. 批量写入整行数据，减少worksheet调用开销
 * 
 * 预期性能提升：从25秒优化到6秒以下 (76%+)
 */
class WorksheetParser : public BaseSAXParser {
public:
    WorksheetParser() = default;
    ~WorksheetParser() = default;
    
    /**
     * @brief 解析工作表XML内容 - 混合架构版本
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

    /**
     * @brief 流式解析工作表XML内容（避免一次性读入大字符串）
     * @param zip_reader ZIP读取器
     * @param internal_path ZIP内部路径，如"xl/worksheets/sheet1.xml"
     * @param worksheet 目标工作表对象
     * @param shared_strings 共享字符串映射
     * @param styles 样式映射
     * @param style_id_mapping 样式ID映射（原始ID -> FormatRepository ID）
     * @param options 工作簿选项（用于列式优化等）
     * @return 是否解析成功
     */
    bool parseStream(archive::ZipReader* zip_reader,
                     const std::string& internal_path,
                     core::Worksheet* worksheet,
                     const std::unordered_map<int, std::string>& shared_strings,
                     const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>& styles,
                     const std::unordered_map<int, int>& style_id_mapping = {},
                     const core::WorkbookOptions* options = nullptr);

private:
    // 高性能单元格数据结构（避免Cell对象开销）
    struct FastCellData {
        uint32_t col = UINT32_MAX;  // 使用最大值表示无效列
        std::string_view value;
        enum Type { Number, String, SharedString, Boolean, Formula, Error } type = Number;
        int style_id = -1;
        bool is_empty = true;
        
        FastCellData() = default;
        FastCellData(uint32_t c, std::string_view v, Type t, int s = -1) 
            : col(c), value(v), type(t), style_id(s), is_empty(false) {}
    };
    
    // 解析状态（简化版，只关注关键信息）
    struct ParseState {
        core::Worksheet* worksheet = nullptr;
        const std::unordered_map<int, std::string>* shared_strings = nullptr;
        const std::unordered_map<int, std::shared_ptr<core::FormatDescriptor>>* styles = nullptr;
        const std::unordered_map<int, int>* style_id_mapping = nullptr;
        const core::WorkbookOptions* options = nullptr;  // 列式优化选项
        
        // 当前行状态
        int current_row = -1;
        bool in_sheet_data = false;
        bool in_row = false;
        
        // 列式优化过滤器
        bool has_column_filter = false;
        bool has_row_limit = false;
        std::unordered_set<uint32_t> projected_columns;
        uint32_t max_rows = 0;
        
        // 行级缓冲区（批量处理）
        std::vector<FastCellData> row_buffer;
        std::string row_xml_buffer;  // 收集完整行XML
        
        // 高性能格式化缓冲区
        mutable std::string format_buffer;  // 复用的格式化缓冲区
        
        void reset() {
            current_row = -1;
            in_sheet_data = false;
            in_row = false;
            row_buffer.clear();
            row_xml_buffer.clear();
            // 列式优化过滤器不重置，因为它们在整个解析过程中保持不变
        }
        
        void setupColumnarOptions() {
            if (options && options->enable_columnar_storage) {
                if (!options->projected_columns.empty()) {
                    has_column_filter = true;
                    for (uint32_t col : options->projected_columns) {
                        projected_columns.insert(col);
                    }
                }
                
                if (options->max_rows > 0) {
                    has_row_limit = true;
                    max_rows = options->max_rows;
                }
            }
        }
        
        bool shouldSkipColumn(uint32_t col) const {
            return has_column_filter && projected_columns.find(col) == projected_columns.end();
        }
        
        bool shouldSkipRow(int row) const {
            return has_row_limit && row >= static_cast<int>(max_rows);
        }
    } state_;
    
    // SAX事件处理（只处理行级事件）
    void onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
    void onText(std::string_view text, int depth) override;
    
    // 核心优化：零分配指针扫描器
    void parseRowWithPointerScan(std::string_view row_xml, std::vector<FastCellData>& cells);
    
    // 快速工具函数
    static bool extractCellInfo(const char*& p, const char* end, FastCellData& cell);
    
    // 批量数据处理
    void processBatchCellData(int row, const std::vector<FastCellData>& cells);
    void setCellValue(int row, uint32_t col, const FastCellData& cell_data);
    
    // 完整的工作表元素处理方法
    void handleColumnElement(span<const xml::XMLAttribute> attributes);
    void handleMergeCellElement(span<const xml::XMLAttribute> attributes);
    void handleRowStartElement(span<const xml::XMLAttribute> attributes);
    
    // 工具方法
    bool isDateFormat(int style_index) const;
    std::string convertExcelDateToString(double excel_date);
};

} // namespace reader
} // namespace fastexcel
