#pragma once

#include "fastexcel/core/columnar/ReadOnlyWorkbook.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/reader/WorksheetParser.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include <string>
#include <memory>
#include <unordered_set>
#include <unordered_map>

namespace fastexcel {
namespace reader {

// 列式工作表解析器 - 专为只读模式优化
class ColumnarWorksheetParser {
private:
    struct ParseState {
        core::columnar::ReadOnlyWorksheet* worksheet = nullptr;
        const std::unordered_map<int, std::string>* shared_strings = nullptr;
        const core::columnar::ReadOnlyOptions* options = nullptr;
        
        // 当前解析状态
        bool in_sheet_data = false;
        bool in_row = false;
        uint32_t current_row = 0;
        
        // 列投影过滤
        std::unordered_set<uint32_t> projected_columns;
        bool has_column_filter = false;
        
        // 行数限制
        uint32_t max_rows = 0;
        bool has_row_limit = false;
        
        void reset() {
            worksheet = nullptr;
            shared_strings = nullptr;
            options = nullptr;
            in_sheet_data = false;
            in_row = false;
            current_row = 0;
            projected_columns.clear();
            has_column_filter = false;
            max_rows = 0;
            has_row_limit = false;
        }
        
        void setupProjection(const core::columnar::ReadOnlyOptions& opts) {
            options = &opts;
            
            if (!opts.projected_columns.empty()) {
                has_column_filter = true;
                for (uint32_t col : opts.projected_columns) {
                    projected_columns.insert(col);
                }
            }
            
            if (opts.max_rows > 0) {
                has_row_limit = true;
                max_rows = opts.max_rows;
            }
        }
        
        bool shouldSkipColumn(uint32_t col) const {
            return has_column_filter && projected_columns.find(col) == projected_columns.end();
        }
        
        bool shouldSkipRow(uint32_t row) const {
            return has_row_limit && row >= max_rows;
        }
    };
    
    ParseState state_;
    
public:
    ColumnarWorksheetParser() = default;
    ~ColumnarWorksheetParser() = default;
    
    // 流式解析 - 直接输出到列式存储
    bool parseToColumnar(archive::ZipReader* zip_reader,
                        const std::string& internal_path,
                        core::columnar::ReadOnlyWorksheet* worksheet,
                        const std::unordered_map<int, std::string>& shared_strings,
                        const core::columnar::ReadOnlyOptions& options = {});
    
    // 内存解析 - 用于小文件
    bool parseToColumnar(const std::string& xml_content,
                        core::columnar::ReadOnlyWorksheet* worksheet,
                        const std::unordered_map<int, std::string>& shared_strings,
                        const core::columnar::ReadOnlyOptions& options = {});
                        
private:
    // XML解析回调
    void handleStartElement(std::string_view name, 
                           xml::span<const xml::XMLAttribute> attributes, 
                           int depth);
    void handleEndElement(std::string_view name, int depth);
    void handleText(std::string_view text, int depth);
    
    // 单元格数据处理
    void processCellElement(xml::span<const xml::XMLAttribute> attributes);
    void processCellValue(std::string_view value, std::string_view cell_type, uint32_t row, uint32_t col);
    
    // 简化的解析方法 - 用于快速实现
    void parseSheetDataSimple(const std::string& sheet_data);
    void parseCellsInRow(const std::string& row_data, uint32_t row);
    void processCellValue(const std::string& value, const std::string& cell_type, uint32_t row, uint32_t col);
    
    // 辅助方法
    uint32_t parseColumnReference(std::string_view cell_ref) const;
    uint32_t parseColumnReference(const std::string& cell_ref) const;
    uint32_t parseRowReference(std::string_view cell_ref) const;
    uint32_t parseRowReference(const std::string& cell_ref) const;
    bool isProjectedColumn(uint32_t col) const;
};

}} // namespace fastexcel::reader