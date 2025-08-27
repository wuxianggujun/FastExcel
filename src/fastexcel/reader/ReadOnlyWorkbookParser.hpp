/**
 * @file ReadOnlyWorkbookParser.hpp
 * @brief 只读模式专用工作簿XML解析器
 */

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include <vector>
#include <string>
#include <unordered_map>

namespace fastexcel {
namespace reader {

/**
 * @brief 只读模式工作表信息
 */
struct ReadOnlySheetInfo {
    std::string name;
    std::string rel_id;
    std::string worksheet_path;
    
    ReadOnlySheetInfo(std::string n, std::string rid) 
        : name(std::move(n)), rel_id(std::move(rid)) {}
};

/**
 * @brief 专用于只读模式的工作簿XML流式解析器
 */
class ReadOnlyWorkbookParser : public BaseSAXParser {
private:
    std::vector<ReadOnlySheetInfo> sheets_;
    std::unordered_map<std::string, std::string> relationships_;
    
    // 解析状态
    bool in_sheets_section_ = false;
    bool in_sheet_element_ = false;
    
protected:
    // 重写SAX事件处理方法
    void onStartElement(std::string_view name, 
                       span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
    void onText(std::string_view data, int depth) override;
    
public:
    ReadOnlyWorkbookParser() = default;
    ~ReadOnlyWorkbookParser() = default;
    
    /**
     * @brief 设置关系映射（从workbook.xml.rels解析）
     */
    void setRelationships(std::unordered_map<std::string, std::string> relationships) {
        relationships_ = std::move(relationships);
    }
    
    /**
     * @brief 获取解析后的工作表信息
     */
    const std::vector<ReadOnlySheetInfo>& getSheets() const {
        return sheets_;
    }
    
    /**
     * @brief 移动工作表信息所有权
     */
    std::vector<ReadOnlySheetInfo> takeSheets() {
        return std::move(sheets_);
    }
    
    /**
     * @brief 清空解析状态
     */
    void reset() {
        state_.reset();
        sheets_.clear();
        in_sheets_section_ = false;
        in_sheet_element_ = false;
    }
};

}} // namespace fastexcel::reader