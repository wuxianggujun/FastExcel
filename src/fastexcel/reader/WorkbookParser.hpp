/**
 * @file WorkbookParser.hpp
 * @brief 通用工作簿XML解析器
 */

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/span.hpp"
#include <string>
#include <vector>
#include <unordered_map>

/**
 * @brief 定义名称信息结构
 */
struct DefinedNameInfo {
    std::string name;           // 名称
    std::string formula;        // 公式/引用内容
    std::string local_sheet_id; // 本地工作表ID (可选)
    std::string comment;        // 注释 (可选)
    bool hidden = false;        // 是否隐藏
    
    DefinedNameInfo(std::string n) : name(std::move(n)) {}
};

using fastexcel::core::span;

namespace fastexcel {
namespace reader {

/**
 * @brief 工作表信息结构
 */
struct WorksheetInfo {
    std::string name;
    std::string sheet_id;
    std::string rel_id;
    std::string worksheet_path;
    
    WorksheetInfo(std::string n, std::string sid, std::string rid) 
        : name(std::move(n)), sheet_id(std::move(sid)), rel_id(std::move(rid)) {}
};

/**
 * @brief 通用工作簿XML流式解析器
 * 用于完整功能的XLSXReader，提供比ReadOnlyWorkbookParser更完整的功能
 */
class WorkbookParser : public BaseSAXParser {
private:
    std::vector<WorksheetInfo> worksheets_;
    std::vector<DefinedNameInfo> defined_names_;
    std::unordered_map<std::string, std::string> relationships_;
    
    // 解析状态
    bool in_sheets_section_ = false;
    bool in_defined_names_section_ = false;
    bool in_defined_name_ = false;
    
    // 当前解析的定义名称
    DefinedNameInfo* current_defined_name_ = nullptr;
    
protected:
    // 重写SAX事件处理方法
    void onStartElement(std::string_view name, 
                       span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
    void onText(std::string_view data, int depth) override;
    
public:
    WorkbookParser() = default;
    ~WorkbookParser() = default;
    
    /**
     * @brief 设置关系映射（从RelationshipsParser获得）
     */
    void setRelationships(std::unordered_map<std::string, std::string> relationships) {
        relationships_ = std::move(relationships);
    }
    
    /**
     * @brief 获取解析后的工作表信息
     */
    const std::vector<WorksheetInfo>& getWorksheets() const {
        return worksheets_;
    }
    
    /**
     * @brief 获取定义名称
     */
    const std::vector<DefinedNameInfo>& getDefinedNames() const {
        return defined_names_;
    }
    
    /**
     * @brief 获取定义名称的简单字符串列表（向后兼容）
     */
    std::vector<std::string> getDefinedNameStrings() const {
        std::vector<std::string> result;
        result.reserve(defined_names_.size());
        for (const auto& def : defined_names_) {
            result.push_back(def.name);
        }
        return result;
    }
    
    /**
     * @brief 移动工作表信息所有权
     */
    std::vector<WorksheetInfo> takeWorksheets() {
        return std::move(worksheets_);
    }
    
    /**
     * @brief 清空解析状态
     */
    void reset() {
        state_.reset();
        worksheets_.clear();
        defined_names_.clear();
        in_sheets_section_ = false;
        in_defined_names_section_ = false;
        in_defined_name_ = false;
        current_defined_name_ = nullptr;
    }
};

}} // namespace fastexcel::reader