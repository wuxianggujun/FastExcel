/**
 * @file WorkbookParser.cpp
 * @brief 通用工作簿XML解析器实现
 */

#include "WorkbookParser.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace reader {

void WorkbookParser::onStartElement(std::string_view name, 
                                   span<const xml::XMLAttribute> attributes, int /* depth */) {
    std::string element_name(name);
    
    if (element_name == "sheets") {
        in_sheets_section_ = true;
        FASTEXCEL_LOG_DEBUG("进入工作表列表解析");
    } else if (element_name == "sheet" && in_sheets_section_) {
        // 提取工作表属性
        std::string sheet_name;
        std::string sheet_id;
        std::string rel_id;
        
        for (const auto& attr : attributes) {
            if (attr.name == "name") {
                sheet_name = std::string(attr.value);
            } else if (attr.name == "sheetId") {
                sheet_id = std::string(attr.value);
            } else if (attr.name == "r:id") {
                rel_id = std::string(attr.value);
            }
        }
        
        if (!sheet_name.empty() && !sheet_id.empty() && !rel_id.empty()) {
            WorksheetInfo worksheet_info(sheet_name, sheet_id, rel_id);
            
            // 从关系映射中查找实际的工作表路径
            auto rel_it = relationships_.find(rel_id);
            if (rel_it != relationships_.end()) {
                worksheet_info.worksheet_path = "xl/" + rel_it->second;
            } else {
                // 回退到默认路径构造方式
                worksheet_info.worksheet_path = "xl/worksheets/sheet" + sheet_id + ".xml";
                FASTEXCEL_LOG_WARN("关系映射中未找到 {}，使用默认路径: {}", 
                                 rel_id, worksheet_info.worksheet_path);
            }
            
            worksheets_.push_back(std::move(worksheet_info));
            FASTEXCEL_LOG_DEBUG("找到工作表: {} (ID: {}) -> {}", 
                              sheet_name, sheet_id, worksheets_.back().worksheet_path);
        } else {
            FASTEXCEL_LOG_WARN("工作表元素缺少必要属性: name='{}', sheetId='{}', r:id='{}'", 
                             sheet_name, sheet_id, rel_id);
        }
    } else if (element_name == "definedNames") {
        in_defined_names_section_ = true;
        FASTEXCEL_LOG_DEBUG("进入定义名称列表解析");
    } else if (element_name == "definedName" && in_defined_names_section_) {
        in_defined_name_ = true;
        
        // 创建新的定义名称对象
        std::string defined_name;
        for (const auto& attr : attributes) {
            if (attr.name == "name") {
                defined_name = std::string(attr.value);
            }
        }
        
        if (!defined_name.empty()) {
            defined_names_.emplace_back(std::move(defined_name));
            current_defined_name_ = &defined_names_.back();
            
            // 提取其他可选属性
            for (const auto& attr : attributes) {
                if (attr.name == "localSheetId") {
                    current_defined_name_->local_sheet_id = std::string(attr.value);
                } else if (attr.name == "comment") {
                    current_defined_name_->comment = std::string(attr.value);
                } else if (attr.name == "hidden") {
                    current_defined_name_->hidden = (std::string(attr.value) == "1");
                }
            }
        } else {
            FASTEXCEL_LOG_WARN("definedName元素缺少name属性");
        }
    }
}

void WorkbookParser::onEndElement(std::string_view name, int /* depth */) {
    std::string element_name(name);
    
    if (element_name == "sheets") {
        in_sheets_section_ = false;
        FASTEXCEL_LOG_DEBUG("完成工作表列表解析，共找到 {} 个工作表", worksheets_.size());
    } else if (element_name == "definedNames") {
        in_defined_names_section_ = false;
        FASTEXCEL_LOG_DEBUG("完成定义名称列表解析，共找到 {} 个定义名称", defined_names_.size());
    } else if (element_name == "definedName" && in_defined_name_) {
        in_defined_name_ = false;
        if (current_defined_name_) {
            FASTEXCEL_LOG_DEBUG("完成定义名称解析: {} = '{}'", 
                              current_defined_name_->name, current_defined_name_->formula);
        }
        current_defined_name_ = nullptr;
    }
}

void WorkbookParser::onText(std::string_view data, int /* depth */) {
    // 处理definedName元素的文本内容（公式/引用）
    if (in_defined_name_ && current_defined_name_) {
        // 累积文本内容，因为可能会分多次调用onText
        current_defined_name_->formula += std::string(data);
    }
    // 其他元素的文本内容通常不重要，在属性中已经处理
}

}} // namespace fastexcel::reader