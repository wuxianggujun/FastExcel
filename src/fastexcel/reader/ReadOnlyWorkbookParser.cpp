/**
 * @file ReadOnlyWorkbookParser.cpp
 * @brief 只读模式专用工作簿XML解析器实现
 */

#include "ReadOnlyWorkbookParser.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace reader {

void ReadOnlyWorkbookParser::onStartElement(std::string_view name, 
                                           span<const xml::XMLAttribute> attributes, int depth) {
    std::string element_name(name);
    
    if (element_name == "sheets") {
        in_sheets_section_ = true;
        FASTEXCEL_LOG_DEBUG("进入工作表列表解析");
    } else if (element_name == "sheet" && in_sheets_section_) {
        in_sheet_element_ = true;
        
        // 提取工作表属性
        std::string sheet_name;
        std::string rel_id;
        
        for (const auto& attr : attributes) {
            if (attr.name == "name") {
                sheet_name = std::string(attr.value);
            } else if (attr.name == "r:id") {
                rel_id = std::string(attr.value);
            }
        }
        
        if (!sheet_name.empty() && !rel_id.empty()) {
            ReadOnlySheetInfo sheet_info(sheet_name, rel_id);
            
            // 从关系映射中查找实际的工作表路径
            auto rel_it = relationships_.find(rel_id);
            if (rel_it != relationships_.end()) {
                sheet_info.worksheet_path = "xl/" + rel_it->second;
            } else {
                // 回退到默认路径构造方式
                if (rel_id.length() > 3 && rel_id.substr(0, 3) == "rId") {
                    std::string sheet_num = rel_id.substr(3);
                    sheet_info.worksheet_path = "xl/worksheets/sheet" + sheet_num + ".xml";
                } else {
                    FASTEXCEL_LOG_WARN("无法确定工作表路径，跳过: {}", sheet_name);
                    return;
                }
            }
            
            sheets_.push_back(std::move(sheet_info));
            FASTEXCEL_LOG_DEBUG("找到工作表: {} -> {}", sheet_name, sheets_.back().worksheet_path);
        } else {
            FASTEXCEL_LOG_WARN("工作表元素缺少必要属性");
        }
    }
}

void ReadOnlyWorkbookParser::onEndElement(std::string_view name, int depth) {
    std::string element_name(name);
    
    if (element_name == "sheets") {
        in_sheets_section_ = false;
        FASTEXCEL_LOG_DEBUG("完成工作表列表解析，共找到 {} 个工作表", sheets_.size());
    } else if (element_name == "sheet" && in_sheet_element_) {
        in_sheet_element_ = false;
    }
}

void ReadOnlyWorkbookParser::onText(std::string_view data, int depth) {
    // 工作簿XML中的字符数据通常不需要特殊处理
    // 主要信息都在元素属性中
}

}} // namespace fastexcel::reader