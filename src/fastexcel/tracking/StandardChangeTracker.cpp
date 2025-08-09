#include "fastexcel/tracking/StandardChangeTracker.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace tracking {

// ========== IChangeTracker 接口实现 ==========

void StandardChangeTracker::markPartDirty(const std::string& part) {
    if (part.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty part as dirty");
        return;
    }
    
    LOG_DEBUG("Marking part as dirty: {}", part);
    dirty_parts_.insert(part);
    
    // 智能标记关联部件
    markRelatedDirty(part);
}

void StandardChangeTracker::markPartClean(const std::string& part) {
    if (part.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty part as clean");
        return;
    }
    
    LOG_DEBUG("Marking part as clean: {}", part);
    dirty_parts_.erase(part);
}

bool StandardChangeTracker::isPartDirty(const std::string& part) const {
    return dirty_parts_.count(part) > 0;
}

std::vector<std::string> StandardChangeTracker::getDirtyParts() const {
    return std::vector<std::string>(dirty_parts_.begin(), dirty_parts_.end());
}

void StandardChangeTracker::clearAll() {
    LOG_DEBUG("Clearing all dirty parts (was {} parts)", dirty_parts_.size());
    dirty_parts_.clear();
}

bool StandardChangeTracker::hasChanges() const {
    return !dirty_parts_.empty();
}

// ========== 私有方法实现 ==========

void StandardChangeTracker::markRelatedDirty(const std::string& part) {
    // 工作表修改影响的关联文件
    if (part.find("xl/worksheets/sheet") == 0) {
        LOG_DEBUG("Worksheet modified, marking related parts dirty");
        
        // 工作表修改可能影响 calcChain
        dirty_parts_.insert("xl/calcChain.xml");
        
        // 工作表修改影响工作簿定义
        dirty_parts_.insert("xl/workbook.xml");
        
        // 工作表修改影响内容类型定义
        dirty_parts_.insert("[Content_Types].xml");
        
        // 工作表修改影响关系定义
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
    }
    
    // 共享字符串修改影响的关联文件
    if (part == "xl/sharedStrings.xml") {
        LOG_DEBUG("SharedStrings modified, marking related parts dirty");
        
        // 共享字符串影响关系文件
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
        
        // 共享字符串影响内容类型定义
        dirty_parts_.insert("[Content_Types].xml");
    }
    
    // 样式修改影响的关联文件
    if (part == "xl/styles.xml") {
        LOG_DEBUG("Styles modified, marking related parts dirty");
        
        // 样式修改影响关系文件
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
        
        // 样式修改影响内容类型定义
        dirty_parts_.insert("[Content_Types].xml");
    }
    
    // 工作簿修改影响的关联文件
    if (part == "xl/workbook.xml") {
        LOG_DEBUG("Workbook modified, marking related parts dirty");
        
        // 工作簿修改影响主关系文件
        dirty_parts_.insert("_rels/.rels");
        
        // 工作簿修改影响内容类型定义
        dirty_parts_.insert("[Content_Types].xml");
    }
    
    // 关系文件修改的连锁反应
    if (part.find("_rels/") != std::string::npos) {
        LOG_DEBUG("Relationships modified, marking content types dirty");
        
        // 关系文件修改影响内容类型定义
        dirty_parts_.insert("[Content_Types].xml");
    }
    
    // 文档属性修改影响的关联文件
    if (part.find("docProps/") == 0) {
        LOG_DEBUG("Document properties modified, marking main rels dirty");
        
        // 文档属性修改影响主关系文件
        dirty_parts_.insert("_rels/.rels");
    }
}

}} // namespace fastexcel::tracking