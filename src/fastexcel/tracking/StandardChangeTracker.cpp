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
    
    dirty_parts_.insert(part);
    markRelatedDirty(part);  // 智能标记关联部件
    LOG_DEBUG("Marked part as dirty: {}", part);
}

void StandardChangeTracker::markPartClean(const std::string& part) {
    if (part.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty part as clean");
        return;
    }
    
    dirty_parts_.erase(part);
    LOG_DEBUG("Marked part as clean: {}", part);
}

bool StandardChangeTracker::isPartDirty(const std::string& part) const {
    return dirty_parts_.count(part) > 0;
}

std::vector<std::string> StandardChangeTracker::getDirtyParts() const {
    return std::vector<std::string>(dirty_parts_.begin(), dirty_parts_.end());
}

void StandardChangeTracker::clearAll() {
    size_t count = dirty_parts_.size();
    dirty_parts_.clear();
    LOG_DEBUG("Cleared all dirty parts ({} total)", count);
}

bool StandardChangeTracker::hasChanges() const {
    return !dirty_parts_.empty();
}

// ========== 私有方法实现 ==========

void StandardChangeTracker::markRelatedDirty(const std::string& part) {
    // 实现Excel包中部件间的智能关联逻辑
    // 根据OPC标准，当某些核心部件变更时，相关部件也需要更新
    
    if (part == "xl/workbook.xml") {
        // 工作簿变更影响：内容类型、关系文件
        dirty_parts_.insert("[Content_Types].xml");
        dirty_parts_.insert("_rels/.rels");
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
    }
    else if (part.find("xl/worksheets/") == 0) {
        // 工作表变更影响：工作簿关系
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
        dirty_parts_.insert("xl/workbook.xml");
    }
    else if (part == "xl/styles.xml") {
        // 样式变更影响：工作簿关系
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
    }
    else if (part == "xl/sharedStrings.xml") {
        // 共享字符串变更影响：工作簿关系
        dirty_parts_.insert("xl/_rels/workbook.xml.rels");
    }
}

} // namespace tracking
} // namespace fastexcel