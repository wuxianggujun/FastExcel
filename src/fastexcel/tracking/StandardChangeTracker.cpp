#include "fastexcel/tracking/StandardChangeTracker.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <chrono>

namespace fastexcel {
namespace tracking {

// ========== IChangeTracker 接口实现 ==========

// 简化接口实现（用于PackageEditor）
void StandardChangeTracker::markPartDirty(const std::string& part) {
    markModified(part);
}

void StandardChangeTracker::markPartClean(const std::string& part) {
    clearChanges(part);
}

bool StandardChangeTracker::isPartDirty(const std::string& part) const {
    return isModified(part);
}

std::vector<std::string> StandardChangeTracker::getDirtyParts() const {
    return getModifiedResources();
}

// 扩展接口实现（用于ChangeTrackerService）
void StandardChangeTracker::markCreated(const std::string& resource_name) {
    if (resource_name.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty resource as created");
        return;
    }
    
    created_resources_.insert(resource_name);
    
    // 从其他集合中移除（创建优先级最高）
    modified_resources_.erase(resource_name);
    deleted_resources_.erase(resource_name);
    
    addToHistory(resource_name, ChangeType::CREATED);
    LOG_DEBUG("Marked resource as created: {}", resource_name);
}

void StandardChangeTracker::markModified(const std::string& resource_name) {
    if (resource_name.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty resource as modified");
        return;
    }
    
    // 如果是新创建的资源，不需要标记为修改
    if (created_resources_.count(resource_name)) {
        LOG_DEBUG("Resource '{}' is newly created, not marking as modified", resource_name);
        return;
    }
    
    modified_resources_.insert(resource_name);
    deleted_resources_.erase(resource_name);  // 取消删除标记
    
    addToHistory(resource_name, ChangeType::MODIFIED);
    markRelatedDirty(resource_name);  // 智能标记关联部件
    LOG_DEBUG("Marked resource as modified: {}", resource_name);
}

void StandardChangeTracker::markDeleted(const std::string& resource_name) {
    if (resource_name.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty resource as deleted");
        return;
    }
    
    deleted_resources_.insert(resource_name);
    
    // 从其他集合中移除
    modified_resources_.erase(resource_name);
    created_resources_.erase(resource_name);
    
    addToHistory(resource_name, ChangeType::DELETED);
    LOG_DEBUG("Marked resource as deleted: {}", resource_name);
}

void StandardChangeTracker::markMoved(const std::string& old_name, const std::string& new_name) {
    if (old_name.empty() || new_name.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to mark empty resource as moved");
        return;
    }
    
    moved_resources_[old_name] = new_name;
    
    // 移动操作也算作修改
    markModified(new_name);
    markDeleted(old_name);
    
    addToHistory(old_name + " -> " + new_name, ChangeType::MOVED);
    LOG_DEBUG("Marked resource as moved: {} -> {}", old_name, new_name);
}

bool StandardChangeTracker::isModified(const std::string& resource_name) const {
    return modified_resources_.count(resource_name) > 0 ||
           created_resources_.count(resource_name) > 0 ||
           deleted_resources_.count(resource_name) > 0;
}

bool StandardChangeTracker::hasChanges() const {
    return !modified_resources_.empty() || 
           !created_resources_.empty() || 
           !deleted_resources_.empty();
}

size_t StandardChangeTracker::getChangeCount() const {
    return modified_resources_.size() + 
           created_resources_.size() + 
           deleted_resources_.size();
}

std::vector<std::string> StandardChangeTracker::getModifiedResources() const {
    std::vector<std::string> result;
    
    // 合并所有变更类型
    for (const auto& name : modified_resources_) {
        result.push_back(name);
    }
    for (const auto& name : created_resources_) {
        result.push_back(name);
    }
    for (const auto& name : deleted_resources_) {
        result.push_back(name);
    }
    
    return result;
}

std::vector<std::string> StandardChangeTracker::getDeletedResources() const {
    return std::vector<std::string>(deleted_resources_.begin(), deleted_resources_.end());
}

std::vector<std::string> StandardChangeTracker::getCreatedResources() const {
    return std::vector<std::string>(created_resources_.begin(), created_resources_.end());
}

std::vector<ChangeRecord> StandardChangeTracker::getAllChanges() const {
    return change_history_;
}

void StandardChangeTracker::clearAll() {
    size_t total = modified_resources_.size() + created_resources_.size() + deleted_resources_.size();
    LOG_DEBUG("Clearing all tracked changes ({} total)", total);
    
    modified_resources_.clear();
    created_resources_.clear();
    deleted_resources_.clear();
    moved_resources_.clear();
    change_history_.clear();
}

void StandardChangeTracker::clearChanges(const std::string& resource_name) {
    if (resource_name.empty()) {
        LOG_WARN("StandardChangeTracker: attempt to clear changes for empty resource name");
        return;
    }
    
    modified_resources_.erase(resource_name);
    created_resources_.erase(resource_name);
    deleted_resources_.erase(resource_name);
    
    // 清理移动记录
    auto it = moved_resources_.find(resource_name);
    if (it != moved_resources_.end()) {
        moved_resources_.erase(it);
    }
    
    LOG_DEBUG("Cleared changes for resource: {}", resource_name);
}

// ========== 私有方法实现 ==========

void StandardChangeTracker::addToHistory(const std::string& resource_name, ChangeType type) {
    if (!enable_history_) return;
    
    auto now = std::chrono::system_clock::now();
    change_history_.emplace_back(resource_name, type, now);
    
    // 限制历史记录大小
    if (change_history_.size() > max_history_size_) {
        change_history_.erase(change_history_.begin());
    }
}

void StandardChangeTracker::markRelatedDirty(const std::string& part) {
    // 实现Excel包中部件间的智能关联逻辑
    // 这里是一个简化的实现，可以根据实际需求扩展
    
    if (part == "xl/workbook.xml") {
        // 工作簿变更影响：内容类型、关系文件
        markModified("[Content_Types].xml");
        markModified("_rels/.rels");
        markModified("xl/_rels/workbook.xml.rels");
    }
    else if (part.find("xl/worksheets/") == 0) {
        // 工作表变更影响：工作簿关系
        markModified("xl/_rels/workbook.xml.rels");
        markModified("xl/workbook.xml");
    }
    else if (part == "xl/styles.xml") {
        // 样式变更影响：工作簿关系
        markModified("xl/_rels/workbook.xml.rels");
    }
    else if (part == "xl/sharedStrings.xml") {
        // 共享字符串变更影响：工作簿关系
        markModified("xl/_rels/workbook.xml.rels");
    }
}

}} // namespace fastexcel::tracking