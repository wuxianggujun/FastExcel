#pragma once

#include "fastexcel/tracking/IChangeTracker.hpp"
#include <optional>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace fastexcel {
namespace tracking {

/**
 * @brief 标准变更跟踪器实现
 * 
 * 职责：
 * - 跟踪Excel包中部件的修改状态
 * - 智能标记关联部件的变更
 * - 提供变更历史和统计信息
 * 
 * 设计原则：
 * - 单一职责：专注于变更状态管理
 * - 智能关联：自动标记相关部件为dirty
 * - 高效查询：使用hash表进行快速状态查询
 */
class StandardChangeTracker : public IChangeTracker {
public:
    StandardChangeTracker() = default;
    ~StandardChangeTracker() = default;

    // 禁用拷贝，支持移动
    StandardChangeTracker(const StandardChangeTracker&) = delete;
    StandardChangeTracker& operator=(const StandardChangeTracker&) = delete;
    StandardChangeTracker(StandardChangeTracker&&) = default;
    StandardChangeTracker& operator=(StandardChangeTracker&&) = default;

    // ========== IChangeTracker 接口实现 ==========
    
    // 简化接口（用于PackageEditor）
    void markPartDirty(const std::string& part) override;
    void markPartClean(const std::string& part) override;
    bool isPartDirty(const std::string& part) const override;
    std::vector<std::string> getDirtyParts() const override;
    void clearAll() override;
    bool hasChanges() const override;
    
    // 扩展接口（用于ChangeTrackerService）
    void markCreated(const std::string& resource_name) override;
    void markModified(const std::string& resource_name) override;
    void markDeleted(const std::string& resource_name) override;
    void markMoved(const std::string& old_name, const std::string& new_name) override;
    
    bool isModified(const std::string& resource_name) const override;
    size_t getChangeCount() const override;
    std::vector<std::string> getModifiedResources() const override;
    std::vector<std::string> getDeletedResources() const override;
    std::vector<std::string> getCreatedResources() const override;
    std::vector<ChangeRecord> getAllChanges() const override;
    
    void clearChanges(const std::string& resource_name) override;

private:
    /**
     * @brief 智能标记关联部件为dirty
     * 
     * 根据Excel包的内部关联关系，当某个部件被修改时，
     * 自动标记其相关部件也需要更新
     * 
     * @param part 被修改的部件路径
     */
    void markRelatedDirty(const std::string& part);

    /**
     * @brief 添加变更记录到历史
     */
    void addToHistory(const std::string& resource_name, ChangeType type);

    // 状态集合
    std::unordered_set<std::string> modified_resources_;
    std::unordered_set<std::string> deleted_resources_;  
    std::unordered_set<std::string> created_resources_;
    std::unordered_map<std::string, std::string> moved_resources_;  // old_name -> new_name
    
    // 变更历史
    std::vector<ChangeRecord> change_history_;
    
    // 配置选项
    bool enable_history_ = true;
    size_t max_history_size_ = 1000;
};

}} // namespace fastexcel::tracking
