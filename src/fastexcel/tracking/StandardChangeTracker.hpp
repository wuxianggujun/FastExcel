#pragma once

#include "fastexcel/tracking/ChangeTrackerService.hpp"
#include <unordered_set>
#include <string>
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
    
    void markPartDirty(const std::string& part) override;
    void markPartClean(const std::string& part) override;
    bool isPartDirty(const std::string& part) const override;
    std::vector<std::string> getDirtyParts() const override;
    void clearAll() override;
    bool hasChanges() const override;

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

    std::unordered_set<std::string> dirty_parts_;
};

}} // namespace fastexcel::tracking