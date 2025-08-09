#pragma once

#include "fastexcel/tracking/IChangeTracker.hpp"
#include <unordered_set>
#include <string>
#include <vector>

namespace fastexcel {
namespace tracking {

/**
 * @brief 标准变更跟踪器实现 - 专门用于PackageEditor
 * 
 * 职责：
 * - 跟踪Excel包中部件的脏标记状态 
 * - 智能标记关联部件的变更
 * - 提供简洁清晰的查询接口
 * 
 * 设计原则：
 * - 单一职责：专注于脏标记管理，不处理复杂的变更历史
 * - 接口隔离：只实现PackageEditor需要的基本功能
 * - 高效查询：使用hash表进行O(1)状态查询
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

    // 脏标记集合 - 简单清晰
    std::unordered_set<std::string> dirty_parts_;
};

} // namespace tracking
} // namespace fastexcel
