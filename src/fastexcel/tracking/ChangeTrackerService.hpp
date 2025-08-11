#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/tracking/IChangeHistory.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <string>
#include <unordered_set>
#include <unordered_map>
#include <vector>
#include <chrono>
#include <functional>

namespace fastexcel {
namespace tracking {

/**
 * @brief 变更历史跟踪服务 - 实现详细的变更历史管理
 * 
 * 职责：
 * - 记录详细的变更历史和时间戳
 * - 支持复杂的变更查询和分析
 * - 提供变更回滚和审计功能
 * 
 * 设计原则：
 * - 单一职责：专注于变更历史管理
 * - 开放封闭：易于扩展新的变更类型
 * - 依赖倒置：实现IChangeHistory抽象接口
 */
class ChangeHistoryTracker : public IChangeHistory {
public:
    ChangeHistoryTracker() = default;
    ~ChangeHistoryTracker() = default;

    // 禁用拷贝，支持移动
    ChangeHistoryTracker(const ChangeHistoryTracker&) = delete;
    ChangeHistoryTracker& operator=(const ChangeHistoryTracker&) = delete;
    ChangeHistoryTracker(ChangeHistoryTracker&&) = default;
    ChangeHistoryTracker& operator=(ChangeHistoryTracker&&) = default;

    // ========== IChangeHistory 接口实现 ==========
    
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
    void clearAll() override;

private:
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

} // namespace tracking
} // namespace fastexcel