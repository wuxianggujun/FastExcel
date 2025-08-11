#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include <string>
#include <vector>
#include <chrono>

namespace fastexcel {
namespace tracking {

/**
 * @brief 变更类型枚举
 */
enum class ChangeType {
    CREATED,    // 新创建
    MODIFIED,   // 修改
    DELETED,    // 删除
    MOVED       // 移动/重命名
};

/**
 * @brief 变更记录结构体
 */
struct ChangeRecord {
    std::string resource_name;
    ChangeType type;
    std::chrono::system_clock::time_point timestamp;
    
    ChangeRecord(const std::string& name, ChangeType t, const std::chrono::system_clock::time_point& time)
        : resource_name(name), type(t), timestamp(time) {}
};

/**
 * @brief 变更历史跟踪器接口
 * 遵循单一职责原则 (SRP)：专门负责详细的变更历史记录
 * 用于 ChangeTrackerService 等需要详细变更历史的场景
 */
class IChangeHistory {
public:
    virtual ~IChangeHistory() = default;
    
    // 变更记录
    virtual void markCreated(const std::string& resource_name) = 0;
    virtual void markModified(const std::string& resource_name) = 0;
    virtual void markDeleted(const std::string& resource_name) = 0;
    virtual void markMoved(const std::string& old_name, const std::string& new_name) = 0;
    
    // 查询接口
    virtual bool isModified(const std::string& resource_name) const = 0;
    virtual size_t getChangeCount() const = 0;
    virtual std::vector<std::string> getModifiedResources() const = 0;
    virtual std::vector<std::string> getDeletedResources() const = 0;
    virtual std::vector<std::string> getCreatedResources() const = 0;
    virtual std::vector<ChangeRecord> getAllChanges() const = 0;
    
    // 清理操作
    virtual void clearChanges(const std::string& resource_name) = 0;
    virtual void clearAll() = 0;
};

} // namespace tracking
} // namespace fastexcel