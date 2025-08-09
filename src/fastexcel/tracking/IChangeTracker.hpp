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
 * @brief 变更跟踪器接口 - 负责跟踪和管理变更状态
 * 遵循接口隔离原则 (ISP)：定义清晰的变更跟踪契约
 */
class IChangeTracker {
public:
    virtual ~IChangeTracker() = default;
    
    // 基本变更标记 - 简化接口（用于PackageEditor）
    virtual void markPartDirty(const std::string& part) = 0;
    virtual void markPartClean(const std::string& part) = 0;
    virtual bool isPartDirty(const std::string& part) const = 0;
    virtual std::vector<std::string> getDirtyParts() const = 0;
    virtual void clearAll() = 0;
    virtual bool hasChanges() const = 0;
    
    // 扩展接口 - 详细变更跟踪（用于ChangeTrackerService）
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
    
    // 重置和清理
    virtual void clearChanges(const std::string& resource_name) = 0;
};

} // namespace tracking
} // namespace fastexcel