#pragma once

#include <string>
#include <vector>

namespace fastexcel {
namespace tracking {

/**
 * @brief 变更跟踪器接口 - 专门用于 PackageEditor
 * 遵循单一职责原则 (SRP)：专注于简单的脏标记功能
 * 遵循接口隔离原则 (ISP)：只提供 PackageEditor 需要的方法
 */
class IChangeTracker {
public:
    virtual ~IChangeTracker() = default;
    
    // 基本变更标记 - 简洁明了的接口
    virtual void markPartDirty(const std::string& part) = 0;
    virtual void markPartClean(const std::string& part) = 0;
    virtual bool isPartDirty(const std::string& part) const = 0;
    virtual std::vector<std::string> getDirtyParts() const = 0;
    virtual void clearAll() = 0;
    virtual bool hasChanges() const = 0;
};

} // namespace tracking
} // namespace fastexcel