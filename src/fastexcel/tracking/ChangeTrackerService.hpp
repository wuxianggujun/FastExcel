#pragma once

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
 * @brief 变更类型枚举
 */
enum class ChangeType {
    CREATED,    // 新创建
    MODIFIED,   // 修改
    DELETED,    // 删除
    MOVED       // 移动/重命名
};

/**
 * @brief 变更记录结构
 */
struct ChangeRecord {
    std::string resource_name;
    ChangeType type;
    std::chrono::system_clock::time_point timestamp;
    std::string checksum_before;  // 修改前的校验和
    std::string checksum_after;   // 修改后的校验和
    
    ChangeRecord(const std::string& name, ChangeType change_type)
        : resource_name(name), type(change_type), timestamp(std::chrono::system_clock::now()) {}
};

/**
 * @brief 变更跟踪器接口
 * 遵循接口隔离原则 (ISP)：只包含变更跟踪相关的方法
 */
class IChangeTracker {
public:
    virtual ~IChangeTracker() = default;
    
    // 基本变更操作
    virtual void markCreated(const std::string& resource_name) = 0;
    virtual void markModified(const std::string& resource_name) = 0;
    virtual void markDeleted(const std::string& resource_name) = 0;
    virtual void markMoved(const std::string& old_name, const std::string& new_name) = 0;
    
    // 状态查询
    virtual bool hasChanges() const = 0;
    virtual bool isModified(const std::string& resource_name) const = 0;
    virtual bool isDeleted(const std::string& resource_name) const = 0;
    virtual bool isCreated(const std::string& resource_name) const = 0;
    
    // 获取变更信息
    virtual std::vector<std::string> getModifiedResources() const = 0;
    virtual std::vector<std::string> getDeletedResources() const = 0;
    virtual std::vector<std::string> getCreatedResources() const = 0;
    virtual std::vector<ChangeRecord> getAllChanges() const = 0;
    
    // 重置和清理
    virtual void clearAll() = 0;
    virtual void clearChanges(const std::string& resource_name) = 0;
};

/**
 * @brief 标准变更跟踪器实现
 * 遵循单一职责原则 (SRP)：专注于跟踪和记录变更
 * 提供高性能的变更检测和历史记录功能
 */
class StandardChangeTracker : public IChangeTracker {
private:
    std::unordered_set<std::string> modified_resources_;
    std::unordered_set<std::string> deleted_resources_;  
    std::unordered_set<std::string> created_resources_;
    std::unordered_map<std::string, std::string> moved_resources_;  // old_name -> new_name
    
    std::vector<ChangeRecord> change_history_;  // 完整的变更历史
    std::unordered_map<std::string, std::string> resource_checksums_; // 用于检测实际变更
    
    // 配置选项
    bool enable_history_ = true;
    bool enable_checksums_ = false;
    size_t max_history_size_ = 1000;
    
    // 添加变更记录到历史
    void addToHistory(const std::string& resource_name, ChangeType type) {
        if (!enable_history_) return;
        
        change_history_.emplace_back(resource_name, type);
        
        // 限制历史记录大小
        if (change_history_.size() > max_history_size_) {
            change_history_.erase(change_history_.begin(), 
                                change_history_.begin() + (change_history_.size() - max_history_size_));
        }
    }
    
    // 计算资源的校验和（如果启用）
    std::string calculateChecksum(const std::string& content) const {
        if (!enable_checksums_) return \"\";
        
        // 简单的哈希函数（实际项目中应该使用更强的哈希算法）
        std::hash<std::string> hasher;
        return std::to_string(hasher(content));
    }
    
public:
    /**
     * @brief 构造函数，可配置跟踪选项
     */
    explicit StandardChangeTracker(bool enable_history = true, bool enable_checksums = false)
        : enable_history_(enable_history), enable_checksums_(enable_checksums) {}
    
    // ========== 基本变更操作 ==========
    
    void markCreated(const std::string& resource_name) override {
        created_resources_.insert(resource_name);
        
        // 从其他集合中移除（创建优先级最高）
        modified_resources_.erase(resource_name);
        deleted_resources_.erase(resource_name);
        
        addToHistory(resource_name, ChangeType::CREATED);
        LOG_DEBUG(\"Marked resource as created: {}\", resource_name);
    }
    
    void markModified(const std::string& resource_name) override {
        // 如果是新创建的资源，不需要标记为修改
        if (created_resources_.count(resource_name)) {
            LOG_DEBUG(\"Resource '{}' is newly created, not marking as modified\", resource_name);
            return;
        }
        
        modified_resources_.insert(resource_name);
        deleted_resources_.erase(resource_name);  // 取消删除标记
        
        addToHistory(resource_name, ChangeType::MODIFIED);
        LOG_DEBUG(\"Marked resource as modified: {}\", resource_name);
    }
    
    void markDeleted(const std::string& resource_name) override {
        deleted_resources_.insert(resource_name);
        
        // 从其他集合中移除
        modified_resources_.erase(resource_name);
        created_resources_.erase(resource_name);
        
        addToHistory(resource_name, ChangeType::DELETED);
        LOG_DEBUG(\"Marked resource as deleted: {}\", resource_name);
    }
    
    void markMoved(const std::string& old_name, const std::string& new_name) override {
        moved_resources_[old_name] = new_name;
        
        // 转移变更状态到新名称
        if (modified_resources_.count(old_name)) {
            modified_resources_.erase(old_name);
            modified_resources_.insert(new_name);
        }
        if (created_resources_.count(old_name)) {
            created_resources_.erase(old_name);
            created_resources_.insert(new_name);
        }
        
        addToHistory(old_name, ChangeType::MOVED);
        LOG_DEBUG(\"Marked resource as moved: {} -> {}\", old_name, new_name);
    }
    
    // ========== 状态查询 ==========
    
    bool hasChanges() const override {
        return !modified_resources_.empty() || 
               !deleted_resources_.empty() || 
               !created_resources_.empty() ||
               !moved_resources_.empty();
    }
    
    bool isModified(const std::string& resource_name) const override {
        return modified_resources_.count(resource_name) > 0;
    }
    
    bool isDeleted(const std::string& resource_name) const override {
        return deleted_resources_.count(resource_name) > 0;
    }
    
    bool isCreated(const std::string& resource_name) const override {
        return created_resources_.count(resource_name) > 0;
    }
    
    // ========== 获取变更信息 ==========
    
    std::vector<std::string> getModifiedResources() const override {
        return std::vector<std::string>(modified_resources_.begin(), modified_resources_.end());
    }
    
    std::vector<std::string> getDeletedResources() const override {
        return std::vector<std::string>(deleted_resources_.begin(), deleted_resources_.end());
    }
    
    std::vector<std::string> getCreatedResources() const override {
        return std::vector<std::string>(created_resources_.begin(), created_resources_.end());
    }
    
    std::vector<ChangeRecord> getAllChanges() const override {
        return change_history_;
    }
    
    // ========== 重置和清理 ==========
    
    void clearAll() override {
        modified_resources_.clear();
        deleted_resources_.clear();
        created_resources_.clear();
        moved_resources_.clear();
        change_history_.clear();
        resource_checksums_.clear();
        
        LOG_DEBUG(\"Cleared all tracked changes\");
    }
    
    void clearChanges(const std::string& resource_name) override {
        modified_resources_.erase(resource_name);
        deleted_resources_.erase(resource_name);
        created_resources_.erase(resource_name);
        
        // 清理移动记录
        for (auto it = moved_resources_.begin(); it != moved_resources_.end();) {
            if (it->first == resource_name || it->second == resource_name) {
                it = moved_resources_.erase(it);
            } else {
                ++it;
            }
        }
        
        resource_checksums_.erase(resource_name);
        
        LOG_DEBUG(\"Cleared changes for resource: {}\", resource_name);
    }
    
    // ========== 高级功能 ==========
    
    /**
     * @brief 获取变更统计信息
     */
    struct ChangeStats {
        size_t created_count = 0;
        size_t modified_count = 0;
        size_t deleted_count = 0;
        size_t moved_count = 0;
        size_t total_changes = 0;
    };
    
    ChangeStats getChangeStats() const {
        ChangeStats stats;
        stats.created_count = created_resources_.size();
        stats.modified_count = modified_resources_.size();
        stats.deleted_count = deleted_resources_.size();
        stats.moved_count = moved_resources_.size();
        stats.total_changes = stats.created_count + stats.modified_count + 
                             stats.deleted_count + stats.moved_count;
        return stats;
    }
};

/**
 * @brief 简单变更跟踪器 - 用于测试和简单场景
 * 最小化的实现，专注于基本功能
 */
class SimpleChangeTracker : public IChangeTracker {
private:
    std::unordered_set<std::string> dirty_resources_;
    
public:
    void markCreated(const std::string& resource_name) override {
        dirty_resources_.insert(resource_name);
    }
    
    void markModified(const std::string& resource_name) override {
        dirty_resources_.insert(resource_name);
    }
    
    void markDeleted(const std::string& resource_name) override {
        dirty_resources_.insert(resource_name);
    }
    
    void markMoved(const std::string& old_name, const std::string& new_name) override {
        dirty_resources_.insert(old_name);
        dirty_resources_.insert(new_name);
    }
    
    bool hasChanges() const override {
        return !dirty_resources_.empty();
    }
    
    bool isModified(const std::string& resource_name) const override {
        return dirty_resources_.count(resource_name) > 0;
    }
    
    bool isDeleted(const std::string& resource_name) const override {
        return dirty_resources_.count(resource_name) > 0;
    }
    
    bool isCreated(const std::string& resource_name) const override {
        return dirty_resources_.count(resource_name) > 0;
    }
    
    std::vector<std::string> getModifiedResources() const override {
        return std::vector<std::string>(dirty_resources_.begin(), dirty_resources_.end());
    }
    
    std::vector<std::string> getDeletedResources() const override {
        return getModifiedResources();  // 简化实现
    }
    
    std::vector<std::string> getCreatedResources() const override {
        return getModifiedResources();  // 简化实现
    }
    
    std::vector<ChangeRecord> getAllChanges() const override {
        return {};  // 简化实现，不保存历史
    }
    
    void clearAll() override {
        dirty_resources_.clear();
    }
    
    void clearChanges(const std::string& resource_name) override {
        dirty_resources_.erase(resource_name);
    }
};

} // namespace tracking
} // namespace fastexcel
