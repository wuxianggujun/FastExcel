#pragma once

#include "fastexcel/utils/Logger.hpp"
#include <map>
#include <set>
#include <string>
#include <vector>
#include <optional>
#include <fmt/format.h>
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace core {

/**
 * @brief 保存策略枚举
 */
enum class SaveStrategy {
    NONE,           // 无需保存
    PURE_CREATE,    // 纯新建模式
    MINIMAL_UPDATE, // 最小更新模式
    SMART_EDIT,     // 智能编辑模式
    FULL_REBUILD    // 完全重建模式
};

/**
 * @brief 脏数据管理器 - 智能追踪文件修改状态
 * 
 * 这个类负责：
 * 1. 追踪各个部件的修改状态
 * 2. 分析修改的影响范围
 * 3. 推荐最优的保存策略
 */
class DirtyManager {
public:
    /**
     * @brief 脏数据级别
     */
    enum class DirtyLevel {
        NONE = 0,       // 未修改
        METADATA = 1,   // 仅元数据改变（如修改时间）
        CONTENT = 2,    // 内容改变
        STRUCTURE = 3   // 结构改变（增删工作表等）
    };
    
    /**
     * @brief 部件脏信息
     */
    struct PartDirtyInfo {
        DirtyLevel level = DirtyLevel::NONE;
        std::set<std::string> affectedPaths;  // 受影响的具体路径
        bool requiresRegeneration = false;     // 是否需要完全重新生成
        
        // 添加受影响路径
        void addAffectedPath(const std::string& path) {
            if (!path.empty()) {
                affectedPaths.insert(path);
            }
        }
        
        // 判断是否影响特定路径
        bool affects(const std::string& path) const {
            return affectedPaths.find(path) != affectedPaths.end();
        }
        
        // 清除状态
        void clear() {
            level = DirtyLevel::NONE;
            affectedPaths.clear();
            requiresRegeneration = false;
        }
    };
    
    /**
     * @brief 变更集合
     */
    class ChangeSet {
    public:
        struct Change {
            std::string part;
            std::string path;
            DirtyLevel level;
        };
        
    private:
        std::vector<Change> changes_;
        bool hasStructuralChanges_ = false;
        
    public:
        void add(const std::string& part, const std::string& path, DirtyLevel level) {
            changes_.push_back({part, path, level});
            if (level == DirtyLevel::STRUCTURE) {
                hasStructuralChanges_ = true;
            }
        }
        
        bool isEmpty() const { return changes_.empty(); }
        bool hasStructuralChanges() const { return hasStructuralChanges_; }
        
        bool isOnlyContentChanges() const {
            if (isEmpty()) return false;
            for (const auto& change : changes_) {
                if (change.level != DirtyLevel::CONTENT) {
                    return false;
                }
            }
            return true;
        }
        
        bool affects(const std::string& path) const {
            for (const auto& change : changes_) {
                if (change.path == path) {
                    return true;
                }
            }
            return false;
        }
        
        const std::vector<Change>& getChanges() const { return changes_; }
        
        size_t size() const { return changes_.size(); }
    };
    
private:
    std::map<std::string, PartDirtyInfo> dirtyParts_;
    bool isNewFile_ = true;  // 是否是新建文件
    
    // 部件依赖关系
    std::map<std::string, std::set<std::string>> dependencies_;
    
public:
    DirtyManager() {
        initializeDependencies();
    }
    
    /**
     * @brief 标记部件为脏
     */
    void markDirty(const std::string& part, DirtyLevel level, 
                   const std::string& affectedPath = "") {
        auto& info = dirtyParts_[part];
        
        // 更新脏级别（取最高级别）
        if (level > info.level) {
            info.level = level;
        }
        
        // 添加受影响路径
        if (!affectedPath.empty()) {
            info.addAffectedPath(affectedPath);
        }
        
        // 结构变更需要重新生成
        if (level == DirtyLevel::STRUCTURE) {
            info.requiresRegeneration = true;
        }
        // 传播脏标记到依赖项
        propagateDirty(part, level);
    }
    
    /**
     * @brief 标记工作表相关部件为脏
     */
    void markWorksheetDirty(size_t index, DirtyLevel level = DirtyLevel::CONTENT) {
        std::string sheetPart = fmt::format("xl/worksheets/sheet{}.xml", index + 1);
        markDirty(sheetPart, level);
        
        // 如果是结构变更，也要更新workbook
        if (level == DirtyLevel::STRUCTURE) {
            markDirty("xl/workbook.xml", DirtyLevel::CONTENT);
            markDirty("xl/_rels/workbook.xml.rels", DirtyLevel::CONTENT);
        }
    }
    
    /**
     * @brief 标记样式为脏
     */
    void markStylesDirty() {
        markDirty("xl/styles.xml", DirtyLevel::CONTENT);
    }
    
    /**
     * @brief 标记主题为脏
     */
    void markThemeDirty() {
        markDirty("xl/theme/theme1.xml", DirtyLevel::CONTENT);
    }
    
    /**
     * @brief 标记共享字符串为脏
     */
    void markSharedStringsDirty() {
        markDirty("xl/sharedStrings.xml", DirtyLevel::CONTENT);
    }
    
    /**
     * @brief 判断部件是否需要更新
     */
    bool shouldUpdate(const std::string& part) const {
        auto it = dirtyParts_.find(part);
        if (it == dirtyParts_.end()) {
            return isNewFile_;  // 新文件需要生成所有部件
        }
        return it->second.level != DirtyLevel::NONE;
    }
    
    /**
     * @brief 获取部件的脏级别
     */
    DirtyLevel getDirtyLevel(const std::string& part) const {
        auto it = dirtyParts_.find(part);
        if (it == dirtyParts_.end()) {
            return isNewFile_ ? DirtyLevel::CONTENT : DirtyLevel::NONE;
        }
        return it->second.level;
    }
    
    /**
     * @brief 获取最优保存策略
     */
    SaveStrategy getOptimalStrategy() const {
        if (isNewFile_) {
            return SaveStrategy::PURE_CREATE;
        }
        
        // 统计各级别的脏部件数量
        int structureCount = 0;
        int contentCount = 0;
        int metadataCount = 0;
        
        for (const auto& [part, info] : dirtyParts_) {
            switch (info.level) {
                case DirtyLevel::STRUCTURE:
                    structureCount++;
                    break;
                case DirtyLevel::CONTENT:
                    contentCount++;
                    break;
                case DirtyLevel::METADATA:
                    metadataCount++;
                    break;
                default:
                    break;
            }
        }
        
        // 根据修改情况选择策略
        if (structureCount > 0) {
            // 有结构变更，需要重建
            return SaveStrategy::FULL_REBUILD;
        } else if (contentCount > 0) {
            // 内容变更
            if (contentCount <= 3) {
                // 少量修改，使用最小更新
                return SaveStrategy::MINIMAL_UPDATE;
            } else {
                // 较多修改，使用智能编辑
                return SaveStrategy::SMART_EDIT;
            }
        } else if (metadataCount > 0) {
            // 仅元数据变更
            return SaveStrategy::MINIMAL_UPDATE;
        }
        
        return SaveStrategy::NONE;
    }
    
    /**
     * @brief 获取所有变更
     */
    ChangeSet getChanges() const {
        ChangeSet changes;
        
        for (const auto& [part, info] : dirtyParts_) {
            if (info.level != DirtyLevel::NONE) {
                // 如果有具体的受影响路径，添加每个路径
                if (!info.affectedPaths.empty()) {
                    for (const auto& path : info.affectedPaths) {
                        changes.add(part, path, info.level);
                    }
                } else {
                    // 否则添加整个部件
                    changes.add(part, part, info.level);
                }
            }
        }
        
        return changes;
    }
    
    /**
     * @brief 清除所有脏标记
     */
    void clear() {
        dirtyParts_.clear();
    }
    
    /**
     * @brief 设置是否为新文件
     */
    void setIsNewFile(bool isNew) {
        isNewFile_ = isNew;
    }
    
    /**
     * @brief 获取脏部件数量
     */
    size_t getDirtyCount() const {
        size_t count = 0;
        for (const auto& [part, info] : dirtyParts_) {
            if (info.level != DirtyLevel::NONE) {
                count++;
            }
        }
        return count;
    }
    
    /**
     * @brief 判断是否有任何脏数据
     */
    bool hasDirtyData() const {
        for (const auto& [part, info] : dirtyParts_) {
            if (info.level != DirtyLevel::NONE) {
                return true;
            }
        }
        return false;
    }
    
private:
    /**
     * @brief 初始化部件依赖关系
     */
    void initializeDependencies() {
        // workbook.xml 依赖于工作表
        dependencies_["xl/workbook.xml"] = {
            "xl/worksheets/sheet*.xml"
        };
        
        // workbook.xml.rels 依赖于所有关系
        dependencies_["xl/_rels/workbook.xml.rels"] = {
            "xl/worksheets/sheet*.xml",
            "xl/theme/theme1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml"
        };
        
        // Content_Types.xml 依赖于所有部件
        dependencies_["[Content_Types].xml"] = {
            "xl/worksheets/sheet*.xml",
            "xl/theme/theme1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "xl/workbook.xml"
        };
    }
    
    /**
     * @brief 传播脏标记到依赖项
     */
    void propagateDirty(const std::string& part, DirtyLevel level) {
        // 查找哪些部件依赖于当前部件
        for (const auto& [dependent, deps] : dependencies_) {
            for (const auto& dep : deps) {
                // 简单的通配符匹配
                if (dep.find('*') != std::string::npos) {
                    // 处理通配符
                    std::string prefix = dep.substr(0, dep.find('*'));
                    if (part.find(prefix) == 0) {
                        // 传播较低级别的脏标记
                        DirtyLevel propagatedLevel = (level == DirtyLevel::STRUCTURE) 
                            ? DirtyLevel::CONTENT 
                            : DirtyLevel::METADATA;
                        
                        auto& depInfo = dirtyParts_[dependent];
                        if (propagatedLevel > depInfo.level) {
                            depInfo.level = propagatedLevel;
                            FASTEXCEL_LOG_DEBUG("Propagated dirty from {} to {} (level: {})",
                                     part, dependent, static_cast<int>(propagatedLevel));
                        }
                    }
                } else if (dep == part) {
                    // 直接依赖
                    DirtyLevel propagatedLevel = (level == DirtyLevel::STRUCTURE) 
                        ? DirtyLevel::CONTENT 
                        : DirtyLevel::METADATA;
                    
                    auto& depInfo = dirtyParts_[dependent];
                    if (propagatedLevel > depInfo.level) {
                        depInfo.level = propagatedLevel;
                        FASTEXCEL_LOG_DEBUG("Propagated dirty from {} to {} (level: {})",
                                 part, dependent, static_cast<int>(propagatedLevel));
                    }
                }
            }
        }
    }
};

} // namespace core
} // namespace fastexcel
