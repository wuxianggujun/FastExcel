#pragma once

#include <memory>
#include <vector>
#include <unordered_map>
#include <string>
#include <algorithm>
#include <functional>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;
class Workbook;

/**
 * @brief 工作表管理器 - 负责管理所有工作表
 * 
 * 设计原则：
 * 1. 单一职责原则(SRP)：只负责工作表集合的管理
 * 2. 高性能：使用索引和哈希表提供O(1)查找
 * 3. 内存安全：使用智能指针管理生命周期
 */
class WorksheetManager {
public:
    using WorksheetPtr = std::shared_ptr<Worksheet>;
    using ConstWorksheetPtr = std::shared_ptr<const Worksheet>;
    using WorksheetPredicate = std::function<bool(const WorksheetPtr&)>;
    
    // 工作表信息（轻量级元数据）
    struct WorksheetInfo {
        int id;
        std::string name;
        size_t index;
        bool visible;
        bool selected;
        size_t row_count;
        size_t col_count;
    };

private:
    std::vector<WorksheetPtr> worksheets_;                    // 有序列表
    std::unordered_map<std::string, size_t> name_index_;     // 名称索引
    std::unordered_map<int, size_t> id_index_;              // ID索引
    
    Workbook* workbook_;                                     // 父工作簿引用
    int next_sheet_id_ = 1;                                  // 下一个工作表ID
    size_t active_index_ = 0;                                // 活动工作表索引
    
    // 配置
    struct Configuration {
        size_t max_sheets = 255;                             // 最大工作表数
        bool auto_generate_names = true;                     // 自动生成名称
        bool check_duplicates = true;                        // 检查重复名称
        std::string default_name_prefix = "Sheet";           // 默认名称前缀
    } config_;
    
    // 统计信息
    mutable struct Statistics {
        size_t total_created = 0;
        size_t total_deleted = 0;
        size_t lookups_by_name = 0;
        size_t lookups_by_id = 0;
        size_t lookups_by_index = 0;
    } stats_;

public:
    explicit WorksheetManager(Workbook* workbook);
    ~WorksheetManager();
    
    // 禁用拷贝，允许移动
    WorksheetManager(const WorksheetManager&) = delete;
    WorksheetManager& operator=(const WorksheetManager&) = delete;
    WorksheetManager(WorksheetManager&&) = default;
    WorksheetManager& operator=(WorksheetManager&&) = default;
    
    // 创建和添加
    
    /**
     * @brief 创建新工作表
     * @param name 工作表名称（可选，自动生成）
     * @return 新工作表
     */
    WorksheetPtr createWorksheet(const std::string& name = "");
    
    /**
     * @brief 添加现有工作表
     * @param worksheet 工作表
     * @return 是否成功
     */
    bool addWorksheet(WorksheetPtr worksheet);
    
    /**
     * @brief 批量创建工作表
     * @param count 数量
     * @param name_prefix 名称前缀
     * @return 创建的工作表列表
     */
    std::vector<WorksheetPtr> createWorksheets(size_t count, const std::string& name_prefix = "");
    
    // 查找和访问
    
    /**
     * @brief 通过名称获取工作表
     * @param name 工作表名称
     * @return 工作表（如果不存在返回nullptr）
     */
    WorksheetPtr getByName(const std::string& name);
    ConstWorksheetPtr getByName(const std::string& name) const;
    
    /**
     * @brief 通过ID获取工作表
     * @param id 工作表ID
     * @return 工作表（如果不存在返回nullptr）
     */
    WorksheetPtr getById(int id);
    ConstWorksheetPtr getById(int id) const;
    
    /**
     * @brief 通过索引获取工作表
     * @param index 索引
     * @return 工作表（如果索引无效返回nullptr）
     */
    WorksheetPtr getByIndex(size_t index);
    ConstWorksheetPtr getByIndex(size_t index) const;
    
    /**
     * @brief 获取所有工作表
     * @return 工作表列表
     */
    std::vector<WorksheetPtr> getAll() { return worksheets_; }
    std::vector<ConstWorksheetPtr> getAll() const;
    
    /**
     * @brief 查找满足条件的工作表
     * @param predicate 条件函数
     * @return 满足条件的工作表列表
     */
    std::vector<WorksheetPtr> findWhere(WorksheetPredicate predicate);
    
    // 删除操作
    
    /**
     * @brief 删除工作表
     * @param name 工作表名称
     * @return 是否成功
     */
    bool removeByName(const std::string& name);
    
    /**
     * @brief 删除工作表
     * @param id 工作表ID
     * @return 是否成功
     */
    bool removeById(int id);
    
    /**
     * @brief 删除工作表
     * @param index 索引
     * @return 是否成功
     */
    bool removeByIndex(size_t index);
    
    /**
     * @brief 清空所有工作表
     * @return 删除的工作表数量
     */
    size_t clear();
    
    // 重命名和移动
    
    /**
     * @brief 重命名工作表
     * @param old_name 旧名称
     * @param new_name 新名称
     * @return 是否成功
     */
    bool rename(const std::string& old_name, const std::string& new_name);
    
    /**
     * @brief 移动工作表位置
     * @param from_index 源索引
     * @param to_index 目标索引
     * @return 是否成功
     */
    bool move(size_t from_index, size_t to_index);
    
    /**
     * @brief 交换两个工作表位置
     * @param index1 第一个索引
     * @param index2 第二个索引
     * @return 是否成功
     */
    bool swap(size_t index1, size_t index2);
    
    // 复制和克隆
    
    /**
     * @brief 复制工作表
     * @param source_name 源工作表名称
     * @param new_name 新工作表名称
     * @return 新工作表（如果失败返回nullptr）
     */
    WorksheetPtr copy(const std::string& source_name, const std::string& new_name);
    
    /**
     * @brief 深度克隆工作表
     * @param source 源工作表
     * @param new_name 新名称
     * @return 克隆的工作表
     */
    WorksheetPtr deepClone(const WorksheetPtr& source, const std::string& new_name);
    
    // 活动工作表管理
    
    /**
     * @brief 设置活动工作表
     * @param index 索引
     */
    void setActive(size_t index);
    
    /**
     * @brief 设置活动工作表
     * @param name 工作表名称
     * @return 是否成功
     */
    bool setActive(const std::string& name);
    
    /**
     * @brief 获取活动工作表
     * @return 活动工作表
     */
    WorksheetPtr getActive();
    ConstWorksheetPtr getActive() const;
    
    /**
     * @brief 获取活动工作表索引
     * @return 索引
     */
    size_t getActiveIndex() const { return active_index_; }
    
    // 查询和统计
    
    /**
     * @brief 获取工作表数量
     * @return 数量
     */
    size_t count() const { return worksheets_.size(); }
    
    /**
     * @brief 检查是否为空
     * @return 是否为空
     */
    bool empty() const { return worksheets_.empty(); }
    
    /**
     * @brief 检查名称是否存在
     * @param name 名称
     * @return 是否存在
     */
    bool exists(const std::string& name) const {
        return name_index_.find(name) != name_index_.end();
    }
    
    /**
     * @brief 获取工作表信息
     * @param index 索引
     * @return 工作表信息
     */
    WorksheetInfo getInfo(size_t index) const;
    
    /**
     * @brief 获取所有工作表信息
     * @return 信息列表
     */
    std::vector<WorksheetInfo> getAllInfo() const;
    
    /**
     * @brief 获取统计信息
     * @return 统计信息
     */
    const Statistics& getStatistics() const { return stats_; }
    
    // 验证和工具
    
    /**
     * @brief 验证工作表名称
     * @param name 名称
     * @return 是否有效
     */
    bool validateName(const std::string& name) const;
    
    /**
     * @brief 生成唯一的工作表名称
     * @param prefix 前缀
     * @return 唯一名称
     */
    std::string generateUniqueName(const std::string& prefix = "");
    
    /**
     * @brief 重建索引（修复索引不一致）
     */
    void rebuildIndexes();
    
    /**
     * @brief 获取配置
     * @return 配置
     */
    Configuration& getConfiguration() { return config_; }
    const Configuration& getConfiguration() const { return config_; }

private:
    void updateIndexes(size_t start_index = 0);
    void removeFromIndexes(const WorksheetPtr& worksheet);
    bool isValidIndex(size_t index) const { return index < worksheets_.size(); }
    size_t getNextAvailableId();
};

}} // namespace fastexcel::core
