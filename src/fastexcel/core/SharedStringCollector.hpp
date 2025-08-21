#pragma once

#include <string>
#include <vector>
#include <unordered_set>
#include <memory>
#include <functional>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;
class Worksheet;
class SharedStringTable;
class Cell;

/**
 * @brief 共享字符串收集器 - 预收集所有字符串，避免动态修改问题
 * 
 * 解决的核心问题：
 * 1. XML生成时动态修改SharedStringTable导致的索引错位
 * 2. 多次遍历导致的性能问题
 * 3. 内存占用优化
 */
class SharedStringCollector {
public:
    // 收集策略
    enum class CollectionStrategy {
        IMMEDIATE,      // 立即收集（默认）
        LAZY,          // 延迟收集
        INCREMENTAL,   // 增量收集
        PARALLEL       // 并行收集（未来功能）
    };
    
    // 收集统计
    struct CollectionStatistics {
        size_t total_strings = 0;        // 总字符串数
        size_t unique_strings = 0;       // 唯一字符串数
        size_t duplicate_strings = 0;    // 重复字符串数
        size_t memory_saved = 0;         // 节省的内存（字节）
        size_t collection_time_ms = 0;   // 收集耗时（毫秒）
        double deduplication_rate = 0.0; // 去重率
    };
    
    // 字符串过滤器
    using StringFilter = std::function<bool(const std::string&)>;
    
    // 字符串转换器
    using StringTransformer = std::function<std::string(const std::string&)>;

private:
    SharedStringTable* sst_;                          // 共享字符串表引用
    std::unordered_set<std::string> collected_set_;  // 已收集的字符串集合（去重）
    std::vector<std::string> collected_strings_;     // 收集的字符串列表（有序）
    
    CollectionStrategy strategy_;                    // 收集策略
    CollectionStatistics stats_;                     // 统计信息
    
    // 配置
    struct Configuration {
        bool enable_deduplication = true;            // 启用去重
        bool enable_compression = false;             // 启用压缩
        bool case_sensitive = true;                  // 大小写敏感
        size_t min_string_length = 0;               // 最小字符串长度
        size_t max_string_length = 32767;           // 最大字符串长度
        size_t batch_size = 1000;                   // 批处理大小
    } config_;
    
    // 过滤器和转换器
    std::vector<StringFilter> filters_;
    std::vector<StringTransformer> transformers_;

public:
    explicit SharedStringCollector(SharedStringTable* sst);
    ~SharedStringCollector();
    
    // 禁用拷贝，允许移动
    SharedStringCollector(const SharedStringCollector&) = delete;
    SharedStringCollector& operator=(const SharedStringCollector&) = delete;
    SharedStringCollector(SharedStringCollector&&) = default;
    SharedStringCollector& operator=(SharedStringCollector&&) = default;
    
    // 核心收集方法
    
    /**
     * @brief 从工作簿收集所有字符串
     * @param workbook 工作簿
     * @return 收集的字符串数量
     */
    size_t collectFromWorkbook(const Workbook* workbook);
    
    /**
     * @brief 从工作表收集字符串
     * @param worksheet 工作表
     * @return 收集的字符串数量
     */
    size_t collectFromWorksheet(const Worksheet* worksheet);
    
    /**
     * @brief 从单元格收集字符串
     * @param cell 单元格
     * @return 是否收集成功
     */
    bool collectFromCell(const Cell* cell);
    
    /**
     * @brief 添加单个字符串
     * @param str 字符串
     * @return 是否成功添加（false表示重复或被过滤）
     */
    bool addString(const std::string& str);
    
    /**
     * @brief 批量添加字符串
     * @param strings 字符串列表
     * @return 成功添加的数量
     */
    size_t addStrings(const std::vector<std::string>& strings);
    
    // 应用到 SharedStringTable
    
    /**
     * @brief 将收集的字符串应用到SharedStringTable
     * @param clear_existing 是否清除现有内容
     * @return 应用的字符串数量
     */
    size_t applyToSharedStringTable(bool clear_existing = true);
    
    /**
     * @brief 执行完整的收集和应用流程
     * @param workbook 工作簿
     * @return 处理的字符串数量
     */
    size_t collectAndApply(const Workbook* workbook);
    
    // 优化和清理
    
    /**
     * @brief 优化收集的字符串（去重、排序等）
     * @return 优化后的字符串数量
     */
    size_t optimize();
    
    /**
     * @brief 清空收集的字符串
     */
    void clear();
    
    /**
     * @brief 重置收集器状态
     */
    void reset();
    
    // 过滤和转换
    
    /**
     * @brief 添加字符串过滤器
     * @param filter 过滤器函数
     */
    void addFilter(StringFilter filter) {
        filters_.push_back(filter);
    }
    
    /**
     * @brief 添加字符串转换器
     * @param transformer 转换器函数
     */
    void addTransformer(StringTransformer transformer) {
        transformers_.push_back(transformer);
    }
    
    /**
     * @brief 清除所有过滤器
     */
    void clearFilters() { filters_.clear(); }
    
    /**
     * @brief 清除所有转换器
     */
    void clearTransformers() { transformers_.clear(); }
    
    // 查询和统计
    
    /**
     * @brief 获取收集的字符串数量
     * @return 字符串数量
     */
    size_t getCollectedCount() const { return collected_strings_.size(); }
    
    /**
     * @brief 获取唯一字符串数量
     * @return 唯一字符串数量
     */
    size_t getUniqueCount() const { return collected_set_.size(); }
    
    /**
     * @brief 检查字符串是否已收集
     * @param str 字符串
     * @return 是否已收集
     */
    bool isCollected(const std::string& str) const {
        return collected_set_.find(str) != collected_set_.end();
    }
    
    /**
     * @brief 获取收集的字符串列表
     * @return 字符串列表
     */
    const std::vector<std::string>& getCollectedStrings() const {
        return collected_strings_;
    }
    
    /**
     * @brief 获取统计信息
     * @return 统计信息
     */
    const CollectionStatistics& getStatistics() const { return stats_; }
    
    /**
     * @brief 获取配置
     * @return 配置
     */
    Configuration& getConfiguration() { return config_; }
    const Configuration& getConfiguration() const { return config_; }
    
    /**
     * @brief 设置收集策略
     * @param strategy 策略
     */
    void setStrategy(CollectionStrategy strategy) { strategy_ = strategy; }
    
    /**
     * @brief 获取收集策略
     * @return 策略
     */
    CollectionStrategy getStrategy() const { return strategy_; }

private:
    bool shouldCollect(const std::string& str) const;
    std::string transformString(const std::string& str) const;
    void updateStatistics();
    size_t calculateMemorySaved() const;
};

}} // namespace fastexcel::core
