#pragma once

#include <memory>
#include <string>
#include <vector>

namespace fastexcel {

// 前向声明 - 将xml命名空间移到core外面
namespace xml {
class UnifiedXMLGenerator;
}

namespace core {

// 前向声明
class Workbook;
class ResourceManager;
class DirtyManager;
class IFileWriter;

/**
 * @brief 工作簿协调器 - 负责协调各个管理器完成工作簿操作
 * 
 * 设计原则：
 * 1. 单一职责原则(SRP)：只负责协调，不处理具体业务逻辑
 * 2. 依赖倒置原则(DIP)：依赖于抽象接口而非具体实现
 * 3. 组合优于继承：通过组合各个管理器实现功能
 * 
 * 性能优化：
 * - 延迟加载：只在需要时创建管理器实例
 * - 缓存机制：避免重复创建XML生成器
 * - 增量保存：只保存修改的部分
 */
class WorkbookCoordinator {
public:
    // 保存策略
    struct SaveStrategy {
        bool use_streaming = false;        // 使用流式写入（大文件）
        bool use_batch = true;             // 使用批量写入（默认）
        bool incremental = true;           // 增量保存
        bool validate_xml = false;         // 验证XML（调试用）
        int compression_level = 6;         // 压缩级别 (0-9)
        bool preserve_resources = true;    // 保留原有资源（图片等）
    };
    
    // 协调器配置
    struct Configuration {
        bool enable_caching = true;        // 启用缓存
        bool enable_lazy_loading = true;   // 延迟加载
        bool enable_parallel = false;      // 并行处理（未来功能）
        size_t batch_size = 100;          // 批处理大小
        size_t cache_size_mb = 50;        // 缓存大小（MB）
        int compression_level = 6;         // 压缩级别 (0-9)
    };

private:
    Workbook* workbook_;                                      // 工作簿引用（不拥有）
    std::unique_ptr<ResourceManager> resource_manager_;       // 资源管理器
    Configuration config_;                                    // 配置
    
    // 缓存的XML生成器（避免重复创建）
    mutable std::unique_ptr<::fastexcel::xml::UnifiedXMLGenerator> xml_generator_cache_;
    
    // 统计信息
    struct Statistics {
        size_t files_written = 0;
        size_t bytes_written = 0;
        size_t time_ms = 0;
        size_t cache_hits = 0;
        size_t cache_misses = 0;
    } stats_;

public:
    explicit WorkbookCoordinator(Workbook* workbook);
    ~WorkbookCoordinator();
    
    // 禁用拷贝，允许移动
    WorkbookCoordinator(const WorkbookCoordinator&) = delete;
    WorkbookCoordinator& operator=(const WorkbookCoordinator&) = delete;
    WorkbookCoordinator(WorkbookCoordinator&&) = default;
    WorkbookCoordinator& operator=(WorkbookCoordinator&&) = default;
    
    // 核心保存流程
    
    /**
     * @brief 执行保存操作
     * @param filename 文件名
     * @param strategy 保存策略
     * @return 是否成功
     */
    bool save(const std::string& filename, const SaveStrategy& strategy = SaveStrategy());
    
    /**
     * @brief 执行另存为操作
     * @param new_filename 新文件名
     * @param strategy 保存策略
     * @return 是否成功
     */
    bool saveAs(const std::string& new_filename, const SaveStrategy& strategy = SaveStrategy());
    
    /**
     * @brief 执行增量保存
     * @param dirty_manager 脏数据管理器
     * @return 是否成功
     */
    bool saveIncremental(const DirtyManager* dirty_manager);
    
    // XML生成协调
    
    /**
     * @brief 生成所有XML文件
     * @param writer 文件写入器
     * @return 是否成功
     */
    bool generateAllXML(IFileWriter& writer);
    
    /**
     * @brief 生成指定的XML文件
     * @param writer 文件写入器
     * @param parts 要生成的部件列表
     * @return 是否成功
     */
    bool generateSpecificXML(IFileWriter& writer, const std::vector<std::string>& parts);
    
    /**
     * @brief 获取或创建XML生成器
     * @return XML生成器
     */
    ::fastexcel::xml::UnifiedXMLGenerator* getOrCreateXMLGenerator();
    
    // 资源管理协调
    
    /**
     * @brief 获取资源管理器
     * @return 资源管理器
     */
    ResourceManager* getResourceManager() { return resource_manager_.get(); }
    const ResourceManager* getResourceManager() const { return resource_manager_.get(); }
    
    /**
     * @brief 准备编辑模式
     * @param original_path 原始文件路径
     * @return 是否成功
     */
    bool prepareForEditing(const std::string& original_path);
    
    /**
     * @brief 执行透传复制
     * @param source_path 源文件路径
     * @return 是否成功
     */
    bool performPassthroughCopy(const std::string& source_path);
    
    // 文件写入器工厂
    
    /**
     * @brief 创建文件写入器
     * @param use_streaming 是否使用流式写入
     * @return 文件写入器
     */
    std::unique_ptr<IFileWriter> createFileWriter(bool use_streaming = false);
    
    // 配置管理
    
    /**
     * @brief 设置配置
     * @param config 配置
     */
    void setConfiguration(const Configuration& config) { config_ = config; }
    
    /**
     * @brief 获取配置
     * @return 配置
     */
    const Configuration& getConfiguration() const { return config_; }
    
    // 统计信息
    
    /**
     * @brief 获取统计信息
     * @return 统计信息
     */
    const Statistics& getStatistics() const { return stats_; }
    
    /**
     * @brief 重置统计信息
     */
    void resetStatistics() { stats_ = Statistics(); }
    
    // 性能优化
    
    /**
     * @brief 预热缓存
     */
    void warmupCache();
    
    /**
     * @brief 清理缓存
     */
    void clearCache();
    
    /**
     * @brief 优化内存使用
     */
    void optimizeMemory();

private:
    // 内部辅助方法
    bool initializeResourceManager(const std::string& filename);
    bool performSave(const SaveStrategy& strategy);
    bool validateBeforeSave();
    void updateStatistics(size_t files, size_t bytes, size_t time_ms);
    std::vector<std::string> determinePartsToGenerate(const DirtyManager* dirty_manager);
    
    // 智能决策方法
    bool shouldUseStreaming(size_t estimated_size) const;
    bool shouldUseIncremental(const DirtyManager* dirty_manager) const;
    int determineOptimalCompressionLevel(size_t file_size) const;
};

}} // namespace fastexcel::core
