#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include <memory>
#include <vector>
#include <string>
#include <mutex>
#include <functional>

namespace fastexcel {
namespace read {

// 前向声明
class ReadWorksheet;

/**
 * @brief 工作簿访问模式
 */
enum class WorkbookAccessMode {
    READ_ONLY,    // 只读访问
    EDITABLE,     // 可编辑
    CREATE_NEW    // 创建新文件
};

/**
 * @brief 只读工作簿接口
 * 
 * 提供轻量级、高性能的只读访问
 * - 延迟加载
 * - 流式处理
 * - 内存优化
 */
class IReadOnlyWorkbook {
public:
    virtual ~IReadOnlyWorkbook() = default;
    
    // 只读查询方法
    virtual size_t getWorksheetCount() const = 0;
    virtual std::vector<std::string> getWorksheetNames() const = 0;
    virtual std::shared_ptr<const ReadWorksheet> getWorksheet(const std::string& name) const = 0;
    virtual std::shared_ptr<const ReadWorksheet> getWorksheet(size_t index) const = 0;
    
    // 状态查询
    virtual WorkbookAccessMode getAccessMode() const = 0;
    virtual bool isReadOnly() const = 0;
    virtual std::string getFilename() const = 0;
    
    // 工作表查找
    virtual bool hasWorksheet(const std::string& name) const = 0;
    virtual int getWorksheetIndex(const std::string& name) const = 0;
};

/**
 * @brief 轻量级只读工作簿实现
 * 
 * 特点：
 * - 基于XLSXReader，直接从ZIP文件读取
 * - 延迟加载：只在需要时加载工作表数据
 * - 内存效率：不缓存不必要的数据
 * - 线程安全：多线程读取支持
 */
class ReadWorkbook : public IReadOnlyWorkbook {
private:
    std::unique_ptr<reader::XLSXReader> reader_;
    mutable std::vector<std::shared_ptr<ReadWorksheet>> worksheets_cache_;
    mutable std::mutex cache_mutex_;
    
    reader::WorkbookMetadata metadata_;
    std::vector<std::string> worksheet_names_;
    
    // 只读状态：明确不可变
    const WorkbookAccessMode access_mode_ = WorkbookAccessMode::READ_ONLY;
    const std::string filename_;
    
    // 两级缓存系统
    struct CacheEntry {
        std::string key;
        std::shared_ptr<void> data;
        std::chrono::steady_clock::time_point last_access;
    };
    
    // L1缓存：超快速环形缓冲（lock-free）
    mutable std::vector<CacheEntry> l1_cache_;
    mutable std::atomic<size_t> l1_index_{0};
    static constexpr size_t L1_CACHE_SIZE = 256;
    
    // L2缓存：LRU哈希表
    mutable std::unordered_map<std::string, CacheEntry> l2_cache_;
    mutable std::mutex l2_mutex_;
    static constexpr size_t L2_CACHE_SIZE = 10000;
    
public:
    /**
     * @brief 构造函数
     * @param path 文件路径
     */
    explicit ReadWorkbook(const core::Path& path);
    
    /**
     * @brief 析构函数
     */
    ~ReadWorkbook();
    
    // 禁用拷贝
    ReadWorkbook(const ReadWorkbook&) = delete;
    ReadWorkbook& operator=(const ReadWorkbook&) = delete;
    
    // 允许移动
    ReadWorkbook(ReadWorkbook&&) = default;
    ReadWorkbook& operator=(ReadWorkbook&&) = default;
    
    // ========== IReadOnlyWorkbook 接口实现 ==========
    
    WorkbookAccessMode getAccessMode() const override { 
        return access_mode_; 
    }
    
    bool isReadOnly() const override { 
        return true; 
    }
    
    std::string getFilename() const override {
        return filename_;
    }
    
    size_t getWorksheetCount() const override {
        return worksheet_names_.size();
    }
    
    std::vector<std::string> getWorksheetNames() const override {
        return worksheet_names_;
    }
    
    std::shared_ptr<const ReadWorksheet> getWorksheet(const std::string& name) const override;
    std::shared_ptr<const ReadWorksheet> getWorksheet(size_t index) const override;
    
    bool hasWorksheet(const std::string& name) const override;
    int getWorksheetIndex(const std::string& name) const override;
    
    // ========== 高性能读取方法 ==========
    
    /**
     * @brief 刷新工作簿（重新读取）
     * @return 是否成功
     */
    bool refresh();
    
    /**
     * @brief 流式读取范围
     * @param worksheet_name 工作表名称
     * @param row_first 起始行
     * @param row_last 结束行
     * @param col_first 起始列
     * @param col_last 结束列
     * @param callback 回调函数
     * @return 错误码
     */
    core::ErrorCode readRange(const std::string& worksheet_name,
                             int row_first, int row_last,
                             int col_first, int col_last,
                             const std::function<void(int row, int col, const core::Cell&)>& callback);
    
    /**
     * @brief 获取元数据
     * @return 工作簿元数据
     */
    const reader::WorkbookMetadata& getMetadata() const { return metadata_; }
    
private:
    /**
     * @brief 加载基本信息
     */
    void loadBasicInfo();
    
    /**
     * @brief 构建快速索引
     */
    void buildQuickIndex();
    
    /**
     * @brief 从缓存获取数据
     * @param key 缓存键
     * @return 缓存的数据（可能为nullptr）
     */
    template<typename T>
    std::shared_ptr<T> getFromCache(const std::string& key) const;
    
    /**
     * @brief 添加到缓存
     * @param key 缓存键
     * @param data 要缓存的数据
     */
    template<typename T>
    void addToCache(const std::string& key, std::shared_ptr<T> data) const;
};

}} // namespace fastexcel::read