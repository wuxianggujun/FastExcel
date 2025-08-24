/**
 * @file WorkbookMemoryManager.hpp
 * @brief Workbook统一内存管理器
 */

#pragma once

#include "CellMemoryPool.hpp"
#include "FormatMemoryPool.hpp"
#include "MultiSizePool.hpp"
#include <memory>

namespace fastexcel {
namespace memory {

/**
 * @brief Workbook统一内存管理器
 * 
 * 提供对所有内存池的统一管理，包括Cell对象池、FormatDescriptor对象池
 * 和字符串池。确保内存的高效分配和回收。
 */
class WorkbookMemoryManager {
private:
    std::unique_ptr<CellMemoryPool> cell_pool_;
    std::unique_ptr<FormatMemoryPool> format_pool_;
    // 字符串池移除：依赖共享字符串表进行落盘去重
    // 统一原始分配：多大小内存池
    std::unique_ptr<MultiSizePool> raw_pool_;
    
public:
    /**
     * @brief 构造函数，创建所有内存池
     */
    WorkbookMemoryManager() 
        : cell_pool_(std::make_unique<CellMemoryPool>())
        , format_pool_(std::make_unique<FormatMemoryPool>())
        , raw_pool_(std::make_unique<MultiSizePool>()) {
    }
    
    ~WorkbookMemoryManager() = default;
    
    // 禁止拷贝，允许移动
    WorkbookMemoryManager(const WorkbookMemoryManager&) = delete;
    WorkbookMemoryManager& operator=(const WorkbookMemoryManager&) = delete;
    WorkbookMemoryManager(WorkbookMemoryManager&&) = default;
    WorkbookMemoryManager& operator=(WorkbookMemoryManager&&) = default;
    
    /**
     * @brief 获取Cell内存池
     * @return CellMemoryPool引用
     */
    CellMemoryPool& getCellPool() { return *cell_pool_; }
    const CellMemoryPool& getCellPool() const { return *cell_pool_; }
    
    /**
     * @brief 获取FormatDescriptor内存池
     * @return FormatMemoryPool引用
     */
    FormatMemoryPool& getFormatPool() { return *format_pool_; }
    const FormatMemoryPool& getFormatPool() const { return *format_pool_; }
    
    /**
     * @brief 获取字符串内存池
     * @return StringMemoryPool引用
     */
    // 字符串池已移除

    /**
     * @brief 原始字节分配（高性能多大小池）。
     */
    void* allocateRaw(std::size_t size, std::size_t alignment = alignof(std::max_align_t)) {
        return raw_pool_ ? raw_pool_->allocate(size, alignment) : nullptr;
    }

    /**
     * @brief 原始字节释放（对应 allocateRaw）。
     */
    void deallocateRaw(void* ptr, std::size_t size, std::size_t alignment = alignof(std::max_align_t)) {
        if (!raw_pool_ || !ptr) return;
        raw_pool_->deallocate(ptr, size, alignment);
    }
    
    /**
     * @brief 创建优化的Cell对象
     * @param args Cell构造函数参数
     * @return unique_ptr<Cell>
     */
    template<typename... Args>
    ::fastexcel::pool_ptr<core::Cell> createOptimizedCell(Args&&... args) {
        return cell_pool_->createCell(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 创建优化的FormatDescriptor对象
     * @param args FormatDescriptor构造函数参数
     * @return unique_ptr<FormatDescriptor>
     */
    template<typename... Args>
    ::fastexcel::pool_ptr<core::FormatDescriptor> createOptimizedFormat(Args&&... args) {
        return format_pool_->createFormat(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 创建基于默认格式的FormatDescriptor对象
     * @return unique_ptr<FormatDescriptor>
     */
    ::fastexcel::pool_ptr<core::FormatDescriptor> createDefaultFormat() {
        return format_pool_->createDefaultFormat();
    }
    
    /**
     * @brief 将字符串加入字符串池
     * @param value 要池化的字符串
     * @return 指向池化字符串的指针
     */
    // 字符串池相关API移除
    
    /**
     * @brief 统一的内存统计信息
     */
    struct MemoryStatistics {
        CellMemoryPool::Statistics cell_stats;
        FormatMemoryPool::Statistics format_stats;
        // String pool removed
        
        // 总计信息
        size_t total_memory_usage = 0;
        size_t total_allocations = 0;
        size_t total_active_objects = 0;
    };
    
    /**
     * @brief 获取所有内存池的统计信息
     * @return 统一的内存统计信息
     */
    MemoryStatistics getMemoryStatistics() const {
        MemoryStatistics stats;
        
        stats.cell_stats = cell_pool_->getStatistics();
        stats.format_stats = format_pool_->getStatistics();
        // string stats removed
        
        // 计算总计
        stats.total_allocations = stats.cell_stats.total_allocations + 
                                 stats.format_stats.total_allocations;
        
        stats.total_active_objects = stats.cell_stats.active_objects + 
                                   stats.format_stats.active_objects;
        
        stats.total_memory_usage = stats.cell_stats.current_usage + 
                                 stats.format_stats.current_usage;
        
        return stats;
    }
    
    /**
     * @brief 预留所有内存池的空间
     * @param cell_capacity 预期的Cell数量
     * @param format_capacity 预期的FormatDescriptor数量
     * @param string_capacity 预期的字符串数量
     */
    void reserve(size_t cell_capacity, size_t format_capacity, size_t string_capacity) {
        cell_pool_->reserve(cell_capacity);
        format_pool_->reserve(format_capacity);
        (void)string_capacity; // no-op
    }
    
    /**
     * @brief 收缩所有内存池，释放未使用的内存
     */
    void shrinkAll() {
        cell_pool_->shrink();
        format_pool_->shrink();
        // no-op
        if (raw_pool_) raw_pool_->shrink();
    }
    
    /**
     * @brief 清空所有内存池
     */
    void clearAll() {
        cell_pool_->clear();
        format_pool_->clear();
        // no-op
        if (raw_pool_) raw_pool_->clear();
    }
    
    /**
     * @brief 清空所有内存池 (clearAll的别名)
     */
    void clear() {
        clearAll();
    }
    
    /**
     * @brief 获取内存使用效率报告
     * @return 格式化的效率报告字符串
     */
    std::string getEfficiencyReport() const {
        auto stats = getMemoryStatistics();
        
        std::string report = "=== Workbook Memory Efficiency Report ===\\n";
        report += "Cell Pool:\\n";
        report += "  - Active objects: " + std::to_string(stats.cell_stats.active_objects) + "\\n";
        report += "  - Peak usage: " + std::to_string(stats.cell_stats.peak_usage) + " bytes\\n";
        
        report += "Format Pool:\\n";
        report += "  - Active objects: " + std::to_string(stats.format_stats.active_objects) + "\\n";
        report += "  - Peak usage: " + std::to_string(stats.format_stats.peak_usage) + " bytes\\n";
        
        // String pool removed
        
        report += "Total:\\n";
        report += "  - Total allocations: " + std::to_string(stats.total_allocations) + "\\n";
        report += "  - Active objects: " + std::to_string(stats.total_active_objects) + "\\n";
        
        return report;
    }
};

}} // namespace fastexcel::memory
