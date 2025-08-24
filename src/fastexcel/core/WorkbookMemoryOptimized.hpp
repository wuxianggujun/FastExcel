/**
 * @file WorkbookMemoryOptimized.hpp
 * @brief 内存优化版本的Workbook类，展示内存管理改进的使用
 */

#pragma once

#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/memory/MemoryPoolOptimized.hpp"
#include "fastexcel/utils/SafeConstruction.hpp"
#include "fastexcel/utils/StringViewOptimized.hpp"
#include "fastexcel/xml/XMLStreamWriterOptimized.hpp"
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 内存优化版本的Workbook类
 * 
 * 主要改进：
 * - 使用内存池管理Cell和FormatDescriptor对象
 * - 异常安全的构造模式
 * - string_view优化减少字符串分配
 * - 使用优化的XMLStreamWriter
 */
class WorkbookMemoryOptimized {
public:
    using CellPool = memory::FixedSizePool<Cell, 2048>;
    using FormatPool = memory::FixedSizePool<FormatDescriptor, 512>;
    
private:
    // 内存池实例（延迟初始化）
    utils::LazyInitializer<CellPool> cell_pool_;
    utils::LazyInitializer<FormatPool> format_pool_;
    
    // 原始Workbook实例
    std::unique_ptr<Workbook> workbook_;
    
    // 字符串池，避免重复字符串分配
    mutable utils::StringPool string_pool_;
    
    // 性能统计
    mutable size_t cell_allocations_ = 0;
    mutable size_t format_allocations_ = 0;
    mutable size_t string_optimizations_ = 0;

public:
    /**
     * @brief 异常安全的构造函数
     */
    explicit WorkbookMemoryOptimized(const std::string& filename = "") {
        // 使用异常安全构造模式
        auto constructor = [&](utils::ResourceManager& rm) -> std::unique_ptr<WorkbookMemoryOptimized> {
            auto instance = std::make_unique<WorkbookMemoryOptimized>();
            
            // 初始化内存池
            instance->initializeMemoryPools();
            rm.addCleanup([&instance]() {
                if (instance) {
                    instance->cleanupMemoryPools();
                }
            });
            
            // 创建原始Workbook
            if (!filename.empty()) {
                instance->workbook_ = std::make_unique<Workbook>(filename);
            } else {
                instance->workbook_ = std::make_unique<Workbook>();
            }
            rm.addResource(instance->workbook_);
            
            // 成功构造，取消清理
            rm.release();
            return instance;
        };
        
        utils::SafeConstructor<WorkbookMemoryOptimized> safe_constructor;
        auto instance = safe_constructor
            .onSuccess([this](WorkbookMemoryOptimized& wb) {
                FASTEXCEL_LOG_DEBUG("WorkbookMemoryOptimized constructed successfully");
            })
            .onFailure([](const std::exception& e) {
                FASTEXCEL_LOG_ERROR("WorkbookMemoryOptimized construction failed: {}", e.what());
            })
            .construct(constructor);
        
        // 移动构造的结果到this
        *this = std::move(*instance);
    }
    
    /**
     * @brief 析构函数
     */
    ~WorkbookMemoryOptimized() {
        try {
            // 打印内存使用统计
            FASTEXCEL_LOG_DEBUG("WorkbookMemoryOptimized destroyed. "
                              "Cell allocations: {}, Format allocations: {}, "
                              "String optimizations: {}",
                              cell_allocations_, format_allocations_, string_optimizations_);
            
            cleanupMemoryPools();
        } catch (...) {
            // 析构函数中不抛出异常
        }
    }
    
    // 禁用拷贝，允许移动
    WorkbookMemoryOptimized(const WorkbookMemoryOptimized&) = delete;
    WorkbookMemoryOptimized& operator=(const WorkbookMemoryOptimized&) = delete;
    
    WorkbookMemoryOptimized(WorkbookMemoryOptimized&& other) noexcept
        : cell_pool_(std::move(other.cell_pool_))
        , format_pool_(std::move(other.format_pool_))
        , workbook_(std::move(other.workbook_))
        , string_pool_(std::move(other.string_pool_))
        , cell_allocations_(other.cell_allocations_)
        , format_allocations_(other.format_allocations_)
        , string_optimizations_(other.string_optimizations_) {
        
        // 重置被移动对象
        other.cell_allocations_ = 0;
        other.format_allocations_ = 0;
        other.string_optimizations_ = 0;
    }
    
    WorkbookMemoryOptimized& operator=(WorkbookMemoryOptimized&& other) noexcept {
        if (this != &other) {
            cleanupMemoryPools();
            
            cell_pool_ = std::move(other.cell_pool_);
            format_pool_ = std::move(other.format_pool_);
            workbook_ = std::move(other.workbook_);
            string_pool_ = std::move(other.string_pool_);
            cell_allocations_ = other.cell_allocations_;
            format_allocations_ = other.format_allocations_;
            string_optimizations_ = other.string_optimizations_;
            
            other.cell_allocations_ = 0;
            other.format_allocations_ = 0;
            other.string_optimizations_ = 0;
        }
        return *this;
    }
    
    /**
     * @brief 内存池优化的Cell创建
     */
    template<typename... Args>
    std::unique_ptr<Cell> createCell(Args&&... args) {
        FASTEXCEL_LAZY_INIT(cell_pool_, CellPool);
        
        Cell* cell = cell_pool_.get().allocate(std::forward<Args>(args)...);
        ++cell_allocations_;
        
        return std::unique_ptr<Cell>(cell, [this](Cell* p) {
            if (cell_pool_.isInitialized()) {
                cell_pool_.get().deallocate(p);
            }
        });
    }
    
    /**
     * @brief 内存池优化的FormatDescriptor创建
     */
    template<typename... Args>
    std::unique_ptr<FormatDescriptor> createFormat(Args&&... args) {
        FASTEXCEL_LAZY_INIT(format_pool_, FormatPool);
        
        FormatDescriptor* format = format_pool_.get().allocate(std::forward<Args>(args)...);
        ++format_allocations_;
        
        return std::unique_ptr<FormatDescriptor>(format, [this](FormatDescriptor* p) {
            if (format_pool_.isInitialized()) {
                format_pool_.get().deallocate(p);
            }
        });
    }
    
    /**
     * @brief 字符串优化的写入方法
     */
    void setCellValueOptimized(int row, int col, std::string_view value) {
        // 使用字符串池避免重复分配
        const std::string* pooled_string = string_pool_.intern(value);
        ++string_optimizations_;
        
        // 使用原始Workbook的方法，但传入池化的字符串
        if (workbook_) {
            workbook_->setCellValue(row, col, *pooled_string);
        }
    }
    
    /**
     * @brief 使用StringJoiner优化的复合值设置
     */
    void setCellComplexValue(int row, int col, 
                           const std::vector<std::string_view>& parts,
                           std::string_view separator = " ") {
        utils::StringViewOptimized::StringJoiner joiner(std::string(separator));
        
        for (const auto& part : parts) {
            joiner.add(part);
        }
        
        setCellValueOptimized(row, col, joiner.build());
    }
    
    /**
     * @brief 使用StringBuilder优化的格式化值设置
     */
    template<typename... Args>
    void setCellFormattedValue(int row, int col, const char* format, Args&&... args) {
        std::string formatted = utils::StringViewOptimized::format(format, std::forward<Args>(args)...);
        setCellValueOptimized(row, col, formatted);
    }
    
    /**
     * @brief 创建内存优化的XML写入器
     */
    std::unique_ptr<xml::XMLStreamWriterOptimized> createOptimizedXMLWriter(const std::string& filename = "") {
        if (filename.empty()) {
            return xml::XMLWriterFactory::createMemoryWriter();
        } else {
            return xml::XMLWriterFactory::createFileWriter(filename);
        }
    }
    
    /**
     * @brief 代理到原始Workbook的方法
     */
    void save(const std::string& filename) {
        if (workbook_) {
            workbook_->save(filename);
        }
    }
    
    Worksheet& createWorksheet(const std::string& name) {
        if (!workbook_) {
            throw core::OperationException(
                "Workbook not initialized", 
                "createWorksheet", 
                core::ErrorCode::InvalidOperation,
                __FILE__, __LINE__
            );
        }
        return workbook_->createWorksheet(name);
    }
    
    const Worksheet& getWorksheet(const std::string& name) const {
        if (!workbook_) {
            throw core::OperationException(
                "Workbook not initialized",
                "getWorksheet",
                core::ErrorCode::InvalidOperation,
                __FILE__, __LINE__
            );
        }
        return workbook_->getWorksheet(name);
    }
    
    /**
     * @brief 获取内存使用统计
     */
    struct MemoryStats {
        size_t cell_allocations;
        size_t format_allocations;
        size_t string_optimizations;
        size_t cell_pool_usage;
        size_t format_pool_usage;
        size_t string_pool_size;
    };
    
    MemoryStats getMemoryStats() const {
        MemoryStats stats{};
        stats.cell_allocations = cell_allocations_;
        stats.format_allocations = format_allocations_;
        stats.string_optimizations = string_optimizations_;
        
        if (cell_pool_.isInitialized()) {
            stats.cell_pool_usage = cell_pool_.get().getCurrentUsage();
        }
        
        if (format_pool_.isInitialized()) {
            stats.format_pool_usage = format_pool_.get().getCurrentUsage();
        }
        
        stats.string_pool_size = string_pool_.size();
        
        return stats;
    }
    
    /**
     * @brief 内存收缩（释放未使用的内存）
     */
    void shrinkMemory() {
        if (cell_pool_.isInitialized()) {
            cell_pool_.get().shrink();
        }
        
        if (format_pool_.isInitialized()) {
            format_pool_.get().shrink();
        }
        
        string_pool_.clear();
        
        FASTEXCEL_LOG_DEBUG("Memory shrinking completed");
    }

private:
    /**
     * @brief 私有构造函数（用于异常安全构造）
     */
    WorkbookMemoryOptimized() = default;
    
    void initializeMemoryPools() {
        // 内存池延迟初始化，只在需要时创建
        FASTEXCEL_LOG_DEBUG("Memory pools ready for initialization");
    }
    
    void cleanupMemoryPools() noexcept {
        try {
            // LazyInitializer会自动清理
            cell_pool_.reset();
            format_pool_.reset();
            string_pool_.clear();
        } catch (...) {
            // 忽略清理时的异常
        }
    }
};

/**
 * @brief 创建内存优化Workbook的工厂函数
 */
inline std::unique_ptr<WorkbookMemoryOptimized> createOptimizedWorkbook(const std::string& filename = "") {
    return FASTEXCEL_SAFE_CONSTRUCT(WorkbookMemoryOptimized, filename);
}

} // namespace core
} // namespace fastexcel