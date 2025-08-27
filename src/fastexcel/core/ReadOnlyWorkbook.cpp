/**
 * @file ReadOnlyWorkbook.cpp
 * @brief 只读工作簿实现 - 完全独立的架构
 */

#include "fastexcel/core/ReadOnlyWorkbook.hpp"
#include "fastexcel/core/ReadOnlyWorksheet.hpp"
#include "fastexcel/core/ColumnarStorageManager.hpp"
#include "fastexcel/reader/ReadOnlyXLSXReader.hpp"  // 使用专用只读解析器
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace core {

// 私有构造函数
ReadOnlyWorkbook::ReadOnlyWorkbook(std::vector<WorksheetInfo> worksheet_infos, 
                                   const WorkbookOptions& options)
    : worksheet_infos_(std::move(worksheet_infos)), options_(options) {
}

// 静态工厂方法 - 完全独立实现
std::unique_ptr<ReadOnlyWorkbook> ReadOnlyWorkbook::fromFile(const std::string& filepath) {
    WorkbookOptions default_options;
    return fromFile(filepath, default_options);
}

std::unique_ptr<ReadOnlyWorkbook> ReadOnlyWorkbook::fromFile(const std::string& filepath, 
                                                           const WorkbookOptions& options) {
    try {
        Path path(filepath);
        if (!path.exists()) {
            FASTEXCEL_LOG_ERROR("File not found for read-only access: {}", filepath);
            return nullptr;
        }
        
        FASTEXCEL_LOG_INFO("Loading read-only workbook: {}", filepath);
        
        // 使用专用只读解析器，直接解析到列式存储
        reader::ReadOnlyXLSXReader reader(path, &options);
        auto result = reader.parse();
        
        if (result != core::ErrorCode::Ok) {
            FASTEXCEL_LOG_ERROR("Failed to parse XLSX file in read-only mode: {}, error code: {}", 
                               filepath, static_cast<int>(result));
            return nullptr;
        }
        
        // 获取解析后的工作表信息
        auto worksheet_infos = reader.takeWorksheetInfos();
        
        if (worksheet_infos.empty()) {
            FASTEXCEL_LOG_WARN("No worksheets found in file: {}", filepath);
            return std::unique_ptr<ReadOnlyWorkbook>(new ReadOnlyWorkbook({}, options));
        }
        
        // 转换为ReadOnlyWorkbook的WorksheetInfo格式
        std::vector<WorksheetInfo> workbook_infos;
        workbook_infos.reserve(worksheet_infos.size());
        
        for (auto& info : worksheet_infos) {
            workbook_infos.emplace_back(
                std::move(info.name),
                info.storage_manager,
                info.first_row,
                info.first_col, 
                info.last_row,
                info.last_col
            );
        }
        
        FASTEXCEL_LOG_INFO("Successfully created read-only workbook with {} worksheets using dedicated parser", 
                          workbook_infos.size());
        
        // 创建ReadOnlyWorkbook实例
        return std::unique_ptr<ReadOnlyWorkbook>(new ReadOnlyWorkbook(std::move(workbook_infos), options));
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception while loading read-only workbook: {}, error: {}", filepath, e.what());
        return nullptr;
    }
}

// 基本信息接口
size_t ReadOnlyWorkbook::getSheetCount() const {
    return worksheet_infos_.size();
}

std::unique_ptr<ReadOnlyWorksheet> ReadOnlyWorkbook::getSheet(size_t index) const {
    if (index >= worksheet_infos_.size()) {
        return nullptr;
    }
    
    const auto& info = worksheet_infos_[index];
    
    // 创建ReadOnlyWorksheet，共享ColumnarStorageManager
    return std::unique_ptr<ReadOnlyWorksheet>(new ReadOnlyWorksheet(
        info.name, 
        info.storage_manager,  // shared_ptr可以安全复制
        info.first_row, 
        info.first_col, 
        info.last_row, 
        info.last_col
    ));
}

std::unique_ptr<ReadOnlyWorksheet> ReadOnlyWorkbook::getSheet(const std::string& name) const {
    for (size_t i = 0; i < worksheet_infos_.size(); ++i) {
        if (worksheet_infos_[i].name == name) {
            return getSheet(i);
        }
    }
    return nullptr;
}

std::vector<std::string> ReadOnlyWorkbook::getSheetNames() const {
    std::vector<std::string> names;
    names.reserve(worksheet_infos_.size());
    
    for (const auto& info : worksheet_infos_) {
        names.push_back(info.name);
    }
    
    return names;
}

bool ReadOnlyWorkbook::hasSheet(const std::string& name) const {
    for (const auto& info : worksheet_infos_) {
        if (info.name == name) {
            return true;
        }
    }
    return false;
}

size_t ReadOnlyWorkbook::getTotalMemoryUsage() const {
    size_t total_memory = 0;
    for (const auto& info : worksheet_infos_) {
        if (info.storage_manager) {
            total_memory += info.storage_manager->getMemoryUsage();
        }
    }
    return total_memory;
}

const WorkbookOptions& ReadOnlyWorkbook::getOptions() const {
    return options_;
}

ReadOnlyWorkbook::Stats ReadOnlyWorkbook::getStats() const {
    Stats stats{};
    
    stats.sheet_count = getSheetCount();
    stats.columnar_optimized = true;  // 只读工作簿总是使用列式优化
    
    // 遍历所有工作表收集统计信息
    for (const auto& info : worksheet_infos_) {
        if (info.storage_manager) {
            stats.total_data_points += info.storage_manager->getDataCount();
            stats.total_memory_usage += info.storage_manager->getMemoryUsage();
        }
    }
    
    // TODO: 获取共享字符串数量（需要保存SST引用）
    stats.sst_string_count = 0;
    
    return stats;
}

std::vector<ReadOnlyWorkbook::Stats> ReadOnlyWorkbook::getBatchStats(const std::vector<size_t>& sheet_indices) const {
    std::vector<Stats> results;
    results.reserve(sheet_indices.size());
    
    for (size_t index : sheet_indices) {
        Stats sheet_stats{};
        if (index < worksheet_infos_.size()) {
            const auto& info = worksheet_infos_[index];
            sheet_stats.sheet_count = 1;
            if (info.storage_manager) {
                sheet_stats.total_data_points = info.storage_manager->getDataCount();
                sheet_stats.total_memory_usage = info.storage_manager->getMemoryUsage();
            }
            sheet_stats.columnar_optimized = true;
        }
        results.push_back(sheet_stats);
    }
    
    return results;
}

bool ReadOnlyWorkbook::isColumnarOptimized() const {
    return true;  // 只读工作簿总是使用列式优化
}

ReadOnlyWorkbook::~ReadOnlyWorkbook() = default;

}} // namespace fastexcel::core
