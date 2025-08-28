#include "fastexcel/core/SharedStringCollector.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Cell.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <chrono>
#include <algorithm>

namespace fastexcel {
namespace core {

SharedStringCollector::SharedStringCollector(SharedStringTable* sst)
    : sst_(sst)
    , strategy_(CollectionStrategy::IMMEDIATE) {
    if (!sst_) {
        throw std::invalid_argument("SharedStringTable cannot be null");
    }
}

SharedStringCollector::~SharedStringCollector() {
    clear();
}

// 核心收集方法

size_t SharedStringCollector::collectFromWorkbook(const Workbook* workbook) {
    if (!workbook) {
        FASTEXCEL_LOG_ERROR("Workbook is null");
        return 0;
    }
    
    auto start_time = std::chrono::steady_clock::now();
    size_t collected_count = 0;
    
    // 遍历所有工作表
    auto sheets = workbook->getAllSheets();
    for (const auto& sheet : sheets) {
        if (sheet) {
            collected_count += collectFromWorksheet(sheet.get());
        }
    }
    
    auto end_time = std::chrono::steady_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    stats_.collection_time_ms = duration.count();
    
    updateStatistics();
    
    FASTEXCEL_LOG_DEBUG("Collected {} strings from workbook in {} ms", 
              collected_count, stats_.collection_time_ms);
    
    return collected_count;
}

size_t SharedStringCollector::collectFromWorksheet(const Worksheet* worksheet) {
    if (!worksheet) {
        FASTEXCEL_LOG_ERROR("Worksheet is null");
        return 0;
    }
    
    size_t collected_count = 0;
    
    // 遍历所有单元格
    // 注意：这里需要根据实际的Worksheet接口调整
    size_t row_count = worksheet->getRowCount();
    size_t col_count = worksheet->getColumnCount();
    
    for (int row = 0; row < static_cast<int>(row_count); ++row) {
        for (int col = 0; col < static_cast<int>(col_count); ++col) {
            const Cell& cell = worksheet->getCell(core::Address(row, col));
            if (!cell.isEmpty() && collectFromCell(&cell)) {
                collected_count++;
            }
        }
    }
    
    FASTEXCEL_LOG_DEBUG("Collected {} strings from worksheet: {}", 
              collected_count, worksheet->getName());
    
    return collected_count;
}

bool SharedStringCollector::collectFromCell(const Cell* cell) {
    if (!cell) {
        return false;
    }
    
    // 只收集字符串类型的单元格
    if (!cell->isString()) {
        return false;
    }
    
    std::string value = cell->getValue<std::string>();
    if (value.empty()) {
        return false;
    }
    
    return addString(value);
}

bool SharedStringCollector::addString(const std::string& str) {
    // 检查是否应该收集
    if (!shouldCollect(str)) {
        return false;
    }
    
    // 应用转换器
    std::string processed = transformString(str);
    
    // 检查是否已存在（去重）
    if (config_.enable_deduplication) {
        if (collected_set_.find(processed) != collected_set_.end()) {
            stats_.duplicate_strings++;
            return false;
        }
        collected_set_.insert(processed);
    }
    
    // 添加到列表
    collected_strings_.push_back(processed);
    stats_.total_strings++;
    
    return true;
}

size_t SharedStringCollector::addStrings(const std::vector<std::string>& strings) {
    size_t added_count = 0;
    
    // 批量处理以提高性能
    if (strings.size() > config_.batch_size) {
        // 预分配空间
        collected_strings_.reserve(collected_strings_.size() + strings.size());
    }
    
    for (const auto& str : strings) {
        if (addString(str)) {
            added_count++;
        }
    }
    
    return added_count;
}

// 应用到 SharedStringTable

size_t SharedStringCollector::applyToSharedStringTable(bool clear_existing) {
    if (!sst_) {
        FASTEXCEL_LOG_ERROR("SharedStringTable is null");
        return 0;
    }
    
    if (clear_existing) {
        sst_->clear();
    }
    
    size_t applied_count = 0;
    
    // 批量添加到SharedStringTable
    for (const auto& str : collected_strings_) {
        sst_->addString(str);
        applied_count++;
    }
    
    FASTEXCEL_LOG_DEBUG("Applied {} strings to SharedStringTable", applied_count);
    
    return applied_count;
}

size_t SharedStringCollector::collectAndApply(const Workbook* workbook) {
    // 清空现有收集
    clear();
    
    // 收集字符串
    size_t collected = collectFromWorkbook(workbook);
    
    if (collected == 0) {
        FASTEXCEL_LOG_WARN("No strings collected from workbook");
        return 0;
    }
    
    // 优化收集的字符串
    optimize();
    
    // 应用到SharedStringTable
    size_t applied = applyToSharedStringTable(true);
    
    FASTEXCEL_LOG_INFO("Collected and applied {} strings (unique: {})", 
             applied, stats_.unique_strings);
    
    return applied;
}

// 优化和清理

size_t SharedStringCollector::optimize() {
    if (collected_strings_.empty()) {
        return 0;
    }
    
    size_t original_size = collected_strings_.size();
    
    // 如果启用去重且还没有去重
    if (config_.enable_deduplication && collected_set_.empty()) {
        std::vector<std::string> unique_strings;
        std::unordered_set<std::string> seen;
        
        for (const auto& str : collected_strings_) {
            if (seen.insert(str).second) {
                unique_strings.push_back(str);
            }
        }
        
        collected_strings_ = std::move(unique_strings);
        collected_set_ = std::move(seen);
    }
    
    // 可选：按字母顺序排序（可能有助于压缩）
    if (config_.enable_compression) {
        std::sort(collected_strings_.begin(), collected_strings_.end());
    }
    
    size_t optimized_size = collected_strings_.size();
    FASTEXCEL_LOG_DEBUG("Optimized strings from {} to {}", original_size, optimized_size);
    
    updateStatistics();
    
    return optimized_size;
}

void SharedStringCollector::clear() {
    collected_strings_.clear();
    collected_set_.clear();
    stats_ = CollectionStatistics();
}

void SharedStringCollector::reset() {
    clear();
    filters_.clear();
    transformers_.clear();
    config_ = Configuration();
    strategy_ = CollectionStrategy::IMMEDIATE;
}

// 私有辅助方法

bool SharedStringCollector::shouldCollect(const std::string& str) const {
    // 检查长度限制
    if (str.length() < config_.min_string_length || 
        str.length() > config_.max_string_length) {
        return false;
    }
    
    // 应用所有过滤器
    for (const auto& filter : filters_) {
        if (!filter(str)) {
            return false;
        }
    }
    
    return true;
}

std::string SharedStringCollector::transformString(const std::string& str) const {
    std::string result = str;
    
    // 应用所有转换器
    for (const auto& transformer : transformers_) {
        result = transformer(result);
    }
    
    // 大小写处理
    if (!config_.case_sensitive) {
        std::transform(result.begin(), result.end(), result.begin(), ::tolower);
    }
    
    return result;
}

void SharedStringCollector::updateStatistics() {
    stats_.unique_strings = collected_set_.empty() ? 
                           collected_strings_.size() : collected_set_.size();
    stats_.duplicate_strings = stats_.total_strings - stats_.unique_strings;
    
    if (stats_.total_strings > 0) {
        stats_.deduplication_rate = 
            static_cast<double>(stats_.duplicate_strings) / stats_.total_strings;
    }
    
    stats_.memory_saved = calculateMemorySaved();
}

size_t SharedStringCollector::calculateMemorySaved() const {
    if (stats_.duplicate_strings == 0) {
        return 0;
    }
    
    // 估算平均字符串长度
    size_t total_length = 0;
    size_t sample_size = std::min(size_t(100), collected_strings_.size());
    
    for (size_t i = 0; i < sample_size; ++i) {
        total_length += collected_strings_[i].length();
    }
    
    if (sample_size == 0) {
        return 0;
    }
    
    size_t avg_length = total_length / sample_size;
    
    // 计算节省的内存（重复字符串 * 平均长度）
    return stats_.duplicate_strings * avg_length;
}

}} // namespace fastexcel::core
