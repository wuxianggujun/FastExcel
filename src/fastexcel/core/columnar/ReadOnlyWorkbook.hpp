#pragma once

#include "ReadOnlyWorksheet.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include <vector>
#include <memory>
#include <string>
#include <unordered_map>

namespace fastexcel {
namespace core {
namespace columnar {

// 只读模式选项
struct ReadOnlyOptions {
    std::vector<uint32_t> projected_columns;  // 只读取指定列，空则读取所有列
    std::vector<std::string> projected_column_names;  // 按名称指定列
    uint32_t max_rows = 0;                    // 最大读取行数，0表示无限制
    bool optimize_memory = true;              // 开启内存优化
    bool cache_strings = true;                // 缓存字符串访问
    
    ReadOnlyOptions() = default;
    
    // 便捷构造函数
    static ReadOnlyOptions columns(const std::vector<uint32_t>& cols) {
        ReadOnlyOptions opts;
        opts.projected_columns = cols;
        return opts;
    }
    
    static ReadOnlyOptions columnNames(const std::vector<std::string>& names) {
        ReadOnlyOptions opts;
        opts.projected_column_names = names;
        return opts;
    }
    
    static ReadOnlyOptions maxRows(uint32_t max_rows) {
        ReadOnlyOptions opts;
        opts.max_rows = max_rows;
        return opts;
    }
};

// 只读工作簿 - 专为大规模数据读取优化
class ReadOnlyWorkbook {
private:
    std::vector<std::unique_ptr<ReadOnlyWorksheet>> worksheets_;
    std::unordered_map<std::string, size_t> worksheet_name_index_;
    
    // 核心组件
    std::unique_ptr<SharedStringTable> sst_;
    std::unique_ptr<FormatRepository> format_repo_;
    
    // 选项
    ReadOnlyOptions options_;
    
    // 元数据
    std::string filename_;
    size_t total_memory_usage_ = 0;
    
public:
    explicit ReadOnlyWorkbook(const ReadOnlyOptions& options = ReadOnlyOptions());
    ~ReadOnlyWorkbook() = default;
    
    // 禁用拷贝，只允许移动
    ReadOnlyWorkbook(const ReadOnlyWorkbook&) = delete;
    ReadOnlyWorkbook& operator=(const ReadOnlyWorkbook&) = delete;
    ReadOnlyWorkbook(ReadOnlyWorkbook&&) = default;
    ReadOnlyWorkbook& operator=(ReadOnlyWorkbook&&) = default;
    
    // 静态工厂方法 - 主要入口点
    static std::unique_ptr<ReadOnlyWorkbook> openReadOnly(const std::string& filename,
                                                         const ReadOnlyOptions& options = ReadOnlyOptions());
    
    // 工作表访问
    size_t getWorksheetCount() const { return worksheets_.size(); }
    
    ReadOnlyWorksheet* getWorksheet(size_t index) const {
        return index < worksheets_.size() ? worksheets_[index].get() : nullptr;
    }
    
    ReadOnlyWorksheet* getWorksheet(const std::string& name) const {
        auto it = worksheet_name_index_.find(name);
        return it != worksheet_name_index_.end() ? worksheets_[it->second].get() : nullptr;
    }
    
    std::vector<std::string> getWorksheetNames() const {
        std::vector<std::string> names;
        names.reserve(worksheets_.size());
        for (const auto& ws : worksheets_) {
            names.push_back(ws->getName());
        }
        return names;
    }
    
    // 内部方法 - 供解析器使用
    ReadOnlyWorksheet* addWorksheet(const std::string& name);
    void setSharedStringTable(std::unique_ptr<SharedStringTable> sst);
    void setFormatRepository(std::unique_ptr<FormatRepository> format_repo);
    
    // 访问器
    const SharedStringTable* getSharedStringTable() const { return sst_.get(); }
    const FormatRepository* getFormatRepository() const { return format_repo_.get(); }
    const ReadOnlyOptions& getOptions() const { return options_; }
    
    // 统计信息
    size_t getTotalCellCount() const {
        size_t total = 0;
        for (const auto& ws : worksheets_) {
            total += ws->getCellCount();
        }
        return total;
    }
    
    size_t getTotalMemoryUsage() const {
        size_t total = sizeof(*this);
        
        if (sst_) {
            total += sst_->getMemoryUsage();
        }
        
        if (format_repo_) {
            total += format_repo_->getMemoryUsage();
        }
        
        for (const auto& ws : worksheets_) {
            total += ws->getMemoryUsage();
        }
        
        return total;
    }
    
    // 便捷访问方法
    template<typename T>
    T getValue(size_t worksheet_index, uint32_t row, uint32_t col) const {
        auto* ws = getWorksheet(worksheet_index);
        if (!ws) {
            if constexpr (std::is_same_v<T, std::string>) {
                return "";
            } else if constexpr (std::is_same_v<T, double>) {
                return 0.0;
            } else if constexpr (std::is_same_v<T, bool>) {
                return false;
            }
        }
        
        if constexpr (std::is_same_v<T, std::string>) {
            return ws->getStringValue(row, col);
        } else if constexpr (std::is_same_v<T, double>) {
            return ws->getNumberValue(row, col);
        } else if constexpr (std::is_same_v<T, bool>) {
            return ws->getBooleanValue(row, col);
        }
    }
    
    template<typename T>
    T getValue(const std::string& worksheet_name, uint32_t row, uint32_t col) const {
        auto* ws = getWorksheet(worksheet_name);
        if (!ws) {
            if constexpr (std::is_same_v<T, std::string>) {
                return "";
            } else if constexpr (std::is_same_v<T, double>) {
                return 0.0;
            } else if constexpr (std::is_same_v<T, bool>) {
                return false;
            }
        }
        
        if constexpr (std::is_same_v<T, std::string>) {
            return ws->getStringValue(row, col);
        } else if constexpr (std::is_same_v<T, double>) {
            return ws->getNumberValue(row, col);
        } else if constexpr (std::is_same_v<T, bool>) {
            return ws->getBooleanValue(row, col);
        }
    }
    
    // 批量访问 - 高性能列式迭代
    template<typename T>
    auto getColumnView(size_t worksheet_index, uint32_t col) const -> decltype(auto) {
        auto* ws = getWorksheet(worksheet_index);
        if (!ws) {
            static ReadOnlyWorksheet empty_ws("");
            return empty_ws.getColumnView<T>(col);
        }
        return ws->getColumnView<T>(col);
    }
    
    template<typename T>
    auto getColumnView(const std::string& worksheet_name, uint32_t col) const -> decltype(auto) {
        auto* ws = getWorksheet(worksheet_name);
        if (!ws) {
            static ReadOnlyWorksheet empty_ws("");
            return empty_ws.getColumnView<T>(col);
        }
        return ws->getColumnView<T>(col);
    }
    
private:
    void updateMemoryUsage() {
        total_memory_usage_ = getTotalMemoryUsage();
    }
};

}}} // namespace fastexcel::core::columnar