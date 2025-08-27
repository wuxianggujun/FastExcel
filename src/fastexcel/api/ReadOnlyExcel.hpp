#pragma once

#include "fastexcel/core/columnar/ReadOnlyWorkbook.hpp"
#include <string>
#include <string_view>
#include <functional>
#include <memory>

namespace fastexcel {
namespace api {

// 轻量级只读接口 - 对外暴露的简洁API
class ReadOnlyExcel {
private:
    std::unique_ptr<core::columnar::ReadOnlyWorkbook> workbook_;
    
public:
    // 构造函数
    explicit ReadOnlyExcel(std::unique_ptr<core::columnar::ReadOnlyWorkbook> workbook)
        : workbook_(std::move(workbook)) {}
    
    // 静态工厂方法 - 主要入口点
    static std::unique_ptr<ReadOnlyExcel> open(const std::string& filename) {
        auto workbook = core::columnar::ReadOnlyWorkbook::openReadOnly(filename);
        if (!workbook) {
            return nullptr;
        }
        return std::make_unique<ReadOnlyExcel>(std::move(workbook));
    }
    
    static std::unique_ptr<ReadOnlyExcel> openWithColumns(const std::string& filename,
                                                         const std::vector<uint32_t>& columns) {
        auto options = core::columnar::ReadOnlyOptions::columns(columns);
        auto workbook = core::columnar::ReadOnlyWorkbook::openReadOnly(filename, options);
        if (!workbook) {
            return nullptr;
        }
        return std::make_unique<ReadOnlyExcel>(std::move(workbook));
    }
    
    static std::unique_ptr<ReadOnlyExcel> openWithMaxRows(const std::string& filename,
                                                         uint32_t max_rows) {
        auto options = core::columnar::ReadOnlyOptions::maxRows(max_rows);
        auto workbook = core::columnar::ReadOnlyWorkbook::openReadOnly(filename, options);
        if (!workbook) {
            return nullptr;
        }
        return std::make_unique<ReadOnlyExcel>(std::move(workbook));
    }
    
    // 工作表访问
    size_t getSheetCount() const {
        return workbook_->getWorksheetCount();
    }
    
    std::vector<std::string> getSheetNames() const {
        return workbook_->getWorksheetNames();
    }
    
    // 基础数据访问 - 零拷贝获取
    std::string getString(size_t sheet_index, uint32_t row, uint32_t col) const {
        return workbook_->getValue<std::string>(sheet_index, row, col);
    }
    
    std::string getString(const std::string& sheet_name, uint32_t row, uint32_t col) const {
        return workbook_->getValue<std::string>(sheet_name, row, col);
    }
    
    double getNumber(size_t sheet_index, uint32_t row, uint32_t col) const {
        return workbook_->getValue<double>(sheet_index, row, col);
    }
    
    double getNumber(const std::string& sheet_name, uint32_t row, uint32_t col) const {
        return workbook_->getValue<double>(sheet_name, row, col);
    }
    
    bool getBoolean(size_t sheet_index, uint32_t row, uint32_t col) const {
        return workbook_->getValue<bool>(sheet_index, row, col);
    }
    
    bool getBoolean(const std::string& sheet_name, uint32_t row, uint32_t col) const {
        return workbook_->getValue<bool>(sheet_name, row, col);
    }
    
    // 高性能列迭代器 - 适合批量处理
    template<typename T, typename Func>
    void forEachInColumn(size_t sheet_index, uint32_t col, Func&& func) const {
        auto column_view = workbook_->getColumnView<T>(sheet_index, col);
        for (auto it = column_view.begin(); it != column_view.end(); ++it) {
            auto [row, value] = *it;
            func(row, value);
        }
    }
    
    template<typename T, typename Func>
    void forEachInColumn(const std::string& sheet_name, uint32_t col, Func&& func) const {
        auto column_view = workbook_->getColumnView<T>(sheet_name, col);
        for (auto it = column_view.begin(); it != column_view.end(); ++it) {
            auto [row, value] = *it;
            func(row, value);
        }
    }
    
    // 范围迭代 - 高效处理大量数据
    template<typename Func>
    void forEachCell(size_t sheet_index, Func&& func) const {
        auto* worksheet = workbook_->getWorksheet(sheet_index);
        if (!worksheet) return;
        
        auto [min_row, min_col, max_row, max_col] = worksheet->getUsedRangeFull();
        
        for (uint32_t row = min_row; row <= max_row; ++row) {
            for (uint32_t col = min_col; col <= max_col; ++col) {
                if (worksheet->hasValue(row, col)) {
                    auto value = worksheet->getValue(row, col);
                    func(row, col, value);
                }
            }
        }
    }
    
    template<typename Func>
    void forEachCell(const std::string& sheet_name, Func&& func) const {
        auto* worksheet = workbook_->getWorksheet(sheet_name);
        if (!worksheet) return;
        
        auto [min_row, min_col, max_row, max_col] = worksheet->getUsedRangeFull();
        
        for (uint32_t row = min_row; row <= max_row; ++row) {
            for (uint32_t col = min_col; col <= max_col; ++col) {
                if (worksheet->hasValue(row, col)) {
                    auto value = worksheet->getValue(row, col);
                    func(row, col, value);
                }
            }
        }
    }
    
    // 统计信息
    size_t getTotalCells() const {
        return workbook_->getTotalCellCount();
    }
    
    size_t getMemoryUsage() const {
        return workbook_->getTotalMemoryUsage();
    }
    
    // 查找功能
    std::vector<std::pair<uint32_t, uint32_t>> findCells(size_t sheet_index,
                                                         const std::string& search_text,
                                                         bool match_case = false,
                                                         bool match_entire_cell = false) const {
        auto* worksheet = workbook_->getWorksheet(sheet_index);
        if (!worksheet) {
            return {};
        }
        return worksheet->findCells(search_text, match_case, match_entire_cell);
    }
    
    std::vector<std::pair<uint32_t, uint32_t>> findCells(const std::string& sheet_name,
                                                         const std::string& search_text,
                                                         bool match_case = false,
                                                         bool match_entire_cell = false) const {
        auto* worksheet = workbook_->getWorksheet(sheet_name);
        if (!worksheet) {
            return {};
        }
        return worksheet->findCells(search_text, match_case, match_entire_cell);
    }
    
    // 便捷方法 - 获取范围信息
    std::tuple<uint32_t, uint32_t, uint32_t, uint32_t> getUsedRange(size_t sheet_index) const {
        auto* worksheet = workbook_->getWorksheet(sheet_index);
        if (!worksheet) {
            return {0, 0, 0, 0};
        }
        return worksheet->getUsedRangeFull();
    }
    
    std::tuple<uint32_t, uint32_t, uint32_t, uint32_t> getUsedRange(const std::string& sheet_name) const {
        auto* worksheet = workbook_->getWorksheet(sheet_name);
        if (!worksheet) {
            return {0, 0, 0, 0};
        }
        return worksheet->getUsedRangeFull();
    }
};

// 便捷的工厂函数
inline std::unique_ptr<ReadOnlyExcel> openExcelReadOnly(const std::string& filename) {
    return ReadOnlyExcel::open(filename);
}

inline std::unique_ptr<ReadOnlyExcel> openExcelReadOnly(const std::string& filename,
                                                       const std::vector<uint32_t>& columns) {
    return ReadOnlyExcel::openWithColumns(filename, columns);
}

inline std::unique_ptr<ReadOnlyExcel> openExcelReadOnly(const std::string& filename,
                                                       uint32_t max_rows) {
    return ReadOnlyExcel::openWithMaxRows(filename, max_rows);
}

}} // namespace fastexcel::api
