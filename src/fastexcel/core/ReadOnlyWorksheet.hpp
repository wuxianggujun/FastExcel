#pragma once

#include <string>
#include <memory>
#include <vector>
#include <unordered_map>
#include <functional>
#include <variant>
#include <cstdint>

namespace fastexcel {
namespace core {

// 前向声明
class ColumnarStorageManager;

/**
 * @brief 只读工作表类 - 专门用于只读操作的优化版本
 * 
 * 这个类只提供读取操作，使用列式存储优化提供最佳性能。
 * 完全避免Cell对象创建，直接操作列式数据结构。
 * 
 * 特点：
 * - 零Cell对象：完全绕过Cell对象创建
 * - 列式存储：数据按类型分列存储
 * - 类型安全：编译期避免编辑操作
 * - 高性能访问：支持列遍历和批量操作
 */
class ReadOnlyWorksheet {
public:
    // === 数据类型定义 ===
    
    /**
     * @brief 单元格数据变体类型 - 使用与ColumnarStorageManager相同的类型
     */
    using CellValue = std::variant<double, uint32_t, bool, std::string>;
    
    /**
     * @brief 列数据映射类型
     * Key: 行号, Value: 数据值
     */
    using ColumnData = std::unordered_map<uint32_t, CellValue>;
    
    /**
     * @brief 列遍历回调函数类型
     */
    using ColumnCallback = std::function<void(uint32_t row, const CellValue& value)>;
    
    // === 基本信息接口 ===
    
    /**
     * @brief 获取工作表名称
     */
    std::string getName() const;
    
    /**
     * @brief 获取使用范围（有数据的范围）
     * @return {max_row, max_col} 最大行列数
     */
    std::pair<int, int> getUsedRange() const;
    
    /**
     * @brief 获取完整使用范围
     * @return {first_row, first_col, last_row, last_col}
     */
    std::tuple<int, int, int, int> getUsedRangeFull() const;
    
    // === 列式数据访问接口 ===
    
    /**
     * @brief 获取指定列的数字数据
     * @param col 列索引
     * @return 该列所有数字数据的映射
     */
    std::unordered_map<uint32_t, double> getNumberColumn(uint32_t col) const;
    
    /**
     * @brief 获取指定列的字符串数据（SST索引）
     * @param col 列索引
     * @return 该列所有字符串SST索引的映射
     */
    std::unordered_map<uint32_t, uint32_t> getStringColumn(uint32_t col) const;
    
    /**
     * @brief 获取指定列的布尔数据
     * @param col 列索引
     * @return 该列所有布尔数据的映射
     */
    std::unordered_map<uint32_t, bool> getBooleanColumn(uint32_t col) const;
    
    /**
     * @brief 获取指定列的错误/内联字符串数据
     * @param col 列索引
     * @return 该列所有错误/文本数据的映射
     */
    std::unordered_map<uint32_t, std::string> getErrorColumn(uint32_t col) const;
    
    // === 列遍历接口 ===
    
    /**
     * @brief 遍历指定列的所有数据
     * @param col 列索引
     * @param callback 遍历回调函数
     */
    void forEachInColumn(uint32_t col, const ColumnCallback& callback) const;
    
    /**
     * @brief 遍历指定列的指定行范围数据
     * @param col 列索引
     * @param start_row 起始行
     * @param end_row 结束行（包含）
     * @param callback 遍历回调函数
     */
    void forEachInColumnRange(uint32_t col, uint32_t start_row, uint32_t end_row,
                             const ColumnCallback& callback) const;
    
    // === 批量数据接口 ===
    
    /**
     * @brief 批量获取多列数据
     * @param columns 列索引列表
     * @return 每列的数据映射
     */
    std::vector<ColumnData> getBatchColumns(const std::vector<uint32_t>& columns) const;
    
    /**
     * @brief 获取指定行范围内所有列的数据
     * @param start_row 起始行
     * @param end_row 结束行（包含）
     * @return 行数据映射，Key: 行号, Value: 列数据映射
     */
    std::unordered_map<uint32_t, std::unordered_map<uint32_t, CellValue>> 
    getRowRangeData(uint32_t start_row, uint32_t end_row) const;
    
    // === 统计信息接口 ===
    
    /**
     * @brief 获取列式数据点总数
     * @return 总数据点数量
     */
    size_t getColumnarDataCount() const;
    
    /**
     * @brief 获取列式存储内存使用量
     * @return 内存使用字节数
     */
    size_t getColumnarMemoryUsage() const;
    
    /**
     * @brief 检查是否启用列式模式
     * @return true如果启用列式存储，false否则
     */
    bool isColumnarMode() const;
    
    /**
     * @brief 获取工作表统计信息
     */
    struct Stats {
        size_t total_data_points;    // 总数据点数
        size_t memory_usage;         // 内存使用量
        size_t number_columns;       // 包含数字的列数
        size_t string_columns;       // 包含字符串的列数
        size_t boolean_columns;      // 包含布尔值的列数
        size_t error_columns;        // 包含错误/文本的列数
        std::pair<int, int> used_range;  // 使用范围
    };
    
    Stats getStats() const;
    
    // === 数据查询接口 ===
    
    /**
     * @brief 检查指定位置是否有数据
     * @param row 行索引
     * @param col 列索引
     * @return true如果有数据，false否则
     */
    bool hasDataAt(uint32_t row, uint32_t col) const;
    
    /**
     * @brief 获取指定列包含数据的行数
     * @param col 列索引
     * @return 该列的数据行数
     */
    size_t getColumnDataCount(uint32_t col) const;
    
    /**
     * @brief 获取包含数据的所有列索引
     * @return 列索引列表
     */
    std::vector<uint32_t> getDataColumns() const;
    
    // === 高级查询接口 ===
    
    /**
     * @brief 根据条件查询数据
     * @param col 列索引
     * @param predicate 查询条件函数
     * @return 符合条件的行数据映射
     */
    template<typename Predicate>
    std::unordered_map<uint32_t, CellValue> queryColumn(uint32_t col, Predicate predicate) const;
    
    /**
     * @brief 统计指定列中符合条件的数据数量
     * @param col 列索引
     * @param predicate 统计条件函数
     * @return 符合条件的数据数量
     */
    template<typename Predicate>
    size_t countColumn(uint32_t col, Predicate predicate) const;
    
    // 析构函数
    ~ReadOnlyWorksheet();

private:
    // 私有构造函数，只能由ReadOnlyWorkbook创建
    ReadOnlyWorksheet(const std::string& name, 
                     std::shared_ptr<ColumnarStorageManager> storage_manager,
                     int first_row, int first_col, int last_row, int last_col);
    
    // 工作表名称
    std::string name_;
    
    // 列式存储管理器 - 使用shared_ptr允许共享
    std::shared_ptr<ColumnarStorageManager> storage_manager_;
    
    // 使用范围信息
    int first_row_;
    int first_col_;
    int last_row_;
    int last_col_;
    
    // 禁用拷贝和赋值
    ReadOnlyWorksheet(const ReadOnlyWorksheet&) = delete;
    ReadOnlyWorksheet& operator=(const ReadOnlyWorksheet&) = delete;
    
    // 允许移动
    ReadOnlyWorksheet(ReadOnlyWorksheet&&) = default;
    ReadOnlyWorksheet& operator=(ReadOnlyWorksheet&&) = default;
    
    // ReadOnlyWorkbook可以创建ReadOnlyWorksheet
    friend class ReadOnlyWorkbook;
};

// === 模板方法实现 ===

template<typename Predicate>
std::unordered_map<uint32_t, ReadOnlyWorksheet::CellValue> 
ReadOnlyWorksheet::queryColumn(uint32_t col, Predicate predicate) const {
    std::unordered_map<uint32_t, CellValue> result;
    
    forEachInColumn(col, [&](uint32_t row, const CellValue& value) {
        if (predicate(value)) {
            result[row] = value;
        }
    });
    
    return result;
}

template<typename Predicate>
size_t ReadOnlyWorksheet::countColumn(uint32_t col, Predicate predicate) const {
    size_t count = 0;
    
    forEachInColumn(col, [&](uint32_t row, const CellValue& value) {
        if (predicate(value)) {
            ++count;
        }
    });
    
    return count;
}

}} // namespace fastexcel::core