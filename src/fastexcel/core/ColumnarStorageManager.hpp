/**
 * @file ColumnarStorageManager.hpp
 * @brief 列式存储管理器 - 专门处理列式数据存储和访问
 * 
 * 职责分离原则：
 * - Worksheet 负责总体协调和传统 Cell 对象管理
 * - ColumnarStorageManager 专门负责列式存储的数据管理
 * - WorksheetParser 根据模式选择调用不同的存储方式
 */

#pragma once

#include "fastexcel/core/WorkbookTypes.hpp"
#include <unordered_map>
#include <functional>
#include <memory>
#include <variant>
#include <cstdint>
#include <string>
#include <ctime>

namespace fastexcel {
namespace core {

/**
 * @brief 列式存储管理器
 * 
 * 专门负责列式数据的存储、访问和管理，从Worksheet中分离出来
 * 减轻Worksheet的职责，提高代码的可维护性
 */
class ColumnarStorageManager {
public:
    // 公式值结构
    struct FormulaValue {
        uint32_t formula_index;  // 公式在FormulaRepository中的索引
        double result;           // 公式计算结果
    };
    
    // 列式数据值的变体类型
    using ColumnarValueVariant = std::variant<
        std::monostate,    // 空值
        double,            // 数字值  
        uint32_t,          // 字符串SST索引
        bool,              // 布尔值
        FormulaValue,      // 公式值
        std::string        // 错误值
    >;

private:
    /**
     * @brief 列式数据存储结构
     * 
     * 按数据类型分别存储，提高内存局部性和访问效率
     */
    struct ColumnarData {
        // 按列存储不同类型的数据
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, double>> number_columns;        // 数字列
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, uint32_t>> string_columns;      // 字符串列(SST索引)
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, bool>> boolean_columns;         // 布尔列
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, double>> datetime_columns;      // 日期时间列(Excel序列号)
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, FormulaValue>> formula_columns; // 公式列
        std::unordered_map<uint32_t, std::unordered_map<uint32_t, std::string>> error_columns;    // 错误列
    };

    std::unique_ptr<ColumnarData> data_;
    const WorkbookOptions* options_;  // 优化选项引用

public:
    ColumnarStorageManager() = default;
    ~ColumnarStorageManager() = default;
    
    // 禁止拷贝，允许移动
    ColumnarStorageManager(const ColumnarStorageManager&) = delete;
    ColumnarStorageManager& operator=(const ColumnarStorageManager&) = delete;
    ColumnarStorageManager(ColumnarStorageManager&&) = default;
    ColumnarStorageManager& operator=(ColumnarStorageManager&&) = default;
    
    /**
     * @brief 启用列式存储模式
     * @param options 工作簿选项，用于列投影等优化
     */
    void enableColumnarStorage(const WorkbookOptions* options = nullptr);
    
    /**
     * @brief 检查是否启用了列式存储
     */
    bool isColumnarEnabled() const { return data_ != nullptr; }
    
    // 数据写入接口
    void setValue(uint32_t row, uint32_t col, double value);
    void setValue(uint32_t row, uint32_t col, uint32_t sst_index);
    void setValue(uint32_t row, uint32_t col, bool value);
    void setValue(uint32_t row, uint32_t col, const std::tm& datetime);
    void setFormula(uint32_t row, uint32_t col, uint32_t formula_index, double result);
    void setError(uint32_t row, uint32_t col, const std::string& error_code);
    
    // 数据查询接口
    bool hasValue(uint32_t row, uint32_t col) const;
    ColumnarValueVariant getValue(uint32_t row, uint32_t col) const;
    
    // 列遍历接口
    void forEachInColumn(uint32_t col, std::function<void(uint32_t row, const ColumnarValueVariant& value)> callback) const;
    
    // 类型化列访问接口
    std::unordered_map<uint32_t, double> getNumberColumn(uint32_t col) const;
    std::unordered_map<uint32_t, uint32_t> getStringColumn(uint32_t col) const;
    std::unordered_map<uint32_t, bool> getBooleanColumn(uint32_t col) const;
    std::unordered_map<uint32_t, double> getDateTimeColumn(uint32_t col) const;
    std::unordered_map<uint32_t, FormulaValue> getFormulaColumn(uint32_t col) const;
    std::unordered_map<uint32_t, std::string> getErrorColumn(uint32_t col) const;
    
    // 统计和管理接口
    size_t getDataCount() const;
    size_t getMemoryUsage() const;
    void clearData();
    
    // 数据范围接口
    uint32_t getFirstRow() const { return first_row_; }
    uint32_t getLastRow() const { return last_row_; }
    uint32_t getFirstColumn() const { return first_col_; }
    uint32_t getLastColumn() const { return last_col_; }
    bool hasData() const { return has_data_; }
    
private:
    /**
     * @brief 检查列是否应该被过滤
     */
    bool shouldSkipColumn(uint32_t col) const;
    
    /**
     * @brief 更新数据范围
     */
    void updateDataRange(uint32_t row, uint32_t col);
    
    // 数据范围跟踪
    uint32_t first_row_ = 0;
    uint32_t last_row_ = 0;
    uint32_t first_col_ = 0;
    uint32_t last_col_ = 0;
    bool has_data_ = false;
};

} // namespace core
} // namespace fastexcel