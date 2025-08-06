#pragma once

#include "../domain/FormatDescriptor.hpp"
#include "../ui/StyleBuilder.hpp"
#include <string>
#include <memory>
#include <vector>
#include <variant>

namespace fastexcel {
namespace api {

// 前向声明
class Workbook;
class WorksheetImpl;

/**
 * @brief 单元格值类型
 */
using CellValue = std::variant<
    std::monostate,    // 空值
    std::string,       // 字符串
    double,            // 数字
    bool,              // 布尔值
    int64_t           // 整数（日期时间的内部表示）
>;

/**
 * @brief Excel工作表 - 重构后的高级API
 * 
 * 提供类型安全、高性能的单元格操作接口。
 * 所有样式操作使用样式ID，避免指针管理的复杂性。
 */
class Worksheet {
private:
    std::unique_ptr<WorksheetImpl> pImpl_;
    friend class Workbook;
    
    // 私有构造函数，只能通过Workbook创建
    Worksheet(const std::string& name, Workbook* parent_workbook);

public:
    ~Worksheet();
    
    // 禁用拷贝，允许移动
    Worksheet(const Worksheet&) = delete;
    Worksheet& operator=(const Worksheet&) = delete;
    Worksheet(Worksheet&&) = default;
    Worksheet& operator=(Worksheet&&) = default;
    
    // ========== 基本属性 ==========
    
    /**
     * @brief 获取工作表名称
     * @return 工作表名称
     */
    std::string getName() const;
    
    /**
     * @brief 设置工作表名称
     * @param name 新名称
     * @return 是否成功
     */
    bool setName(const std::string& name);
    
    /**
     * @brief 获取父工作簿
     * @return 工作簿指针
     */
    Workbook* getWorkbook() const;
    
    /**
     * @brief 获取工作表索引
     * @return 在工作簿中的索引
     */
    size_t getIndex() const;
    
    // ========== 单元格写入操作 ==========
    
    /**
     * @brief 写入字符串
     * @param row 行号（从0开始）
     * @param col 列号（从0开始）
     * @param value 字符串值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeString(size_t row, size_t col, const std::string& value, int style_id = 0);
    
    /**
     * @brief 写入数字
     * @param row 行号
     * @param col 列号
     * @param value 数字值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeNumber(size_t row, size_t col, double value, int style_id = 0);
    
    /**
     * @brief 写入整数
     * @param row 行号
     * @param col 列号
     * @param value 整数值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeInteger(size_t row, size_t col, int64_t value, int style_id = 0);
    
    /**
     * @brief 写入布尔值
     * @param row 行号
     * @param col 列号
     * @param value 布尔值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeBool(size_t row, size_t col, bool value, int style_id = 0);
    
    /**
     * @brief 写入日期时间
     * @param row 行号
     * @param col 列号
     * @param timestamp Unix时间戳
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeDateTime(size_t row, size_t col, int64_t timestamp, int style_id = 0);
    
    /**
     * @brief 写入公式
     * @param row 行号
     * @param col 列号
     * @param formula 公式字符串（不包含等号）
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeFormula(size_t row, size_t col, const std::string& formula, int style_id = 0);
    
    /**
     * @brief 写入空值（清空单元格）
     * @param row 行号
     * @param col 列号
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeBlank(size_t row, size_t col, int style_id = 0);
    
    /**
     * @brief 写入任意值（自动推断类型）
     * @param row 行号
     * @param col 列号
     * @param value 单元格值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool writeValue(size_t row, size_t col, const CellValue& value, int style_id = 0);
    
    // ========== 批量写入操作 ==========
    
    /**
     * @brief 写入行数据
     * @param row 行号
     * @param start_col 起始列号
     * @param values 值列表
     * @param style_id 样式ID（应用于所有单元格）
     * @return 写入的单元格数量
     */
    size_t writeRow(size_t row, size_t start_col, const std::vector<CellValue>& values, int style_id = 0);
    
    /**
     * @brief 写入行数据（不同样式）
     * @param row 行号
     * @param start_col 起始列号
     * @param values 值列表
     * @param style_ids 样式ID列表
     * @return 写入的单元格数量
     */
    size_t writeRow(size_t row, size_t start_col, const std::vector<CellValue>& values, 
                   const std::vector<int>& style_ids);
    
    /**
     * @brief 写入列数据
     * @param col 列号
     * @param start_row 起始行号
     * @param values 值列表
     * @param style_id 样式ID（应用于所有单元格）
     * @return 写入的单元格数量
     */
    size_t writeColumn(size_t col, size_t start_row, const std::vector<CellValue>& values, int style_id = 0);
    
    /**
     * @brief 写入二维数据
     * @param start_row 起始行号
     * @param start_col 起始列号
     * @param data 二维数据
     * @param style_id 样式ID（应用于所有单元格）
     * @return 写入的单元格数量
     */
    size_t writeRange(size_t start_row, size_t start_col, 
                     const std::vector<std::vector<CellValue>>& data, int style_id = 0);
    
    // ========== 单元格读取操作 ==========
    
    /**
     * @brief 读取单元格值
     * @param row 行号
     * @param col 列号
     * @return 单元格值
     */
    CellValue readValue(size_t row, size_t col) const;
    
    /**
     * @brief 读取单元格样式ID
     * @param row 行号
     * @param col 列号
     * @return 样式ID
     */
    int readStyleId(size_t row, size_t col) const;
    
    /**
     * @brief 检查单元格是否为空
     * @param row 行号
     * @param col 列号
     * @return 是否为空
     */
    bool isEmpty(size_t row, size_t col) const;
    
    // ========== 样式操作 ==========
    
    /**
     * @brief 设置单元格样式
     * @param row 行号
     * @param col 列号
     * @param style_id 样式ID
     * @return 是否成功
     */
    bool setCellStyle(size_t row, size_t col, int style_id);
    
    /**
     * @brief 设置范围样式
     * @param start_row 起始行号
     * @param start_col 起始列号
     * @param end_row 结束行号（不包含）
     * @param end_col 结束列号（不包含）
     * @param style_id 样式ID
     * @return 设置的单元格数量
     */
    size_t setRangeStyle(size_t start_row, size_t start_col, size_t end_row, size_t end_col, int style_id);
    
    /**
     * @brief 设置行样式
     * @param row 行号
     * @param style_id 样式ID
     * @return 设置的单元格数量
     */
    size_t setRowStyle(size_t row, int style_id);
    
    /**
     * @brief 设置列样式
     * @param col 列号
     * @param style_id 样式ID
     * @return 设置的单元格数量
     */
    size_t setColumnStyle(size_t col, int style_id);
    
    // ========== 行列操作 ==========
    
    /**
     * @brief 设置行高
     * @param row 行号
     * @param height 高度（点数）
     */
    void setRowHeight(size_t row, double height);
    
    /**
     * @brief 设置列宽
     * @param col 列号
     * @param width 宽度（字符数）
     */
    void setColumnWidth(size_t col, double width);
    
    /**
     * @brief 隐藏行
     * @param row 行号
     * @param hidden 是否隐藏
     */
    void setRowHidden(size_t row, bool hidden = true);
    
    /**
     * @brief 隐藏列
     * @param col 列号
     * @param hidden 是否隐藏
     */
    void setColumnHidden(size_t col, bool hidden = true);
    
    /**
     * @brief 插入行
     * @param row 插入位置
     * @param count 插入行数
     * @return 是否成功
     */
    bool insertRows(size_t row, size_t count = 1);
    
    /**
     * @brief 插入列
     * @param col 插入位置
     * @param count 插入列数
     * @return 是否成功
     */
    bool insertColumns(size_t col, size_t count = 1);
    
    /**
     * @brief 删除行
     * @param row 起始行号
     * @param count 删除行数
     * @return 是否成功
     */
    bool deleteRows(size_t row, size_t count = 1);
    
    /**
     * @brief 删除列
     * @param col 起始列号
     * @param count 删除列数
     * @return 是否成功
     */
    bool deleteColumns(size_t col, size_t count = 1);
    
    // ========== 合并单元格 ==========
    
    /**
     * @brief 合并单元格区域
     * @param start_row 起始行号
     * @param start_col 起始列号
     * @param end_row 结束行号
     * @param end_col 结束列号
     * @return 是否成功
     */
    bool mergeCells(size_t start_row, size_t start_col, size_t end_row, size_t end_col);
    
    /**
     * @brief 取消合并单元格
     * @param start_row 起始行号
     * @param start_col 起始列号
     * @param end_row 结束行号
     * @param end_col 结束列号
     * @return 是否成功
     */
    bool unmergeCells(size_t start_row, size_t start_col, size_t end_row, size_t end_col);
    
    // ========== 冻结窗格 ==========
    
    /**
     * @brief 冻结窗格
     * @param row 冻结行位置
     * @param col 冻结列位置
     */
    void freezePanes(size_t row, size_t col);
    
    /**
     * @brief 取消冻结窗格
     */
    void unfreezePanes();
    
    // ========== 工作表保护 ==========
    
    /**
     * @brief 保护工作表
     * @param password 密码（空则无密码保护）
     * @param options 保护选项
     */
    void protect(const std::string& password = "");
    
    /**
     * @brief 取消保护
     */
    void unprotect();
    
    /**
     * @brief 检查是否受保护
     * @return 是否受保护
     */
    bool isProtected() const;
    
    // ========== 工作表属性 ==========
    
    /**
     * @brief 设置工作表选中状态
     * @param selected 是否选中
     */
    void setSelected(bool selected);
    
    /**
     * @brief 检查是否选中
     * @return 是否选中
     */
    bool isSelected() const;
    
    /**
     * @brief 设置缩放比例
     * @param zoom 缩放比例（10-400）
     */
    void setZoom(int zoom);
    
    /**
     * @brief 获取缩放比例
     * @return 缩放比例
     */
    int getZoom() const;
    
    // ========== 数据范围 ==========
    
    /**
     * @brief 获取使用的行数
     * @return 行数
     */
    size_t getUsedRowCount() const;
    
    /**
     * @brief 获取使用的列数
     * @return 列数
     */
    size_t getUsedColumnCount() const;
    
    /**
     * @brief 获取最大行号
     * @return 最大行号（如果没有数据则返回0）
     */
    size_t getMaxRow() const;
    
    /**
     * @brief 获取最大列号
     * @return 最大列号（如果没有数据则返回0）
     */
    size_t getMaxColumn() const;
    
    /**
     * @brief 清空所有数据
     */
    void clear();
    
    /**
     * @brief 清空指定范围
     * @param start_row 起始行号
     * @param start_col 起始列号
     * @param end_row 结束行号
     * @param end_col 结束列号
     */
    void clearRange(size_t start_row, size_t start_col, size_t end_row, size_t end_col);
    
    // ========== 性能优化 ==========
    
    /**
     * @brief 启用批量写入模式
     * 
     * 在批量写入模式下，某些操作会被缓存以提高性能。
     * 完成批量操作后应调用endBatchWrite()。
     */
    void beginBatchWrite();
    
    /**
     * @brief 结束批量写入模式
     */
    void endBatchWrite();
    
    /**
     * @brief 预分配存储空间
     * @param rows 预期最大行数
     * @param cols 预期最大列数
     */
    void reserve(size_t rows, size_t cols);
    
    // ========== 便捷方法 ==========
    
    /**
     * @brief 使用A1格式写入值
     * @param cell_ref A1格式的单元格引用（如"A1"）
     * @param value 值
     * @param style_id 样式ID（可选）
     * @return 是否成功
     */
    bool write(const std::string& cell_ref, const CellValue& value, int style_id = 0);
    
    /**
     * @brief 使用A1格式读取值
     * @param cell_ref A1格式的单元格引用
     * @return 单元格值
     */
    CellValue read(const std::string& cell_ref) const;
    
    /**
     * @brief A1引用转换为行列坐标
     * @param cell_ref A1格式引用
     * @return {row, col} 坐标对
     */
    static std::pair<size_t, size_t> parseA1Reference(const std::string& cell_ref);
    
    /**
     * @brief 行列坐标转换为A1引用
     * @param row 行号
     * @param col 列号
     * @return A1格式引用
     */
    static std::string toA1Reference(size_t row, size_t col);
};

}} // namespace fastexcel::api