#pragma once

#include <string>
#include <vector>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;

/**
 * @brief 链式调用助手类
 * 
 * 提供流畅的链式调用API，简化连续的工作表操作
 * 
 * @example
 * worksheet.chain()
 *     .setValue("A1", std::string("Hello"))
 *     .setValue("B1", 123.45)
 *     .setValue("C1", true)
 *     .setColumnWidth(0, 15.0)
 *     .setRowHeight(0, 20.0)
 *     .mergeCells(1, 0, 1, 2);
 */
class WorksheetChain {
private:
    Worksheet& worksheet_;
    
public:
    /**
     * @brief 构造函数
     * @param worksheet 工作表引用
     */
    explicit WorksheetChain(Worksheet& worksheet) : worksheet_(worksheet) {}
    
    /**
     * @brief 设置单元格值（链式调用）
     * @tparam T 值类型
     * @param row 行号
     * @param col 列号
     * @param value 值
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setValue(int row, int col, const T& value);
    
    /**
     * @brief 通过Excel地址设置单元格值（链式调用）
     * @tparam T 值类型
     * @param address Excel地址
     * @param value 值
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setValue(const std::string& address, const T& value);
    
    /**
     * @brief 设置范围值（链式调用）
     * @tparam T 值类型
     * @param start_row 开始行
     * @param start_col 开始列
     * @param data 二维数据
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setRange(int start_row, int start_col, const std::vector<std::vector<T>>& data);
    
    /**
     * @brief 通过Excel地址设置范围值（链式调用）
     * @tparam T 值类型
     * @param range Excel范围地址
     * @param data 二维数据
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setRange(const std::string& range, const std::vector<std::vector<T>>& data);
    
    /**
     * @brief 设置列宽（链式调用）
     * @param col 列号
     * @param width 宽度
     * @return 链式调用对象
     */
    WorksheetChain& setColumnWidth(int col, double width);
    
    /**
     * @brief 设置行高（链式调用）
     * @param row 行号
     * @param height 高度
     * @return 链式调用对象
     */
    WorksheetChain& setRowHeight(int row, double height);
    
    /**
     * @brief 合并单元格（链式调用）
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @return 链式调用对象
     */
    WorksheetChain& mergeCells(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 获取原始Worksheet引用
     * @return Worksheet引用
     */
    Worksheet& worksheet() { return worksheet_; }
    const Worksheet& worksheet() const { return worksheet_; }
};

}} // namespace fastexcel::core