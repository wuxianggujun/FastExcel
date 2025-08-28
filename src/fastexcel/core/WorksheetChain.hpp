#pragma once

#include "CellAddress.hpp"
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
     * @param address 单元格地址（支持Address对象、字符串地址或坐标构造）
     * @param value 值
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setValue(const Address& address, const T& value);
    
    /**
     * @brief 设置范围值（链式调用）
     * @tparam T 值类型
     * @param range 单元格范围（支持CellRange对象、字符串范围或坐标构造）
     * @param data 二维数据
     * @return 链式调用对象
     */
    template<typename T>
    WorksheetChain& setRange(const CellRange& range, const std::vector<std::vector<T>>& data);
    
    /**
     * @brief 设置列宽（链式调用）
     * @param col 列号（可以是整数或Address对象）
     * @param width 宽度
     * @return 链式调用对象
     */
    WorksheetChain& setColumnWidth(const Address& col, double width);
    
    /**
     * @brief 设置行高（链式调用）
     * @param row 行号（可以是整数或Address对象）
     * @param height 高度
     * @return 链式调用对象
     */
    WorksheetChain& setRowHeight(const Address& row, double height);
    
    /**
     * @brief 合并单元格（链式调用）
     * @param range 合并范围（支持CellRange对象、字符串范围或坐标构造）
     * @return 链式调用对象
     */
    WorksheetChain& mergeCells(const CellRange& range);
    
    /**
     * @brief 获取原始Worksheet引用
     * @return Worksheet引用
     */
    Worksheet& worksheet() { return worksheet_; }
    const Worksheet& worksheet() const { return worksheet_; }
};

}} // namespace fastexcel::core