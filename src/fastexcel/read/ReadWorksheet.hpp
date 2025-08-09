#pragma once

#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include <string>
#include <memory>
#include <functional>
#include <utility>

namespace fastexcel {

// 前向声明
namespace reader {
    class XLSXReader;
}

namespace read {

/**
 * @brief 只读工作表接口
 * 
 * 提供高性能的只读访问
 * - 流式处理
 * - 范围读取
 * - 内存优化
 */
class IReadOnlyWorksheet {
public:
    virtual ~IReadOnlyWorksheet() = default;
    
    // 基本信息
    virtual std::string getName() const = 0;
    virtual size_t getRowCount() const = 0;
    virtual size_t getColumnCount() const = 0;
    
    // 数据读取
    virtual std::string readString(int row, int col) const = 0;
    virtual double readNumber(int row, int col) const = 0;
    virtual bool readBoolean(int row, int col) const = 0;
    virtual core::CellType getCellType(int row, int col) const = 0;
    
    // 范围操作
    virtual bool hasData(int row, int col) const = 0;
    virtual std::pair<int, int> getUsedRange() const = 0;
};

/**
 * @brief 只读工作表实现
 * 
 * 特点：
 * - 基于SAX流式解析
 * - 不保存完整数据在内存
 * - 支持范围读取
 */
class ReadWorksheet : public IReadOnlyWorksheet {
private:
    reader::XLSXReader* reader_;  // 不拥有，由ReadWorkbook管理
    std::string name_;
    std::string path_;
    mutable int row_count_ = -1;
    mutable int col_count_ = -1;
    mutable bool range_calculated_ = false;
    
public:
    /**
     * @brief 构造函数
     * @param reader XLSXReader指针
     * @param name 工作表名称
     */
    ReadWorksheet(reader::XLSXReader* reader, const std::string& name);
    
    /**
     * @brief 析构函数
     */
    ~ReadWorksheet() = default;
    
    // ========== IReadOnlyWorksheet 接口实现 ==========
    
    std::string getName() const override { return name_; }
    size_t getRowCount() const override;
    size_t getColumnCount() const override;
    
    std::string readString(int row, int col) const override;
    double readNumber(int row, int col) const override;
    bool readBoolean(int row, int col) const override;
    core::CellType getCellType(int row, int col) const override;
    
    bool hasData(int row, int col) const override;
    std::pair<int, int> getUsedRange() const override;
    
    // ========== 高性能读取方法 ==========
    
    /**
     * @brief 行迭代器
     * 
     * 提供SAX风格的流式行迭代
     */
    class RowIterator {
    private:
        ReadWorksheet* worksheet_;
        int current_row_;
        int max_row_;
        
    public:
        RowIterator(ReadWorksheet* worksheet, int start_row = 0, int end_row = -1);
        
        bool hasNext() const;
        void next();
        int currentRow() const { return current_row_; }
        
        /**
         * @brief 读取当前行的单元格
         * @param callback 回调函数 (col, cell)
         */
        void readCells(const std::function<void(int, const core::Cell&)>& callback);
    };
    
    /**
     * @brief 创建行迭代器
     * @param start_row 起始行（默认0）
     * @param end_row 结束行（默认-1表示到最后）
     * @return 行迭代器
     */
    RowIterator createRowIterator(int start_row = 0, int end_row = -1);
    
    /**
     * @brief 流式读取范围
     * @param row_first 起始行
     * @param row_last 结束行
     * @param col_first 起始列
     * @param col_last 结束列
     * @param callback 回调函数
     * @return 错误码
     */
    core::ErrorCode readRange(int row_first, int row_last,
                             int col_first, int col_last,
                             const std::function<void(int row, int col, const core::Cell&)>& callback);
    
    /**
     * @brief 检查工作表是否有效
     * @return 是否有效
     */
    bool isValid() const { return reader_ != nullptr && !name_.empty(); }
    
private:
    /**
     * @brief 计算使用范围
     */
    void calculateUsedRange() const;
    
    /**
     * @brief 读取单元格（内部方法）
     * @param row 行号
     * @param col 列号
     * @param cell 输出的单元格
     * @return 是否成功
     */
    bool readCellInternal(int row, int col, core::Cell& cell) const;
};

}} // namespace fastexcel::read