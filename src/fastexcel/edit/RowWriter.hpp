#pragma once

#include "fastexcel/core/Cell.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <memory>
#include <vector>
#include <string>
#include <functional>

namespace fastexcel {

// 前向声明
namespace core {
    class Worksheet;
    class SharedStringTable;
}

namespace edit {

// 前向声明
class EditSession;
class EditWorksheet;

/**
 * @brief 行写入器
 * 
 * 提供流式写入能力，常量内存模式
 * - 直接写入ZIP流
 * - 不保留内存
 * - 高性能批量写入
 */
class RowWriter {
private:
    EditWorksheet* worksheet_;
    EditSession* session_;
    core::SharedStringTable* sst_;
    xml::XMLStreamWriter* xml_writer_;
    
    int current_row_ = 0;
    bool streaming_mode_ = true;
    size_t buffer_size_ = 1000;
    
    // 行缓冲（批量模式）
    struct RowData {
        int row_num;
        std::vector<core::Cell> cells;
    };
    std::vector<RowData> row_buffer_;
    
public:
    /**
     * @brief 构造函数
     * @param worksheet 工作表
     * @param session 编辑会话
     */
    RowWriter(EditWorksheet* worksheet, EditSession* session);
    
    /**
     * @brief 析构函数
     */
    ~RowWriter();
    
    /**
     * @brief 写入一行数据
     * @param row 行数据
     */
    void writeRow(const std::vector<core::Cell>& row);
    
    /**
     * @brief 写入一行字符串数据
     * @param row 字符串数组
     */
    void writeRow(const std::vector<std::string>& row);
    
    /**
     * @brief 写入一行数字数据
     * @param row 数字数组
     */
    void writeRow(const std::vector<double>& row);
    
    /**
     * @brief 写入一行混合数据
     * @tparam Args 参数类型
     * @param args 单元格值
     */
    template<typename... Args>
    void writeRow(Args&&... args) {
        std::vector<core::Cell> row;
        row.reserve(sizeof...(args));
        (row.emplace_back(std::forward<Args>(args)), ...);
        writeRow(row);
    }
    
    /**
     * @brief 开始新行
     * @param row_num 行号（可选，默认自动递增）
     */
    void beginRow(int row_num = -1);
    
    /**
     * @brief 写入单元格
     * @param col 列号
     * @param value 值
     */
    void writeCell(int col, const std::string& value);
    void writeCell(int col, double value);
    void writeCell(int col, bool value);
    void writeCell(int col, const core::Cell& cell);
    
    /**
     * @brief 结束当前行
     */
    void endRow();
    
    /**
     * @brief 刷新缓冲区
     */
    void flush();
    
    /**
     * @brief 设置流式模式
     * @param enable 是否启用流式模式
     */
    void setStreamingMode(bool enable) { streaming_mode_ = enable; }
    
    /**
     * @brief 设置缓冲区大小
     * @param size 缓冲区大小
     */
    void setBufferSize(size_t size) { buffer_size_ = size; }
    
    /**
     * @brief 获取当前行号
     * @return 当前行号
     */
    int getCurrentRow() const { return current_row_; }
    
    /**
     * @brief 获取写入的行数
     * @return 行数
     */
    size_t getRowCount() const { return current_row_; }
    
private:
    /**
     * @brief 写入行到XML流
     * @param row_data 行数据
     */
    void writeRowToStream(const RowData& row_data);
    
    /**
     * @brief 处理共享字符串
     * @param str 字符串
     * @return 字符串索引
     */
    int processSharedString(const std::string& str);
};

/**
 * @brief 批量行写入器
 * 
 * 用于高性能批量数据写入
 */
class BatchRowWriter {
private:
    std::vector<std::unique_ptr<RowWriter>> writers_;
    EditSession* session_;
    
public:
    /**
     * @brief 构造函数
     * @param session 编辑会话
     */
    explicit BatchRowWriter(EditSession* session);
    
    /**
     * @brief 添加工作表写入器
     * @param worksheet_name 工作表名称
     * @return 行写入器
     */
    RowWriter* addWorksheet(const std::string& worksheet_name);
    
    /**
     * @brief 并行写入数据
     * @tparam DataProvider 数据提供者类型
     * @param provider 数据提供者
     */
    template<typename DataProvider>
    void parallelWrite(DataProvider& provider) {
        // 使用线程池并行写入多个工作表
        // 实现略
    }
    
    /**
     * @brief 完成写入
     */
    void finish();
};

}} // namespace fastexcel::edit