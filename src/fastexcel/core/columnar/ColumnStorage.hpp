#pragma once

#include "ColumnImplementations.hpp"
#include <memory>
#include <string_view>

namespace fastexcel {
namespace core {
namespace columnar {

// 列存储容器 - 管理单个列的所有数据
class ColumnStorage {
private:
    std::unique_ptr<ColumnBase> column_;
    uint32_t column_index_;
    
public:
    explicit ColumnStorage(uint32_t column_index);
    ~ColumnStorage() = default;
    
    // 禁用拷贝，只允许移动
    ColumnStorage(const ColumnStorage&) = delete;
    ColumnStorage& operator=(const ColumnStorage&) = delete;
    ColumnStorage(ColumnStorage&&) = default;
    ColumnStorage& operator=(ColumnStorage&&) = default;
    
    // 设置值 - 根据类型自动创建相应的列
    void setValue(uint32_t row, double value);
    void setValue(uint32_t row, uint32_t sst_index);
    void setValue(uint32_t row, bool value);
    void setValue(uint32_t row, std::string_view value);
    
    // 获取值 - 类型安全的访问
    template<typename T>
    T getValue(uint32_t row) const;
    
    bool hasValue(uint32_t row) const;
    ColumnType getColumnType() const;
    uint32_t getColumnIndex() const { return column_index_; }
    
    // 统计信息
    size_t getRowCount() const;
    size_t getMemoryUsage() const;
    bool isEmpty() const;
    void clear();
    
    // 内部访问 - 用于批量操作
    const ColumnBase* getColumnBase() const { return column_.get(); }
};

}}} // namespace fastexcel::core::columnar