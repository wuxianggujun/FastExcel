#pragma once

#include "ColumnTypes.hpp"
#include <vector>
#include <string>
#include <string_view>

namespace fastexcel {
namespace core {
namespace columnar {

// 数值列 - 存储double值
class NumberColumn : public ColumnBase {
private:
    std::vector<double> values_;
    ValidityBitmap validity_;
    
public:
    ColumnType getType() const override { return ColumnType::Number; }
    
    void setValue(uint32_t row, double value);
    double getValue(uint32_t row) const;
    bool hasValue(uint32_t row) const;
    
    size_t getRowCount() const override;
    size_t getMemoryUsage() const override;
    void clear() override;
    bool isEmpty() const override;
    
    // 批量访问接口
    const std::vector<double>& getValues() const { return values_; }
    const ValidityBitmap& getValidityBitmap() const { return validity_; }
};

// 字符串索引列 - 存储SST索引
class StringIndexColumn : public ColumnBase {
private:
    std::vector<uint32_t> indices_;
    ValidityBitmap validity_;
    
public:
    ColumnType getType() const override { return ColumnType::SharedStringIndex; }
    
    void setValue(uint32_t row, uint32_t sst_index);
    uint32_t getValue(uint32_t row) const;
    bool hasValue(uint32_t row) const;
    
    size_t getRowCount() const override;
    size_t getMemoryUsage() const override;
    void clear() override;
    bool isEmpty() const override;
    
    // 批量访问接口
    const std::vector<uint32_t>& getIndices() const { return indices_; }
    const ValidityBitmap& getValidityBitmap() const { return validity_; }
};

// 布尔列
class BooleanColumn : public ColumnBase {
private:
    std::vector<uint8_t> values_;  // 使用uint8_t存储boolean，节省空间
    ValidityBitmap validity_;
    
public:
    ColumnType getType() const override { return ColumnType::Boolean; }
    
    void setValue(uint32_t row, bool value);
    bool getValue(uint32_t row) const;
    bool hasValue(uint32_t row) const;
    
    size_t getRowCount() const override;
    size_t getMemoryUsage() const override;
    void clear() override;
    bool isEmpty() const override;
    
    // 批量访问接口
    const std::vector<uint8_t>& getValues() const { return values_; }
    const ValidityBitmap& getValidityBitmap() const { return validity_; }
};

// 内联字符串列 - 对于短字符串
class InlineStringColumn : public ColumnBase {
private:
    std::vector<std::string> values_;
    ValidityBitmap validity_;
    
public:
    ColumnType getType() const override { return ColumnType::InlineString; }
    
    void setValue(uint32_t row, std::string_view value);
    std::string_view getValue(uint32_t row) const;
    bool hasValue(uint32_t row) const;
    
    size_t getRowCount() const override;
    size_t getMemoryUsage() const override;
    void clear() override;
    bool isEmpty() const override;
    
    // 批量访问接口
    const std::vector<std::string>& getValues() const { return values_; }
    const ValidityBitmap& getValidityBitmap() const { return validity_; }
};

}}} // namespace fastexcel::core::columnar