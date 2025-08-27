#pragma once

#include <vector>
#include <cstdint>

namespace fastexcel {
namespace core {
namespace columnar {

// 列数据类型枚举
enum class ColumnType : uint8_t {
    Empty = 0,
    Number = 1,
    SharedStringIndex = 2,  // SST索引
    Boolean = 3,
    InlineString = 4       // 短字符串内联
};

// 列的有效性位图 - 压缩存储哪些行有数据
class ValidityBitmap {
private:
    std::vector<uint64_t> bits_;
    uint32_t max_row_ = 0;
    
public:
    void setBit(uint32_t row);
    bool getBit(uint32_t row) const;
    void clear();
    uint32_t getMaxRow() const { return max_row_; }
    size_t getMemoryUsage() const;
};

// 基础列接口
class ColumnBase {
public:
    virtual ~ColumnBase() = default;
    virtual ColumnType getType() const = 0;
    virtual size_t getRowCount() const = 0;
    virtual size_t getMemoryUsage() const = 0;
    virtual void clear() = 0;
    virtual bool isEmpty() const = 0;
};

}}} // namespace fastexcel::core::columnar