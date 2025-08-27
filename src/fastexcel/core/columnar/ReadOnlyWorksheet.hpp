#pragma once

#include "ColumnStorage.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include <unordered_map>
#include <memory>
#include <string>
#include <string_view>
#include <optional>

namespace fastexcel {
namespace core {
namespace columnar {

// 只读模式的单元格值 - 轻量级值对象，不持有数据
struct ReadOnlyValue {
    enum Type : uint8_t {
        Empty = 0,
        Number = 1,
        String = 2,
        Boolean = 3
    };
    
    Type type = Empty;
    
    union {
        double number_value;
        uint32_t string_index;  // SST索引
        bool boolean_value;
    };
    
    // 构造函数
    ReadOnlyValue() : type(Empty), number_value(0.0) {}
    ReadOnlyValue(double value) : type(Number), number_value(value) {}
    ReadOnlyValue(uint32_t sst_index) : type(String), string_index(sst_index) {}
    ReadOnlyValue(bool value) : type(Boolean), boolean_value(value) {}
    
    bool isEmpty() const { return type == Empty; }
    bool isNumber() const { return type == Number; }
    bool isString() const { return type == String; }
    bool isBoolean() const { return type == Boolean; }
    
    double asNumber() const { return type == Number ? number_value : 0.0; }
    uint32_t asStringIndex() const { return type == String ? string_index : 0; }
    bool asBoolean() const { return type == Boolean ? boolean_value : false; }
};

// 列视图 - 提供对单列数据的高效访问
template<typename T>
class ColumnView {
private:
    const ColumnStorage* storage_;
    const SharedStringTable* sst_;  // 用于字符串索引解引用
    
public:
    ColumnView(const ColumnStorage* storage, const SharedStringTable* sst = nullptr)
        : storage_(storage), sst_(sst) {}
    
    // 获取指定行的值
    T getValue(uint32_t row) const;
    bool hasValue(uint32_t row) const;
    
    // 批量迭代器
    class Iterator {
    private:
        const ColumnStorage* storage_;
        const SharedStringTable* sst_;
        uint32_t current_row_;
        uint32_t max_row_;
        
    public:
        Iterator(const ColumnStorage* storage, const SharedStringTable* sst, uint32_t row)
            : storage_(storage), sst_(sst), current_row_(row) {
            max_row_ = storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0;
            skipToNextValid();
        }
        
        bool operator!=(const Iterator& other) const {
            return current_row_ != other.current_row_;
        }
        
        Iterator& operator++() {
            ++current_row_;
            skipToNextValid();
            return *this;
        }
        
        std::pair<uint32_t, T> operator*() const {
            return {current_row_, storage_->getValue<T>(current_row_)};
        }
        
    private:
        void skipToNextValid() {
            while (current_row_ < max_row_ && !storage_->hasValue(current_row_)) {
                ++current_row_;
            }
        }
    };
    
    Iterator begin() const {
        return Iterator(storage_, sst_, 0);
    }
    
    Iterator end() const {
        uint32_t max_row = storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0;
        return Iterator(storage_, sst_, max_row);
    }
    
    // 统计信息
    uint32_t getRowCount() const { 
        return storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0; 
    }
    bool isEmpty() const { return !storage_ || storage_->isEmpty(); }
};

// 字符串列视图特化 - 自动解引用SST
template<>
class ColumnView<std::string> {
private:
    const ColumnStorage* storage_;
    const SharedStringTable* sst_;
    
public:
    ColumnView(const ColumnStorage* storage, const SharedStringTable* sst)
        : storage_(storage), sst_(sst) {}
    
    std::string getValue(uint32_t row) const {
        if (!storage_ || !storage_->hasValue(row)) {
            return "";
        }
        
        if (storage_->getColumnType() == ColumnType::SharedStringIndex) {
            uint32_t index = storage_->getValue<uint32_t>(row);
            return sst_ ? sst_->getString(index) : "";
        } else if (storage_->getColumnType() == ColumnType::InlineString) {
            auto sv = storage_->getValue<std::string_view>(row);
            return std::string(sv);
        }
        
        return "";
    }
    
    bool hasValue(uint32_t row) const {
        return storage_ && storage_->hasValue(row);
    }
    
    // 批量迭代器 - 字符串特化版本
    class Iterator {
    private:
        const ColumnStorage* storage_;
        const SharedStringTable* sst_;
        uint32_t current_row_;
        uint32_t max_row_;
        
    public:
        Iterator(const ColumnStorage* storage, const SharedStringTable* sst, uint32_t row)
            : storage_(storage), sst_(sst), current_row_(row) {
            max_row_ = storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0;
            skipToNextValid();
        }
        
        bool operator!=(const Iterator& other) const {
            return current_row_ != other.current_row_;
        }
        
        Iterator& operator++() {
            ++current_row_;
            skipToNextValid();
            return *this;
        }
        
        std::pair<uint32_t, std::string> operator*() const {
            if (!storage_ || !storage_->hasValue(current_row_)) {
                return {current_row_, ""};
            }
            
            if (storage_->getColumnType() == ColumnType::SharedStringIndex) {
                uint32_t index = storage_->getValue<uint32_t>(current_row_);
                return {current_row_, sst_ ? sst_->getString(index) : ""};
            } else if (storage_->getColumnType() == ColumnType::InlineString) {
                auto sv = storage_->getValue<std::string_view>(current_row_);
                return {current_row_, std::string(sv)};
            }
            
            return {current_row_, ""};
        }
        
    private:
        void skipToNextValid() {
            while (current_row_ < max_row_ && (!storage_ || !storage_->hasValue(current_row_))) {
                ++current_row_;
            }
        }
    };
    
    Iterator begin() const {
        return Iterator(storage_, sst_, 0);
    }
    
    Iterator end() const {
        uint32_t max_row = storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0;
        return Iterator(storage_, sst_, max_row);
    }
    
    // 统计信息
    uint32_t getRowCount() const { 
        return storage_ ? static_cast<uint32_t>(storage_->getRowCount()) : 0; 
    }
    bool isEmpty() const { return !storage_ || storage_->isEmpty(); }
};

// 只读工作表 - 列式存储，零拷贝访问
class ReadOnlyWorksheet {
private:
    std::string name_;
    std::unordered_map<uint32_t, std::unique_ptr<ColumnStorage>> columns_;
    const SharedStringTable* sst_;
    const FormatRepository* format_repo_;
    
    // 使用范围缓存
    mutable std::optional<std::pair<uint32_t, uint32_t>> cached_used_range_;
    mutable bool used_range_dirty_ = true;
    
public:
    explicit ReadOnlyWorksheet(const std::string& name,
                              const SharedStringTable* sst = nullptr,
                              const FormatRepository* format_repo = nullptr);
    
    ~ReadOnlyWorksheet() = default;
    
    // 禁用拷贝，只允许移动
    ReadOnlyWorksheet(const ReadOnlyWorksheet&) = delete;
    ReadOnlyWorksheet& operator=(const ReadOnlyWorksheet&) = delete;
    ReadOnlyWorksheet(ReadOnlyWorksheet&&) = default;
    ReadOnlyWorksheet& operator=(ReadOnlyWorksheet&&) = default;
    
    // 基本属性
    const std::string& getName() const { return name_; }
    
    // 数据写入接口 - 仅在解析时使用
    void setValue(uint32_t row, uint32_t col, double value);
    void setValue(uint32_t row, uint32_t col, uint32_t sst_index);
    void setValue(uint32_t row, uint32_t col, bool value);
    void setValue(uint32_t row, uint32_t col, std::string_view value);
    
    // 只读访问接口
    ReadOnlyValue getValue(uint32_t row, uint32_t col) const;
    bool hasValue(uint32_t row, uint32_t col) const;
    
    // 类型化访问
    double getNumberValue(uint32_t row, uint32_t col) const;
    std::string getStringValue(uint32_t row, uint32_t col) const;
    bool getBooleanValue(uint32_t row, uint32_t col) const;
    
    // 列视图访问 - 高效批量处理
    template<typename T>
    ColumnView<T> getColumnView(uint32_t col) const {
        auto it = columns_.find(col);
        if (it != columns_.end()) {
            return ColumnView<T>(it->second.get(), sst_);
        }
        return ColumnView<T>(nullptr, sst_);
    }
    
    // 范围信息
    std::pair<uint32_t, uint32_t> getUsedRange() const;
    std::tuple<uint32_t, uint32_t, uint32_t, uint32_t> getUsedRangeFull() const;
    
    // 统计信息
    size_t getCellCount() const;
    size_t getColumnCount() const { return columns_.size(); }
    size_t getMemoryUsage() const;
    
    // 查找功能
    std::vector<std::pair<uint32_t, uint32_t>> findCells(const std::string& search_text,
                                                         bool match_case = false,
                                                         bool match_entire_cell = false) const;
    
    // 清理
    void clear();
    
private:
    ColumnStorage* getOrCreateColumn(uint32_t col);
    void invalidateUsedRangeCache() const { used_range_dirty_ = true; }
    void updateUsedRangeCache() const;
};

}}} // namespace fastexcel::core::columnar