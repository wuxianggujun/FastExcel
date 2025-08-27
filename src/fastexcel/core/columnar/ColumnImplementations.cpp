#include "ColumnImplementations.hpp"

namespace fastexcel {
namespace core {
namespace columnar {

// NumberColumn 实现
void NumberColumn::setValue(uint32_t row, double value) {
    if (row >= values_.size()) {
        values_.resize(row + 1, 0.0);
    }
    values_[row] = value;
    validity_.setBit(row);
}

double NumberColumn::getValue(uint32_t row) const {
    if (row >= values_.size() || !validity_.getBit(row)) {
        return 0.0;  // 默认值
    }
    return values_[row];
}

bool NumberColumn::hasValue(uint32_t row) const {
    return validity_.getBit(row);
}

size_t NumberColumn::getRowCount() const {
    return validity_.getMaxRow();
}

size_t NumberColumn::getMemoryUsage() const {
    return values_.capacity() * sizeof(double) + validity_.getMemoryUsage();
}

void NumberColumn::clear() {
    values_.clear();
    validity_.clear();
}

bool NumberColumn::isEmpty() const {
    return values_.empty();
}

// StringIndexColumn 实现
void StringIndexColumn::setValue(uint32_t row, uint32_t sst_index) {
    if (row >= indices_.size()) {
        indices_.resize(row + 1, 0);
    }
    indices_[row] = sst_index;
    validity_.setBit(row);
}

uint32_t StringIndexColumn::getValue(uint32_t row) const {
    if (row >= indices_.size() || !validity_.getBit(row)) {
        return 0;  // 默认值
    }
    return indices_[row];
}

bool StringIndexColumn::hasValue(uint32_t row) const {
    return validity_.getBit(row);
}

size_t StringIndexColumn::getRowCount() const {
    return validity_.getMaxRow();
}

size_t StringIndexColumn::getMemoryUsage() const {
    return indices_.capacity() * sizeof(uint32_t) + validity_.getMemoryUsage();
}

void StringIndexColumn::clear() {
    indices_.clear();
    validity_.clear();
}

bool StringIndexColumn::isEmpty() const {
    return indices_.empty();
}

// BooleanColumn 实现
void BooleanColumn::setValue(uint32_t row, bool value) {
    if (row >= values_.size()) {
        values_.resize(row + 1, 0);
    }
    values_[row] = value ? 1 : 0;
    validity_.setBit(row);
}

bool BooleanColumn::getValue(uint32_t row) const {
    if (row >= values_.size() || !validity_.getBit(row)) {
        return false;  // 默认值
    }
    return values_[row] != 0;
}

bool BooleanColumn::hasValue(uint32_t row) const {
    return validity_.getBit(row);
}

size_t BooleanColumn::getRowCount() const {
    return validity_.getMaxRow();
}

size_t BooleanColumn::getMemoryUsage() const {
    return values_.capacity() * sizeof(uint8_t) + validity_.getMemoryUsage();
}

void BooleanColumn::clear() {
    values_.clear();
    validity_.clear();
}

bool BooleanColumn::isEmpty() const {
    return values_.empty();
}

// InlineStringColumn 实现
void InlineStringColumn::setValue(uint32_t row, std::string_view value) {
    if (row >= values_.size()) {
        values_.resize(row + 1);
    }
    values_[row] = value;
    validity_.setBit(row);
}

std::string_view InlineStringColumn::getValue(uint32_t row) const {
    if (row >= values_.size() || !validity_.getBit(row)) {
        static const std::string empty_string;
        return empty_string;
    }
    return values_[row];
}

bool InlineStringColumn::hasValue(uint32_t row) const {
    return validity_.getBit(row);
}

size_t InlineStringColumn::getRowCount() const {
    return validity_.getMaxRow();
}

size_t InlineStringColumn::getMemoryUsage() const {
    size_t total = validity_.getMemoryUsage();
    total += values_.capacity() * sizeof(std::string);
    for (const auto& str : values_) {
        total += str.capacity();
    }
    return total;
}

void InlineStringColumn::clear() {
    values_.clear();
    validity_.clear();
}

bool InlineStringColumn::isEmpty() const {
    return values_.empty();
}

}}} // namespace fastexcel::core::columnar