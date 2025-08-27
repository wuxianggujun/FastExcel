#include "ColumnStorage.hpp"
#include <stdexcept>

namespace fastexcel {
namespace core {
namespace columnar {

// ColumnStorage 实现
ColumnStorage::ColumnStorage(uint32_t column_index) 
    : column_index_(column_index) {
}

void ColumnStorage::setValue(uint32_t row, double value) {
    if (!column_ || column_->getType() != ColumnType::Number) {
        column_ = std::make_unique<NumberColumn>();
    }
    static_cast<NumberColumn*>(column_.get())->setValue(row, value);
}

void ColumnStorage::setValue(uint32_t row, uint32_t sst_index) {
    if (!column_ || column_->getType() != ColumnType::SharedStringIndex) {
        column_ = std::make_unique<StringIndexColumn>();
    }
    static_cast<StringIndexColumn*>(column_.get())->setValue(row, sst_index);
}

void ColumnStorage::setValue(uint32_t row, bool value) {
    if (!column_ || column_->getType() != ColumnType::Boolean) {
        column_ = std::make_unique<BooleanColumn>();
    }
    static_cast<BooleanColumn*>(column_.get())->setValue(row, value);
}

void ColumnStorage::setValue(uint32_t row, std::string_view value) {
    if (!column_ || column_->getType() != ColumnType::InlineString) {
        column_ = std::make_unique<InlineStringColumn>();
    }
    static_cast<InlineStringColumn*>(column_.get())->setValue(row, value);
}

template<>
double ColumnStorage::getValue<double>(uint32_t row) const {
    if (!column_ || column_->getType() != ColumnType::Number) {
        return 0.0;
    }
    return static_cast<const NumberColumn*>(column_.get())->getValue(row);
}

template<>
uint32_t ColumnStorage::getValue<uint32_t>(uint32_t row) const {
    if (!column_ || column_->getType() != ColumnType::SharedStringIndex) {
        return 0;
    }
    return static_cast<const StringIndexColumn*>(column_.get())->getValue(row);
}

template<>
bool ColumnStorage::getValue<bool>(uint32_t row) const {
    if (!column_ || column_->getType() != ColumnType::Boolean) {
        return false;
    }
    return static_cast<const BooleanColumn*>(column_.get())->getValue(row);
}

template<>
std::string_view ColumnStorage::getValue<std::string_view>(uint32_t row) const {
    if (!column_ || column_->getType() != ColumnType::InlineString) {
        static const std::string empty_string;
        return empty_string;
    }
    return static_cast<const InlineStringColumn*>(column_.get())->getValue(row);
}

bool ColumnStorage::hasValue(uint32_t row) const {
    if (!column_) {
        return false;
    }
    
    switch (column_->getType()) {
        case ColumnType::Number:
            return static_cast<const NumberColumn*>(column_.get())->hasValue(row);
        case ColumnType::SharedStringIndex:
            return static_cast<const StringIndexColumn*>(column_.get())->hasValue(row);
        case ColumnType::Boolean:
            return static_cast<const BooleanColumn*>(column_.get())->hasValue(row);
        case ColumnType::InlineString:
            return static_cast<const InlineStringColumn*>(column_.get())->hasValue(row);
        default:
            return false;
    }
}

ColumnType ColumnStorage::getColumnType() const {
    if (!column_) {
        return ColumnType::Empty;
    }
    return column_->getType();
}

size_t ColumnStorage::getRowCount() const {
    if (!column_) {
        return 0;
    }
    return column_->getRowCount();
}

size_t ColumnStorage::getMemoryUsage() const {
    if (!column_) {
        return sizeof(*this);
    }
    return sizeof(*this) + column_->getMemoryUsage();
}

bool ColumnStorage::isEmpty() const {
    return !column_ || column_->isEmpty();
}

void ColumnStorage::clear() {
    if (column_) {
        column_->clear();
    }
}

}}} // namespace fastexcel::core::columnar