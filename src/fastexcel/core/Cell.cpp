#include "fastexcel/core/Cell.hpp"
#include <stdexcept>

namespace fastexcel {
namespace core {

void Cell::setValue(const std::string& value) {
    value_ = value;
    type_ = CellType::String;
    formula_.clear();
}

void Cell::setValue(double value) {
    value_ = value;
    type_ = CellType::Number;
    formula_.clear();
}

void Cell::setValue(bool value) {
    value_ = value;
    type_ = CellType::Boolean;
    formula_.clear();
}

void Cell::setFormula(const std::string& formula) {
    formula_ = formula;
    type_ = CellType::Formula;
    // 保持原有的值不变
}

std::string Cell::getStringValue() const {
    if (type_ == CellType::String) {
        return std::get<std::string>(value_);
    }
    return "";
}

double Cell::getNumberValue() const {
    if (type_ == CellType::Number) {
        return std::get<double>(value_);
    }
    return 0.0;
}

bool Cell::getBooleanValue() const {
    if (type_ == CellType::Boolean) {
        return std::get<bool>(value_);
    }
    return false;
}

void Cell::clear() {
    type_ = CellType::Empty;
    value_ = std::monostate{};
    formula_.clear();
    hyperlink_.clear();
    format_.reset();
}

Cell::Cell(const Cell& other) {
    type_ = other.type_;
    value_ = other.value_;
    format_ = other.format_;
    formula_ = other.formula_;
    hyperlink_ = other.hyperlink_;
}

Cell& Cell::operator=(const Cell& other) {
    if (this != &other) {
        type_ = other.type_;
        value_ = other.value_;
        format_ = other.format_;
        formula_ = other.formula_;
        hyperlink_ = other.hyperlink_;
    }
    return *this;
}

}} // namespace fastexcel::core