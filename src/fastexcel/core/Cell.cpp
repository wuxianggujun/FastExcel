#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Format.hpp"
#include <cstring>
#include <stdexcept>

namespace fastexcel {
namespace core {

Cell::Cell() : extended_(nullptr) {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    flags_.reserved = 0;
}

Cell::~Cell() {
    clearExtended();
}

void Cell::setValue(double value) {
    clear();
    flags_.type = CellType::Number;
    value_.number = value;
}

void Cell::setValue(bool value) {
    clear();
    flags_.type = CellType::Boolean;
    value_.boolean = value;
}

void Cell::setValue(const std::string& value) {
    clear();
    
    // 短字符串内联存储优化 - 借鉴libxlsxwriter的思路
    if (value.length() < sizeof(value_.inline_string) - 1) {  // 留一位给\0
        flags_.type = CellType::InlineString;
        std::strcpy(value_.inline_string, value.c_str());
        value_.string_id = static_cast<int32_t>(value.length());
    } else {
        flags_.type = CellType::String;
        ensureExtended();
        if (!extended_->long_string) {
            extended_->long_string = new std::string();
        }
        *extended_->long_string = value;
        // 这里可以实现SST索引逻辑，暂时设为0
        value_.string_id = 0;
    }
}

void Cell::setFormula(const std::string& formula, double result) {
    if (flags_.type != CellType::Formula) {
        clear();
        flags_.type = CellType::Formula;
    }
    
    ensureExtended();
    if (!extended_->formula) {
        extended_->formula = new std::string();
    }
    *extended_->formula = formula;
    extended_->formula_result = result;
    flags_.has_formula_result = true;
}

double Cell::getNumberValue() const {
    if (flags_.type == CellType::Number) {
        return value_.number;
    }
    if (flags_.type == CellType::Formula && flags_.has_formula_result && extended_) {
        return extended_->formula_result;
    }
    return 0.0;
}

bool Cell::getBooleanValue() const {
    return (flags_.type == CellType::Boolean) ? value_.boolean : false;
}

std::string Cell::getStringValue() const {
    switch (flags_.type) {
        case CellType::InlineString:
            return std::string(value_.inline_string, value_.string_id);
        case CellType::String:
            if (extended_ && extended_->long_string) {
                return *extended_->long_string;
            }
            break;
        default:
            break;
    }
    return "";
}

std::string Cell::getFormula() const {
    if (flags_.type == CellType::Formula && extended_ && extended_->formula) {
        return *extended_->formula;
    }
    return "";
}

double Cell::getFormulaResult() const {
    if (flags_.type == CellType::Formula && flags_.has_formula_result && extended_) {
        return extended_->formula_result;
    }
    return 0.0;
}

void Cell::setFormat(std::shared_ptr<Format> format) {
    if (format) {
        ensureExtended();
        extended_->format = format.get();
        format_holder_ = format;  // 保持shared_ptr引用
        flags_.has_format = true;
    } else {
        flags_.has_format = false;
        format_holder_.reset();
        if (extended_) {
            extended_->format = nullptr;
        }
    }
}

void Cell::setFormat(Format* format) {
    if (format) {
        ensureExtended();
        extended_->format = format;
        flags_.has_format = true;
        // 不设置format_holder_，表示外部管理生命周期
    } else {
        flags_.has_format = false;
        if (extended_) {
            extended_->format = nullptr;
        }
    }
}

std::shared_ptr<Format> Cell::getFormat() const {
    if (flags_.has_format && format_holder_) {
        return format_holder_;
    }
    return nullptr;
}

Format* Cell::getFormatPtr() const {
    return (flags_.has_format && extended_) ? extended_->format : nullptr;
}

void Cell::setHyperlink(const std::string& url) {
    if (!url.empty()) {
        ensureExtended();
        if (!extended_->hyperlink) {
            extended_->hyperlink = new std::string();
        }
        *extended_->hyperlink = url;
        flags_.has_hyperlink = true;
    } else {
        flags_.has_hyperlink = false;
        if (extended_ && extended_->hyperlink) {
            delete extended_->hyperlink;
            extended_->hyperlink = nullptr;
        }
    }
}

std::string Cell::getHyperlink() const {
    if (flags_.has_hyperlink && extended_ && extended_->hyperlink) {
        return *extended_->hyperlink;
    }
    return "";
}

void Cell::clear() {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    value_.number = 0.0;
    format_holder_.reset();
    clearExtended();
}

size_t Cell::getMemoryUsage() const {
    size_t usage = sizeof(Cell);
    if (extended_) {
        usage += sizeof(ExtendedData);
        if (extended_->long_string) usage += extended_->long_string->capacity();
        if (extended_->formula) usage += extended_->formula->capacity();
        if (extended_->hyperlink) usage += extended_->hyperlink->capacity();
    }
    return usage;
}

void Cell::ensureExtended() {
    if (!extended_) {
        extended_ = new ExtendedData();
    }
}

void Cell::clearExtended() {
    if (extended_) {
        delete extended_->long_string;
        delete extended_->formula;
        delete extended_->hyperlink;
        delete extended_;
        extended_ = nullptr;
    }
}

Cell::Cell(Cell&& other) noexcept {
    flags_ = other.flags_;
    value_ = other.value_;
    extended_ = other.extended_;
    format_holder_ = std::move(other.format_holder_);
    
    other.flags_.type = CellType::Empty;
    other.extended_ = nullptr;
}

Cell& Cell::operator=(Cell&& other) noexcept {
    if (this != &other) {
        clearExtended();
        format_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        extended_ = other.extended_;
        format_holder_ = std::move(other.format_holder_);
        
        other.flags_.type = CellType::Empty;
        other.extended_ = nullptr;
    }
    return *this;
}

Cell::Cell(const Cell& other) : extended_(nullptr) {
    flags_ = other.flags_;
    value_ = other.value_;
    format_holder_ = other.format_holder_;
    
    // 深拷贝ExtendedData
    if (other.extended_) {
        ensureExtended();
        if (other.extended_->long_string) {
            extended_->long_string = new std::string(*other.extended_->long_string);
        }
        if (other.extended_->formula) {
            extended_->formula = new std::string(*other.extended_->formula);
        }
        if (other.extended_->hyperlink) {
            extended_->hyperlink = new std::string(*other.extended_->hyperlink);
        }
        extended_->format = other.extended_->format;
        extended_->formula_result = other.extended_->formula_result;
    }
}

Cell& Cell::operator=(const Cell& other) {
    if (this != &other) {
        clearExtended();
        format_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        format_holder_ = other.format_holder_;
        
        // 深拷贝ExtendedData
        if (other.extended_) {
            ensureExtended();
            if (other.extended_->long_string) {
                extended_->long_string = new std::string(*other.extended_->long_string);
            }
            if (other.extended_->formula) {
                extended_->formula = new std::string(*other.extended_->formula);
            }
            if (other.extended_->hyperlink) {
                extended_->hyperlink = new std::string(*other.extended_->hyperlink);
            }
            extended_->format = other.extended_->format;
            extended_->formula_result = other.extended_->formula_result;
        }
    }
    return *this;
}

}} // namespace fastexcel::core