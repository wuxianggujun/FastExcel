#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Format.hpp"
#include <cstring>
#include <stdexcept>

namespace fastexcel {
namespace core {

// 私有辅助方法：初始化标志位
void Cell::initializeFlags() {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    flags_.reserved = 0;
}

Cell::Cell() : extended_(nullptr) {
    initializeFlags();
}

// 便利构造函数实现 - 使用委托构造函数
Cell::Cell(const std::string& value) : Cell() {
    setValue(value);
}

Cell::Cell(const char* value) : Cell() {
    setValue(std::string(value));
}

Cell::Cell(double value) : Cell() {
    setValue(value);
}

Cell::Cell(int value) : Cell() {
    setValue(static_cast<double>(value));
}

Cell::Cell(bool value) : Cell() {
    setValue(value);
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
        flags_.type = CellType::InlineString;  // 短字符串使用InlineString类型
        // 内联存储，但不修改string_id，因为它们共享内存
        std::memset(value_.inline_string, 0, sizeof(value_.inline_string));
        std::strcpy(value_.inline_string, value.c_str());
        // 使用extended_来标记是否为内联字符串
        // 如果extended_为空，说明是内联字符串
    } else {
        flags_.type = CellType::String;  // 长字符串使用String类型
        // 长字符串存储
        ensureExtended();
        if (!extended_->long_string) {
            extended_->long_string = new std::string();
        }
        *extended_->long_string = value;
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
    if (flags_.type == CellType::String) {
        if (extended_ && extended_->long_string) {
            // 长字符串
            return *extended_->long_string;
        } else {
            // 内联字符串（extended_为空表示内联）
            return std::string(value_.inline_string);
        }
    } else if (flags_.type == CellType::InlineString) {
        // 内联字符串
        return std::string(value_.inline_string);
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
        if (extended_->long_string) {
            usage += sizeof(std::string) + extended_->long_string->capacity();
        }
        if (extended_->formula) {
            usage += sizeof(std::string) + extended_->formula->capacity();
            usage += sizeof(double);  // formula_result
        }
        if (extended_->hyperlink) {
            usage += sizeof(std::string) + extended_->hyperlink->capacity();
        }
        if (extended_->format) {
            usage += sizeof(Format*);  // 格式指针
        }
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

Cell::Cell(Cell&& other) noexcept : extended_(nullptr) {
    flags_ = other.flags_;
    value_ = other.value_;
    extended_ = other.extended_;
    format_holder_ = std::move(other.format_holder_);
    
    // 重置源对象
    other.resetToEmpty();
}

Cell& Cell::operator=(Cell&& other) noexcept {
    if (this != &other) {
        clearExtended();
        format_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        extended_ = other.extended_;
        format_holder_ = std::move(other.format_holder_);
        
        // 重置源对象
        other.resetToEmpty();
    }
    return *this;
}

// 私有辅助方法：深拷贝ExtendedData
void Cell::deepCopyExtendedData(const Cell& other) {
    if (other.extended_) {
        ensureExtended();
        copyStringField(extended_->long_string, other.extended_->long_string);
        copyStringField(extended_->formula, other.extended_->formula);
        copyStringField(extended_->hyperlink, other.extended_->hyperlink);
        extended_->format = other.extended_->format;
        extended_->formula_result = other.extended_->formula_result;
    }
}

// 私有辅助方法：拷贝字符串字段
void Cell::copyStringField(std::string*& dest, const std::string* src) {
    if (src) {
        if (!dest) {
            dest = new std::string();
        }
        *dest = *src;
    }
}

// 私有辅助方法：重置为空状态
void Cell::resetToEmpty() {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    extended_ = nullptr;
    format_holder_.reset();
}

Cell::Cell(const Cell& other) : extended_(nullptr) {
    flags_ = other.flags_;
    value_ = other.value_;
    format_holder_ = other.format_holder_;
    deepCopyExtendedData(other);
}

Cell& Cell::operator=(const Cell& other) {
    if (this != &other) {
        clearExtended();
        format_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        format_holder_ = other.format_holder_;
        deepCopyExtendedData(other);
    }
    return *this;
}

}} // namespace fastexcel::core