#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/FormatDescriptor.hpp"
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
    flags_.is_shared_formula = false;
}

Cell::Cell() : extended_(nullptr) {
    initializeFlags();
}

// 便利构造函数实现 - 使用模板API
Cell::Cell(const std::string& value) : Cell() {
    setValue<std::string>(value);
}

Cell::Cell(const char* value) : Cell() {
    setValue<std::string>(std::string(value));
}

Cell::Cell(double value) : Cell() {
    setValue<double>(value);
}

Cell::Cell(int value) : Cell() {
    setValue<int>(value);
}

Cell::Cell(bool value) : Cell() {
    setValue<bool>(value);
}

Cell::~Cell() {
    clearExtended();
}

// ========== V3风格赋值运算 ==========

Cell& Cell::operator=(double value) {
    setValue<double>(value);
    return *this;
}

Cell& Cell::operator=(int value) {
    setValue<int>(value);
    return *this;
}

Cell& Cell::operator=(bool value) {
    setValue<bool>(value);
    return *this;
}

Cell& Cell::operator=(const std::string& value) {
    setValue<std::string>(value);
    return *this;
}

Cell& Cell::operator=(std::string_view value) {
    setValue<std::string>(std::string(value));
    return *this;
}

Cell& Cell::operator=(const char* value) {
    setValue<std::string>(std::string(value));
    return *this;
}

// ========== 内部实现方法（由模板API调用）==========

void Cell::setValueImpl(double value) {
    clear();
    flags_.type = CellType::Number;
    value_.number = value;
}

void Cell::setValueImpl(bool value) {
    clear();
    flags_.type = CellType::Boolean;
    value_.boolean = value;
}

void Cell::setValueImpl(const std::string& value) {
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
        extended_->long_string = std::make_unique<std::string>(value);
    }
}

void Cell::setFormula(const std::string& formula, double result) {
    if (flags_.type != CellType::Formula) {
        clear();
        flags_.type = CellType::Formula;
    }
    
    ensureExtended();
    extended_->formula = std::make_unique<std::string>(formula);
    extended_->formula_result = result;
    flags_.has_formula_result = true;
}

// ========== 内部get方法实现 ==========

double Cell::getNumberValue() const {
    if (flags_.type == CellType::Number) {
        return value_.number;
    }
    if ((flags_.type == CellType::Formula || flags_.type == CellType::SharedFormula) && 
        flags_.has_formula_result && extended_) {
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
    if ((flags_.type == CellType::Formula || flags_.type == CellType::SharedFormula) && 
        extended_ && extended_->formula) {
        return *extended_->formula;
    }
    return "";
}

double Cell::getFormulaResult() const {
    if ((flags_.type == CellType::Formula || flags_.type == CellType::SharedFormula) && 
        flags_.has_formula_result && extended_) {
        return extended_->formula_result;
    }
    return 0.0;
}

// ========== 共享公式支持方法 ==========

void Cell::setSharedFormula(int shared_index, double result) {
    // 如果当前不是公式类型，则清除并设置为共享公式
    if (flags_.type != CellType::Formula) {
        clear();
        flags_.type = CellType::SharedFormula;
    } else {
        // 如果已经是公式，转换为共享公式但保留公式文本
        flags_.type = CellType::SharedFormula;
    }
    
    flags_.is_shared_formula = true;
    
    ensureExtended();
    extended_->shared_formula_index = shared_index;
    extended_->formula_result = result;
    flags_.has_formula_result = true;
}

void Cell::setSharedFormulaReference(int shared_index) {
    clear();
    flags_.type = CellType::SharedFormula;
    flags_.is_shared_formula = true;
    
    ensureExtended();
    extended_->shared_formula_index = shared_index;
}

int Cell::getSharedFormulaIndex() const {
    if (flags_.type == CellType::SharedFormula && extended_) {
        return extended_->shared_formula_index;
    }
    return -1;
}

bool Cell::isSharedFormula() const {
    return flags_.type == CellType::SharedFormula && flags_.is_shared_formula;
}

// FormatDescriptor架构格式操作
void Cell::setFormat(std::shared_ptr<const FormatDescriptor> format) {
    if (format) {
        format_descriptor_holder_ = format;
        flags_.has_format = true;
    } else {
        flags_.has_format = false;
        format_descriptor_holder_.reset();
    }
}

std::shared_ptr<const FormatDescriptor> Cell::getFormatDescriptor() const {
    if (flags_.has_format && format_descriptor_holder_) {
        return format_descriptor_holder_;
    }
    return nullptr;
}

void Cell::setHyperlink(const std::string& url) {
    if (!url.empty()) {
        ensureExtended();
        if (!extended_->hyperlink) {
            extended_->hyperlink = std::make_unique<std::string>();
        }
        *extended_->hyperlink = url;
        flags_.has_hyperlink = true;
    } else {
        flags_.has_hyperlink = false;
        if (extended_ && extended_->hyperlink) {
            extended_->hyperlink.reset();
        }
    }
}

std::string Cell::getHyperlink() const {
    if (flags_.has_hyperlink && extended_ && extended_->hyperlink) {
        return *extended_->hyperlink;
    }
    return "";
}

// 批注（注释）
void Cell::setComment(const std::string& comment) {
    if (!comment.empty()) {
        ensureExtended();
        if (!extended_->comment) {
            extended_->comment = std::make_unique<std::string>();
        }
        *extended_->comment = comment;
    } else if (extended_ && extended_->comment) {
        extended_->comment.reset();
    }
}

std::string Cell::getComment() const {
    if (extended_ && extended_->comment) {
        return *extended_->comment;
    }
    return "";
}

void Cell::clear() {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    value_.number = 0.0;
    // 重置FormatDescriptor
    format_descriptor_holder_.reset();
    // smart pointer 自动清理
    extended_.reset();
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
        if (extended_->comment) {
            usage += sizeof(std::string) + extended_->comment->capacity();
        }
        // Format相关内存计算已移除，现在使用FormatDescriptor
    }
    return usage;
}

void Cell::ensureExtended() {
    if (!extended_) {
        extended_ = std::make_unique<ExtendedData>();
    }
}

void Cell::clearExtended() {
    // smart pointer会自动清理，不需要手动delete
    extended_.reset();
}

Cell::Cell(Cell&& other) noexcept : extended_(nullptr) {
    flags_ = other.flags_;
    value_ = other.value_;
    extended_ = std::move(other.extended_);
    format_descriptor_holder_ = std::move(other.format_descriptor_holder_);
    
    // 重置源对象
    other.resetToEmpty();
}

Cell& Cell::operator=(Cell&& other) noexcept {
    if (this != &other) {
        // smart pointer会自动清理
        extended_.reset();
        format_descriptor_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        extended_ = std::move(other.extended_);
        format_descriptor_holder_ = std::move(other.format_descriptor_holder_);
        
        // 重置源对象
        other.resetToEmpty();
    }
    return *this;
}

// 私有辅助方法：深拷贝ExtendedData
void Cell::deepCopyExtendedData(const Cell& other) {
    if (other.extended_) {
        ensureExtended();
        
        // 使用 smart pointer 安全拷贝
        if (other.extended_->long_string) {
            extended_->long_string = std::make_unique<std::string>(*other.extended_->long_string);
        }
        if (other.extended_->formula) {
            extended_->formula = std::make_unique<std::string>(*other.extended_->formula);
        }
        if (other.extended_->hyperlink) {
            extended_->hyperlink = std::make_unique<std::string>(*other.extended_->hyperlink);
        }
        if (other.extended_->comment) {
            extended_->comment = std::make_unique<std::string>(*other.extended_->comment);
        }
        
        // 拷贝基础数据
        extended_->formula_result = other.extended_->formula_result;
        extended_->shared_formula_index = other.extended_->shared_formula_index;
    }
}


// 私有辅助方法：重置为空状态
void Cell::resetToEmpty() {
    flags_.type = CellType::Empty;
    flags_.has_format = false;
    flags_.has_hyperlink = false;
    flags_.has_formula_result = false;
    extended_.reset();  // smart pointer 自动清理
    // 重置FormatDescriptor
    format_descriptor_holder_.reset();
}

Cell::Cell(const Cell& other) : extended_(nullptr) {
    flags_ = other.flags_;
    value_ = other.value_;
    format_descriptor_holder_ = other.format_descriptor_holder_;
    deepCopyExtendedData(other);
}

Cell& Cell::operator=(const Cell& other) {
    if (this != &other) {
        // smart pointer会自动清理
        extended_.reset();
        format_descriptor_holder_.reset();
        
        flags_ = other.flags_;
        value_ = other.value_;
        format_descriptor_holder_ = other.format_descriptor_holder_;
        deepCopyExtendedData(other);
    }
    return *this;
}

}} // namespace fastexcel::core