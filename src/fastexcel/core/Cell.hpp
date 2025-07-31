#pragma once

#include <string>
#include <memory>
#include <variant>

namespace fastexcel {
namespace core {

// 前向声明
class Format;

enum class CellType {
    Empty,
    String,
    Number,
    Boolean,
    Date,
    Formula,
    Error
};

class Cell {
private:
    CellType type_ = CellType::Empty;
    std::variant<std::monostate, std::string, double, bool> value_;
    std::shared_ptr<Format> format_;
    std::string formula_;
    
public:
    Cell() = default;
    ~Cell() = default;
    
    // 设置值
    void setValue(const std::string& value);
    void setValue(double value);
    void setValue(bool value);
    void setValue(int value) { setValue(static_cast<double>(value)); }
    void setFormula(const std::string& formula);
    
    // 获取值
    CellType getType() const { return type_; }
    std::string getStringValue() const;
    double getNumberValue() const;
    bool getBooleanValue() const;
    std::string getFormula() const { return formula_; }
    
    // 格式设置
    void setFormat(std::shared_ptr<Format> format) { format_ = format; }
    std::shared_ptr<Format> getFormat() const { return format_; }
    
    // 检查状态
    bool isEmpty() const { return type_ == CellType::Empty; }
    bool isString() const { return type_ == CellType::String; }
    bool isNumber() const { return type_ == CellType::Number; }
    bool isBoolean() const { return type_ == CellType::Boolean; }
    bool isFormula() const { return type_ == CellType::Formula; }
    
    // 清空单元格
    void clear();
    
    // 复制和赋值
    Cell(const Cell& other);
    Cell& operator=(const Cell& other);
    Cell(Cell&& other) noexcept = default;
    Cell& operator=(Cell&& other) noexcept = default;
};

}} // namespace fastexcel::core