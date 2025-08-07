#pragma once

#include <string>
#include <memory>
#include <cstdint>

namespace fastexcel {
namespace core {

// 前向声明
class Format;
class FormatDescriptor;

enum class CellType : uint8_t {
    Empty = 0,
    Number = 1,
    String = 2,
    Boolean = 3,
    Formula = 4,
    Date = 5,
    Error = 6,
    Hyperlink = 7,
    InlineString = 8    // 短字符串内联存储（内部使用）
};

class Cell {
private:
    // 使用位域压缩标志 - 借鉴libxlsxwriter的优化思路
    struct {
        CellType type : 4;           // 4位足够存储类型
        bool has_format : 1;         // 是否有格式
        bool has_hyperlink : 1;      // 是否有超链接
        bool has_formula_result : 1; // 公式是否有缓存结果
        uint8_t reserved : 1;        // 保留位
    } flags_;
    
    // 使用union节省内存 - 核心优化点
    union CellValue {
        double number;
        int32_t string_id;           // SST索引或内联字符串长度
        bool boolean;
        uint32_t error_code;
        char inline_string[16];      // 短字符串内联存储
        
        CellValue() : number(0.0) {}
        ~CellValue() {}
    } value_;
    
    // 可选字段指针（只在需要时分配） - 延迟分配策略
    struct ExtendedData {
        std::string* long_string;    // 长字符串
        std::string* formula;        // 公式
        std::string* hyperlink;      // 超链接
        Format* format;              // 格式（原始指针，避免shared_ptr开销）
        double formula_result;       // 公式计算结果
        
        ExtendedData() : long_string(nullptr), formula(nullptr),
                        hyperlink(nullptr), format(nullptr), formula_result(0.0) {}
    };
    
    ExtendedData* extended_;  // 只在需要时分配
    
    // 辅助方法
    void ensureExtended();
    void clearExtended();
    void initializeFlags();
    void deepCopyExtendedData(const Cell& other);
    void copyStringField(std::string*& dest, const std::string* src);
    void resetToEmpty();
    
public:
    Cell();
    ~Cell();
    
    // 便利构造函数
    explicit Cell(const std::string& value);
    explicit Cell(const char* value);
    explicit Cell(double value);
    explicit Cell(int value);
    explicit Cell(bool value);
    
    // 基本值设置
    void setValue(double value);
    void setValue(bool value);
    void setValue(const std::string& value);
    void setValue(const char* value) { setValue(std::string(value)); }  // 避免隐式转换到bool
    void setValue(int value) { setValue(static_cast<double>(value)); }
    
    // 公式设置
    void setFormula(const std::string& formula, double result = 0.0);
    
    // 获取值
    CellType getType() const {
        // 对外API统一：InlineString也显示为String
        return (flags_.type == CellType::InlineString) ? CellType::String : flags_.type;
    }
    
    // 内部方法：获取真实的类型（用于测试和内部逻辑）
    CellType getInternalType() const { return flags_.type; }
    double getNumberValue() const;
    bool getBooleanValue() const;
    std::string getStringValue() const;
    std::string getFormula() const;
    double getFormulaResult() const;
    
    // 格式操作 - 兼容原有API
    void setFormat(std::shared_ptr<Format> format);
    void setFormat(Format* format);  // 新增：直接指针版本
    std::shared_ptr<Format> getFormat() const;
    Format* getFormatPtr() const;    // 新增：获取原始指针
    
    // 新架构格式操作 - FormatDescriptor支持
    void setFormat(std::shared_ptr<const FormatDescriptor> format);
    std::shared_ptr<const FormatDescriptor> getFormatDescriptor() const;
    
    bool hasFormat() const { return flags_.has_format; }
    
    // 超链接操作
    void setHyperlink(const std::string& url);
    std::string getHyperlink() const;
    bool hasHyperlink() const { return flags_.has_hyperlink; }
    
    // 状态检查
    bool isEmpty() const { return flags_.type == CellType::Empty; }
    bool isNumber() const { return flags_.type == CellType::Number; }
    bool isString() const { return flags_.type == CellType::String || flags_.type == CellType::InlineString; }
    bool isBoolean() const { return flags_.type == CellType::Boolean; }
    bool isFormula() const { return flags_.type == CellType::Formula; }
    bool isDate() const { return flags_.type == CellType::Date; }
    
    // 清空
    void clear();
    
    // 内存使用统计 - 调试用
    size_t getMemoryUsage() const;
    
    // 移动语义优化
    Cell(Cell&& other) noexcept;
    Cell& operator=(Cell&& other) noexcept;
    
    // 拷贝构造和赋值（保持兼容性）
    Cell(const Cell& other);
    Cell& operator=(const Cell& other);

private:
    // 格式的shared_ptr持有者（为了兼容性）
    mutable std::shared_ptr<Format> format_holder_;
    mutable std::shared_ptr<const FormatDescriptor> format_descriptor_holder_;
};

}} // namespace fastexcel::core