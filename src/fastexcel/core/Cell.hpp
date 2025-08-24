#pragma once

#include <string>
#include <string_view>
#include <optional>
#include <memory>
#include <cstdint>
#include <type_traits>

namespace fastexcel {
namespace xml {
    class WorksheetXMLGenerator; // 前向声明
}
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
    InlineString = 8,   // 短字符串内联存储（内部使用）
    SharedFormula = 9   // 共享公式类型
};

class Cell {
    friend class Worksheet;  // 让Worksheet能访问private方法
    friend class Workbook;   // 让Workbook能访问private方法
    friend class ::fastexcel::xml::WorksheetXMLGenerator;  // 让XML生成器能访问private方法
private:
    // 使用位域压缩标志 - 借鉴libxlsxwriter的优化思路
    struct {
        CellType type : 4;           // 4位足够存储类型
        bool has_format : 1;         // 是否有格式
        bool has_hyperlink : 1;      // 是否有超链接
        bool has_formula_result : 1; // 公式是否有缓存结果
        bool is_shared_formula : 1;  // 是否为共享公式
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
        std::unique_ptr<std::string> long_string;    // 长字符串
        std::unique_ptr<std::string> formula;        // 公式
        std::unique_ptr<std::string> hyperlink;      // 超链接
        std::unique_ptr<std::string> comment;        // 批注
        // 格式相关字段由 FormatDescriptor 管理
        double formula_result;       // 公式计算结果
        int shared_formula_index;    // 共享公式索引（-1表示不是共享公式）
        
        ExtendedData() : formula_result(0.0),
                        shared_formula_index(-1) {}
    };
    
    std::unique_ptr<ExtendedData> extended_;  // 只在需要时分配
    
    // 辅助方法
    void ensureExtended();
    void clearExtended();
    void initializeFlags();
    void deepCopyExtendedData(const Cell& other);
    void resetToEmpty();
    
public:
    Cell();
    virtual ~Cell();  // Make destructor virtual for memory pool compatibility
    
    // 便利构造函数
    explicit Cell(const std::string& value);
    explicit Cell(const char* value);
    explicit Cell(double value);
    explicit Cell(int value);
    explicit Cell(bool value);

    // 赋值运算（V3风格便捷API）
    Cell& operator=(double value);
    Cell& operator=(int value);
    Cell& operator=(bool value);
    Cell& operator=(const std::string& value);
    Cell& operator=(std::string_view value);
    Cell& operator=(const char* value);
    
    // 公式设置
    void setFormula(const std::string& formula, double result = 0.0);
    
    // 共享公式设置
    void setSharedFormula(int shared_index, double result = 0.0);
    void setSharedFormulaReference(int shared_index);
    
    // 获取值
    CellType getType() const {
        // 对外API统一：InlineString也显示为String, SharedFormula显示为Formula
        if (flags_.type == CellType::InlineString) return CellType::String;
        if (flags_.type == CellType::SharedFormula) return CellType::Formula;
        return flags_.type;
    }
    
    // 内部方法：获取真实的类型（用于测试和内部逻辑）
    CellType getInternalType() const { return flags_.type; }
    std::string getFormula() const;
    double getFormulaResult() const;
    
    // 共享公式获取
    int getSharedFormulaIndex() const;
    bool isSharedFormula() const;
    
    // 格式操作 - FormatDescriptor架构
    void setFormat(std::shared_ptr<const FormatDescriptor> format);
    std::shared_ptr<const FormatDescriptor> getFormatDescriptor() const;
    
    bool hasFormat() const { return flags_.has_format; }
    
    // 超链接操作
    void setHyperlink(const std::string& url);
    std::string getHyperlink() const;
    bool hasHyperlink() const { return flags_.has_hyperlink; }

    // 批注（注释）操作
    void setComment(const std::string& comment);
    std::string getComment() const;
    bool hasComment() const { return extended_ && extended_->comment; }
    
    // 模板化的值获取和设置
    template<typename T>
    T getValue() const {
        if constexpr (std::is_same_v<T, std::string>) {
            return getStringValue();
        } else if constexpr (std::is_floating_point_v<T>) {
            return static_cast<T>(getNumberValue());
        } else if constexpr (std::is_integral_v<T> && !std::is_same_v<T, bool>) {
            return static_cast<T>(getNumberValue());
        } else if constexpr (std::is_same_v<T, bool>) {
            return getBooleanValue();
        } else {
            static_assert(std::is_same_v<T, std::string>, 
                          "Unsupported type for Cell::getValue<T>()");
        }
    }
    
    template<typename T>
    void setValue(const T& value) {
        if constexpr (std::is_arithmetic_v<T> && !std::is_same_v<T, bool>) {
            setValueImpl(static_cast<double>(value));
        } else if constexpr (std::is_same_v<T, bool>) {
            setValueImpl(value);
        } else if constexpr (std::is_convertible_v<T, std::string>) {
            setValueImpl(std::string(value));
        } else {
            static_assert(std::is_arithmetic_v<T>, 
                          "Unsupported type for Cell::setValue<T>()");
        }
    }
    
    // 安全访问方法
    template<typename T>
    std::optional<T> tryGetValue() const noexcept {
        try {
            return getValue<T>();
        } catch (...) {
            return std::nullopt;
        }
    }
    
    template<typename T>
    T getValueOr(const T& default_value) const noexcept {
        return tryGetValue<T>().value_or(default_value);
    }
    
    // 状态检查
    bool isEmpty() const { return flags_.type == CellType::Empty; }
    bool isNumber() const { return flags_.type == CellType::Number; }
    bool isString() const { return flags_.type == CellType::String || flags_.type == CellType::InlineString; }
    bool isBoolean() const { return flags_.type == CellType::Boolean; }
    bool isFormula() const { return flags_.type == CellType::Formula || flags_.type == CellType::SharedFormula; }
    bool isDate() const { return flags_.type == CellType::Date; }
    
    // 便捷访问方法（统一命名风格）
    /**
     * @brief 将单元格值转为字符串
     * @return 字符串表示的值
     */
    std::string asString() const { return getValue<std::string>(); }
    
    /**
     * @brief 将单元格值转为数字
     * @return 数字值，如果无法转换则抛出异常
     */
    double asNumber() const { return getValue<double>(); }
    
    /**
     * @brief 将单元格值转为布尔值
     * @return 布尔值，如果无法转换则抛出异常
     */
    bool asBool() const { return getValue<bool>(); }
    
    /**
     * @brief 将单元格值转为整数
     * @return 整数值，如果无法转换则抛出异常
     */
    int asInt() const { return getValue<int>(); }
    
    // 安全的类型转换方法
    /**
     * @brief 检查是否可以转换为指定类型
     * @tparam T 目标类型
     * @return 是否可以安全转换
     */
    template<typename T>
    bool canConvertTo() const noexcept {
        return tryGetValue<T>().has_value();
    }
    
    /**
     * @brief 安全类型转换
     * @tparam T 目标类型
     * @return 转换结果的optional，失败时返回nullopt
     */
    template<typename T>
    std::optional<T> safeCast() const noexcept {
        return tryGetValue<T>();
    }
    
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
    // FormatDescriptor的shared_ptr持有者
    mutable std::shared_ptr<const FormatDescriptor> format_descriptor_holder_;
    
    // 内部实现方法（由模板API调用）
    void setValueImpl(double value);
    void setValueImpl(bool value);
    void setValueImpl(const std::string& value);
    double getNumberValue() const;
    bool getBooleanValue() const;
    std::string getStringValue() const;
};

}} // namespace fastexcel::core
