#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file FormatTypes.hpp
 * @brief Excel格式相关的所有枚举类型定义
 * 
 * 这个文件集中定义了Excel格式化所需的所有枚举类型，
 * 避免在多个文件中重复定义，提高代码组织性和维护性。
 */

// 字体相关枚举

/**
 * @brief 下划线类型
 */
enum class UnderlineType : uint8_t {
    None = 0,
    Single = 1,
    Double = 2,
    SingleAccounting = 3,
    DoubleAccounting = 4
};

/**
 * @brief 字体脚本类型（上标/下标）
 */
enum class FontScript : uint8_t {
    None = 0,
    Superscript = 1,
    Subscript = 2
};

// 对齐相关枚举

/**
 * @brief 水平对齐方式
 */
enum class HorizontalAlign : uint8_t {
    None = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterAcross = 6,
    Distributed = 7
};

/**
 * @brief 垂直对齐方式
 */
enum class VerticalAlign : uint8_t {
    Top = 0,
    Center = 1,
    Bottom = 2,
    Justify = 3,
    Distributed = 4
};

// 填充相关枚举

/**
 * @brief 填充模式
 */
enum class PatternType : uint8_t {
    None = 0,
    Solid = 1,
    MediumGray = 2,
    DarkGray = 3,
    LightGray = 4,
    DarkHorizontal = 5,
    DarkVertical = 6,
    DarkDown = 7,
    DarkUp = 8,
    DarkGrid = 9,
    DarkTrellis = 10,
    LightHorizontal = 11,
    LightVertical = 12,
    LightDown = 13,
    LightUp = 14,
    LightGrid = 15,
    LightTrellis = 16,
    Gray125 = 17,
    Gray0625 = 18
};

// 边框相关枚举

/**
 * @brief 边框样式
 */
enum class BorderStyle : uint8_t {
    None = 0,
    Thin = 1,
    Medium = 2,
    Thick = 3,
    Double = 4,
    Hair = 5,
    Dotted = 6,
    Dashed = 7,
    DashDot = 8,
    DashDotDot = 9,
    MediumDashed = 10,
    MediumDashDot = 11,
    MediumDashDotDot = 12,
    SlantDashDot = 13
};

/**
 * @brief 对角线边框类型
 */
enum class DiagonalBorderType : uint8_t {
    None = 0,
    Up = 1,     // 左下到右上
    Down = 2,   // 左上到右下
    Both = 3    // 两条对角线
};

// 数字格式相关枚举

/**
 * @brief 数字格式类型（用于快速设置常见格式）
 */
enum class NumberFormatType : uint8_t {
    General = 0,
    Number = 1,
    Decimal = 2,
    Currency = 3,
    Accounting = 4,
    Date = 14,
    Time = 21,
    Percentage = 10,
    Fraction = 12,
    Scientific = 11,
    Text = 49
};

// 枚举实用函数

/**
 * @brief 将枚举转换为字符串（用于调试）
 */
const char* toString(UnderlineType type);
const char* toString(FontScript script);
const char* toString(HorizontalAlign align);
const char* toString(VerticalAlign align);
const char* toString(PatternType pattern);
const char* toString(BorderStyle style);
const char* toString(DiagonalBorderType type);
const char* toString(NumberFormatType type);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(UnderlineType type);
bool isValid(FontScript script);
bool isValid(HorizontalAlign align);
bool isValid(VerticalAlign align);
bool isValid(PatternType pattern);
bool isValid(BorderStyle style);
bool isValid(DiagonalBorderType type);
bool isValid(NumberFormatType type);

}} // namespace fastexcel::core
