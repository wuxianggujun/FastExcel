#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file NumberFormatTypes.hpp
 * @brief 数字格式相关的枚举类型定义
 * 
 * 包含数字格式化所需的枚举类型：
 * - 数字格式类型（用于快速设置常见格式）
 */

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
const char* toString(NumberFormatType type);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(NumberFormatType type);

}} // namespace fastexcel::core