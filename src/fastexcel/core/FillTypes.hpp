#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file FillTypes.hpp
 * @brief 填充相关的枚举类型定义
 * 
 * 包含单元格填充格式化所需的枚举类型：
 * - 填充模式类型
 */

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

// 枚举实用函数

/**
 * @brief 将枚举转换为字符串（用于调试）
 */
const char* toString(PatternType pattern);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(PatternType pattern);

}} // namespace fastexcel::core