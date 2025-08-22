#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file BorderTypes.hpp
 * @brief 边框相关的枚举类型定义
 * 
 * 包含单元格边框格式化所需的所有枚举类型：
 * - 边框样式
 * - 对角线边框类型
 */

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

// 枚举实用函数

/**
 * @brief 将枚举转换为字符串（用于调试）
 */
const char* toString(BorderStyle style);
const char* toString(DiagonalBorderType type);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(BorderStyle style);
bool isValid(DiagonalBorderType type);

}} // namespace fastexcel::core