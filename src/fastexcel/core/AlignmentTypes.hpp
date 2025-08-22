#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file AlignmentTypes.hpp
 * @brief 对齐相关的枚举类型定义
 * 
 * 包含单元格对齐所需的所有枚举类型：
 * - 水平对齐方式
 * - 垂直对齐方式
 */

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

// 枚举实用函数

/**
 * @brief 将枚举转换为字符串（用于调试）
 */
const char* toString(HorizontalAlign align);
const char* toString(VerticalAlign align);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(HorizontalAlign align);
bool isValid(VerticalAlign align);

}} // namespace fastexcel::core