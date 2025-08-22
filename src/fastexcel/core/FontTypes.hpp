#pragma once

#include <cstdint>

namespace fastexcel {
namespace core {

/**
 * @file FontTypes.hpp
 * @brief 字体相关的枚举类型定义
 * 
 * 包含字体格式化所需的所有枚举类型：
 * - 下划线类型
 * - 字体脚本类型（上标/下标）
 */

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

// 枚举实用函数

/**
 * @brief 将枚举转换为字符串（用于调试）
 */
const char* toString(UnderlineType type);
const char* toString(FontScript script);

/**
 * @brief 检查枚举值是否有效
 */
bool isValid(UnderlineType type);
bool isValid(FontScript script);

}} // namespace fastexcel::core