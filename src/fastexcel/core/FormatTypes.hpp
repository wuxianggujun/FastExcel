#pragma once

/**
 * @file FormatTypes.hpp
 * @brief Excel格式相关的所有枚举类型定义（统一导入）
 * 
 * 这个文件作为统一的导入点，包含所有格式化相关的枚举类型。
 * 原有的枚举定义已经按功能模块化拆分到独立的文件中：
 * 
 * - FontTypes.hpp: 字体相关枚举（UnderlineType, FontScript）
 * - AlignmentTypes.hpp: 对齐相关枚举（HorizontalAlign, VerticalAlign）
 * - FillTypes.hpp: 填充相关枚举（PatternType）
 * - BorderTypes.hpp: 边框相关枚举（BorderStyle, DiagonalBorderType）
 * - NumberFormatTypes.hpp: 数字格式相关枚举（NumberFormatType）
 * 
 * 这种拆分提高了代码的模块化程度，便于维护和理解。
 */

// 导入所有格式类型定义
#include "FontTypes.hpp"
#include "AlignmentTypes.hpp"
#include "FillTypes.hpp"
#include "BorderTypes.hpp"
#include "NumberFormatTypes.hpp"
