#pragma once

#include "Worksheet.hpp"
#include "StyleBuilder.hpp"
#include "FormatTypes.hpp"
#include "Color.hpp"
#include <string>

namespace fastexcel {
namespace core {

/**
 * @brief 快速格式化工具类
 * 
 * 提供一组静态方法，用于快速应用常见的格式化需求。
 * 这些方法封装了复杂的格式化逻辑，提供简单易用的API。
 * 
 * 设计原则：
 * - KISS：每个方法专注于一个特定的格式化任务
 * - 方便性：提供开箱即用的常用格式
 * - 性能：内部使用智能API进行优化
 * - 灵活性：支持可选参数进行定制
 */
class QuickFormat {
public:
    // 财务格式化
    
    /**
     * @brief 格式化为货币
     * @param worksheet 目标工作表
     * @param range Excel地址字符串（如"A1:C10"）
     * @param symbol 货币符号（如"$", "¥", "€"等）
     * @param decimal_places 小数位数（默认2位）
     * @param thousands_separator 是否使用千位分隔符（默认true）
     */
    static void formatAsCurrency(Worksheet& worksheet, const std::string& range, 
                                const std::string& symbol = "$",
                                int decimal_places = 2,
                                bool thousands_separator = true);
    
    /**
     * @brief 格式化为百分比
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param decimal_places 小数位数（默认2位）
     */
    static void formatAsPercentage(Worksheet& worksheet, const std::string& range, 
                                  int decimal_places = 2);
    
    /**
     * @brief 格式化为会计格式
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param symbol 货币符号
     */
    static void formatAsAccounting(Worksheet& worksheet, const std::string& range,
                                  const std::string& symbol = "$");
    
    // 数字格式化
    
    /**
     * @brief 格式化为数字（带千位分隔符）
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param decimal_places 小数位数（默认2位）
     */
    static void formatAsNumber(Worksheet& worksheet, const std::string& range,
                              int decimal_places = 2);
    
    /**
     * @brief 格式化为科学计数法
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param decimal_places 小数位数（默认2位）
     */
    static void formatAsScientific(Worksheet& worksheet, const std::string& range,
                                  int decimal_places = 2);
    
    // 日期时间格式化
    
    /**
     * @brief 格式化为日期
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param format 日期格式（如"yyyy-mm-dd", "m/d/yyyy"等，默认为系统默认）
     */
    static void formatAsDate(Worksheet& worksheet, const std::string& range,
                            const std::string& format = "m/d/yyyy");
    
    /**
     * @brief 格式化为时间
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param format 时间格式（如"hh:mm:ss", "h:mm AM/PM"等）
     */
    static void formatAsTime(Worksheet& worksheet, const std::string& range,
                            const std::string& format = "h:mm:ss AM/PM");
    
    /**
     * @brief 格式化为日期时间
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param format 日期时间格式
     */
    static void formatAsDateTime(Worksheet& worksheet, const std::string& range,
                                const std::string& format = "m/d/yyyy h:mm");
    
    // 表格格式化
    
    /**
     * @brief 格式化为表格
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param has_headers 是否包含标题行（默认true）
     * @param zebra_striping 是否使用斑马纹（默认true）
     * @param style_name 表格样式名称（可选）
     */
    static void formatAsTable(Worksheet& worksheet, const std::string& range,
                             bool has_headers = true,
                             bool zebra_striping = true,
                             const std::string& style_name = "TableStyleMedium2");
    
    /**
     * @brief 格式化数据列表
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param has_headers 是否包含标题行
     */
    static void formatAsDataList(Worksheet& worksheet, const std::string& range,
                                bool has_headers = true);
    
    // 标题和文本格式化
    
    /**
     * @brief 格式化为主标题
     * @param worksheet 目标工作表
     * @param row 行号
     * @param col 列号
     * @param text 标题文本（可选，如果提供则会设置单元格值）
     * @param font_size 字体大小（默认18）
     */
    static void formatAsTitle(Worksheet& worksheet, int row, int col, 
                             const std::string& text = "", 
                             double font_size = 18.0);
    
    /**
     * @brief 格式化为表头
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param style 表头样式类型（Modern, Classic, Bold等）
     */
    enum class HeaderStyle {
        Modern,     // 现代风格：蓝色背景，白色文字
        Classic,    // 经典风格：灰色背景，黑色文字
        Bold,       // 粗体风格：无背景，粗体文字
        Colorful    // 彩色风格：彩色背景，白色文字
    };
    
    static void formatAsHeader(Worksheet& worksheet, const std::string& range,
                              HeaderStyle style = HeaderStyle::Modern);
    
    /**
     * @brief 格式化为注释文本
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     */
    static void formatAsComment(Worksheet& worksheet, const std::string& range);
    
    // 数据突出显示
    
    /**
     * @brief 突出显示重要数据
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param color 突出显示颜色（默认黄色）
     */
    static void highlight(Worksheet& worksheet, const std::string& range,
                         Color color = Color::YELLOW);
    
    /**
     * @brief 格式化为警告样式
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     */
    static void formatAsWarning(Worksheet& worksheet, const std::string& range);
    
    /**
     * @brief 格式化为错误样式
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     */
    static void formatAsError(Worksheet& worksheet, const std::string& range);
    
    /**
     * @brief 格式化为成功样式
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     */
    static void formatAsSuccess(Worksheet& worksheet, const std::string& range);
    
    // 预定义样式套餐
    
    /**
     * @brief 应用财务报表样式套餐
     * @param worksheet 目标工作表
     * @param data_range 数据区域
     * @param header_range 标题区域（可选）
     * @param title_cell 主标题单元格（可选，格式："A1"）
     */
    static void applyFinancialReportStyle(Worksheet& worksheet,
                                         const std::string& data_range,
                                         const std::string& header_range = "",
                                         const std::string& title_cell = "");
    
    /**
     * @brief 应用现代简洁样式套餐
     * @param worksheet 目标工作表
     * @param data_range 数据区域
     * @param header_range 标题区域（可选）
     */
    static void applyModernStyle(Worksheet& worksheet,
                                const std::string& data_range,
                                const std::string& header_range = "");
    
    /**
     * @brief 应用经典样式套餐
     * @param worksheet 目标工作表
     * @param data_range 数据区域
     * @param header_range 标题区域（可选）
     */
    static void applyClassicStyle(Worksheet& worksheet,
                                 const std::string& data_range,
                                 const std::string& header_range = "");

private:
    // 私有辅助方法
    static std::string buildNumberFormat(const std::string& symbol, 
                                        int decimal_places, 
                                        bool thousands_separator);
    static Color getStyleColor(HeaderStyle style);
    static Color getTextColor(HeaderStyle style);
};

}} // namespace fastexcel::core
