#pragma once

#include "FormatDescriptor.hpp"
#include "StyleBuilder.hpp"
#include "Color.hpp"
#include "FormatTypes.hpp"
#include <string>
#include <memory>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;

/**
 * @brief 范围格式化器 - 批量设置单元格格式
 * 
 * 提供流畅的API来设置单元格范围的格式，支持：
 * - Excel地址字符串（"A1:C10"）和数字坐标
 * - 表格样式应用
 * - 边框快捷设置
 * - 批量格式应用
 * 
 * 设计原则：
 * - KISS：简单直观的方法名
 * - 链式调用：支持流畅的API体验  
 * - 性能优化：内部批量处理减少开销
 * - 延迟执行：调用apply()时才真正应用格式
 */
class RangeFormatter {
private:
    Worksheet* worksheet_ = nullptr;
    int start_row_ = -1;
    int start_col_ = -1;
    int end_row_ = -1;
    int end_col_ = -1;
    
    // 待应用的格式
    std::unique_ptr<FormatDescriptor> pending_format_;
    
    // 表格样式配置
    std::string table_style_name_;
    bool has_headers_ = false;
    bool row_banding_ = true;
    bool col_banding_ = false;
    
    // 边框配置
    enum class BorderTarget {
        None,
        All,
        Outside,
        Inside
    };
    BorderTarget border_target_ = BorderTarget::None;
    BorderStyle border_style_ = BorderStyle::Thin;
    core::Color border_color_ = core::Color::BLACK;
    
    // 内部辅助方法
    bool parseRange(const std::string& range);
    void validateRange() const;
    void applyFormatToRange();
    void applyBordersToRange();
    void applyTableStyle();
    
    // 获取当前单元格格式
    std::shared_ptr<const FormatDescriptor> getCellFormatDescriptor(int row, int col) const;
    
    // Excel地址转换辅助函数
    static int columnLetterToNumber(const std::string& col_str);

public:
    /**
     * @brief 构造函数
     * @param worksheet 目标工作表
     */
    explicit RangeFormatter(Worksheet* worksheet);
    
    // 禁用拷贝，支持移动
    RangeFormatter(const RangeFormatter&) = delete;
    RangeFormatter& operator=(const RangeFormatter&) = delete;
    RangeFormatter(RangeFormatter&&) = default;
    RangeFormatter& operator=(RangeFormatter&&) = default;
    
    // 范围设置
    
    /**
     * @brief 设置格式化范围（数字坐标）
     * @param start_row 起始行（0-based）
     * @param start_col 起始列（0-based）
     * @param end_row 结束行（0-based，包含）
     * @param end_col 结束列（0-based，包含）
     * @return RangeFormatter引用，支持链式调用
     */
    RangeFormatter& setRange(int start_row, int start_col, int end_row, int end_col);
    
    /**
     * @brief 设置格式化范围（Excel地址）
     * @param range Excel地址字符串，如"A1:C10"
     * @return RangeFormatter引用，支持链式调用
     */
    RangeFormatter& setRange(const std::string& range);
    
    /**
     * @brief 设置单行范围
     * @param row 行号（0-based）
     * @param start_col 起始列（可选，默认0）
     * @param end_col 结束列（可选，默认-1表示到最后使用的列）
     * @return RangeFormatter引用
     */
    RangeFormatter& setRow(int row, int start_col = 0, int end_col = -1);
    
    /**
     * @brief 设置单列范围
     * @param col 列号（0-based）
     * @param start_row 起始行（可选，默认0）
     * @param end_row 结束行（可选，默认-1表示到最后使用的行）
     * @return RangeFormatter引用
     */
    RangeFormatter& setColumn(int col, int start_row = 0, int end_row = -1);
    
    // 格式应用
    
    /**
     * @brief 应用格式描述符
     * @param format 要应用的格式
     * @return RangeFormatter引用
     */
    RangeFormatter& applyFormat(const FormatDescriptor& format);
    
    /**
     * @brief 应用样式构建器
     * @param builder 样式构建器
     * @return RangeFormatter引用
     */
    RangeFormatter& applyStyle(const StyleBuilder& builder);
    
    /**
     * @brief 使用共享格式指针（性能优化）
     * @param format 共享的格式指针
     * @return RangeFormatter引用
     */
    RangeFormatter& applySharedFormat(std::shared_ptr<const FormatDescriptor> format);
    
    // 表格样式
    
    /**
     * @brief 应用表格样式
     * @param style_name 表格样式名称（如"TableStyleMedium2"）
     * @return RangeFormatter引用
     */
    RangeFormatter& asTable(const std::string& style_name = "TableStyleMedium2");
    
    /**
     * @brief 设置是否包含标题行
     * @param has_headers 是否包含标题行
     * @return RangeFormatter引用
     */
    RangeFormatter& withHeaders(bool has_headers = true);
    
    /**
     * @brief 设置行列带状格式
     * @param row_banding 行带状格式
     * @param col_banding 列带状格式（可选）
     * @return RangeFormatter引用
     */
    RangeFormatter& withBanding(bool row_banding = true, bool col_banding = false);
    
    // 边框快捷方法
    
    /**
     * @brief 设置所有边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return RangeFormatter引用
     */
    RangeFormatter& allBorders(BorderStyle style = BorderStyle::Thin, 
                              core::Color color = core::Color::BLACK);
    
    /**
     * @brief 设置外边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return RangeFormatter引用
     */
    RangeFormatter& outsideBorders(BorderStyle style = BorderStyle::Medium,
                                  core::Color color = core::Color::BLACK);
    
    /**
     * @brief 设置内边框
     * @param style 边框样式
     * @param color 边框颜色（可选）
     * @return RangeFormatter引用
     */
    RangeFormatter& insideBorders(BorderStyle style = BorderStyle::Thin,
                                 core::Color color = core::Color::BLACK);
    
    /**
     * @brief 清除所有边框
     * @return RangeFormatter引用
     */
    RangeFormatter& noBorders();
    
    // 快捷格式方法
    
    /**
     * @brief 设置背景色
     * @param color 背景色
     * @return RangeFormatter引用
     */
    RangeFormatter& backgroundColor(core::Color color);
    
    /**
     * @brief 设置字体颜色
     * @param color 字体颜色
     * @return RangeFormatter引用
     */
    RangeFormatter& fontColor(core::Color color);
    
    /**
     * @brief 设置粗体
     * @param bold 是否粗体
     * @return RangeFormatter引用
     */
    RangeFormatter& bold(bool bold = true);
    
    /**
     * @brief 设置对齐方式
     * @param horizontal 水平对齐
     * @param vertical 垂直对齐（可选）
     * @return RangeFormatter引用
     */
    RangeFormatter& align(HorizontalAlign horizontal, 
                         VerticalAlign vertical = VerticalAlign::Bottom);
    
    /**
     * @brief 居中对齐
     * @return RangeFormatter引用
     */
    RangeFormatter& centerAlign();
    
    /**
     * @brief 右对齐
     * @return RangeFormatter引用
     */
    RangeFormatter& rightAlign();
    
    // 执行操作
    
    /**
     * @brief 应用所有设置到工作表
     * 这是唯一真正修改工作表的方法，之前的所有调用都是配置
     * @return 成功处理的单元格数量
     */
    int apply();
    
    /**
     * @brief 预览将要应用的格式（调试用）
     * @return 格式化范围的描述字符串
     */
    std::string preview() const;
    
    // 静态工厂方法
    
    /**
     * @brief 创建范围格式化器
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @return RangeFormatter对象
     */
    static RangeFormatter create(Worksheet& worksheet, const std::string& range);
    
    /**
     * @brief 创建范围格式化器
     * @param worksheet 目标工作表
     * @param start_row 起始行
     * @param start_col 起始列
     * @param end_row 结束行
     * @param end_col 结束列
     * @return RangeFormatter对象
     */
    static RangeFormatter create(Worksheet& worksheet, 
                                int start_row, int start_col, 
                                int end_row, int end_col);
};

}} // namespace fastexcel::core
