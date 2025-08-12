#pragma once

#include "Worksheet.hpp"
#include "FormatDescriptor.hpp"
#include "Color.hpp"
#include <string>
#include <vector>
#include <memory>
#include <optional>

namespace fastexcel {
namespace core {

/**
 * @brief 格式工具类
 * 
 * 提供格式复制、清除、检查、比较等实用功能。
 * 这些工具方法帮助用户更好地管理和操作Excel格式。
 * 
 * 设计原则：
 * - 实用性：解决实际格式管理问题
 * - 安全性：提供错误检查和异常处理
 * - 灵活性：支持多种操作模式
 * - 性能：批量操作优化
 */
class FormatUtils {
public:
    // ========== 格式复制 ==========
    
    /**
     * @brief 复制单个单元格的格式
     * @param worksheet 目标工作表
     * @param src_row 源单元格行号
     * @param src_col 源单元格列号
     * @param dest_row 目标单元格行号
     * @param dest_col 目标单元格列号
     * @return 是否成功复制
     */
    static bool copyFormat(Worksheet& worksheet, 
                          int src_row, int src_col,
                          int dest_row, int dest_col);
    
    /**
     * @brief 复制范围的格式（Excel地址版本）
     * @param worksheet 目标工作表
     * @param src_range 源范围（如"A1:C3"）
     * @param dest_range 目标范围（如"E1:G3"）
     * @return 成功复制的单元格数量
     */
    static int copyFormat(Worksheet& worksheet, 
                         const std::string& src_range,
                         const std::string& dest_range);
    
    /**
     * @brief 复制格式到多个目标位置
     * @param worksheet 目标工作表
     * @param src_row 源单元格行号
     * @param src_col 源单元格列号
     * @param dest_positions 目标位置列表（行列对）
     * @return 成功复制的位置数量
     */
    static int copyFormatToMultiple(Worksheet& worksheet,
                                   int src_row, int src_col,
                                   const std::vector<std::pair<int, int>>& dest_positions);
    
    /**
     * @brief 智能格式复制（自动调整范围大小）
     * @param worksheet 目标工作表
     * @param src_range 源范围
     * @param dest_start_cell 目标起始单元格（如"E1"）
     * @return 成功复制的单元格数量
     */
    static int smartCopyFormat(Worksheet& worksheet,
                              const std::string& src_range,
                              const std::string& dest_start_cell);
    
    // ========== 格式清除 ==========
    
    /**
     * @brief 清除单个单元格的格式
     * @param worksheet 目标工作表
     * @param row 行号
     * @param col 列号
     */
    static void clearFormat(Worksheet& worksheet, int row, int col);
    
    /**
     * @brief 清除范围的格式
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @return 清除格式的单元格数量
     */
    static int clearFormat(Worksheet& worksheet, const std::string& range);
    
    /**
     * @brief 清除工作表的所有格式
     * @param worksheet 目标工作表
     * @return 清除格式的单元格数量
     */
    static int clearAllFormats(Worksheet& worksheet);
    
    /**
     * @brief 选择性清除格式（只清除指定类型的格式）
     * @param worksheet 目标工作表
     * @param range Excel地址字符串
     * @param clear_font 是否清除字体格式
     * @param clear_fill 是否清除填充格式
     * @param clear_border 是否清除边框格式
     * @param clear_alignment 是否清除对齐格式
     * @param clear_number 是否清除数字格式
     * @return 处理的单元格数量
     */
    static int selectiveClearFormat(Worksheet& worksheet, const std::string& range,
                                   bool clear_font = true,
                                   bool clear_fill = true,
                                   bool clear_border = true,
                                   bool clear_alignment = true,
                                   bool clear_number = true);
    
    // ========== 格式检查 ==========
    
    /**
     * @brief 检查单元格是否有格式
     * @param worksheet 工作表
     * @param row 行号
     * @param col 列号
     * @return 是否有格式
     */
    static bool hasFormat(const Worksheet& worksheet, int row, int col);
    
    /**
     * @brief 获取单元格的格式
     * @param worksheet 工作表
     * @param row 行号
     * @param col 列号
     * @return 格式描述符（可能为空）
     */
    static std::optional<FormatDescriptor> getFormat(const Worksheet& worksheet, 
                                                    int row, int col);
    
    /**
     * @brief 检查范围内是否所有单元格都有相同格式
     * @param worksheet 工作表
     * @param range Excel地址字符串
     * @return 是否格式一致
     */
    static bool hasUniformFormat(const Worksheet& worksheet, const std::string& range);
    
    /**
     * @brief 获取范围内的所有不同格式
     * @param worksheet 工作表
     * @param range Excel地址字符串
     * @return 不同格式的列表
     */
    static std::vector<FormatDescriptor> getUniqueFormats(const Worksheet& worksheet,
                                                         const std::string& range);
    
    /**
     * @brief 统计格式使用情况
     * @param worksheet 工作表
     * @param range Excel地址字符串（可选，默认整个工作表）
     * @return 格式统计信息
     */
    struct FormatStats {
        int total_cells = 0;
        int formatted_cells = 0;
        int unique_formats = 0;
        std::vector<std::pair<FormatDescriptor, int>> format_usage; // 格式和使用次数
    };
    
    static FormatStats getFormatStats(const Worksheet& worksheet, 
                                     const std::string& range = "");
    
    // ========== 格式比较 ==========
    
    /**
     * @brief 比较两个格式是否相同
     * @param format1 第一个格式
     * @param format2 第二个格式
     * @return 是否相同
     */
    static bool formatsMatch(const FormatDescriptor& format1, 
                            const FormatDescriptor& format2);
    
    /**
     * @brief 比较两个格式的差异
     * @param format1 第一个格式
     * @param format2 第二个格式
     * @return 差异描述
     */
    struct FormatDifference {
        bool font_different = false;
        bool fill_different = false;
        bool border_different = false;
        bool alignment_different = false;
        bool number_format_different = false;
        
        bool hasDifferences() const {
            return font_different || fill_different || border_different || 
                   alignment_different || number_format_different;
        }
        
        std::string toString() const;
    };
    
    static FormatDifference compareFormats(const FormatDescriptor& format1,
                                          const FormatDescriptor& format2);
    
    /**
     * @brief 比较两个单元格的格式
     * @param worksheet 工作表
     * @param row1 第一个单元格行号
     * @param col1 第一个单元格列号
     * @param row2 第二个单元格行号
     * @param col2 第二个单元格列号
     * @return 格式差异（如果任一单元格无格式则返回空）
     */
    static std::optional<FormatDifference> compareCellFormats(const Worksheet& worksheet,
                                                             int row1, int col1,
                                                             int row2, int col2);
    
    // ========== 格式查找和替换 ==========
    
    /**
     * @brief 查找具有特定格式的单元格
     * @param worksheet 工作表
     * @param target_format 目标格式
     * @param search_range 搜索范围（可选）
     * @return 匹配的单元格位置列表
     */
    static std::vector<std::pair<int, int>> findCellsWithFormat(
        const Worksheet& worksheet,
        const FormatDescriptor& target_format,
        const std::string& search_range = "");
    
    /**
     * @brief 替换格式
     * @param worksheet 目标工作表
     * @param old_format 要替换的格式
     * @param new_format 新格式
     * @param search_range 搜索范围（可选）
     * @return 替换的单元格数量
     */
    static int replaceFormat(Worksheet& worksheet,
                            const FormatDescriptor& old_format,
                            const FormatDescriptor& new_format,
                            const std::string& search_range = "");
    
    // ========== 格式导入导出 ==========
    
    /**
     * @brief 导出格式到描述字符串
     * @param format 格式描述符
     * @return JSON格式的描述字符串
     */
    static std::string exportFormat(const FormatDescriptor& format);
    
    /**
     * @brief 从描述字符串导入格式
     * @param format_string JSON格式的描述字符串
     * @return 格式描述符（可能失败）
     */
    static std::optional<FormatDescriptor> importFormat(const std::string& format_string);
    
    /**
     * @brief 导出范围内的所有格式
     * @param worksheet 工作表
     * @param range Excel地址字符串
     * @return 格式映射（位置 -> 格式字符串）
     */
    static std::string exportRangeFormats(const Worksheet& worksheet, 
                                         const std::string& range);
    
    /**
     * @brief 导入范围格式
     * @param worksheet 目标工作表
     * @param formats_data 格式数据字符串
     * @param dest_start_cell 目标起始位置
     * @return 成功导入的格式数量
     */
    static int importRangeFormats(Worksheet& worksheet,
                                 const std::string& formats_data,
                                 const std::string& dest_start_cell);

private:
    // 私有辅助方法
    static std::pair<int, int> parseCell(const std::string& cell_address);
    static std::pair<std::pair<int, int>, std::pair<int, int>> parseRange(const std::string& range);
    static std::string formatToJson(const FormatDescriptor& format);
    static FormatDescriptor jsonToFormat(const std::string& json);
    static bool isValidCellPosition(int row, int col);
};

}} // namespace fastexcel::core