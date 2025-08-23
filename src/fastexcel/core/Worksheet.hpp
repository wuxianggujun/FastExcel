#pragma once

#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/core/FormatRepository.hpp"
#include "fastexcel/core/WorksheetTypes.hpp"
#include "fastexcel/core/CellRangeManager.hpp"
#include "fastexcel/core/SharedFormula.hpp"
#include "fastexcel/core/RangeFormatter.hpp"
#include "fastexcel/core/Image.hpp"
#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/ColumnWidthManager.hpp"
#include "fastexcel/core/managers/CellDataProcessor.hpp"
// WorksheetLayoutManager和WorksheetImageManager功能已直接集成到Worksheet类中
#include "fastexcel/core/managers/WorksheetCSVHandler.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/AddressParser.hpp"
#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include "fastexcel/core/CellAddress.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <utility>
#include <ctime>
#include <unordered_map>
#include <set>
#include <functional>
#include <sstream>
#include <type_traits>
#include <optional>

namespace fastexcel {
namespace xml {
    class WorksheetXMLGenerator; // 前向声明
}
namespace core {

// 前向声明
class Workbook;
class SharedStringTable;
class FormatRepository;
class SharedFormulaManager;
class RangeFormatter;

// WorksheetChain类在独立的头文件中定义
class WorksheetChain;

// 列信息结构 - 使用WorksheetLayoutManager中的定义
// struct ColumnInfo 已在 WorksheetLayoutManager.hpp 中定义

/**
 * @brief Worksheet类 - Excel工作表
 * 
 * 提供完整的Excel工作表功能，包括：
 * - 单元格数据读写
 * - 行列格式设置
 * - 合并单元格
 * - 自动筛选
 * - 冻结窗格
 * - 打印设置
 * - 数据验证
 * - 条件格式
 * - 图表支持
 */
class Worksheet {
    friend class ::fastexcel::xml::WorksheetXMLGenerator;  // 让XML生成器能访问private方法
public:
    // 友元类需要访问的方法
    const std::vector<std::unique_ptr<Image>>& getImages() const { 
        return images_; 
    }
    
    std::pair<int, int> getUsedRange() const;
    std::tuple<int, int, int, int> getUsedRangeFull() const;
    
    bool hasCellAt(int row, int col) const;
    
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
    
    void setCellFormat(int row, int col, const core::FormatDescriptor& format);
    void setCellFormat(int row, int col, std::shared_ptr<const core::FormatDescriptor> format);
    void setCellFormat(int row, int col, const core::StyleBuilder& builder);
    
    const std::string& getName() const { return name_; }
    void setName(const std::string& name) { name_ = name; }
    
    /**
     * @brief 获取父工作簿
     * @return 父工作簿指针
     */
    std::shared_ptr<Workbook> getParentWorkbook() const { return parent_workbook_; }
    
    int findAndReplace(const std::string& find_text, const std::string& replace_text,
                       bool match_case = false, bool match_entire_cell = false);
    
    std::vector<std::pair<int, int>> findCells(const std::string& search_text,
                                               bool match_case = false,
                                               bool match_entire_cell = false) const;
    
    size_t getCellCount() const { return cells_.size(); }
    
    /**
     * @brief 获取工作表ID
     * @return 工作表ID
     */
    int getSheetId() const { return sheet_id_; }
    
    /**
     * @brief 获取总行数（有数据的行）
     * @return 总行数
     */
    int getRowCount() const;
    
    /**
     * @brief 获取总列数（有数据的列）
     * @return 总列数
     */
    int getColumnCount() const;
    
    /**
     * @brief 检查选项卡是否选中
     * @return 选项卡是否选中
     */
    bool isTabSelected() const { return sheet_view_.tab_selected; }
    
    /**
     * @brief 复制单元格
     * @param src_row 源行号
     * @param src_col 源列号
     * @param dst_row 目标行号
     * @param dst_col 目标列号
     * @param copy_format 是否复制格式
     * @param copy_row_height 是否复制行高
     */
    void copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format = true, bool copy_row_height = false);
    
    /**
     * @brief 复制范围
     * @param src_first_row 源起始行
     * @param src_first_col 源起始列
     * @param src_last_row 源结束行
     * @param src_last_col 源结束列
     * @param dst_row 目标起始行
     * @param dst_col 目标起始列
     * @param copy_format 是否复制格式
     */
    void copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col, bool copy_format = true);
    
    /**
     * @brief 排序范围
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @param sort_column 排序列（相对于范围的列索引）
     * @param ascending 是否升序
     * @param has_header 是否有标题行
     */
    void sortRange(int first_row, int first_col, int last_row, int last_col,
                   int sort_column = 0, bool ascending = true, bool has_header = false);
    
    /**
     * @brief 设置单元格公式
     * @param row 行号（0基索引）
     * @param col 列号（0基索引） 
     * @param formula 公式字符串
     * @param result 公式计算结果（可选，默认0.0）
     */
    void setFormula(int row, int col, const std::string& formula, double result = 0.0);
    
    /**
     * @brief 设置单元格公式（地址版本）
     * @param address 单元格地址
     * @param formula 公式字符串
     * @param result 公式计算结果（可选，默认0.0）
     */
    void setFormula(const Address& address, const std::string& formula, double result = 0.0);
    
    /**
     * @brief 创建共享公式
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @param formula 基础公式
     * @return 共享索引，失败返回-1
     */
    int createSharedFormula(int first_row, int first_col, int last_row, int last_col, const std::string& formula);

private:
    std::string name_;
    std::map<std::pair<int, int>, Cell> cells_; // (row, col) -> Cell
    std::shared_ptr<Workbook> parent_workbook_;
    int sheet_id_;
    
    // 优化组件
    SharedStringTable* sst_ = nullptr;
    FormatRepository* format_repo_ = nullptr;
    bool optimize_mode_ = false;
    
    // 共享公式管理器
    std::unique_ptr<SharedFormulaManager> shared_formula_manager_;
    
    // 优化模式下的行缓存
    struct WorksheetRow {
        int row_num;
        std::map<int, Cell> cells;
        double height = -1.0;
        bool hidden = false;
        bool data_changed = false;
        
        explicit WorksheetRow(int row) : row_num(row) {}
    };
    std::unique_ptr<WorksheetRow> current_row_;
    std::vector<Cell> row_buffer_;
    
    // 使用范围跟踪
    CellRangeManager range_manager_;
    
    // 管理器委托模式（简化后）
    std::unique_ptr<CellDataProcessor> cell_processor_;
    std::unique_ptr<WorksheetCSVHandler> csv_handler_;
    
    // 行列信息（直接管理，不再通过Manager）
    std::unordered_map<int, ColumnInfo> column_info_;
    std::unordered_map<int, RowInfo> row_info_;
    
    // 合并单元格（直接管理）
    std::vector<MergeRange> merge_ranges_;
    
    // 自动筛选（直接管理）
    std::unique_ptr<AutoFilterRange> autofilter_;
    
    // 冻结窗格（直接管理）
    std::unique_ptr<FreezePanes> freeze_panes_;
    
    // 页面视图
    SheetView sheet_view_;
    
    // 默认行高和列宽
    double default_row_height_ = 15.0;
    double default_col_width_ = 8.43;
    
    // 工作表保护
    bool protected_ = false;
    std::string protection_password_;
    
    // 选中范围
    std::string selection_ = "A1";
    
    // 活动单元格
    std::string active_cell_ = "A1";
    
    // 图片管理（直接管理，不再通过Manager）
    std::vector<std::unique_ptr<Image>> images_;
    int next_image_id_ = 1;
    
    // 列宽管理器
    std::unique_ptr<ColumnWidthManager> column_width_manager_;
    
    // 字体信息获取辅助方法
    std::string getWorkbookDefaultFont() const;
    double getWorkbookDefaultFontSize() const;
    
    // 内部辅助方法（从Manager类移过来）
    void validateCellPosition(int row, int col) const;
    void validateRange(int first_row, int first_col, int last_row, int last_col) const;
    std::string generateNextImageId();
    
    // 内部状态管理
    void syncLayoutManagerState(); // 用于向后兼容，如果需要

public:
    explicit Worksheet(const std::string& name, std::shared_ptr<Workbook> workbook, int sheet_id = 1);
    ~Worksheet() = default;
    
    // 禁用拷贝构造和赋值
    Worksheet(const Worksheet&) = delete;
    Worksheet& operator=(const Worksheet&) = delete;
    
    // 允许移动构造和赋值
    Worksheet(Worksheet&&) = default;
    Worksheet& operator=(Worksheet&&) = default;
    
    // 优化功能
    
    /**
     * @brief 设置共享字符串表
     * @param sst 共享字符串表指针
     */
    void setSharedStringTable(SharedStringTable* sst) { sst_ = sst; }
    
    /**
     * @brief 设置格式仓储
     * @param format_repo 格式仓储指针
     */
    void setFormatRepository(FormatRepository* format_repo) { 
        format_repo_ = format_repo;
        
        // 初始化列宽管理器
        if (format_repo_ && !column_width_manager_) {
            column_width_manager_ = std::make_unique<ColumnWidthManager>(format_repo_);
            // 同步工作簿Normal字体的MDW，按默认格式估算（与Excel对齐）
            const auto& def = core::FormatDescriptor::getDefault();
            int mdw = utils::ColumnWidthCalculator::estimateMDW(def.getFontName(), def.getFontSize());
            column_width_manager_->setWorkbookNormalMDW(mdw);
        }
    }
    
    /**
     * @brief 启用/禁用优化模式
     * @param enable 是否启用优化模式
     */
    void setOptimizeMode(bool enable);
    
    /**
     * @brief 检查是否启用了优化模式
     * @return 是否启用优化模式
     */
    bool isOptimizeMode() const { return optimize_mode_; }
    
    /**
     * @brief 刷新当前行缓存
     */
    void flushCurrentRow();
    
    /**
     * @brief 获取内存使用情况
     * @return 内存使用字节数
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 获取性能统计信息
     */
    struct PerformanceStats {
        size_t total_cells = 0;
        size_t memory_usage = 0;
        size_t sst_strings = 0;
        double sst_compression_ratio = 0.0;
        size_t unique_formats = 0;
        double format_deduplication_ratio = 0.0;
    };
    PerformanceStats getPerformanceStats() const;
    
    // 基本单元格操作
    
    /**
     * @brief 获取单元格引用
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 单元格引用
     */
private:
public:
    
    /**
     * @brief 获取单元格（支持多种地址格式）
     * @param address 单元格地址（支持隐式转换）
     *   - 字符串: "A1", "B2", "Sheet1!C3"
     *   - 坐标: Address(0, 0), Address{1, 2}
     * @return 单元格引用
     * 
     * @example
     * auto& cell1 = worksheet.getCell("A1");        // 字符串地址
     * auto& cell2 = worksheet.getCell({0, 1});      // 坐标地址
     * auto value = worksheet.getCell("B2").getValue<std::string>();
     */
    Cell& getCell(const core::Address& address);
    const Cell& getCell(const core::Address& address) const;
    
    // 模板化的单元格值获取和设置
    /**
     * @brief 模板化获取单元格值
     * @tparam T 返回值类型
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 指定类型的值
     * 
     * @example
     * auto str_value = worksheet.getValue<std::string>(0, 0);  // 获取A1的字符串值
     * auto num_value = worksheet.getValue<double>(1, 1);       // 获取B2的数字值
     * auto bool_value = worksheet.getValue<bool>(2, 2);        // 获取C3的布尔值
     */
public:
    
    /**
     * @brief 模板化获取单元格值
     * @tparam T 返回值类型（支持int, double, string, bool等基础类型）
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 指定类型的值
     * 
     * @example
     * auto str_value = worksheet.getValue<std::string>(0, 0);  // 获取A1的字符串值
     * auto num_value = worksheet.getValue<double>(1, 1);       // 获取B2的数字值
     * auto bool_value = worksheet.getValue<bool>(2, 2);        // 获取C3的布尔值
     */
    template<typename T>
    T getValue(int row, int col) const {
        return getCell(row, col).getValue<T>();
    }
public:
    
    /**
     * @brief 模板化获取单元格值（统一地址接口）
     * @tparam T 返回值类型
     * @param address 单元格地址（支持隐式转换）
     * @return 单元格值
     * 
     * 示例：
     * auto name = worksheet.getValue<std::string>("A1");
     * auto age = worksheet.getValue<int>("B2");
     * auto salary = worksheet.getValue<double>(Address(1, 3));
     */
    template<typename T>
    T getValue(const core::Address& address) const {
        return getValue<T>(address.getRow(), address.getCol());
    }
    
    /**
     * @brief 模板化设置单元格值
     * @tparam T 值类型
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @param value 要设置的值
     * 
     * @example
     * worksheet.setValue(0, 0, std::string("Hello"));  // 设置A1为字符串
     * worksheet.setValue(1, 1, 123.45);                // 设置B2为数字
     * worksheet.setValue(2, 2, true);                   // 设置C3为布尔值
     */
public:
    
    /**
     * @brief 设置单元格值（模板化方法）
     * @tparam T 值类型（支持int, double, string, bool等基础类型）
     * @param row 行号（0开始）
     * @param col 列号（0开始）  
     * @param value 要设置的值
     * 
     * @example
     * worksheet.setValue(0, 0, std::string("Hello"));  // 设置A1为字符串
     * worksheet.setValue(1, 1, 123.45);                // 设置B2为数字
     * worksheet.setValue(2, 2, true);                   // 设置C3为布尔值
     */
    template<typename T>
    void setValue(int row, int col, const T& value) {
        cell_processor_->setValue(row, col, value);
    }
public:
    
    /**
     * @brief 模板化设置单元格值（统一地址接口）
     * @tparam T 值类型
     * @param address 单元格地址（支持多种格式的隐式转换）
     *   - 行列坐标: Address(0, 0) 或直接 (0, 0)
     *   - 字符串: Address("A1") 或直接 "A1" 
     *   - C字符串: "A1"
     * @param value 要设置的值
     * 
     * 示例：
     * worksheet.setValue("A1", std::string("Hello"));  // 字符串地址
     * worksheet.setValue("B2", 123.45);                // 字符串地址
     * worksheet.setValue(Address(0, 0), "Hello");      // Address对象
     */
    template<typename T>
    void setValue(const core::Address& address, const T& value) {
        setValue<T>(address.getRow(), address.getCol(), value);
    }
    
    /**
     * @brief setCellValue - setValue的语义化别名
     * 为了API一致性，提供更明确的方法名
     */
    template<typename T>
    void setCellValue(int row, int col, const T& value) {
        setValue<T>(row, col, value);
    }
public:
    
    // 智能单元格格式设置 API
    
    /**
     * @brief 设置单元格格式（智能优化版）
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @param format 格式描述符
     * 
     * @details 设置指定单元格的显示格式，内部自动FormatRepository优化。
     *          格式可能被多个单元格共享以节省内存。
     * @example worksheet.setCellFormat(0, 0, format);
     */
private:
public:
    
    /**
     * @brief 设置单元格格式（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @param format 格式描述符
     * 
     * 示例：
     * worksheet.setCellFormat("A1", format);           // 字符串地址
     * worksheet.setCellFormat({0, 1}, format);         // 坐标地址
     * worksheet.setCellFormat(Address("B2"), format);  // Address对象
     */
    void setCellFormat(const core::Address& address, const core::FormatDescriptor& format);
    void setCellFormat(const core::Address& address, std::shared_ptr<const core::FormatDescriptor> format);
    void setCellFormat(const core::Address& address, const core::StyleBuilder& builder);
    
    // 范围格式化API
    
    /**
     * @brief 创建范围格式化器
     * @param range Excel地址字符串（如"A1:C10"）
     * @return RangeFormatter对象，支持链式调用
     * 
     * @example 
     * worksheet.rangeFormatter("A1:C10")
     *     .backgroundColor(Color::YELLOW)
     *     .allBorders()
     *     .apply();
     */
    RangeFormatter rangeFormatter(const std::string& range);
    
    /**
     * @brief 创建范围格式化器（坐标版本）
     * @param start_row 起始行（0-based）
     * @param start_col 起始列（0-based）
     * @param end_row 结束行（0-based，包含）
     * @param end_col 结束列（0-based，包含）
     * @return RangeFormatter对象，支持链式调用
     */
private:
    RangeFormatter rangeFormatter(int start_row, int start_col, int end_row, int end_col);
public:
    
    /**
     * @brief 创建范围格式化器（统一范围接口）
     * @param range 范围地址（支持隐式转换）
     *   - 字符串范围: "A1:C10", "Sheet1!A1:C10"
     *   - 坐标范围: CellRange(0, 0, 9, 2), CellRange{0, 0, 9, 2}
     * @return RangeFormatter对象，支持链式调用
     * 
     * 示例：
     * worksheet.rangeFormatter("A1:C10").backgroundColor(Color::YELLOW).apply();
     * worksheet.rangeFormatter({0, 0, 9, 2}).allBorders().apply();
     */
    RangeFormatter rangeFormatter(const core::CellRange& range) {
        return rangeFormatter(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    /**
     * @brief 安全获取单元格格式
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 格式描述符的可选值，失败时返回std::nullopt
     */
private:
    std::optional<std::shared_ptr<const FormatDescriptor>> tryGetCellFormat(int row, int col) const noexcept {
        try {
            if (!hasCellAt(row, col)) {
                return std::nullopt;
            }
            const auto& cell = getCell(row, col);
            auto format = cell.getFormatDescriptor();
            return format ? std::make_optional(format) : std::nullopt;
        } catch (...) {
            return std::nullopt;
        }
    }
public:
    
    /**
     * @brief 安全获取列宽
     * @param col 列号（0开始）
     * @return 列宽的可选值，失败时返回std::nullopt
     */
private:
    std::optional<double> tryGetColumnWidth(int col) const noexcept {
        try {
            if (col < 0) return std::nullopt;
            return std::make_optional(getColumnWidth(col));
        } catch (...) {
            return std::nullopt;
        }
    }
public:
    
    /**
     * @brief 安全获取行高
     * @param row 行号（0开始）
     * @return 行高的可选值，失败时返回std::nullopt
     */
private:
    std::optional<double> tryGetRowHeight(int row) const noexcept {
        try {
            if (row < 0) return std::nullopt;
            return std::make_optional(getRowHeight(row));
        } catch (...) {
            return std::nullopt;
        }
    }
public:
    
    /**
     * @brief 安全获取使用范围
     * @return 使用范围的可选值 (最大行, 最大列)，失败或无数据时返回std::nullopt
     */
    std::optional<std::pair<int, int>> tryGetUsedRange() const noexcept {
        try {
            auto range = getUsedRange();
            if (range.first == -1 || range.second == -1) {
                return std::nullopt;
            }
            return std::make_optional(range);
        } catch (...) {
            return std::nullopt;
        }
    }
    
    /**
     * @brief 安全获取单元格值（不抛异常）
     * @tparam T 返回值类型
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 可选值，失败时返回std::nullopt
     */
private:
    template<typename T>
    std::optional<T> tryGetValue(int row, int col) const noexcept {
        if (!hasCellAt(row, col)) {
            return std::nullopt;
        }
        return getCell(row, col).tryGetValue<T>();
    }
public:
    
    /**
     * @brief 安全获取单元格值（统一地址接口）
     * @tparam T 返回值类型
     * @param address 单元格地址（支持隐式转换）
     * @return 可选值，失败时返回std::nullopt
     * 
     * 示例：
     * auto value = worksheet.tryGetValue<std::string>("A1");
     * if (value.has_value()) {
     *     std::cout << value.value() << std::endl;
     * }
     */
    template<typename T>
    std::optional<T> tryGetValue(const core::Address& address) const noexcept {
        return tryGetValue<T>(address.getRow(), address.getCol());
    }
    
    /**
     * @brief 获取单元格值或默认值
     * @tparam T 返回值类型
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @param default_value 默认值
     * @return 单元格值或默认值
     */
    /**
     * @brief 获取单元格值或默认值（兼容旧版接口）
     * @tparam T 返回值类型
     * @param row 行号
     * @param col 列号
     * @param default_value 默认值
     * @return 单元格值或默认值
     */
    template<typename T>
    T getValueOr(int row, int col, const T& default_value) const noexcept {
        if (!hasCellAt(row, col)) {
            return default_value;
        }
        return getCell(row, col).getValueOr<T>(default_value);
    }
public:
    
    /**
     * @brief 获取单元格值或默认值（统一地址接口）
     * @tparam T 返回值类型
     * @param address 单元格地址（支持隐式转换）
     * @param default_value 默认值
     * @return 单元格值或默认值
     * 
     * 示例：
     * auto value = worksheet.getValueOr<std::string>("A1", "默认值");
     * auto number = worksheet.getValueOr<double>({0, 1}, 0.0);
     */
    template<typename T>
    T getValueOr(const core::Address& address, const T& default_value) const noexcept {
        return getValueOr<T>(address.getRow(), address.getCol(), default_value);
    }
    
    /**
     * @brief 写入日期时间
     * @param row 行号
     * @param col 列号
     * @param datetime 日期时间
     */
private:
    void writeDateTime(int row, int col, const std::tm& datetime);
public:
    
    /**
     * @brief 写入日期时间（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @param datetime 日期时间
     * 
     * 示例：
     * worksheet.writeDateTime("A1", datetime);     // 字符串地址
     * worksheet.writeDateTime({0, 0}, datetime);   // 坐标地址
     */
    void writeDateTime(const core::Address& address, const std::tm& datetime);
    
    /**
     * @brief 写入URL链接
     * @param row 行号
     * @param col 列号
     * @param url URL地址
     * @param string 显示文本（可选）
     */
private:
    void writeUrl(int row, int col, const std::string& url, const std::string& string = "");
public:
    
    /**
     * @brief 写入URL链接（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @param url URL地址
     * @param string 显示文本（可选）
     * 
     * 示例：
     * worksheet.writeUrl("A1", "https://example.com", "点击这里");
     * worksheet.writeUrl({0, 0}, "https://example.com");
     */
    void writeUrl(const core::Address& address, const std::string& url, const std::string& string = "");

    // 视图与窗格（统一地址接口）
    void freezePanes(const core::Address& split_cell);
    void freezePanes(const core::Address& split_cell, const core::Address& top_left_cell);
    
    /**
     * @brief 冻结窗格（使用行列坐标）
     * @param row 分割行
     * @param col 分割列
     */
    void freezePanes(int row, int col);
    
    /**
     * @brief 冻结窗格（使用行列坐标，指定左上角单元格）
     * @param row 分割行
     * @param col 分割列
     * @param top_left_row 左上角行
     * @param top_left_col 左上角列
     */
    void freezePanes(int row, int col, int top_left_row, int top_left_col);
    void splitPanes(const core::Address& split_cell);
    
    // 批量数据操作
    
    // 模板化范围操作
    /**
     * @brief 获取范围内的所有值
     * @tparam T 返回值类型
     * @param start_row 开始行
     * @param start_col 开始列
     * @param end_row 结束行
     * @param end_col 结束列
     * @return 二维数组，包含范围内的所有值
     * 
     * @example
     * auto data = worksheet.getRange<std::string>(0, 0, 2, 2);  // 获取A1:C3范围的字符串值
     * auto numbers = worksheet.getRange<double>(1, 1, 3, 3);    // 获取B2:D4范围的数字值
     */
    template<typename T>
    std::vector<std::vector<T>> getRange(int start_row, int start_col, int end_row, int end_col) const {
        std::vector<std::vector<T>> result;
        result.reserve(end_row - start_row + 1);
        
        for (int row = start_row; row <= end_row; ++row) {
            std::vector<T> row_data;
            row_data.reserve(end_col - start_col + 1);
            
            for (int col = start_col; col <= end_col; ++col) {
                if (hasCellAt(row, col)) {
                    row_data.push_back(getCell(row, col).getValue<T>());
                } else {
                    // 空单元格返回默认值
                    if constexpr (std::is_same_v<T, std::string>) {
                        row_data.push_back("");
                    } else if constexpr (std::is_arithmetic_v<T>) {
                        row_data.push_back(static_cast<T>(0));
                    } else if constexpr (std::is_same_v<T, bool>) {
                        row_data.push_back(false);
                    }
                }
            }
            result.push_back(std::move(row_data));
        }
        return result;
    }
    
    /**
     * @brief 获取范围内的所有值（统一范围接口）
     * @tparam T 返回值类型
     * @param range 范围地址（支持隐式转换）
     *   - 字符串范围: "A1:C3", "Sheet1!A1:C3"
     *   - 坐标范围: CellRange(0, 0, 2, 2), CellRange{0, 0, 2, 2}
     * @return 二维数组，包含范围内的所有值
     * 
     * 示例：
     * auto data = worksheet.getRange<std::string>("A1:C3");    // 字符串范围
     * auto numbers = worksheet.getRange<double>({0, 0, 2, 2}); // 坐标范围
     */
    template<typename T>
    std::vector<std::vector<T>> getRange(const core::CellRange& range) const {
        return getRange<T>(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    /**
     * @brief 通过Excel地址获取范围内的所有值
     * @tparam T 返回值类型
     * @param range Excel范围地址（如"A1:C3"）
     * @return 二维数组，包含范围内的所有值
     * 
     * @example
     * auto data = worksheet.getRange<std::string>("A1:C3");  // 获取A1:C3范围的字符串值
     * auto numbers = worksheet.getRange<double>("B2:D4");    // 获取B2:D4范围的数字值
     */
    template<typename T>
    std::vector<std::vector<T>> getRange(const std::string& range) const {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        return getRange<T>(start_row, start_col, end_row, end_col);
    }
    
    /**
     * @brief 设置范围内的所有值
     * @tparam T 值类型
     * @param start_row 开始行
     * @param start_col 开始列
     * @param data 二维数据数组
     * 
     * @example
     * std::vector<std::vector<std::string>> data = {{"A", "B"}, {"C", "D"}};
     * worksheet.setRange(0, 0, data);  // 设置A1:B2的值
     */
    template<typename T>
    void setRange(int start_row, int start_col, const std::vector<std::vector<T>>& data) {
        for (size_t row_idx = 0; row_idx < data.size(); ++row_idx) {
            for (size_t col_idx = 0; col_idx < data[row_idx].size(); ++col_idx) {
                int target_row = static_cast<int>(start_row + row_idx);
                int target_col = static_cast<int>(start_col + col_idx);
                setValue<T>(target_row, target_col, data[row_idx][col_idx]);
            }
        }
    }
    
    /**
     * @brief 设置范围内的所有值（统一范围接口）
     * @tparam T 值类型
     * @param range 范围地址（支持隐式转换）
     *   - 字符串范围: "A1:B2", "Sheet1!A1:B2"
     *   - 坐标范围: CellRange(0, 0, 1, 1), CellRange{0, 0, 1, 1}
     * @param data 二维数据数组
     * 
     * 示例：
     * std::vector<std::vector<std::string>> data = {{"A", "B"}, {"C", "D"}};
     * worksheet.setRange("A1:B2", data);          // 字符串范围
     * worksheet.setRange({0, 0, 1, 1}, data);     // 坐标范围
     */
    template<typename T>
    void setRange(const core::CellRange& range, const std::vector<std::vector<T>>& data) {
        setRange<T>(range.getStartRow(), range.getStartCol(), data);
    }
    
    /**
     * @brief 通过Excel地址设置范围内的所有值
     * @tparam T 值类型  
     * @param range Excel范围地址（如"A1:C3"）
     * @param data 二维数据数组
     * 
     * @example
     * std::vector<std::vector<std::string>> data = {{"A", "B"}, {"C", "D"}};
     * worksheet.setRange("A1:B2", data);  // 设置A1:B2的值
     */
    template<typename T>
    void setRange(const std::string& range, const std::vector<std::vector<T>>& data) {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        setRange<T>(start_row, start_col, data);
    }
    
    // 链式调用支持
    /**
     * @brief 获取链式调用对象
     * @return 链式调用助手对象
     * 
     * @example
     * worksheet.chain()
     *     .setValue("A1", std::string("Hello"))
     *     .setValue("B1", 123.45)
     *     .setValue("C1", true)
     *     .setColumnWidth(0, 15.0)
     *     .setRowHeight(0, 20.0)
     *     .mergeCells(1, 0, 1, 2);
     */
    WorksheetChain chain();
    
    // 行列操作
    
    /**
     * @brief 设置列宽（简化版本）
     * @param col 列号
     * @param width 列宽值
     * @return 实际设置的列宽
     * 
     * @details 简单直接的列宽设置，不涉及格式系统，无需FormatRepository。
     *          适合纯列宽调整场景，性能高效。
     * 
     * @example
     * worksheet.setColumnWidth(0, 4.5);  // 设置第1列宽度为4.5
     * worksheet.setColumnWidth(1, 10.0); // 设置第2列宽度为10.0
     */
    double setColumnWidth(int col, double width);
    
    /**
     * @brief 设置列宽同时绑定字体格式（高级版本）
     * @param col 列号
     * @param width 列宽值
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @return 实际设置的列宽和格式ID
     * 
     * @details 高级功能，同时设置列宽和字体格式。需要FormatRepository。
     * 
     * @example
     * auto [width, format_id] = worksheet.setColumnWidthWithFont(0, 4.5, "微软雅黑", 11);
     */
    std::pair<double, int> setColumnWidthWithFont(int col, double width, 
                                                  const std::string& font_name, 
                                                  double font_size = 11.0);
    
    /**
     * @brief 智能列宽设置（专家版本）
     * @param col 列号
     * @param target_width 目标列宽
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @param strategy 列宽策略
     * @param cell_contents 单元格内容（可选）
     * @return 实际设置的列宽和格式ID
     * 
     * @details 专家级功能，提供完整的智能列宽管理。
     * 
     * @example
     * auto result = worksheet.setColumnWidthAdvanced(0, 4.0, "微软雅黑", 12, 
     *                                               ColumnWidthManager::WidthStrategy::EXACT);
     * 
     * // 内容感知模式
     * std::vector<std::string> contents = {"Hello", "世界", "Mixed内容"};
     * auto result = worksheet.setColumnWidthAdvanced(0, 5.0, "", 0, 
     *                                               ColumnWidthManager::WidthStrategy::CONTENT_AWARE,
     *                                               contents);
     */
    std::pair<double, int> setColumnWidthAdvanced(int col, double target_width,
                                                  const std::string& font_name,
                                                  double font_size,
                                                  ColumnWidthManager::WidthStrategy strategy,
                                                  const std::vector<std::string>& cell_contents = {});
    
    /**
     * @brief 批量智能列宽设置（高性能）
     * @param configs 列宽配置映射 (列号 -> 配置)
     * @return 实际设置结果 (列号 -> (宽度, 格式ID))
     * 
     * @details 高性能批量操作，内部按字体分组处理以提高效率
     * 
     * @example
     * std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig> configs;
     * configs[0] = {4.0, "微软雅黑", 11, ColumnWidthManager::WidthStrategy::EXACT};
     * configs[1] = {5.0, "Calibri", 11, ColumnWidthManager::WidthStrategy::EXACT};
     * configs[2] = {6.0, "", 0, ColumnWidthManager::WidthStrategy::ADAPTIVE};
     * 
     * auto results = worksheet.setColumnWidthsBatch(configs);
     */
    std::unordered_map<int, std::pair<double, int>> setColumnWidthsBatch(
        const std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig>& configs);
    
    /**
     * @brief 预计算列宽（不实际设置）
     * @param target_width 目标宽度
     * @param font_name 字体名称
     * @param font_size 字体大小
     * @return 优化后的实际列宽值
     */
    double calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const;
    
    /**
     * @brief 获取列宽管理器的缓存统计（性能监控）
     * @return 缓存统计信息
     */
    ColumnWidthManager::CacheStats getColumnWidthCacheStats() const {
        if (column_width_manager_) {
            return column_width_manager_->getCacheStats();
        }
        return {0, 0, 0};
    }
    
    /**
     * @brief 清理列宽管理器缓存（内存优化）
     */
    void clearColumnWidthCache() {
        if (column_width_manager_) {
            column_width_manager_->clearCache();
        }
    }
    
    /**
     * @brief 设置列格式
     * @param col 列号
     * @param format_id FormatRepository中的格式ID
     */
    void setColumnFormatId(int col, int format_id);
    
    /**
     * @brief 设置列格式范围
     * @param first_col 起始列
     * @param last_col 结束列
     * @param format_id FormatRepository中的格式ID
     */
    void setColumnFormatId(int first_col, int last_col, int format_id);
    
    /**
     * @brief 设置列格式
     * @param col 列号
     * @param format 格式描述符
     */
    void setColumnFormat(int col, const core::FormatDescriptor& format);
    void setColumnFormat(int col, std::shared_ptr<const core::FormatDescriptor> format);
    
    /**
     * @brief 设置列格式范围
     * @param first_col 起始列
     * @param last_col 结束列
     * @param format 格式描述符
     */
    void setColumnFormat(int first_col, int last_col, const core::FormatDescriptor& format);
    void setColumnFormat(int first_col, int last_col, std::shared_ptr<const core::FormatDescriptor> format);
    
    /**
     * @brief 隐藏列
     * @param col 列号
     */
    void hideColumn(int col);
    
    /**
     * @brief 隐藏列范围
     * @param first_col 起始列
     * @param last_col 结束列
     */
    void hideColumn(int first_col, int last_col);
    
    /**
     * @brief 设置行高
     * @param row 行号
     * @param height 高度
     */
    void setRowHeight(int row, double height);
    
    /**
     * @brief 设置行格式
     * @param row 行号
     * @param format 格式描述符
     */
    void setRowFormat(int row, const core::FormatDescriptor& format);
    void setRowFormat(int row, std::shared_ptr<const core::FormatDescriptor> format);
    
    /**
     * @brief 隐藏行
     * @param row 行号
     */
    void hideRow(int row);
    
    /**
     * @brief 隐藏行范围
     * @param first_row 起始行
     * @param last_row 结束行
     */
    void hideRow(int first_row, int last_row);
    
    // 合并单元格
    
    /**
     * @brief 合并单元格
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 合并单元格（支持多种范围格式）
     * @param range 单元格范围（支持隐式转换）
     *   - 字符串范围: "A1:C3", "Sheet1!A1:C3"
     *   - 坐标范围: CellRange(0, 0, 2, 2), CellRange{0, 0, 2, 2}
     *   - 单个地址: Address("A1") -> 1x1范围
     * 
     * @example
     * worksheet.mergeCells("A1:C3");           // 字符串范围
     * worksheet.mergeCells({0, 0, 2, 2});      // 坐标范围
     * worksheet.mergeCells(Address("B2"));     // 单个地址转为1x1范围
     */
    void mergeCells(const core::CellRange& range) {
        mergeCells(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    /**
     * @brief 合并单元格并写入内容
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @param value 内容
     */
    
    // 自动筛选
    
    /**
     * @brief 设置自动筛选
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void setAutoFilter(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 设置自动筛选（支持多种范围格式）
     * @param range 筛选范围（支持隐式转换）
     *   - 字符串范围: "A1:D10", "Sheet1!A1:D10"
     *   - 坐标范围: CellRange(0, 0, 9, 3), CellRange{0, 0, 9, 3}
     * 
     * @example
     * worksheet.setAutoFilter("A1:D10");       // 字符串范围
     * worksheet.setAutoFilter({0, 0, 9, 3});   // 坐标范围
     */
    void setAutoFilter(const core::CellRange& range) {
        setAutoFilter(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    /**
     * @brief 移除自动筛选
     */
    void removeAutoFilter();
    
    // 冻结窗格
    
    // 冻结窗格和分割窗格已在上面声明
    
    // 打印设置
    
    // 工作表保护
    
    /**
     * @brief 保护工作表
     * @param password 密码（可选）
     */
    void protect(const std::string& password = "");
    
    /**
     * @brief 取消保护
     */
    void unprotect();
    
    /**
     * @brief 检查是否受保护
     * @return 是否受保护
     */
    bool isProtected() const { return protected_; }
    
    // 视图设置
    
    /**
     * @brief 设置缩放比例
     * @param scale 缩放比例（10-400）
     */
    void setZoom(int scale);
    
    /**
     * @brief 显示/隐藏网格线
     * @param show 是否显示
     */
    void showGridlines(bool show = true);
    
    /**
     * @brief 显示/隐藏行列标题
     * @param show 是否显示
     */
    void showRowColHeaders(bool show = true);
    
    /**
     * @brief 设置从右到左显示
     * @param rtl 是否从右到左
     */
    void setRightToLeft(bool rtl = true);
    
    /**
     * @brief 设置选中状态
     * @param selected 是否选中
     */
    void setTabSelected(bool selected = true);
    
    /**
     * @brief 设置活动单元格
     * @param row 行号
     * @param col 列号
     */
private:
    void setActiveCell(int row, int col);
    
    /**
     * @brief 设置活动单元格（支持多种地址格式）
     * @param address 活动单元格地址（支持隐式转换）
     *   - 字符串: "B2", "C3", "Sheet1!D4"
     *   - 坐标: Address(1, 1), Address{2, 2}
     * 
     * @example
     * worksheet.setActiveCell("B2");           // 字符串地址
     * worksheet.setActiveCell({1, 1});         // 坐标地址
     */
    void setActiveCell(const core::Address& address);

    // 图片插入（统一地址/范围接口）
    std::string insertImage(const core::Address& address, const std::string& image_path) {
        return insertImage(address.getRow(), address.getCol(), image_path);
    }

    std::string insertImage(const core::Address& address, std::unique_ptr<Image> image) {
        return insertImage(address.getRow(), address.getCol(), std::move(image));
    }

    std::string insertImage(const core::CellRange& range, const std::string& image_path) {
        return insertImage(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol(), image_path);
    }
    
    /**
     * @brief 设置选中范围
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void setSelection(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 设置选中范围（支持多种范围格式）
     * @param range 选中范围（支持隐式转换）
     *   - 字符串范围: "A1:C5", "Sheet1!A1:C5"
     *   - 坐标范围: CellRange(0, 0, 4, 2), CellRange{0, 0, 4, 2}
     *   - 单个地址: Address("B2") -> 1x1范围
     * 
     * @example
     * worksheet.setSelection("A1:C5");         // 字符串范围
     * worksheet.setSelection({0, 0, 4, 2});    // 坐标范围
     * worksheet.setSelection(Address("B2"));   // 单个地址
     */
    void setSelection(const core::CellRange& range) {
        setSelection(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    // 获取信息
    
    // 其他原本在private的方法已移动到public区域
    
    // getUsedRange, hasCellAt, getCellCount 已在 public 区域声明
    
    // 便捷的工作表状态检查方法
    /**
     * @brief 检查工作表是否为空（无任何单元格数据）
     * @return 是否为空
     * 
     * @example
     * if (worksheet.isEmpty()) {
     *     std::cout << "工作表为空" << std::endl;
     * }
     */
    bool isEmpty() const { return cells_.empty(); }
    
    /**
     * @brief 检查工作表是否有数据
     * @return 是否有数据
     */
    bool hasData() const { return !cells_.empty(); }
    
    // getRowCount 和 getColumnCount 已在 public 区域声明
    
    /**
     * @brief 获取指定行的单元格数量
     * @param row 行号
     * @return 该行的单元格数量
     */
    int getCellCountInRow(int row) const;
    
    /**
     * @brief 获取指定列的单元格数量
     * @param col 列号
     * @return 该列的单元格数量
     */
    int getCellCountInColumn(int col) const;
    
    // hasCellAt 已在 public 区域声明
    
    /**
     * @brief 检查指定地址是否存在单元格（支持多种地址格式）
     * @param address 单元格地址（支持隐式转换）
     *   - 字符串: "A1", "B2", "Sheet1!C3"
     *   - 坐标: Address(0, 0), Address{1, 2}
     * @return 是否存在
     * 
     * @example
     * if (worksheet.hasCellAt("A1")) {         // 字符串地址
     *     auto value = worksheet.getCell("A1").getValue<std::string>();
     * }
     * if (worksheet.hasCellAt({0, 1})) {       // 坐标地址
     *     // 处理B1单元格
     * }
     */
    bool hasCellAt(const core::Address& address) const { return hasCellAt(address.getRow(), address.getCol()); }

    // 统一地址版：安全获取单元格格式
    std::optional<std::shared_ptr<const FormatDescriptor>>
    tryGetCellFormat(const core::Address& address) const noexcept {
        return tryGetCellFormat(address.getRow(), address.getCol());
    }
    
    /**
     * @brief 获取列宽
     * @param col 列号
     * @return 列宽
     */
    double getColumnWidth(int col) const;
    
    /**
     * @brief 获取行高
     * @param row 行号
     * @return 行高
     */
    double getRowHeight(int row) const;
    
    /**
     * @brief 获取列格式
     * @param col 列号
     * @return 列格式描述符
     */
    std::shared_ptr<const core::FormatDescriptor> getColumnFormat(int col) const;
    
    /**
     * @brief 获取行格式
     * @param row 行号
     * @return 行格式描述符
     */
    std::shared_ptr<const core::FormatDescriptor> getRowFormat(int row) const;
    
    /**
     * @brief 获取列格式ID
     * @param col 列号
     * @return 列格式ID，-1表示无格式
     */
    int getColumnFormatId(int col) const;
    
    /**
     * @brief 获取所有列信息
     * @return 列信息映射
     */
    const std::unordered_map<int, ColumnInfo>& getColumnInfo() const { return column_info_; }
    
    /**
     * @brief 检查列是否隐藏
     * @param col 列号
     * @return 是否隐藏
     */
    bool isColumnHidden(int col) const;
    
    /**
     * @brief 检查行是否隐藏
     * @param row 行号
     * @return 是否隐藏
     */
    bool isRowHidden(int row) const;
    
    /**
     * @brief 获取合并单元格范围
     * @return 合并单元格范围列表
     */
    const std::vector<MergeRange>& getMergeRanges() const { return merge_ranges_; }
    
    /**
     * @brief 检查是否有自动筛选
     * @return 是否有自动筛选
     */
    bool hasAutoFilter() const { return autofilter_ != nullptr; }
    
    /**
     * @brief 获取自动筛选范围
     * @return 自动筛选范围
     */
    AutoFilterRange getAutoFilterRange() const;
    
    /**
     * @brief 检查是否有冻结窗格
     * @return 是否有冻结窗格
     */
    bool hasFrozenPanes() const { return freeze_panes_ != nullptr; }
    
    /**
     * @brief 获取冻结窗格信息
     * @return 冻结窗格信息
     */
    FreezePanes getFreezeInfo() const;
    
    /**
     * @brief 获取保护密码
     * @return 保护密码
     */
    const std::string& getProtectionPassword() const { return protection_password_; }
    
    /**
     * @brief 获取缩放比例
     * @return 缩放比例
     */
    int getZoom() const { return sheet_view_.zoom_scale; }
    
    /**
     * @brief 检查网格线是否可见
     * @return 网格线是否可见
     */
    bool isGridlinesVisible() const { return sheet_view_.show_gridlines; }
    
    /**
     * @brief 检查行列标题是否可见
     * @return 行列标题是否可见
     */
    bool isRowColHeadersVisible() const { return sheet_view_.show_row_col_headers; }
    
    /**
     * @brief 检查是否从右到左
     * @return 是否从右到左
     */
    bool isRightToLeft() const { return sheet_view_.right_to_left; }
    
    // isTabSelected 已在 public 区域声明
    
    /**
     * @brief 获取活动单元格
     * @return 活动单元格引用
     */
    const std::string& getActiveCell() const { return active_cell_; }
    
    /**
     * @brief 获取选中范围
     * @return 选中范围引用
     */
    const std::string& getSelection() const { return selection_; }
    
    // XML生成
    
    /**
     * @brief 生成工作表XML到回调函数（使用UnifiedXMLGenerator）
     * @param callback 数据写入回调函数
     */
    void generateXML(const std::function<void(const char*, size_t)>& callback) const;
    
    // 工具方法
    
    /**
     * @brief 清空工作表
     */
    void clear();
    
    /**
     * @brief 清空指定区域
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void clearRange(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 插入行
     * @param row 插入位置
     * @param count 插入数量
     */
    void insertRows(int row, int count = 1);
    
    /**
     * @brief 插入列
     * @param col 插入位置
     * @param count 插入数量
     */
    void insertColumns(int col, int count = 1);
    
    /**
     * @brief 删除行
     * @param row 删除位置
     * @param count 删除数量
     */
    void deleteRows(int row, int count = 1);
    
    /**
     * @brief 删除列
     * @param col 删除位置
     * @param count 删除数量
     */
    void deleteColumns(int col, int count = 1);
    
    // 单元格编辑功能
    
    /**
     * @brief 修改现有单元格的值
     * @param row 行号
     * @param col 列号
     * @param value 新值
     * @param preserve_format 是否保留原有格式
     */
    void editCellValue(int row, int col, const std::string& value, bool preserve_format = true);
    void editCellValue(int row, int col, double value, bool preserve_format = true);
    void editCellValue(int row, int col, bool value, bool preserve_format = true);
    
    /**
     * @brief 修改现有单元格的值（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @param value 新值
     * @param preserve_format 是否保留原有格式
     * 
     * 示例：
     * worksheet.editCellValue("A1", "新值", true);      // 字符串地址
     * worksheet.editCellValue({0, 0}, 123.45, false);  // 坐标地址
     */
    void editCellValue(const core::Address& address, const std::string& value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    
    void editCellValue(const core::Address& address, double value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    
    void editCellValue(const core::Address& address, bool value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    
    /**
     * @brief 修改单元格格式（新架构 - 推荐）
     * @param row 行号
     * @param col 列号
     * @param format 新格式描述符
     */
    void editCellFormat(int row, int col, const core::FormatDescriptor& format);
    void editCellFormat(int row, int col, std::shared_ptr<const core::FormatDescriptor> format);
    
    /**
     * @brief 修改单元格格式（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @param format 新格式描述符
     * 
     * 示例：
     * worksheet.editCellFormat("A1", format);      // 字符串地址
     * worksheet.editCellFormat({0, 0}, format);    // 坐标地址
     */
    void editCellFormat(const core::Address& address, const core::FormatDescriptor& format) {
        editCellFormat(address.getRow(), address.getCol(), format);
    }
    
    void editCellFormat(const core::Address& address, std::shared_ptr<const core::FormatDescriptor> format) {
        editCellFormat(address.getRow(), address.getCol(), format);
    }
    
    /**
     * @deprecated 使用 FormatDescriptor 版本替代
     * @brief 修改单元格格式（旧架构 - 兼容性保留）
     * @param row 行号
     * @param col 列号
     * @param format 新格式
     */
    [[deprecated("Use FormatDescriptor version instead")]]
    void editCellFormat(int row, int col, std::shared_ptr<Format> format);
    
    // copyCell, copyRange, sortRange 已移动到 public 区域
    
    /**
     * @brief 移动单元格
     * @param src_row 源行号
     * @param src_col 源列号
     * @param dst_row 目标行号
     * @param dst_col 目标列号
     */
    void moveCell(int src_row, int src_col, int dst_row, int dst_col);
    
    // copyRange 已移动到 public 区域
    
    /**
     * @brief 移动范围
     * @param src_first_row 源起始行
     * @param src_first_col 源起始列
     * @param src_last_row 源结束行
     * @param src_last_col 源结束列
     * @param dst_row 目标起始行
     * @param dst_col 目标起始列
     */
    void moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col);
    
    /**
     * @brief 查找并替换
     * @param find_text 查找的文本
     * @param replace_text 替换的文本
     * @param match_case 是否区分大小写
     * @param match_entire_cell 是否匹配整个单元格
     * @return 替换的数量
     */
    
    /**
     * @brief 查找单元格
     * @param search_text 搜索文本
     * @param match_case 是否区分大小写
     * @param match_entire_cell 是否匹配整个单元格
     * @return 匹配的单元格位置列表 (row, col)
     */
    
    // sortRange 已移动到 public 区域
    
    // 共享公式管理
    
    /**
     * @brief 创建共享公式 - 支持地址范围类
     * @param range 范围地址（支持 CellRange("A1:C3") 或坐标）
     * @param formula 基础公式
     * @return 共享索引，失败返回-1
     * 
     * @example
     * worksheet->createSharedFormula("A1:A10", "B{row}*C{row}");
     * worksheet->createSharedFormula(CellRange(0, 0, 9, 0), "B{row}*C{row}");
     */
    int createSharedFormula(const CellRange& range, const std::string& formula);
    
    /**
     * @brief 获取共享公式管理器
     * @return 共享公式管理器指针
     */
    const SharedFormulaManager* getSharedFormulaManager() const { return shared_formula_manager_.get(); }
    
    // 便捷的行列操作方法
    /**
     * @brief 追加行数据
     * @tparam T 数据类型
     * @param data 行数据
     * @return 新行的行号
     * 
     * @example
     * std::vector<std::string> row_data = {"Name", "Age", "Score"};
     * int row_num = worksheet.appendRow(row_data);
     */
    template<typename T>
    int appendRow(const std::vector<T>& data) {
        auto [max_row, max_col] = getUsedRange();
        int new_row = max_row + 1;
        
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(new_row, static_cast<int>(i), data[i]);
        }
        
        return new_row;
    }
    
    /**
     * @brief 获取整行的数据
     * @tparam T 返回类型
     * @param row 行号
     * @return 行数据向量
     * 
     * @example
     * auto row_data = worksheet.getRowData<std::string>(0);
     * for (const auto& cell : row_data) {
     *     std::cout << cell << " ";
     * }
     */
    template<typename T>
    std::vector<T> getRowData(int row) const {
        std::vector<T> result;
        auto [max_row, max_col] = getUsedRange();
        
        for (int col = 0; col <= max_col; ++col) {
            if (hasCellAt(row, col)) {
                result.push_back(getValue<T>(row, col));
            } else {
                result.push_back(T{}); // 默认值
            }
        }
        
        return result;
    }
    
    /**
     * @brief 获取整列的数据
     * @tparam T 返回类型
     * @param col 列号
     * @return 列数据向量
     */
    template<typename T>
    std::vector<T> getColumnData(int col) const {
        std::vector<T> result;
        auto [max_row, max_col] = getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            if (hasCellAt(row, col)) {
                result.push_back(getValue<T>(row, col));
            } else {
                result.push_back(T{}); // 默认值
            }
        }
        
        return result;
    }
    
    /**
     * @brief 清空指定行的所有数据
     * @param row 行号
     */
    void clearRow(int row);
    
    /**
     * @brief 清空指定列的所有数据
     * @param col 列号
     */
    void clearColumn(int col);
    
    /**
     * @brief 清空所有单元格数据
     */
    void clearAll();
    
    /**
     * @brief 批量设置行数据
     * @tparam T 数据类型
     * @param row 行号
     * @param data 数据向量
     * @param start_col 起始列号（默认0）
     */
    template<typename T>
    void setRowData(int row, const std::vector<T>& data, int start_col = 0) {
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(row, start_col + static_cast<int>(i), data[i]);
        }
    }
    
    /**
     * @brief 批量设置列数据
     * @tparam T 数据类型
     * @param col 列号
     * @param data 数据向量
     * @param start_row 起始行号（默认0）
     */
    template<typename T>
    void setColumnData(int col, const std::vector<T>& data, int start_row = 0) {
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(start_row + static_cast<int>(i), col, data[i]);
        }
    }
    
    // 图片插入功能
    
    /**
     * @brief 插入图片到指定单元格
     * @param row 行号（0-based）
     * @param col 列号（0-based）
     * @param image_path 图片文件路径
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(int row, int col, const std::string& image_path);
    
    /**
     * @brief 插入图片到指定单元格（使用Image对象）
     * @param row 行号（0-based）
     * @param col 列号（0-based）
     * @param image 图片对象
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(int row, int col, std::unique_ptr<Image> image);
    
    /**
     * @brief 插入图片到指定范围
     * @param from_row 起始行号
     * @param from_col 起始列号
     * @param to_row 结束行号
     * @param to_col 结束列号
     * @param image_path 图片文件路径
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           const std::string& image_path);
    
    /**
     * @brief 插入图片到指定范围（使用Image对象）
     * @param from_row 起始行号
     * @param from_col 起始列号
     * @param to_row 结束行号
     * @param to_col 结束列号
     * @param image 图片对象
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           std::unique_ptr<Image> image);
    
    /**
     * @brief 插入图片到绝对位置
     * @param x 绝对X坐标（像素）
     * @param y 绝对Y坐标（像素）
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
     * @param image_path 图片文件路径
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImageAt(double x, double y, double width, double height,
                             const std::string& image_path);
    
    /**
     * @brief 插入图片到绝对位置（使用Image对象）
     * @param x 绝对X坐标（像素）
     * @param y 绝对Y坐标（像素）
     * @param width 图片宽度（像素）
     * @param height 图片高度（像素）
     * @param image 图片对象
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImageAt(double x, double y, double width, double height,
                             std::unique_ptr<Image> image);
    
    /**
     * @brief 通过Excel地址插入图片
     * @param address Excel地址（如"A1", "B2"）
     * @param image_path 图片文件路径
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(const std::string& address, const std::string& image_path);
    
    /**
     * @brief 通过Excel地址插入图片（使用Image对象）
     * @param address Excel地址（如"A1", "B2"）
     * @param image 图片对象
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImage(const std::string& address, std::unique_ptr<Image> image);
    
    /**
     * @brief 通过Excel范围插入图片
     * @param range Excel范围（如"A1:C3"）
     * @param image_path 图片文件路径
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImageRange(const std::string& range, const std::string& image_path);
    
    /**
     * @brief 通过Excel范围插入图片（使用Image对象）
     * @param range Excel范围（如"A1:C3"）
     * @param image 图片对象
     * @return 图片ID，失败时返回空字符串
     */
    std::string insertImageRange(const std::string& range, std::unique_ptr<Image> image);
    
    // 图片管理功能
    
    /**
     * @brief 获取所有图片
     * @return 图片列表的常量引用
     */
    
    /**
     * @brief 获取图片数量
     * @return 图片数量
     */
    size_t getImageCount() const { 
        return images_.size(); 
    }
    
    /**
     * @brief 根据ID查找图片
     * @param image_id 图片ID
     * @return 图片指针，未找到时返回nullptr
     */
    const Image* findImage(const std::string& image_id) const;
    
    /**
     * @brief 根据ID查找图片（非常量版本）
     * @param image_id 图片ID
     * @return 图片指针，未找到时返回nullptr
     */
    Image* findImage(const std::string& image_id);
    
    /**
     * @brief 删除指定ID的图片
     * @param image_id 图片ID
     * @return 是否成功删除
     */
    bool removeImage(const std::string& image_id);
    
    /**
     * @brief 清空所有图片
     */
    void clearImages();
    
    /**
     * @brief 检查是否包含图片
     * @return 是否包含图片
     */
    bool hasImages() const { 
        return !images_.empty(); 
    }
    
    /**
     * @brief 获取图片占用的内存大小
     * @return 内存大小（字节）
     */
    size_t getImagesMemoryUsage() const;

    // CSV功能
    
    /**
     * @brief 从CSV文件加载数据到当前工作表
     * @param filepath CSV文件路径
     * @param options CSV解析选项
     * @return 解析结果信息
     */
    CSVParseInfo loadFromCSV(const std::string& filepath, 
                            const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 从CSV字符串加载数据到当前工作表
     * @param csv_content CSV内容字符串
     * @param options CSV解析选项
     * @return 解析结果信息
     */
    CSVParseInfo loadFromCSVString(const std::string& csv_content,
                                  const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 将当前工作表导出为CSV文件
     * @param filepath 目标文件路径
     * @param options CSV导出选项
     * @return 是否成功
     */
    bool saveAsCSV(const std::string& filepath,
                   const CSVOptions& options = CSVOptions::standard()) const;
    
    /**
     * @brief 将当前工作表转换为CSV字符串
     * @param options CSV导出选项
     * @return CSV内容字符串
     */
    std::string toCSVString(const CSVOptions& options = CSVOptions::standard()) const;
    
    /**
     * @brief 将指定范围导出为CSV字符串
     * @param start_row 起始行（0-based）
     * @param start_col 起始列（0-based）
     * @param end_row 结束行（0-based）
     * @param end_col 结束列（0-based）
     * @param options CSV导出选项
     * @return CSV内容字符串
     */
    std::string rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                               const CSVOptions& options = CSVOptions::standard()) const;
    
    /**
     * @brief 预览CSV文件的结构（不加载数据）
     * @param filepath CSV文件路径
     * @param options CSV解析选项
     * @return 文件结构信息
     */
    static CSVParseInfo previewCSV(const std::string& filepath,
                                  const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 自动检测CSV文件的最佳解析选项
     * @param filepath CSV文件路径
     * @return 推荐的解析选项
     */
    static CSVOptions detectCSVOptions(const std::string& filepath);
    
    /**
     * @brief 检查文件是否为CSV格式
     * @param filepath 文件路径
     * @return 是否为CSV文件
     */
    static bool isCSVFile(const std::string& filepath);
    
    /**
     * @brief 获取单元格的显示值（用于CSV导出）
     * @param row 行号（0-based）
     * @param col 列号（0-based）
     * @return 单元格的字符串表示
     */
    std::string getCellDisplayValue(int row, int col) const;
    
    /**
     * @brief 获取单元格的显示值（统一地址接口）
     * @param address 单元格地址（支持隐式转换）
     * @return 单元格的字符串表示
     * 
     * 示例：
     * auto display = worksheet.getCellDisplayValue("A1");     // 字符串地址
     * auto display = worksheet.getCellDisplayValue({0, 0});   // 坐标地址
     */
    std::string getCellDisplayValue(const core::Address& address) const {
        return getCellDisplayValue(address.getRow(), address.getCol());
    }

private:
    // 模板化的单元格操作辅助方法
    template<typename T>
    void editCellValueImpl(int row, int col, T&& value, bool preserve_format);
    
    // 优化相关辅助方法
    void ensureCurrentRow(int row_num);
    void switchToNewRow(int row_num);
    void writeOptimizedCell(int row, int col, Cell&& cell);
    void updateUsedRangeOptimized(int row, int col);
    
    // XML生成辅助方法 - 已移至UnifiedXMLGenerator
    // 保留这些方法声明用于向后兼容，但实际实现由 UnifiedXMLGenerator 负责
    void generateXMLBatch(const std::function<void(const char*, size_t)>& callback) const;
    void generateXMLStreaming(const std::function<void(const char*, size_t)>& callback) const;
    
    // 内部状态管理
    void updateUsedRange(int row, int col);
    void shiftCellsForRowInsertion(int row, int count);
    void shiftCellsForColumnInsertion(int col, int count);
    void shiftCellsForRowDeletion(int row, int count);
    void shiftCellsForColumnDeletion(int col, int count);
};

}} // namespace fastexcel::core
