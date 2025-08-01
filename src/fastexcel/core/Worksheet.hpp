#pragma once

#include "fastexcel/core/Cell.hpp"
#include "fastexcel/core/Format.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <utility>
#include <ctime>
#include <unordered_map>
#include <set>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;

// 列信息结构
struct ColumnInfo {
    double width = -1.0;           // 列宽，-1表示默认
    std::shared_ptr<Format> format; // 列格式
    bool hidden = false;           // 是否隐藏
    bool collapsed = false;        // 是否折叠
    uint8_t outline_level = 0;     // 大纲级别
};

// 行信息结构
struct RowInfo {
    double height = -1.0;          // 行高，-1表示默认
    std::shared_ptr<Format> format; // 行格式
    bool hidden = false;           // 是否隐藏
    bool collapsed = false;        // 是否折叠
    uint8_t outline_level = 0;     // 大纲级别
};

// 合并单元格范围
struct MergeRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
    
    MergeRange(int fr, int fc, int lr, int lc) 
        : first_row(fr), first_col(fc), last_row(lr), last_col(lc) {}
};

// 自动筛选范围
struct AutoFilterRange {
    int first_row;
    int first_col;
    int last_row;
    int last_col;
    
    AutoFilterRange(int fr, int fc, int lr, int lc) 
        : first_row(fr), first_col(fc), last_row(lr), last_col(lc) {}
};

// 冻结窗格信息
struct FreezePanes {
    int row = 0;
    int col = 0;
    int top_left_row = 0;
    int top_left_col = 0;
    
    FreezePanes() = default;
    FreezePanes(int r, int c, int tlr = 0, int tlc = 0) 
        : row(r), col(c), top_left_row(tlr), top_left_col(tlc) {}
};

// 打印设置
struct PrintSettings {
    // 打印区域
    int print_area_first_row = -1;
    int print_area_first_col = -1;
    int print_area_last_row = -1;
    int print_area_last_col = -1;
    
    // 重复行/列
    int repeat_rows_first = -1;
    int repeat_rows_last = -1;
    int repeat_cols_first = -1;
    int repeat_cols_last = -1;
    
    // 页面设置
    bool landscape = false;        // 横向打印
    double left_margin = 0.7;      // 左边距（英寸）
    double right_margin = 0.7;     // 右边距
    double top_margin = 0.75;      // 上边距
    double bottom_margin = 0.75;   // 下边距
    double header_margin = 0.3;    // 页眉边距
    double footer_margin = 0.3;    // 页脚边距
    
    // 缩放
    int scale = 100;               // 缩放百分比
    int fit_to_pages_wide = 0;     // 适合页面宽度
    int fit_to_pages_tall = 0;     // 适合页面高度
    
    // 其他选项
    bool print_gridlines = false;  // 打印网格线
    bool print_headings = false;   // 打印行列标题
    bool center_horizontally = false; // 水平居中
    bool center_vertically = false;   // 垂直居中
};

// 页面视图设置
struct SheetView {
    bool show_gridlines = true;    // 显示网格线
    bool show_row_col_headers = true; // 显示行列标题
    bool show_zeros = true;        // 显示零值
    bool right_to_left = false;    // 从右到左
    bool tab_selected = false;     // 选项卡选中
    bool show_ruler = true;        // 显示标尺
    bool show_outline_symbols = true; // 显示大纲符号
    bool show_white_space = true;  // 显示空白
    int zoom_scale = 100;          // 缩放比例
    int zoom_scale_normal = 100;   // 正常缩放比例
};

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
private:
    std::string name_;
    std::map<std::pair<int, int>, Cell> cells_; // (row, col) -> Cell
    std::shared_ptr<Workbook> parent_workbook_;
    int sheet_id_;
    
    // 行列信息
    std::unordered_map<int, ColumnInfo> column_info_;
    std::unordered_map<int, RowInfo> row_info_;
    
    // 合并单元格
    std::vector<MergeRange> merge_ranges_;
    
    // 自动筛选
    std::unique_ptr<AutoFilterRange> autofilter_;
    
    // 冻结窗格
    std::unique_ptr<FreezePanes> freeze_panes_;
    
    // 打印设置
    PrintSettings print_settings_;
    
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

public:
    explicit Worksheet(const std::string& name, std::shared_ptr<Workbook> workbook, int sheet_id = 1);
    ~Worksheet() = default;
    
    // 禁用拷贝构造和赋值
    Worksheet(const Worksheet&) = delete;
    Worksheet& operator=(const Worksheet&) = delete;
    
    // 允许移动构造和赋值
    Worksheet(Worksheet&&) = default;
    Worksheet& operator=(Worksheet&&) = default;
    
    // ========== 基本单元格操作 ==========
    
    /**
     * @brief 获取单元格引用
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 单元格引用
     */
    Cell& getCell(int row, int col);
    const Cell& getCell(int row, int col) const;
    
    /**
     * @brief 写入字符串
     * @param row 行号
     * @param col 列号
     * @param value 字符串值
     * @param format 格式（可选）
     */
    void writeString(int row, int col, const std::string& value, std::shared_ptr<Format> format = nullptr);
    
    /**
     * @brief 写入数字
     * @param row 行号
     * @param col 列号
     * @param value 数字值
     * @param format 格式（可选）
     */
    void writeNumber(int row, int col, double value, std::shared_ptr<Format> format = nullptr);
    
    /**
     * @brief 写入布尔值
     * @param row 行号
     * @param col 列号
     * @param value 布尔值
     * @param format 格式（可选）
     */
    void writeBoolean(int row, int col, bool value, std::shared_ptr<Format> format = nullptr);
    
    /**
     * @brief 写入公式
     * @param row 行号
     * @param col 列号
     * @param formula 公式（不包含=号）
     * @param format 格式（可选）
     */
    void writeFormula(int row, int col, const std::string& formula, std::shared_ptr<Format> format = nullptr);
    
    /**
     * @brief 写入日期时间
     * @param row 行号
     * @param col 列号
     * @param datetime 日期时间
     * @param format 格式（可选）
     */
    void writeDateTime(int row, int col, const std::tm& datetime, std::shared_ptr<Format> format = nullptr);
    
    /**
     * @brief 写入URL链接
     * @param row 行号
     * @param col 列号
     * @param url URL地址
     * @param string 显示文本（可选）
     * @param format 格式（可选）
     */
    void writeUrl(int row, int col, const std::string& url, const std::string& string = "", 
                  std::shared_ptr<Format> format = nullptr);
    
    // ========== 批量数据操作 ==========
    
    /**
     * @brief 批量写入字符串数据
     * @param start_row 起始行
     * @param start_col 起始列
     * @param data 二维字符串数组
     */
    void writeRange(int start_row, int start_col, const std::vector<std::vector<std::string>>& data);
    
    /**
     * @brief 批量写入数字数据
     * @param start_row 起始行
     * @param start_col 起始列
     * @param data 二维数字数组
     */
    void writeRange(int start_row, int start_col, const std::vector<std::vector<double>>& data);
    
    /**
     * @brief 模板化批量写入
     * @tparam T 数据类型
     * @param start_row 起始行
     * @param start_col 起始列
     * @param data 二维数据数组
     */
    template<typename T>
    void writeRange(int start_row, int start_col, const std::vector<std::vector<T>>& data);
    
    // ========== 行列操作 ==========
    
    /**
     * @brief 设置列宽
     * @param col 列号
     * @param width 宽度
     */
    void setColumnWidth(int col, double width);
    
    /**
     * @brief 设置列宽范围
     * @param first_col 起始列
     * @param last_col 结束列
     * @param width 宽度
     */
    void setColumnWidth(int first_col, int last_col, double width);
    
    /**
     * @brief 设置列格式
     * @param col 列号
     * @param format 格式
     */
    void setColumnFormat(int col, std::shared_ptr<Format> format);
    
    /**
     * @brief 设置列格式范围
     * @param first_col 起始列
     * @param last_col 结束列
     * @param format 格式
     */
    void setColumnFormat(int first_col, int last_col, std::shared_ptr<Format> format);
    
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
     * @param format 格式
     */
    void setRowFormat(int row, std::shared_ptr<Format> format);
    
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
    
    // ========== 合并单元格 ==========
    
    /**
     * @brief 合并单元格
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 合并单元格并写入内容
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @param value 内容
     * @param format 格式（可选）
     */
    void mergeRange(int first_row, int first_col, int last_row, int last_col, 
                    const std::string& value, std::shared_ptr<Format> format = nullptr);
    
    // ========== 自动筛选 ==========
    
    /**
     * @brief 设置自动筛选
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void setAutoFilter(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 移除自动筛选
     */
    void removeAutoFilter();
    
    // ========== 冻结窗格 ==========
    
    /**
     * @brief 冻结窗格
     * @param row 冻结行位置
     * @param col 冻结列位置
     */
    void freezePanes(int row, int col);
    
    /**
     * @brief 冻结窗格（指定左上角单元格）
     * @param row 冻结行位置
     * @param col 冻结列位置
     * @param top_left_row 左上角行
     * @param top_left_col 左上角列
     */
    void freezePanes(int row, int col, int top_left_row, int top_left_col);
    
    /**
     * @brief 分割窗格
     * @param row 分割行位置
     * @param col 分割列位置
     */
    void splitPanes(int row, int col);
    
    // ========== 打印设置 ==========
    
    /**
     * @brief 设置打印区域
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void setPrintArea(int first_row, int first_col, int last_row, int last_col);
    
    /**
     * @brief 设置重复打印行
     * @param first_row 起始行
     * @param last_row 结束行
     */
    void setRepeatRows(int first_row, int last_row);
    
    /**
     * @brief 设置重复打印列
     * @param first_col 起始列
     * @param last_col 结束列
     */
    void setRepeatColumns(int first_col, int last_col);
    
    /**
     * @brief 设置页面方向
     * @param landscape 是否横向
     */
    void setLandscape(bool landscape = true);
    
    /**
     * @brief 设置纸张大小
     * @param paper_size 纸张大小代码
     */
    void setPaperSize(int paper_size);
    
    /**
     * @brief 设置页边距
     * @param left 左边距
     * @param right 右边距
     * @param top 上边距
     * @param bottom 下边距
     */
    void setMargins(double left, double right, double top, double bottom);
    
    /**
     * @brief 设置页眉页脚边距
     * @param header 页眉边距
     * @param footer 页脚边距
     */
    void setHeaderFooterMargins(double header, double footer);
    
    /**
     * @brief 设置打印缩放
     * @param scale 缩放百分比
     */
    void setPrintScale(int scale);
    
    /**
     * @brief 适合页面打印
     * @param width 页面宽度
     * @param height 页面高度
     */
    void setFitToPages(int width, int height);
    
    /**
     * @brief 设置打印网格线
     * @param print 是否打印
     */
    void setPrintGridlines(bool print = true);
    
    /**
     * @brief 设置打印标题
     * @param print 是否打印
     */
    void setPrintHeadings(bool print = true);
    
    /**
     * @brief 设置页面居中
     * @param horizontal 水平居中
     * @param vertical 垂直居中
     */
    void setCenterOnPage(bool horizontal, bool vertical);
    
    // ========== 工作表保护 ==========
    
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
    
    // ========== 视图设置 ==========
    
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
    void setActiveCell(int row, int col);
    
    /**
     * @brief 设置选中范围
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     */
    void setSelection(int first_row, int first_col, int last_row, int last_col);
    
    // ========== 获取信息 ==========
    
    /**
     * @brief 获取工作表名称
     * @return 工作表名称
     */
    const std::string& getName() const { return name_; }
    
    /**
     * @brief 设置工作表名称
     * @param name 新名称
     */
    void setName(const std::string& name) { name_ = name; }
    
    /**
     * @brief 获取工作表ID
     * @return 工作表ID
     */
    int getSheetId() const { return sheet_id_; }
    
    /**
     * @brief 获取使用范围
     * @return (最大行, 最大列)
     */
    std::pair<int, int> getUsedRange() const;
    
    /**
     * @brief 获取单元格数量
     * @return 单元格数量
     */
    size_t getCellCount() const { return cells_.size(); }
    
    /**
     * @brief 检查单元格是否存在
     * @param row 行号
     * @param col 列号
     * @return 是否存在
     */
    bool hasCellAt(int row, int col) const;
    
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
     * @return 列格式
     */
    std::shared_ptr<Format> getColumnFormat(int col) const;
    
    /**
     * @brief 获取行格式
     * @param row 行号
     * @return 行格式
     */
    std::shared_ptr<Format> getRowFormat(int row) const;
    
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
     * @brief 获取打印区域
     * @return 打印区域
     */
    AutoFilterRange getPrintArea() const;
    
    /**
     * @brief 获取重复行范围
     * @return (起始行, 结束行)
     */
    std::pair<int, int> getRepeatRows() const;
    
    /**
     * @brief 获取重复列范围
     * @return (起始列, 结束列)
     */
    std::pair<int, int> getRepeatColumns() const;
    
    /**
     * @brief 检查是否横向打印
     * @return 是否横向打印
     */
    bool isLandscape() const { return print_settings_.landscape; }
    
    /**
     * @brief 获取页边距
     * @return 页边距结构
     */
    struct Margins {
        double left, right, top, bottom;
    };
    Margins getMargins() const;
    
    /**
     * @brief 获取打印缩放
     * @return 缩放百分比
     */
    int getPrintScale() const { return print_settings_.scale; }
    
    /**
     * @brief 获取适应页面设置
     * @return (宽度, 高度)
     */
    std::pair<int, int> getFitToPages() const;
    
    /**
     * @brief 检查是否打印网格线
     * @return 是否打印网格线
     */
    bool isPrintGridlines() const { return print_settings_.print_gridlines; }
    
    /**
     * @brief 检查是否打印标题
     * @return 是否打印标题
     */
    bool isPrintHeadings() const { return print_settings_.print_headings; }
    
    /**
     * @brief 检查是否水平居中
     * @return 是否水平居中
     */
    bool isCenterHorizontally() const { return print_settings_.center_horizontally; }
    
    /**
     * @brief 检查是否垂直居中
     * @return 是否垂直居中
     */
    bool isCenterVertically() const { return print_settings_.center_vertically; }
    
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
    
    /**
     * @brief 检查选项卡是否选中
     * @return 选项卡是否选中
     */
    bool isTabSelected() const { return sheet_view_.tab_selected; }
    
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
    
    // ========== XML生成 ==========
    
    /**
     * @brief 生成工作表XML
     * @return XML字符串
     */
    std::string generateXML() const;
    
    /**
     * @brief 生成工作表关系XML
     * @return XML字符串
     */
    std::string generateRelsXML() const;
    
    // ========== 工具方法 ==========
    
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

private:
    // 内部辅助方法
    std::string columnToLetter(int col) const;
    std::string cellReference(int row, int col) const;
    std::string rangeReference(int first_row, int first_col, int last_row, int last_col) const;
    void validateCellPosition(int row, int col) const;
    void validateRange(int first_row, int first_col, int last_row, int last_col) const;
    
    // XML生成辅助方法
    std::string generateSheetDataXML() const;
    std::string generateColumnsXML() const;
    std::string generateRowsXML() const;
    std::string generateMergeCellsXML() const;
    std::string generateAutoFilterXML() const;
    std::string generateSheetViewsXML() const;
    std::string generatePageSetupXML() const;
    std::string generatePrintOptionsXML() const;
    std::string generatePageMarginsXML() const;
    std::string generateSheetProtectionXML() const;
    
    // 内部状态管理
    void updateUsedRange(int row, int col);
    void shiftCellsForRowInsertion(int row, int count);
    void shiftCellsForColumnInsertion(int col, int count);
    void shiftCellsForRowDeletion(int row, int count);
    void shiftCellsForColumnDeletion(int col, int count);
};

// 模板方法实现
template<typename T>
void Worksheet::writeRange(int start_row, int start_col, const std::vector<std::vector<T>>& data) {
    for (size_t row = 0; row < data.size(); ++row) {
        for (size_t col = 0; col < data[row].size(); ++col) {
            int target_row = static_cast<int>(start_row + row);
            int target_col = static_cast<int>(start_col + col);
            
            if constexpr (std::is_same_v<T, std::string>) {
                writeString(target_row, target_col, data[row][col]);
            } else if constexpr (std::is_arithmetic_v<T>) {
                writeNumber(target_row, target_col, static_cast<double>(data[row][col]));
            } else if constexpr (std::is_same_v<T, bool>) {
                writeBoolean(target_row, target_col, data[row][col]);
            }
        }
    }
}

}} // namespace fastexcel::core
