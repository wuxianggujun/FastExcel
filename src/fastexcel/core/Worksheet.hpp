#pragma once

#include "fastexcel/core/BlockSparseMatrix.hpp"
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
#include "fastexcel/core/ColumnarStorageManager.hpp"  // 列式存储管理器
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
    
    std::shared_ptr<Workbook> getParentWorkbook() const { return parent_workbook_; }
    
    int findAndReplace(const std::string& find_text, const std::string& replace_text,
                       bool match_case = false, bool match_entire_cell = false);
    
    std::vector<std::pair<int, int>> findCells(const std::string& search_text,
                                               bool match_case = false,
                                               bool match_entire_cell = false) const;
    
    size_t getCellCount() const { return cells_.size(); }
    
    int getSheetId() const { return sheet_id_; }
    int getRowCount() const;
    int getColumnCount() const;
    bool isTabSelected() const { return sheet_view_.tab_selected; }
    
    void copyCell(int src_row, int src_col, int dst_row, int dst_col, bool copy_format = true, bool copy_row_height = false);
    void copyRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col, bool copy_format = true);
    void sortRange(int first_row, int first_col, int last_row, int last_col,
                   int sort_column = 0, bool ascending = true, bool has_header = false);
    
    void setFormula(int row, int col, const std::string& formula, double result = 0.0);
    void setFormula(const Address& address, const std::string& formula, double result = 0.0);
    int createSharedFormula(int first_row, int first_col, int last_row, int last_col, const std::string& formula);
    
    // === 列式存储接口 (通过ColumnarStorageManager) ===
    void enableColumnarMode(const WorkbookOptions* options = nullptr);
    bool isColumnarMode() const;
    
    // 直接存储到列式系统（完全绕过Cell创建）
    void setColumnarValue(uint32_t row, uint32_t col, double value);
    void setColumnarValue(uint32_t row, uint32_t col, uint32_t sst_index);  // SST索引
    void setColumnarValue(uint32_t row, uint32_t col, bool value);
    void setColumnarValue(uint32_t row, uint32_t col, const std::tm& datetime);
    void setColumnarFormula(uint32_t row, uint32_t col, uint32_t formula_index, double result);
    void setColumnarError(uint32_t row, uint32_t col, const std::string& error_code);
    
    // 高效列式数据访问（委托给管理器）
    bool hasColumnarValue(uint32_t row, uint32_t col) const;
    using ColumnarValueVariant = std::variant<std::monostate, double, uint32_t, bool, 
                                            ColumnarStorageManager::FormulaValue, std::string>;
    ColumnarValueVariant getColumnarValue(uint32_t row, uint32_t col) const;
    void forEachInColumn(uint32_t col, std::function<void(uint32_t row, const ColumnarValueVariant& value)> callback) const;
    
    // 类型化列访问
    std::unordered_map<uint32_t, double> getNumberColumn(uint32_t col) const;
    std::unordered_map<uint32_t, uint32_t> getStringColumn(uint32_t col) const;
    std::unordered_map<uint32_t, bool> getBooleanColumn(uint32_t col) const;
    std::unordered_map<uint32_t, double> getDateTimeColumn(uint32_t col) const;
    std::unordered_map<uint32_t, ColumnarStorageManager::FormulaValue> getFormulaColumn(uint32_t col) const;
    std::unordered_map<uint32_t, std::string> getErrorColumn(uint32_t col) const;
    
    // 列式存储统计信息
    size_t getColumnarDataCount() const;
    size_t getColumnarMemoryUsage() const;
    void clearColumnarData();

private:
    std::string name_;
    BlockSparseMatrix cells_; // 传统的分块稀疏矩阵存储
    std::shared_ptr<Workbook> parent_workbook_;
    int sheet_id_;
    
    // === 列式存储系统 (委托给专门的管理器) ===
    std::unique_ptr<ColumnarStorageManager> columnar_storage_;  // 列式存储管理器
    
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
    
    void setSharedStringTable(SharedStringTable* sst) { sst_ = sst; }
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
    void setOptimizeMode(bool enable);
    bool isOptimizeMode() const { return optimize_mode_; }
    void flushCurrentRow();
    size_t getMemoryUsage() const;
    
    struct PerformanceStats {
        size_t total_cells = 0;
        size_t memory_usage = 0;
        size_t sst_strings = 0;
        double sst_compression_ratio = 0.0;
        size_t unique_formats = 0;
        double format_deduplication_ratio = 0.0;
    };
    PerformanceStats getPerformanceStats() const;

public:
    
    Cell& getCell(const core::Address& address);
    const Cell& getCell(const core::Address& address) const;
    
    template<typename T>
    T getValue(int row, int col) const {
        return getCell(row, col).getValue<T>();
    }
    
    template<typename T>
    T getValue(const core::Address& address) const {
        return getValue<T>(address.getRow(), address.getCol());
    }
    
    template<typename T>
    void setValue(int row, int col, const T& value) {
        cell_processor_->setValue(row, col, value);
    }
    
    template<typename T>
    void setValue(const core::Address& address, const T& value) {
        setValue<T>(address.getRow(), address.getCol(), value);
    }
    
    template<typename T>
    void setCellValue(int row, int col, const T& value) {
        setValue<T>(row, col, value);
    }
    
    void setCellFormat(const core::Address& address, const core::FormatDescriptor& format);
    void setCellFormat(const core::Address& address, std::shared_ptr<const core::FormatDescriptor> format);
    void setCellFormat(const core::Address& address, const core::StyleBuilder& builder);
    
    RangeFormatter rangeFormatter(const std::string& range);
    
private:
    RangeFormatter rangeFormatter(int start_row, int start_col, int end_row, int end_col);
public:
    
    RangeFormatter rangeFormatter(const core::CellRange& range) {
        return rangeFormatter(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
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
    
    std::optional<double> tryGetColumnWidth(int col) const noexcept {
        try {
            if (col < 0) return std::nullopt;
            return std::make_optional(getColumnWidth(col));
        } catch (...) {
            return std::nullopt;
        }
    }
    
    std::optional<double> tryGetRowHeight(int row) const noexcept {
        try {
            if (row < 0) return std::nullopt;
            return std::make_optional(getRowHeight(row));
        } catch (...) {
            return std::nullopt;
        }
    }
    
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
    
private:
    template<typename T>
    std::optional<T> tryGetValue(int row, int col) const noexcept {
        if (!hasCellAt(row, col)) {
            return std::nullopt;
        }
        return getCell(row, col).tryGetValue<T>();
    }
public:
    
    template<typename T>
    std::optional<T> tryGetValue(const core::Address& address) const noexcept {
        return tryGetValue<T>(address.getRow(), address.getCol());
    }
    
    template<typename T>
    T getValueOr(int row, int col, const T& default_value) const noexcept {
        if (!hasCellAt(row, col)) {
            return default_value;
        }
        return getCell(row, col).getValueOr<T>(default_value);
    }
    
    template<typename T>
    T getValueOr(const core::Address& address, const T& default_value) const noexcept {
        return getValueOr<T>(address.getRow(), address.getCol(), default_value);
    }
    
private:
    void writeDateTime(int row, int col, const std::tm& datetime);
public:
    
    void writeDateTime(const core::Address& address, const std::tm& datetime);
    
private:
    void writeUrl(int row, int col, const std::string& url, const std::string& string = "");
public:
    
    void writeUrl(const core::Address& address, const std::string& url, const std::string& string = "");

    // 视图与窗格（统一地址接口）
    void freezePanes(const core::Address& split_cell);
    void freezePanes(const core::Address& split_cell, const core::Address& top_left_cell);
    
    void freezePanes(int row, int col);
    void freezePanes(int row, int col, int top_left_row, int top_left_col);
    void splitPanes(const core::Address& split_cell);
    
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
    
    template<typename T>
    std::vector<std::vector<T>> getRange(const core::CellRange& range) const {
        return getRange<T>(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    template<typename T>
    std::vector<std::vector<T>> getRange(const std::string& range) const {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        return getRange<T>(start_row, start_col, end_row, end_col);
    }
    
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
    
    template<typename T>
    void setRange(const core::CellRange& range, const std::vector<std::vector<T>>& data) {
        setRange<T>(range.getStartRow(), range.getStartCol(), data);
    }
    
    template<typename T>
    void setRange(const std::string& range, const std::vector<std::vector<T>>& data) {
        auto [sheet, start_row, start_col, end_row, end_col] = utils::AddressParser::parseRange(range);
        setRange<T>(start_row, start_col, data);
    }
    
    WorksheetChain chain();
    
    // 行列操作
    
    double setColumnWidth(int col, double width);
    std::pair<double, int> setColumnWidthWithFont(int col, double width,
                                                  const std::string& font_name,
                                                  double font_size = 11.0);
    std::pair<double, int> setColumnWidthAdvanced(int col, double target_width,
                                                  const std::string& font_name,
                                                  double font_size,
                                                  ColumnWidthManager::WidthStrategy strategy,
                                                  const std::vector<std::string>& cell_contents = {});
    std::unordered_map<int, std::pair<double, int>> setColumnWidthsBatch(
        const std::unordered_map<int, ColumnWidthManager::ColumnWidthConfig>& configs);
    double calculateOptimalWidth(double target_width, const std::string& font_name, double font_size) const;
    ColumnWidthManager::CacheStats getColumnWidthCacheStats() const {
        if (column_width_manager_) {
            return column_width_manager_->getCacheStats();
        }
        return {0, 0, 0};
    }
    
    void clearColumnWidthCache() {
        if (column_width_manager_) {
            column_width_manager_->clearCache();
        }
    }
    
    void setColumnFormatId(int col, int format_id);
    void setColumnFormatId(int first_col, int last_col, int format_id);
    void setColumnFormat(int col, const core::FormatDescriptor& format);
    void setColumnFormat(int col, std::shared_ptr<const core::FormatDescriptor> format);
    void setColumnFormat(int first_col, int last_col, const core::FormatDescriptor& format);
    void setColumnFormat(int first_col, int last_col, std::shared_ptr<const core::FormatDescriptor> format);
    void hideColumn(int col);
    void hideColumn(int first_col, int last_col);
    void setRowHeight(int row, double height);
    void setRowFormat(int row, const core::FormatDescriptor& format);
    void setRowFormat(int row, std::shared_ptr<const core::FormatDescriptor> format);
    void hideRow(int row);
    void hideRow(int first_row, int last_row);
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    void mergeCells(const core::CellRange& range) {
        mergeCells(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    
    void setAutoFilter(int first_row, int first_col, int last_row, int last_col);
    void setAutoFilter(const core::CellRange& range) {
        setAutoFilter(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    void removeAutoFilter();
    void protect(const std::string& password = "");
    void unprotect();
    bool isProtected() const { return protected_; }
    void setZoom(int scale);
    void showGridlines(bool show = true);
    void showRowColHeaders(bool show = true);
    void setRightToLeft(bool rtl = true);
    void setTabSelected(bool selected = true);

private:
    
    void setActiveCell(const core::Address& address);
    std::string insertImage(const core::Address& address, const std::string& image_path) {
        return insertImage(address.getRow(), address.getCol(), image_path);
    }
    std::string insertImage(const core::Address& address, std::unique_ptr<Image> image) {
        return insertImage(address.getRow(), address.getCol(), std::move(image));
    }
    std::string insertImage(const core::CellRange& range, const std::string& image_path) {
        return insertImage(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol(), image_path);
    }
    void setSelection(int first_row, int first_col, int last_row, int last_col);
    void setSelection(const core::CellRange& range) {
        setSelection(range.getStartRow(), range.getStartCol(), range.getEndRow(), range.getEndCol());
    }
    
    bool isEmpty() const { return cells_.empty(); }
    bool hasData() const { return !cells_.empty(); }
    int getCellCountInRow(int row) const;
    int getCellCountInColumn(int col) const;
    bool hasCellAt(const core::Address& address) const { return hasCellAt(address.getRow(), address.getCol()); }
    std::optional<std::shared_ptr<const FormatDescriptor>>
    tryGetCellFormat(const core::Address& address) const noexcept {
        return tryGetCellFormat(address.getRow(), address.getCol());
    }
    double getColumnWidth(int col) const;
    double getRowHeight(int row) const;
    std::shared_ptr<const core::FormatDescriptor> getColumnFormat(int col) const;
    std::shared_ptr<const core::FormatDescriptor> getRowFormat(int row) const;
    int getColumnFormatId(int col) const;
    const std::unordered_map<int, ColumnInfo>& getColumnInfo() const { return column_info_; }
    bool isColumnHidden(int col) const;
    bool isRowHidden(int row) const;
    const std::vector<MergeRange>& getMergeRanges() const { return merge_ranges_; }
    bool hasAutoFilter() const { return autofilter_ != nullptr; }
    AutoFilterRange getAutoFilterRange() const;
    bool hasFrozenPanes() const { return freeze_panes_ != nullptr; }
    FreezePanes getFreezeInfo() const;
    const std::string& getProtectionPassword() const { return protection_password_; }
    int getZoom() const { return sheet_view_.zoom_scale; }
    bool isGridlinesVisible() const { return sheet_view_.show_gridlines; }
    bool isRowColHeadersVisible() const { return sheet_view_.show_row_col_headers; }
    bool isRightToLeft() const { return sheet_view_.right_to_left; }
    const std::string& getActiveCell() const { return active_cell_; }
    const std::string& getSelection() const { return selection_; }
    void generateXML(const std::function<void(const std::string&)>& callback) const;
    void clear();
    void clearRange(int first_row, int first_col, int last_row, int last_col);
    void insertRows(int row, int count = 1);
    void insertColumns(int col, int count = 1);
    void deleteRows(int row, int count = 1);
    void deleteColumns(int col, int count = 1);
    void editCellValue(int row, int col, const std::string& value, bool preserve_format = true);
    void editCellValue(int row, int col, double value, bool preserve_format = true);
    void editCellValue(int row, int col, bool value, bool preserve_format = true);
    void editCellValue(const core::Address& address, const std::string& value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    void editCellValue(const core::Address& address, double value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    void editCellValue(const core::Address& address, bool value, bool preserve_format = true) {
        editCellValue(address.getRow(), address.getCol(), value, preserve_format);
    }
    void editCellFormat(int row, int col, const core::FormatDescriptor& format);
    void editCellFormat(int row, int col, std::shared_ptr<const core::FormatDescriptor> format);
    void editCellFormat(const core::Address& address, const core::FormatDescriptor& format) {
        editCellFormat(address.getRow(), address.getCol(), format);
    }
    void editCellFormat(const core::Address& address, std::shared_ptr<const core::FormatDescriptor> format) {
        editCellFormat(address.getRow(), address.getCol(), format);
    }
    [[deprecated("Use FormatDescriptor version instead")]]
    void editCellFormat(int row, int col, std::shared_ptr<Format> format);
    void moveCell(int src_row, int src_col, int dst_row, int dst_col);
    void moveRange(int src_first_row, int src_first_col, int src_last_row, int src_last_col,
                   int dst_row, int dst_col);
    int createSharedFormula(const CellRange& range, const std::string& formula);
    const SharedFormulaManager* getSharedFormulaManager() const { return shared_formula_manager_.get(); }
    template<typename T>
    int appendRow(const std::vector<T>& data) {
        auto [max_row, max_col] = getUsedRange();
        int new_row = max_row + 1;
        
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(new_row, static_cast<int>(i), data[i]);
        }
        
        return new_row;
    }
    template<typename T>
    std::vector<T> getRowData(int row) const {
        std::vector<T> result;
        auto [max_row, max_col] = getUsedRange();
        
        for (int col = 0; col <= max_col; ++col) {
            if (hasCellAt(row, col)) {
                result.push_back(getValue<T>(row, col));
            } else {
                result.push_back(T{});
            }
        }
        
        return result;
    }
    template<typename T>
    std::vector<T> getColumnData(int col) const {
        std::vector<T> result;
        auto [max_row, max_col] = getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            if (hasCellAt(row, col)) {
                result.push_back(getValue<T>(row, col));
            } else {
                result.push_back(T{});
            }
        }
        
        return result;
    }
    void clearRow(int row);
    void clearColumn(int col);
    void clearAll();
    template<typename T>
    void setRowData(int row, const std::vector<T>& data, int start_col = 0) {
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(row, start_col + static_cast<int>(i), data[i]);
        }
    }
    template<typename T>
    void setColumnData(int col, const std::vector<T>& data, int start_row = 0) {
        for (size_t i = 0; i < data.size(); ++i) {
            setValue(start_row + static_cast<int>(i), col, data[i]);
        }
    }
    std::string insertImage(int row, int col, const std::string& image_path);
    std::string insertImage(int row, int col, std::unique_ptr<Image> image);
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           const std::string& image_path);
    std::string insertImage(int from_row, int from_col, int to_row, int to_col,
                           std::unique_ptr<Image> image);
    std::string insertImageAt(double x, double y, double width, double height,
                             const std::string& image_path);
    std::string insertImageAt(double x, double y, double width, double height,
                             std::unique_ptr<Image> image);
    std::string insertImage(const std::string& address, const std::string& image_path);
    std::string insertImage(const std::string& address, std::unique_ptr<Image> image);
    std::string insertImageRange(const std::string& range, const std::string& image_path);
    std::string insertImageRange(const std::string& range, std::unique_ptr<Image> image);
    size_t getImageCount() const {
        return images_.size();
    }
    const Image* findImage(const std::string& image_id) const;
    Image* findImage(const std::string& image_id);
    bool removeImage(const std::string& image_id);
    void clearImages();
    bool hasImages() const {
        return !images_.empty();
    }
    size_t getImagesMemoryUsage() const;
    CSVParseInfo loadFromCSV(const std::string& filepath,
                            const CSVOptions& options = CSVOptions::standard());
    CSVParseInfo loadFromCSVString(const std::string& csv_content,
                                  const CSVOptions& options = CSVOptions::standard());
    bool saveAsCSV(const std::string& filepath,
                   const CSVOptions& options = CSVOptions::standard()) const;
    std::string toCSVString(const CSVOptions& options = CSVOptions::standard()) const;
    std::string rangeToCSVString(int start_row, int start_col, int end_row, int end_col,
                               const CSVOptions& options = CSVOptions::standard()) const;
    static CSVParseInfo previewCSV(const std::string& filepath,
                                  const CSVOptions& options = CSVOptions::standard());
    static CSVOptions detectCSVOptions(const std::string& filepath);
    static bool isCSVFile(const std::string& filepath);
    std::string getCellDisplayValue(int row, int col) const;
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
    
    // 内部状态管理
    void updateUsedRange(int row, int col);
    void shiftCellsForRowInsertion(int row, int count);
    void shiftCellsForColumnInsertion(int col, int count);
    void shiftCellsForRowDeletion(int row, int count);
    void shiftCellsForColumnDeletion(int col, int count);
};

}} // namespace fastexcel::core
