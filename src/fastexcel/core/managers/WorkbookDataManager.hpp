#pragma once

#include <string>
#include <memory>
#include <vector>
#include <functional>
#include <optional>
#include "fastexcel/core/CSVProcessor.hpp"

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;
class Worksheet;

/**
 * @brief 工作簿数据管理器 - 负责数据导入导出功能
 * 
 * 设计原则：
 * - 单一职责：专注于数据导入导出
 * - 格式支持：支持多种数据格式
 * - 性能优化：大数据处理优化
 */
class WorkbookDataManager {
public:
    // 数据格式类型
    enum class DataFormat {
        CSV,
        TSV,
        JSON,
        XML,
        TXT
    };
    
    // 导入结果
    struct ImportResult {
        bool success = false;
        size_t rows_imported = 0;
        size_t cols_imported = 0;
        std::string error_message;
        std::shared_ptr<Worksheet> worksheet;  // 导入到的工作表
    };
    
    // 导出结果
    struct ExportResult {
        bool success = false;
        size_t rows_exported = 0;
        size_t cols_exported = 0;
        size_t bytes_written = 0;
        std::string error_message;
        std::string output_path;
    };
    
    // 批量操作进度回调
    using ProgressCallback = std::function<void(size_t current, size_t total, const std::string& status)>;

private:
    Workbook* workbook_;  // 父工作簿引用（不拥有）
    
    // CSV处理器
    std::unique_ptr<CSVProcessor> csv_processor_;
    
    // 配置
    struct Configuration {
        size_t max_import_rows = 1048576;     // 最大导入行数（Excel限制）
        size_t max_import_cols = 16384;       // 最大导入列数（Excel限制）
        size_t batch_size = 1000;             // 批处理大小
        bool skip_empty_rows = true;          // 跳过空行
        bool auto_detect_types = true;        // 自动检测数据类型
        bool preserve_formatting = false;     // 保留格式（仅支持的格式）
    } config_;

public:
    explicit WorkbookDataManager(Workbook* workbook);
    ~WorkbookDataManager() = default;
    
    // 禁用拷贝，允许移动
    WorkbookDataManager(const WorkbookDataManager&) = delete;
    WorkbookDataManager& operator=(const WorkbookDataManager&) = delete;
    WorkbookDataManager(WorkbookDataManager&&) = default;
    WorkbookDataManager& operator=(WorkbookDataManager&&) = default;
    
    // === CSV 导入导出功能 ===
    
    /**
     * @brief 从CSV文件导入数据到新工作表
     * @param filepath CSV文件路径
     * @param sheet_name 工作表名称（空则自动生成）
     * @param options CSV解析选项
     * @param progress 进度回调
     * @return 导入结果
     */
    ImportResult importCSV(const std::string& filepath, 
                          const std::string& sheet_name,
                          const CSVOptions& options,
                          ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    ImportResult importCSV(const std::string& filepath);
    
    /**
     * @brief 从CSV字符串导入数据到新工作表
     * @param csv_content CSV内容字符串
     * @param sheet_name 工作表名称
     * @param options CSV解析选项
     * @param progress 进度回调
     * @return 导入结果
     */
    ImportResult importCSVString(const std::string& csv_content,
                                 const std::string& sheet_name,
                                 const CSVOptions& options,
                                 ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    ImportResult importCSVString(const std::string& csv_content);
    
    /**
     * @brief 导出工作表为CSV文件
     * @param sheet_index 工作表索引
     * @param filepath 输出文件路径
     * @param options CSV导出选项
     * @param progress 进度回调
     * @return 导出结果
     */
    ExportResult exportCSV(size_t sheet_index, const std::string& filepath,
                          const CSVOptions& options,
                          ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    ExportResult exportCSV(size_t sheet_index, const std::string& filepath);
    
    /**
     * @brief 导出工作表为CSV文件（按名称）
     * @param sheet_name 工作表名称
     * @param filepath 输出文件路径
     * @param options CSV导出选项
     * @param progress 进度回调
     * @return 导出结果
     */
    ExportResult exportCSVByName(const std::string& sheet_name, const std::string& filepath,
                                const CSVOptions& options,
                                ProgressCallback progress = nullptr);
    ExportResult exportCSV(const std::string& sheet_name, const std::string& filepath,
                          const CSVOptions& options,
                          ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    ExportResult exportCSV(const std::string& sheet_name, const std::string& filepath);
    
    /**
     * @brief 导出工作表为CSV字符串
     * @param sheet_index 工作表索引
     * @param options CSV导出选项
     * @return CSV内容字符串
     */
    std::string exportCSVString(size_t sheet_index, 
                               const CSVOptions& options);
    // 便捷重载：默认参数
    std::string exportCSVString(size_t sheet_index);
    
    /**
     * @brief 导出工作表为CSV字符串（按名称）
     * @param sheet_name 工作表名称
     * @param options CSV导出选项
     * @return CSV内容字符串
     */
    std::string exportCSVString(const std::string& sheet_name, 
                               const CSVOptions& options);
    // 便捷重载：默认参数
    std::string exportCSVString(const std::string& sheet_name);
    
    // === 其他格式支持 ===
    
    /**
     * @brief 导入TSV（Tab分隔值）文件
     * @param filepath 文件路径
     * @param sheet_name 工作表名称
     * @param progress 进度回调
     * @return 导入结果
     */
    ImportResult importTSV(const std::string& filepath, 
                          const std::string& sheet_name = "",
                          ProgressCallback progress = nullptr);
    
    /**
     * @brief 导出为TSV文件
     * @param sheet_index 工作表索引
     * @param filepath 输出文件路径
     * @param progress 进度回调
     * @return 导出结果
     */
    ExportResult exportTSV(size_t sheet_index, const std::string& filepath,
                          ProgressCallback progress = nullptr);
    
    /**
     * @brief 导入纯文本文件（固定宽度分列）
     * @param filepath 文件路径
     * @param column_widths 列宽度数组
     * @param sheet_name 工作表名称
     * @param progress 进度回调
     * @return 导入结果
     */
    ImportResult importFixedWidth(const std::string& filepath,
                                 const std::vector<size_t>& column_widths,
                                 const std::string& sheet_name = "",
                                 ProgressCallback progress = nullptr);
    
    // === 批量操作 ===
    
    /**
     * @brief 批量导入多个CSV文件
     * @param filepaths 文件路径列表
     * @param options CSV选项
     * @param progress 进度回调
     * @return 导入结果列表
     */
    std::vector<ImportResult> batchImportCSV(const std::vector<std::string>& filepaths,
                                            const CSVOptions& options,
                                            ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    std::vector<ImportResult> batchImportCSV(const std::vector<std::string>& filepaths);
    
    /**
     * @brief 批量导出工作表为CSV
     * @param export_configs 导出配置列表 (工作表名, 输出路径)
     * @param options CSV选项
     * @param progress 进度回调
     * @return 导出结果列表
     */
    std::vector<ExportResult> batchExportCSV(
        const std::vector<std::pair<std::string, std::string>>& export_configs,
        const CSVOptions& options,
        ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    std::vector<ExportResult> batchExportCSV(
        const std::vector<std::pair<std::string, std::string>>& export_configs);
    
    /**
     * @brief 导出所有工作表为CSV文件
     * @param output_directory 输出目录
     * @param filename_prefix 文件名前缀
     * @param options CSV选项
     * @param progress 进度回调
     * @return 导出结果列表
     */
    std::vector<ExportResult> exportAllSheetsAsCSV(const std::string& output_directory,
                                                   const std::string& filename_prefix,
                                                   const CSVOptions& options,
                                                   ProgressCallback progress = nullptr);
    // 便捷重载：默认参数
    std::vector<ExportResult> exportAllSheetsAsCSV(const std::string& output_directory);
    
    // === 数据验证和预览 ===
    
    /**
     * @brief 预览CSV文件内容
     * @param filepath 文件路径
     * @param max_rows 最大预览行数
     * @param options CSV选项
     * @return 预览数据（行->列->值）
     */
    std::vector<std::vector<std::string>> previewCSV(const std::string& filepath,
                                                    size_t max_rows,
                                                    const CSVOptions& options);
    // 便捷重载：默认参数
    std::vector<std::vector<std::string>> previewCSV(const std::string& filepath);
    
    /**
     * @brief 检测CSV文件信息
     * @param filepath 文件路径
     * @return 文件信息 (行数, 列数, 分隔符等)
     */
    struct CSVInfo {
        size_t estimated_rows = 0;
        size_t estimated_cols = 0;
        char detected_delimiter = ',';
        char detected_quote = '"';
        bool has_header = false;
        size_t file_size_bytes = 0;
        std::string encoding = "UTF-8";
    };
    CSVInfo detectCSVInfo(const std::string& filepath);
    
    /**
     * @brief 验证数据格式是否支持
     * @param filepath 文件路径
     * @return 支持的数据格式，DataFormat 枚举值，如果不支持返回 nullopt
     */
    std::optional<DataFormat> detectDataFormat(const std::string& filepath);
    
    /**
     * @brief 检查文件是否为CSV格式
     * @param filepath 文件路径
     * @return 是否为CSV文件
     */
    static bool isCSVFile(const std::string& filepath);
    
    // === 数据转换和处理 ===
    
    /**
     * @brief 数据清理选项
     */
    struct DataCleaningOptions {
        bool trim_whitespace = true;        // 去除首尾空白
        bool remove_empty_rows = true;      // 移除空行
        bool remove_empty_cols = false;     // 移除空列
        bool normalize_line_endings = true; // 标准化行尾符
        std::string null_value_replacement = "";  // 空值替换
    };
    
    /**
     * @brief 清理导入的数据
     * @param worksheet 工作表
     * @param options 清理选项
     * @return 是否成功
     */
    bool cleanImportedData(std::shared_ptr<Worksheet> worksheet, 
                          const DataCleaningOptions& options);
    // 便捷重载：默认参数
    bool cleanImportedData(std::shared_ptr<Worksheet> worksheet);
    
    /**
     * @brief 自动检测并转换数据类型
     * @param worksheet 工作表
     * @param start_row 开始行（0-based）
     * @param end_row 结束行（-1表示到最后）
     * @return 转换的单元格数量
     */
    size_t autoConvertDataTypes(std::shared_ptr<Worksheet> worksheet, 
                               size_t start_row = 0, 
                               int end_row = -1);
    
    // === 配置管理 ===
    
    Configuration& getConfiguration() { return config_; }
    const Configuration& getConfiguration() const { return config_; }
    
    /**
     * @brief 设置批处理大小
     * @param batch_size 批处理大小
     */
    void setBatchSize(size_t batch_size) { config_.batch_size = batch_size; }
    
    /**
     * @brief 设置最大导入限制
     * @param max_rows 最大行数
     * @param max_cols 最大列数
     */
    void setImportLimits(size_t max_rows, size_t max_cols) {
        config_.max_import_rows = max_rows;
        config_.max_import_cols = max_cols;
    }
    
    // === 统计信息 ===
    
    struct Statistics {
        size_t total_imports = 0;
        size_t total_exports = 0;
        size_t total_rows_processed = 0;
        size_t total_bytes_processed = 0;
        size_t failed_operations = 0;
    };
    
    const Statistics& getStatistics() const { return stats_; }
    void resetStatistics() { stats_ = Statistics(); }

private:
    mutable Statistics stats_;
    
    // 内部辅助方法
    ImportResult importData(const std::string& data, DataFormat format,
                           const std::string& sheet_name,
                           const CSVOptions* csv_options = nullptr,
                           ProgressCallback progress = nullptr);
    
    ExportResult exportData(std::shared_ptr<const Worksheet> worksheet,
                           const std::string& filepath, DataFormat format,
                           const CSVOptions* csv_options = nullptr,
                           ProgressCallback progress = nullptr);
    
    std::string generateUniqueSheetName(const std::string& base_name, const std::string& file_extension);
    
    bool isNumeric(const std::string& str) const;
    bool validateImportLimits(size_t rows, size_t cols) const;
    
    void updateStatistics(const ImportResult& result, size_t bytes_processed = 0);
    void updateStatistics(const ExportResult& result);
    
    CSVOptions createTSVOptions() const;  // 创建TSV选项
};

}} // namespace fastexcel::core
