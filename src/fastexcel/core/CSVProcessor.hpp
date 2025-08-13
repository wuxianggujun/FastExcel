#pragma once

#include <string>
#include <vector>
#include <memory>
#include <functional>
#include <optional>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;

/**
 * @brief CSV处理选项配置
 */
struct CSVOptions {
    // === 基本分隔符配置 ===
    char delimiter = ',';              // 字段分隔符（逗号、分号、制表符等）
    char quote_char = '"';             // 引号字符
    char escape_char = '"';            // 转义字符（双引号转义）
    std::string line_terminator = "\n"; // 行结束符
    
    // === 编码和格式配置 ===
    std::string encoding = "UTF-8";    // 文件编码（UTF-8、GBK、ASCII等）
    bool has_header = true;            // 是否包含标题行
    bool auto_detect_delimiter = true; // 自动检测分隔符
    bool auto_detect_encoding = true;  // 自动检测编码
    
    // === 数据处理选项 ===
    bool skip_empty_lines = true;      // 跳过空行
    bool trim_whitespace = true;       // 去除字段前后空白
    bool ignore_case = false;          // 忽略大小写（用于标题匹配）
    size_t max_field_size = 32768;     // 单个字段最大长度（32KB）
    size_t preview_lines = 100;        // 用于自动检测的预览行数
    
    // === 类型推断配置 ===
    bool auto_detect_types = true;     // 自动检测数据类型
    bool parse_dates = true;           // 解析日期格式
    bool parse_numbers = true;         // 解析数字格式
    std::vector<std::string> date_formats = {  // 支持的日期格式
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"
    };
    
    // === 性能和内存配置 ===
    size_t buffer_size = 8192;         // 读取缓冲区大小
    bool use_streaming = false;        // 是否使用流式读取（大文件）
    size_t streaming_threshold = 50 * 1024 * 1024; // 超过50MB使用流式读取
    
    // === 错误处理配置 ===
    bool strict_mode = false;          // 严格模式（遇到错误立即停止）
    bool log_warnings = true;          // 记录警告信息
    size_t max_errors = 100;           // 最大容忍错误数
    
    /**
     * @brief 获取常用的预设配置
     */
    static CSVOptions standard() { return CSVOptions{}; }
    static CSVOptions excel() { 
        CSVOptions opts;
        opts.delimiter = ',';
        opts.has_header = true;
        return opts;
    }
    static CSVOptions europeanStyle() {
        CSVOptions opts;
        opts.delimiter = ';';
        opts.has_header = true;
        return opts;
    }
    static CSVOptions tabDelimited() {
        CSVOptions opts;
        opts.delimiter = '\t';
        return opts;
    }
};

/**
 * @brief CSV解析结果信息
 */
struct CSVParseInfo {
    size_t rows_processed = 0;         // 处理的行数
    size_t columns_detected = 0;       // 检测到的列数
    char detected_delimiter = ',';     // 检测到的分隔符
    std::string detected_encoding = "UTF-8"; // 检测到的编码
    bool has_header_detected = false;  // 是否检测到标题行
    std::vector<std::string> column_names; // 列名列表
    std::vector<std::string> warnings; // 警告信息
    std::vector<std::string> errors;   // 错误信息
    
    bool isSuccess() const { return errors.empty(); }
    bool hasWarnings() const { return !warnings.empty(); }
};

/**
 * @brief CSV读取器 - 负责解析CSV文件到工作表
 */
class CSVReader {
public:
    /**
     * @brief 从文件加载CSV到工作表
     * @param filepath CSV文件路径
     * @param worksheet 目标工作表
     * @param options 解析选项
     * @return 解析结果信息
     */
    static CSVParseInfo loadFromFile(const std::string& filepath, 
                                    Worksheet& worksheet, 
                                    const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 从字符串加载CSV到工作表
     * @param csv_content CSV内容字符串
     * @param worksheet 目标工作表
     * @param options 解析选项
     * @return 解析结果信息
     */
    static CSVParseInfo loadFromString(const std::string& csv_content,
                                      Worksheet& worksheet,
                                      const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 预览CSV文件结构（不加载数据）
     * @param filepath CSV文件路径
     * @param options 解析选项
     * @return 文件结构信息
     */
    static CSVParseInfo previewFile(const std::string& filepath,
                                   const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 自动检测CSV文件的最佳解析选项
     * @param filepath CSV文件路径
     * @return 推荐的解析选项
     */
    static CSVOptions detectOptions(const std::string& filepath);

private:
    /**
     * @brief 解析CSV内容的核心实现
     */
    static CSVParseInfo parseContent(const std::string& content,
                                    Worksheet& worksheet,
                                    const CSVOptions& options);
    
    /**
     * @brief 自动检测分隔符
     */
    static char detectDelimiter(const std::string& sample);
    
    /**
     * @brief 自动检测编码
     */
    static std::string detectEncoding(const std::vector<uint8_t>& data);
    
    /**
     * @brief 解析单行CSV
     */
    static std::vector<std::string> parseLine(const std::string& line, 
                                             const CSVOptions& options);
    
    /**
     * @brief 推断数据类型并设置单元格值
     */
    static void setCellValue(Worksheet& worksheet, int row, int col, 
                           const std::string& value, const CSVOptions& options);
};

/**
 * @brief CSV写入器 - 负责将工作表导出为CSV
 */
class CSVWriter {
public:
    /**
     * @brief 将工作表保存为CSV文件
     * @param worksheet 源工作表
     * @param filepath 目标文件路径
     * @param options 导出选项
     * @return 是否成功
     */
    static bool saveToFile(const Worksheet& worksheet,
                          const std::string& filepath,
                          const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 将工作表转换为CSV字符串
     * @param worksheet 源工作表
     * @param options 导出选项
     * @return CSV内容字符串
     */
    static std::string saveToString(const Worksheet& worksheet,
                                   const CSVOptions& options = CSVOptions::standard());
    
    /**
     * @brief 将工作表的指定范围导出为CSV
     * @param worksheet 源工作表
     * @param start_row 起始行（0-based）
     * @param start_col 起始列（0-based）
     * @param end_row 结束行（0-based）
     * @param end_col 结束列（0-based）
     * @param options 导出选项
     * @return CSV内容字符串
     */
    static std::string saveRangeToString(const Worksheet& worksheet,
                                        int start_row, int start_col,
                                        int end_row, int end_col,
                                        const CSVOptions& options = CSVOptions::standard());

private:
    /**
     * @brief 转义CSV字段值
     */
    static std::string escapeField(const std::string& field, const CSVOptions& options);
    
    /**
     * @brief 检查字段是否需要引号包围
     */
    static bool needsQuoting(const std::string& field, const CSVOptions& options);
};

/**
 * @brief CSV处理工具类
 */
class CSVUtils {
public:
    /**
     * @brief 检测文件是否为CSV格式
     */
    static bool isCSVFile(const std::string& filepath);
    
    /**
     * @brief 从文件扩展名推断CSV选项
     */
    static CSVOptions optionsFromExtension(const std::string& filepath);
    
    /**
     * @brief 验证CSV选项的合理性
     */
    static bool validateOptions(const CSVOptions& options, std::string& error_message);
    
    /**
     * @brief 获取常见分隔符的描述
     */
    static std::string getDelimiterName(char delimiter);
    
    /**
     * @brief 统计文件行数（快速预览）
     */
    static size_t countLines(const std::string& filepath);
};

}} // namespace fastexcel::core