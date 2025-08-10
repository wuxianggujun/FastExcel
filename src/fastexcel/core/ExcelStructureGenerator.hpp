#pragma once

#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <memory>
#include <functional>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;

/**
 * @brief Excel结构生成器 - 统一的Excel文件生成逻辑
 * 
 * 这个类实现了你文档中设计的核心思想：
 * - 消除批量模式和流式模式的重复代码
 * - 使用策略模式通过IFileWriter抽象写入方式
 * - 提供统一的XML生成流程
 * - 支持智能混合模式（小文件批量，大文件流式）
 */
class ExcelStructureGenerator {
private:
    const Workbook* workbook_;
    std::unique_ptr<IFileWriter> writer_;
    
    // 生成选项
    struct GenerationOptions {
        bool enable_progress_callback = false;
        bool optimize_for_size = false;
        bool validate_xml = false;
        size_t streaming_threshold = 10000; // 单元格数量阈值
        bool parallel_worksheet_generation = false; // 并行生成工作表
        size_t max_memory_limit = 0; // 最大内存限制（0表示不限制）
    } options_;
    
    // 性能统计
    struct PerformanceStats {
        std::chrono::milliseconds total_time;
        std::chrono::milliseconds basic_files_time;
        std::chrono::milliseconds worksheets_time;
        std::chrono::milliseconds finalize_time;
        size_t peak_memory_usage = 0;
    } perf_stats_;
    
    // 进度回调
    std::function<void(const std::string&, int, int)> progress_callback_;
    
public:
    /**
     * @brief 构造函数
     * @param workbook 工作簿指针
     * @param writer 文件写入器（策略模式）
     */
    ExcelStructureGenerator(const Workbook* workbook, std::unique_ptr<IFileWriter> writer);
    
    /**
     * @brief 析构函数
     */
    ~ExcelStructureGenerator();
    
    /**
     * @brief 生成完整的Excel文件结构
     * @return 是否成功
     */
    bool generate();
    
    /**
     * @brief 设置进度回调函数
     * @param callback 回调函数 (stage_name, current, total)
     */
    void setProgressCallback(std::function<void(const std::string&, int, int)> callback);
    
    /**
     * @brief 设置生成选项
     * @param options 生成选项
     */
    void setOptions(const GenerationOptions& options) { options_ = options; }
    
    /**
     * @brief 获取写入器统计信息
     * @return 统计信息
     */
    IFileWriter::WriteStats getWriterStats() const;
    
    /**
     * @brief 获取生成器类型名称
     * @return 类型名称
     */
    std::string getGeneratorType() const;
    
    /**
     * @brief 获取性能统计信息
     * @return 性能统计
     */
    PerformanceStats getPerformanceStats() const { return perf_stats_; }

private:
    // ========== 核心生成方法 ==========
    
    /**
     * @brief 生成基础Excel文件
     * @return 是否成功
     */
    bool generateBasicFiles();
    
    /**
     * @brief 生成工作表文件
     * @return 是否成功
     */
    bool generateWorksheets();
    
    /**
     * @brief 最终化生成过程
     * @return 是否成功
     */
    bool finalize();
    
    // ========== 辅助方法 ==========
    
    /**
     * @brief 报告进度
     * @param stage 阶段名称
     * @param current 当前进度
     * @param total 总进度
     */
    void reportProgress(const std::string& stage, int current, int total);
    
    // 仅保留进度上报等通用辅助
};

}} // namespace fastexcel::core
