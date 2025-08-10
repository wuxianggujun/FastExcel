#pragma once

#include "fastexcel/FastExcel.hpp"
#include <chrono>
#include <vector>
#include <map>
#include <string>
#include <functional>
#include <memory>
#include <fstream>

namespace fastexcel {
namespace test {

/**
 * @brief 性能测试结果
 */
struct PerformanceResult {
    std::string test_name;
    double execution_time_ms;
    size_t memory_usage_kb;
    size_t peak_memory_kb;
    size_t operations_count;
    double operations_per_second;
    size_t file_size_bytes;
    std::map<std::string, double> custom_metrics;
    
    // 计算操作速率
    void calculateOperationRate() {
        if (execution_time_ms > 0) {
            operations_per_second = (operations_count * 1000.0) / execution_time_ms;
        }
    }
};

/**
 * @brief 性能比较结果
 */
struct PerformanceComparison {
    std::string metric_name;
    double baseline_value;
    double current_value;
    double improvement_ratio;  // 正数表示改进，负数表示退化
    bool is_significant;       // 是否有显著差异
};

/**
 * @brief 内存使用监控器
 */
class MemoryMonitor {
private:
    size_t initial_memory_;
    size_t peak_memory_;
    bool monitoring_;

public:
    MemoryMonitor();
    ~MemoryMonitor();
    
    void startMonitoring();
    void stopMonitoring();
    size_t getCurrentMemoryUsage();
    size_t getPeakMemoryUsage();
    void recordMemorySnapshot(const std::string& checkpoint);
    
private:
    size_t getProcessMemoryUsage();
    std::vector<std::pair<std::string, size_t>> memory_snapshots_;
};

/**
 * @brief 性能测试基础类
 */
class PerformanceTestBase {
protected:
    std::string test_suite_name_;
    std::vector<PerformanceResult> results_;
    std::unique_ptr<MemoryMonitor> memory_monitor_;

public:
    explicit PerformanceTestBase(const std::string& suite_name);
    virtual ~PerformanceTestBase() = default;

    // 执行性能测试
    template<typename Func>
    PerformanceResult measurePerformance(const std::string& test_name, 
                                        size_t operations_count,
                                        Func&& test_function);

    // 生成测试报告
    void generateReport(const std::string& output_file = "");
    void exportToCSV(const std::string& csv_file);
    void exportToJSON(const std::string& json_file);
    
    // 与基准对比
    std::vector<PerformanceComparison> compareWithBaseline(const std::string& baseline_file);
    
    // 获取结果
    const std::vector<PerformanceResult>& getResults() const { return results_; }
    
protected:
    void setupTest();
    void teardownTest();
    void logResult(const PerformanceResult& result);
};

/**
 * @brief 读取性能测试类
 */
class ReadPerformanceTest : public PerformanceTestBase {
public:
    ReadPerformanceTest() : PerformanceTestBase("ReadPerformance") {}

    // 基础读取测试
    void testBasicFileReading();
    void testLargeFileReading();
    void testMultipleFilesReading();
    
    // 共享公式读取测试
    void testSharedFormulaReading();
    void testComplexFormulaReading();
    
    // 不同文件大小的读取测试
    void testReadingByFileSize();
    
    // 内存使用测试
    void testMemoryUsageWhileReading();
    
    // 运行所有读取测试
    void runAllTests();

private:
    void createTestFiles();
    std::vector<std::string> test_files_;
};

/**
 * @brief 写入性能测试类
 */
class WritePerformanceTest : public PerformanceTestBase {
public:
    WritePerformanceTest() : PerformanceTestBase("WritePerformance") {}

    // 基础写入测试
    void testBasicFileWriting();
    void testLargeFileWriting();
    void testBatchWriting();
    void testStreamingWriting();
    
    // 共享公式写入测试
    void testSharedFormulaWriting();
    void testFormulaOptimizationWriting();
    
    // 不同数据量的写入测试
    void testWritingByDataSize();
    
    // 批量模式 vs 流式模式对比
    void testBatchVsStreamingMode();
    
    // 运行所有写入测试
    void runAllTests();

private:
    void generateTestData(size_t rows, size_t cols, double formula_ratio = 0.3);
    std::unique_ptr<fastexcel::core::Workbook> test_workbook_;
};

/**
 * @brief 解析性能测试类
 */
class ParsingPerformanceTest : public PerformanceTestBase {
public:
    ParsingPerformanceTest() : PerformanceTestBase("ParsingPerformance") {}

    // XML解析测试
    void testXMLParsingSpeed();
    void testLargeXMLParsing();
    
    // 样式解析测试
    void testStylesParsing();
    void testComplexStylesParsing();
    
    // 共享字符串解析测试
    void testSharedStringsParsing();
    
    // 工作表解析测试
    void testWorksheetParsing();
    void testMultipleWorksheetsParsing();
    
    // 公式解析测试
    void testFormulaParsingSpeed();
    void testComplexFormulaParsing();
    void testSharedFormulaParsing();
    
    // 运行所有解析测试
    void runAllTests();

private:
    void prepareTestXMLData();
    std::map<std::string, std::string> test_xml_data_;
};

/**
 * @brief 共享公式性能测试类
 */
class SharedFormulaPerformanceTest : public PerformanceTestBase {
public:
    SharedFormulaPerformanceTest() : PerformanceTestBase("SharedFormulaPerformance") {}

    // 共享公式创建性能
    void testSharedFormulaCreation();
    void testLargeSharedFormulaCreation();
    
    // 公式优化性能
    void testFormulaOptimizationSpeed();
    void testBatchOptimization();
    
    // 模式检测性能
    void testPatternDetectionSpeed();
    void testComplexPatternDetection();
    
    // 内存使用对比
    void testMemoryUsageComparison();
    
    // XML生成性能
    void testSharedFormulaXMLGeneration();
    
    // 完整优化流程性能
    void testFullOptimizationWorkflow();
    
    // 运行所有共享公式测试
    void runAllTests();

private:
    void createFormulaTestData(size_t formula_count);
    std::map<std::pair<int, int>, std::string> test_formulas_;
};

/**
 * @brief 综合性能测试套件
 */
class ComprehensivePerformanceTestSuite {
private:
    std::unique_ptr<ReadPerformanceTest> read_test_;
    std::unique_ptr<WritePerformanceTest> write_test_;
    std::unique_ptr<ParsingPerformanceTest> parsing_test_;
    std::unique_ptr<SharedFormulaPerformanceTest> shared_formula_test_;
    
    std::string output_directory_;

public:
    ComprehensivePerformanceTestSuite(const std::string& output_dir = "performance_results");
    ~ComprehensivePerformanceTestSuite() = default;

    // 运行所有测试
    void runAllTests();
    
    // 运行指定类别的测试
    void runReadTests();
    void runWriteTests();  
    void runParsingTests();
    void runSharedFormulaTests();
    
    // 生成综合报告
    void generateComprehensiveReport();
    
    // 与历史数据对比
    void compareWithHistoricalData(const std::string& baseline_dir);
    
    // 性能回归检测
    bool detectPerformanceRegression(double threshold_percent = 10.0);
    
    // 导出性能趋势数据
    void exportPerformanceTrends(const std::string& trends_file);

private:
    void setupOutputDirectory();
    void generateSummaryReport();
    void generateDetailedReport();
    void generatePerformanceCharts();
};

/**
 * @brief 性能基准数据管理
 */
class PerformanceBenchmarkManager {
public:
    // 保存当前测试结果为基准
    static void saveAsBaseline(const std::vector<PerformanceResult>& results, 
                              const std::string& baseline_file);
    
    // 加载基准数据
    static std::vector<PerformanceResult> loadBaseline(const std::string& baseline_file);
    
    // 更新历史记录
    static void updateHistory(const std::vector<PerformanceResult>& results);
    
    // 获取性能趋势
    static std::map<std::string, std::vector<double>> getPerformanceTrends(const std::string& metric);
    
    // 性能警报
    static std::vector<std::string> checkPerformanceAlerts(
        const std::vector<PerformanceResult>& current_results,
        const std::vector<PerformanceResult>& baseline_results,
        double threshold_percent = 10.0);
};

} // namespace test
} // namespace fastexcel