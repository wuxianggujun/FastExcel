/**
 * @file large_file_reading_demo.cpp
 * @brief 大文件读取性能演示 (60MB+)
 * 
 * 展示FastExcel读取超大Excel文件的能力：
 * - 内存池优化读取
 * - 流式数据处理
 * - 内存使用监控
 * - 性能指标统计
 * - 大文件数据分析
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/ReadOnlyWorkbook.hpp"
#include "fastexcel/core/ReadOnlyWorksheet.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/memory/PoolManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <chrono>
#include <iomanip>
#include <fstream>
#include <memory>
#include <thread>
#include <fmt/format.h>

#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#else
#include <sys/resource.h>
#include <unistd.h>
#endif

using namespace fastexcel;
using namespace fastexcel::core;

/**
 * @brief 内存使用情况监控器 - 优化版
 */
class MemoryMonitor {
private:
    size_t initial_memory_kb_;
    
    // 转换为合适的单位显示
    std::string formatMemorySize(size_t kb) const {
        if (kb >= 1024 * 1024) {
            double gb = static_cast<double>(kb) / (1024 * 1024);
            return fmt::format("{:.2f} GB", gb);
        } else if (kb >= 1024) {
            double mb = static_cast<double>(kb) / 1024;
            return fmt::format("{:.1f} MB", mb);
        } else {
            return fmt::format("{} KB", kb);
        }
    }
    
public:
    MemoryMonitor() {
        initial_memory_kb_ = getCurrentMemoryUsage();
        std::cout << "🔍 初始内存: " << formatMemorySize(initial_memory_kb_) << "\n";
    }
    
    ~MemoryMonitor() {
        size_t final_memory = getCurrentMemoryUsage();
        size_t growth = final_memory - initial_memory_kb_;
        std::cout << "🔍 最终内存: " << formatMemorySize(final_memory) << "\n";
        std::cout << "🔍 内存增长: " << formatMemorySize(growth) << "\n";
    }
    
    size_t getCurrentMemoryUsage() const {
#ifdef _WIN32
        PROCESS_MEMORY_COUNTERS pmc;
        if (GetProcessMemoryInfo(GetCurrentProcess(), &pmc, sizeof(pmc))) {
            return pmc.WorkingSetSize / 1024; // 转换为KB
        }
        return 0;
#else
        struct rusage usage;
        getrusage(RUSAGE_SELF, &usage);
        return usage.ru_maxrss; // Linux下已经是KB
#endif
    }
    
    void printCurrentUsage(const std::string& stage) const {
        size_t current = getCurrentMemoryUsage();
        size_t growth = current - initial_memory_kb_;
        std::cout << "🔍 [" << stage << "] 内存: " << formatMemorySize(current)
                  << " (增长: " << formatMemorySize(growth) << ")\n";
    }
    
    // 简化版本，不打印详细信息
    void checkMemoryQuiet() const {
        // 只返回当前内存使用量，不打印
    }
};

/**
 * @brief 性能计时器
 */
class PerformanceTimer {
private:
    std::chrono::high_resolution_clock::time_point start_time_;
    std::string operation_name_;
    
public:
    PerformanceTimer(const std::string& name) 
        : operation_name_(name), start_time_(std::chrono::high_resolution_clock::now()) {
        std::cout << "⏱️  开始 " << operation_name_ << "...\n";
    }
    
    ~PerformanceTimer() {
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time_);
        std::cout << "✅ " << operation_name_ << " 完成，耗时: " << duration.count() << " ms\n";
    }
    
    long long getElapsedMs() const {
        auto current = std::chrono::high_resolution_clock::now();
        return std::chrono::duration_cast<std::chrono::milliseconds>(current - start_time_).count();
    }
};


/**
 * @brief 预热内存池
 */
void prewarmMemoryPools() {
    std::cout << "\n🔥 预热内存池...\n";
    
    auto& pool_manager = fastexcel::memory::PoolManager::getInstance();
    
    // 预热常用类型的内存池
    pool_manager.prewarmPools<
        std::string,
        double,
        int,
        bool,
        Cell,
        Worksheet
    >();
    
    std::cout << "✅ 内存池预热完成，当前池数量: " << pool_manager.getPoolCount() << "\n";
}

/**
 * @brief 读取大文件并分析
 */
void readLargeFile(const std::string& filepath) {
    MemoryMonitor memory_monitor;
    
    std::cout << "\n📖 开始读取大文件: " << filepath << "\n";
    std::cout << "========================================\n";
    
    // 步骤2: 预热内存池
    prewarmMemoryPools();
    
    std::unique_ptr<ReadOnlyWorkbook> workbook;
    
    // 步骤3: 打开文件
    {
        PerformanceTimer timer("文件打开");
        workbook = fastexcel::openReadOnly(filepath);
        
        if (!workbook) {
            std::cout << "❌ 无法打开Excel文件: " << filepath << "\n";
            return;
        }
        
        // 只在打开完成后检查一次内存
        memory_monitor.printCurrentUsage("文件打开后");
    }
    
    // 步骤4: 获取基本信息
    {
        PerformanceTimer timer("基本信息获取");
        
        auto sheet_names = workbook->getSheetNames();
        std::cout << "\n📊 工作簿信息:\n";
        std::cout << "  工作表数量: " << sheet_names.size() << "\n";
        
        // 显示工作表名称
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            std::cout << "    " << (i + 1) << ". " << sheet_names[i] << "\n";
        }
    }
    
    // 步骤5: 分析每个工作表
    for (size_t sheet_idx = 0; sheet_idx < workbook->getSheetCount(); ++sheet_idx) {
        std::cout << "\n📋 分析工作表 " << (sheet_idx + 1) << ":\n";
        std::cout << "----------------------------------------\n";
        
        auto worksheet = workbook->getSheet(sheet_idx);
        if (!worksheet) {
            std::cout << "❌ 无法加载工作表 " << (sheet_idx + 1) << "\n";
            continue;
        }
        
        std::string sheet_name = worksheet->getName();
        std::cout << "📋 工作表名称: " << sheet_name << "\n";
        
        // 获取使用范围
        auto [max_row, max_col] = worksheet->getUsedRange();
        int total_rows = max_row + 1;
        int total_cols = max_col + 1;
        
        std::cout << "📐 数据范围: " << total_rows << " 行 × " << total_cols << " 列\n";
        
        if (total_rows == 0 || total_cols == 0) {
            std::cout << "⚠️  空工作表，跳过分析\n";
            continue;
        }
        
        // 获取工作表统计信息（使用ReadOnlyWorksheet的优化统计方法）
        {
            PerformanceTimer timer("列式数据分析 - " + sheet_name);
            
            auto stats = worksheet->getStats();
            
            std::cout << "\n📈 数据统计 (" << sheet_name << "):\n";
            std::cout << "  总数据点数: " << stats.total_data_points << "\n";
            std::cout << "  内存使用: " << stats.memory_usage / 1024 << " KB\n";
            std::cout << "  数字列数: " << stats.number_columns << "\n";
            std::cout << "  字符串列数: " << stats.string_columns << "\n";
            std::cout << "  布尔列数: " << stats.boolean_columns << "\n";
            std::cout << "  错误/文本列数: " << stats.error_columns << "\n";
            
            if (stats.total_data_points > 0) {
                std::cout << "  数据类型分布:\n";
                double number_pct = (double)stats.number_columns / (stats.number_columns + stats.string_columns + stats.boolean_columns + stats.error_columns) * 100;
                double string_pct = (double)stats.string_columns / (stats.number_columns + stats.string_columns + stats.boolean_columns + stats.error_columns) * 100;
                std::cout << "    数字列: " << std::fixed << std::setprecision(1) << number_pct << "%\n";
                std::cout << "    字符串列: " << std::fixed << std::setprecision(1) << string_pct << "%\n";
            }
        }
        
        // 显示数据样例（使用列式存储API）
        std::cout << "\n📝 数据样例 (前5行×5列):\n";
        std::cout << std::setw(8) << "行\\列";
        for (int col = 0; col < std::min(5, total_cols); ++col) {
            std::cout << std::setw(15) << ("列" + std::to_string(col + 1));
        }
        std::cout << "\n";
        
        // 获取前5行的数据范围
        int sample_rows = std::min(5, total_rows);
        int sample_cols = std::min(5, total_cols);
        
        auto sample_data = worksheet->getRowRangeData(0, sample_rows - 1);
        
        for (int row = 0; row < sample_rows; ++row) {
            std::cout << std::setw(8) << ("行" + std::to_string(row + 1));
            
            for (int col = 0; col < sample_cols; ++col) {
                std::string cell_value = "(空)";
                
                auto row_it = sample_data.find(row);
                if (row_it != sample_data.end()) {
                    auto col_it = row_it->second.find(col);
                    if (col_it != row_it->second.end()) {
                        // 处理CellValue变体类型
                        std::visit([&cell_value](const auto& value) {
                            using T = std::decay_t<decltype(value)>;
                            if constexpr (std::is_same_v<T, double>) {
                                std::ostringstream oss;
                                oss << std::fixed << std::setprecision(2) << value;
                                cell_value = oss.str();
                            } else if constexpr (std::is_same_v<T, uint32_t>) {
                                cell_value = std::to_string(value);
                            } else if constexpr (std::is_same_v<T, bool>) {
                                cell_value = value ? "TRUE" : "FALSE";
                            } else if constexpr (std::is_same_v<T, std::string>) {
                                std::string str_value = value;
                                if (str_value.length() > 12) {
                                    str_value = str_value.substr(0, 9) + "...";
                                }
                                cell_value = "\"" + str_value + "\"";
                            }
                        }, col_it->second);
                    }
                }
                
                std::cout << std::setw(15) << cell_value;
            }
            std::cout << "\n";
        }
    }
    
    // 步骤6: 工作簿整体统计
    {
        PerformanceTimer timer("整体统计计算");
        
        auto stats = workbook->getStats();
        std::cout << "\n📊 工作簿整体统计:\n";
        std::cout << "  总工作表数: " << stats.sheet_count << "\n";
        std::cout << "  总数据点数: " << stats.total_data_points << "\n";
        std::cout << "  内存使用: " << stats.total_memory_usage / 1024 / 1024 << " MB\n";
        std::cout << "  共享字符串数: " << stats.sst_string_count << "\n";
        std::cout << "  列式存储优化: " << (stats.columnar_optimized ? "启用" : "禁用") << "\n";
    }
    
    // 步骤7: 内存池统计
    {
        auto& pool_manager = fastexcel::memory::PoolManager::getInstance();
        std::cout << "\n🏊 内存池统计:\n";
        std::cout << "  活跃内存池数量: " << pool_manager.getPoolCount() << "\n";
        
        // 尝试收缩内存池释放多余内存
        pool_manager.shrinkAll();
        std::cout << "  内存池收缩完成\n";
    }
    
    // 步骤8: 关闭文件
    {
        PerformanceTimer timer("文件关闭");
        // ReadOnlyWorkbook 会在析构时自动清理，无需显式关闭
        workbook.reset();
    }
    
    std::cout << "\n🎉 大文件读取分析完成!\n";
}

/**
 * @brief 性能基准测试
 */
void performBenchmark(const std::string& filepath) {
    std::cout << "\n🏁 性能基准测试\n";
    std::cout << "========================================\n";
    
    const int NUM_RUNS = 3; // 运行3次取平均值
    std::vector<long long> open_times;
    std::vector<size_t> memory_usage;
    
    for (int run = 0; run < NUM_RUNS; ++run) {
        std::cout << "\n🔄 第 " << (run + 1) << " 次测试...\n";
        
        MemoryMonitor monitor;
        PerformanceTimer timer("完整读取测试 #" + std::to_string(run + 1));
        
        auto workbook = fastexcel::openReadOnly(filepath);
        if (!workbook) {
            std::cout << "❌ 文件打开失败\n";
            continue;
        }
        
        // 快速遍历所有工作表
        for (size_t i = 0; i < workbook->getSheetCount(); ++i) {
            auto ws = workbook->getSheet(i);
            if (ws) {
                auto [rows, cols] = ws->getUsedRange();
                // 只是获取范围，不实际读取数据
            }
        }
        
        auto stats = workbook->getStats();
        open_times.push_back(timer.getElapsedMs());
        memory_usage.push_back(monitor.getCurrentMemoryUsage());
        
        // ReadOnlyWorkbook 会在析构时自动清理
    }
    
    // 计算平均值
    if (!open_times.empty()) {
        long long avg_time = 0;
        size_t avg_memory = 0;
        
        for (size_t i = 0; i < open_times.size(); ++i) {
            avg_time += open_times[i];
            avg_memory += memory_usage[i];
        }
        
        avg_time /= open_times.size();
        avg_memory /= memory_usage.size();
        
        std::cout << "\n📊 基准测试结果 (平均值):\n";
        std::cout << "  平均打开时间: " << avg_time << " ms\n";
        std::cout << "  平均内存使用: ";
        if (avg_memory >= 1024 * 1024) {
            double gb = static_cast<double>(avg_memory) / (1024 * 1024);
            std::cout << std::fixed << std::setprecision(2) << gb << " GB\n";
        } else {
            double mb = static_cast<double>(avg_memory) / 1024;
            std::cout << std::fixed << std::setprecision(1) << mb << " MB\n";
        }
        
        // 性能评级
        if (avg_time < 5000) {
            std::cout << "  性能评级: 🟢 优秀 (< 5秒)\n";
        } else if (avg_time < 15000) {
            std::cout << "  性能评级: 🟡 良好 (5-15秒)\n";
        } else {
            std::cout << "  性能评级: 🔴 需要优化 (> 15秒)\n";
        }
        
        if (avg_memory / 1024 < 1000) {
            std::cout << "  内存效率: 🟢 优秀 (< 1GB)\n";
        } else if (avg_memory / 1024 < 2000) {
            std::cout << "  内存效率: 🟡 良好 (1-2GB)\n";
        } else {
            std::cout << "  内存效率: 🔴 需要优化 (> 2GB)\n";
        }
    }
}

int main(int argc, char* argv[]) {
    std::cout << "🚀 FastExcel 大文件读取性能演示\n";
    std::cout << "=================================\n";
    std::cout << "本演示将测试读取60MB+大型Excel文件的性能\n\n";
    
    // 获取文件路径
    std::string filepath = "C:\\Users\\wuxianggujun\\CodeSpace\\CMakeProjects\\FastExcel\\test_xlsx\\合并去年和今年的数据.xlsx";
    
    try {
        // 初始化FastExcel
        if (!fastexcel::initialize("logs/large_file_demo.log", true)) {
            std::cout << "❌ FastExcel初始化失败\n";
            return 1;
        }
        
        std::cout << "✅ FastExcel初始化成功\n";
        
        // 主要读取测试
        readLargeFile(filepath);
        
        // 性能基准测试
        performBenchmark(filepath);
        
        std::cout << "\n🎯 演示完成! 检查日志文件: logs/large_file_demo.log\n";
        
        // 清理
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cout << "❌ 程序异常: " << e.what() << "\n";
        return 1;
    } catch (...) {
        std::cout << "❌ 未知异常\n";
        return 1;
    }
    
    return 0;
}
