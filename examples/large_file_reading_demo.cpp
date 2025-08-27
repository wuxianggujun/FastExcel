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
    
    std::unique_ptr<Workbook> workbook;
    
    // 步骤3: 打开文件
    {
        PerformanceTimer timer("文件打开");
        workbook = Workbook::openReadOnly(Path(filepath));
        
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
        
        // 获取文档属性
        const auto& doc_props = workbook->getDocumentProperties();
        if (!doc_props.author.empty()) {
            std::cout << "  作者: " << doc_props.author << "\n";
        }
        if (!doc_props.company.empty()) {
            std::cout << "  公司: " << doc_props.company << "\n";
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
        
        // 流式读取大文件数据
        {
            PerformanceTimer timer("流式数据分析 - " + sheet_name);
            
            // 统计信息
            size_t empty_cells = 0;
            size_t string_cells = 0;
            size_t number_cells = 0;
            size_t boolean_cells = 0;
            size_t formula_cells = 0;
            size_t processed_cells = 0;
            
            const int SAMPLE_BATCH_SIZE = 1000; // 每次处理1000行
            const int PROGRESS_INTERVAL = 15000; // 每15000行显示进度（减少频率）
            
            for (int batch_start = 0; batch_start < total_rows; batch_start += SAMPLE_BATCH_SIZE) {
                int batch_end = std::min(batch_start + SAMPLE_BATCH_SIZE, total_rows);
                
                // 处理当前批次
                for (int row = batch_start; row < batch_end; ++row) {
                    for (int col = 0; col < total_cols; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            const auto& cell = worksheet->getCell(row, col);
                            processed_cells++;
                            
                            switch (cell.getType()) {
                                case CellType::String:
                                    string_cells++;
                                    break;
                                case CellType::Number:
                                    number_cells++;
                                    break;
                                case CellType::Boolean:
                                    boolean_cells++;
                                    break;
                                case CellType::Formula:
                                    formula_cells++;
                                    break;
                                default:
                                    empty_cells++;
                                    break;
                            }
                        } else {
                            empty_cells++;
                        }
                    }
                    
                    // 显示进度（每15000行）
                    if ((row + 1) % PROGRESS_INTERVAL == 0) {
                        double progress = static_cast<double>(row + 1) / total_rows * 100.0;
                        std::cout << "  🔄 进度: " << std::fixed << std::setprecision(1) 
                                  << progress << "% (" << (row + 1) << "/" << total_rows << " 行)\n";
                        
                        // 减少内存监控频率，仅在关键进度点检查
                        if (progress >= 50.0 && progress < 55.0) {
                            memory_monitor.printCurrentUsage("数据读取进度 50%");
                        }
                    }
                }
                
                // 给系统一点喘息时间，避免过度占用CPU
                if (batch_start > 0 && batch_start % (SAMPLE_BATCH_SIZE * 10) == 0) {
                    std::this_thread::sleep_for(std::chrono::milliseconds(10));
                }
            }
            
            // 输出统计结果
            std::cout << "\n📈 数据统计 (" << sheet_name << "):\n";
            std::cout << "  总单元格数: " << (total_rows * total_cols) << "\n";
            std::cout << "  已处理单元格: " << processed_cells << "\n";
            std::cout << "  字符串单元格: " << string_cells << " (" 
                      << std::fixed << std::setprecision(1) 
                      << (double)string_cells/processed_cells*100 << "%)\n";
            std::cout << "  数字单元格: " << number_cells << " (" 
                      << std::fixed << std::setprecision(1) 
                      << (double)number_cells/processed_cells*100 << "%)\n";
            std::cout << "  布尔单元格: " << boolean_cells << "\n";
            std::cout << "  公式单元格: " << formula_cells << "\n";
            std::cout << "  空单元格: " << empty_cells << "\n";
        }
        
        // 显示数据样例（前5行×5列）
        std::cout << "\n📝 数据样例 (前5行×5列):\n";
        std::cout << std::setw(8) << "行\\列";
        for (int col = 0; col < std::min(5, total_cols); ++col) {
            std::cout << std::setw(15) << ("列" + std::to_string(col + 1));
        }
        std::cout << "\n";
        
        for (int row = 0; row < std::min(5, total_rows); ++row) {
            std::cout << std::setw(8) << ("行" + std::to_string(row + 1));
            
            for (int col = 0; col < std::min(5, total_cols); ++col) {
                std::string cell_value = "(空)";
                
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    
                    try {
                        switch (cell.getType()) {
                            case CellType::String: {
                                std::string value = cell.getValue<std::string>();
                                if (value.length() > 12) {
                                    value = value.substr(0, 9) + "...";
                                }
                                cell_value = "\"" + value + "\"";
                                break;
                            }
                            case CellType::Number: {
                                double value = cell.getValue<double>();
                                std::ostringstream oss;
                                oss << std::fixed << std::setprecision(2) << value;
                                cell_value = oss.str();
                                break;
                            }
                            case CellType::Boolean:
                                cell_value = cell.getValue<bool>() ? "TRUE" : "FALSE";
                                break;
                            case CellType::Formula:
                                cell_value = "=" + cell.getFormula();
                                if (cell_value.length() > 12) {
                                    cell_value = cell_value.substr(0, 9) + "...";
                                }
                                break;
                            default:
                                cell_value = "(未知)";
                                break;
                        }
                    } catch (const std::exception& e) {
                        cell_value = "(错误)";
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
        
        auto stats = workbook->getStatistics();
        std::cout << "\n📊 工作簿整体统计:\n";
        std::cout << "  总工作表数: " << stats.total_worksheets << "\n";
        std::cout << "  总单元格数: " << stats.total_cells << "\n";
        std::cout << "  总格式数: " << stats.total_formats << "\n";
        std::cout << "  内存使用: " << stats.memory_usage / 1024 / 1024 << " MB\n";
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
        workbook->close();
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
        
        auto workbook = Workbook::openReadOnly(Path(filepath));
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
        
        auto stats = workbook->getStatistics();
        open_times.push_back(timer.getElapsedMs());
        memory_usage.push_back(monitor.getCurrentMemoryUsage());
        
        workbook->close();
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
