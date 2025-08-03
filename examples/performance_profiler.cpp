#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <chrono>
#include <random>
#include <iomanip>
#include <fstream>
#include <thread>

// 性能监控器
class PerformanceProfiler {
private:
    struct TimingData {
        std::chrono::high_resolution_clock::time_point start_time;
        std::chrono::high_resolution_clock::time_point end_time;
        std::string operation_name;
        bool completed = false;
    };
    
    std::map<std::string, TimingData> timings_;
    std::map<std::string, std::vector<double>> operation_history_;
    
public:
    void startTimer(const std::string& operation) {
        timings_[operation].start_time = std::chrono::high_resolution_clock::now();
        timings_[operation].operation_name = operation;
        timings_[operation].completed = false;
    }
    
    void endTimer(const std::string& operation) {
        auto it = timings_.find(operation);
        if (it != timings_.end()) {
            it->second.end_time = std::chrono::high_resolution_clock::now();
            it->second.completed = true;
            
            auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(
                it->second.end_time - it->second.start_time).count();
            operation_history_[operation].push_back(duration);
        }
    }
    
    double getOperationTime(const std::string& operation) const {
        auto it = timings_.find(operation);
        if (it != timings_.end() && it->second.completed) {
            return std::chrono::duration_cast<std::chrono::milliseconds>(
                it->second.end_time - it->second.start_time).count();
        }
        return 0.0;
    }
    
    void printReport() const {
        std::cout << "\n=== 性能分析报告 ===" << std::endl;
        std::cout << std::left << std::setw(25) << "操作" 
                  << std::setw(15) << "耗时(ms)" 
                  << std::setw(15) << "占比(%)" 
                  << "建议" << std::endl;
        std::cout << std::string(80, '-') << std::endl;
        
        double total_time = 0;
        for (const auto& [name, timing] : timings_) {
            if (timing.completed) {
                total_time += getOperationTime(name);
            }
        }
        
        for (const auto& [name, timing] : timings_) {
            if (timing.completed) {
                double op_time = getOperationTime(name);
                double percentage = (op_time / total_time) * 100;
                
                std::cout << std::left << std::setw(25) << name
                          << std::setw(15) << std::fixed << std::setprecision(2) << op_time
                          << std::setw(15) << std::fixed << std::setprecision(1) << percentage;
                
                // 提供优化建议
                if (percentage > 50) {
                    std::cout << "🔴 主要瓶颈，优先优化";
                } else if (percentage > 20) {
                    std::cout << "🟡 次要瓶颈，可以优化";
                } else if (percentage > 5) {
                    std::cout << "🟢 性能良好";
                } else {
                    std::cout << "✅ 已优化";
                }
                std::cout << std::endl;
            }
        }
    }
    
    void saveReport(const std::string& filename) const {
        std::ofstream file(filename);
        file << "Operation,Time(ms),Percentage\n";
        
        double total_time = 0;
        for (const auto& [name, timing] : timings_) {
            if (timing.completed) {
                total_time += getOperationTime(name);
            }
        }
        
        for (const auto& [name, timing] : timings_) {
            if (timing.completed) {
                double op_time = getOperationTime(name);
                double percentage = (op_time / total_time) * 100;
                file << name << "," << op_time << "," << percentage << "\n";
            }
        }
        file.close();
        std::cout << "性能报告已保存到: " << filename << std::endl;
    }
};

// 内存使用监控
class MemoryMonitor {
private:
    size_t peak_memory_ = 0;
    
public:
    size_t getCurrentMemoryUsage() {
#ifdef _WIN32
        PROCESS_MEMORY_COUNTERS pmc;
        if (GetProcessMemoryInfo(GetCurrentProcess(), &pmc, sizeof(pmc))) {
            return pmc.WorkingSetSize;
        }
#else
        std::ifstream status("/proc/self/status");
        std::string line;
        while (std::getline(status, line)) {
            if (line.substr(0, 6) == "VmRSS:") {
                std::istringstream iss(line);
                std::string key, value, unit;
                iss >> key >> value >> unit;
                return std::stoul(value) * 1024; // Convert KB to bytes
            }
        }
#endif
        return 0;
    }
    
    void updatePeakMemory() {
        size_t current = getCurrentMemoryUsage();
        if (current > peak_memory_) {
            peak_memory_ = current;
        }
    }
    
    void printMemoryReport() const {
        std::cout << "\n=== 内存使用报告 ===" << std::endl;
        std::cout << "当前内存使用: " << (getCurrentMemoryUsage() / 1024.0 / 1024.0) << " MB" << std::endl;
        std::cout << "峰值内存使用: " << (peak_memory_ / 1024.0 / 1024.0) << " MB" << std::endl;
        
        if (peak_memory_ > 1024 * 1024 * 1024) { // > 1GB
            std::cout << "🔴 内存使用较高，建议优化" << std::endl;
        } else if (peak_memory_ > 512 * 1024 * 1024) { // > 512MB
            std::cout << "🟡 内存使用中等" << std::endl;
        } else {
            std::cout << "🟢 内存使用良好" << std::endl;
        }
    }
};

int main() {
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/performance_profiler.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    PerformanceProfiler profiler;
    MemoryMonitor memory_monitor;
    
    std::cout << "FastExcel 性能分析器" << std::endl;
    std::cout << "===================" << std::endl;
    
    try {
        // 测试不同数据量的性能
        std::vector<std::pair<int, int>> test_cases = {
            {1000, 10},    // 1万单元格
            {5000, 20},    // 10万单元格
            {10000, 30},   // 30万单元格
            {20000, 25}    // 50万单元格
        };
        
        for (const auto& [rows, cols] : test_cases) {
            int total_cells = rows * cols;
            std::cout << "\n测试数据量: " << rows << "行 x " << cols << "列 = " << total_cells << "个单元格" << std::endl;
            
            profiler.startTimer("总体耗时");
            memory_monitor.updatePeakMemory();
            
            // 工作簿创建和配置
            profiler.startTimer("工作簿创建");
            auto workbook = std::make_shared<fastexcel::core::Workbook>("profiler_test_" + std::to_string(total_cells) + ".xlsx");
            if (!workbook->open()) {
                std::cerr << "Failed to open workbook" << std::endl;
                continue;
            }
            
            // 测试不同配置的性能
            auto& options = workbook->getOptions();
            options.compression_level = 0; // 无压缩以突出其他瓶颈
            
            auto worksheet = workbook->addWorksheet("性能测试");
            profiler.endTimer("工作簿创建");
            memory_monitor.updatePeakMemory();
            
            // 数据生成
            profiler.startTimer("数据生成");
            std::random_device rd;
            std::mt19937 gen(rd());
            std::uniform_int_distribution<> int_dist(1, 1000);
            std::uniform_real_distribution<> real_dist(1.0, 1000.0);
            
            std::vector<std::vector<std::string>> data(rows, std::vector<std::string>(cols));
            for (int row = 0; row < rows; ++row) {
                for (int col = 0; col < cols; ++col) {
                    if (col % 3 == 0) {
                        data[row][col] = "Text_" + std::to_string(row) + "_" + std::to_string(col);
                    } else if (col % 3 == 1) {
                        data[row][col] = std::to_string(int_dist(gen));
                    } else {
                        data[row][col] = std::to_string(real_dist(gen));
                    }
                }
            }
            profiler.endTimer("数据生成");
            memory_monitor.updatePeakMemory();
            
            // 数据写入
            profiler.startTimer("数据写入");
            for (int row = 0; row < rows; ++row) {
                for (int col = 0; col < cols; ++col) {
                    if (col % 3 == 0) {
                        worksheet->writeString(row, col, data[row][col]);
                    } else {
                        worksheet->writeNumber(row, col, std::stod(data[row][col]));
                    }
                }
                
                // 每1000行更新内存监控
                if (row % 1000 == 0) {
                    memory_monitor.updatePeakMemory();
                }
            }
            profiler.endTimer("数据写入");
            memory_monitor.updatePeakMemory();
            
            // XML生成和保存
            profiler.startTimer("文件保存");
            bool success = workbook->save();
            profiler.endTimer("文件保存");
            memory_monitor.updatePeakMemory();
            
            workbook->close();
            profiler.endTimer("总体耗时");
            
            if (success) {
                double total_time = profiler.getOperationTime("总体耗时");
                double cells_per_second = total_cells / (total_time / 1000.0);
                
                std::cout << "✅ 测试完成" << std::endl;
                std::cout << "总耗时: " << total_time << " ms" << std::endl;
                std::cout << "处理速度: " << std::fixed << std::setprecision(0) << cells_per_second << " 单元格/秒" << std::endl;
                
                // 分析各阶段耗时
                profiler.printReport();
                memory_monitor.printMemoryReport();
                
                // 保存详细报告
                profiler.saveReport("performance_report_" + std::to_string(total_cells) + ".csv");
                
            } else {
                std::cout << "❌ 测试失败" << std::endl;
            }
            
            std::cout << std::string(80, '=') << std::endl;
        }
        
        // 生成优化建议
        std::cout << "\n🎯 优化建议:" << std::endl;
        std::cout << "1. 如果'文件保存'占比>60%，建议实现并行压缩" << std::endl;
        std::cout << "2. 如果'数据写入'占比>30%，建议实现批量写入" << std::endl;
        std::cout << "3. 如果内存使用>1GB，建议优化内存管理" << std::endl;
        std::cout << "4. 如果'数据生成'占比>10%，建议优化数据结构" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Exception occurred: " << e.what() << std::endl;
        return 1;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
    
    std::cout << "\n性能分析完成！请查看生成的CSV报告文件。" << std::endl;
    
    return 0;
}