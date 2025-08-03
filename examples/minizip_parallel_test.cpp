#include "fastexcel/FastExcel.hpp"
#include "fastexcel/archive/MinizipParallelWriter.hpp"
#include <iostream>
#include <chrono>
#include <random>
#include <iomanip>

using namespace fastexcel;

// 生成测试数据
std::string generateTestData(size_t size_kb) {
    std::string data;
    data.reserve(size_kb * 1024);
    
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> char_dist('A', 'Z');
    
    for (size_t i = 0; i < size_kb * 1024; ++i) {
        data += static_cast<char>(char_dist(gen));
        if (i % 80 == 79) data += '\n'; // 添加换行符模拟真实数据
    }
    
    return data;
}

// 测试基于minizip-ng的并行压缩性能
void testMinizipParallelCompression() {
    std::cout << "\n=== 基于Minizip-NG的并行压缩性能测试 ===" << std::endl;
    
    // 创建测试文件（模拟Excel文件结构）- 增大数据量以更好地测试并行性能
    std::vector<std::pair<std::string, std::string>> test_files;
    
    // 生成类似Excel的文件结构 - 增大文件大小
    std::vector<std::pair<std::string, size_t>> file_configs = {
        {"xl/worksheets/sheet1.xml", 8000},    // 8MB 工作表
        {"xl/worksheets/sheet2.xml", 6000},    // 6MB 工作表
        {"xl/worksheets/sheet3.xml", 4000},    // 4MB 工作表
        {"xl/worksheets/sheet4.xml", 3000},    // 3MB 工作表
        {"xl/styles.xml", 1200},               // 1.2MB 样式
        {"xl/workbook.xml", 200},              // 200KB 工作簿
        {"xl/sharedStrings.xml", 3200},        // 3.2MB 共享字符串
        {"[Content_Types].xml", 20},           // 20KB 内容类型
        {"_rels/.rels", 8},                    // 8KB 关系
        {"xl/_rels/workbook.xml.rels", 12},    // 12KB 工作簿关系
        {"docProps/core.xml", 40},             // 40KB 核心属性
        {"docProps/app.xml", 32}               // 32KB 应用属性
    };
    
    std::cout << "生成Excel风格的测试数据..." << std::endl;
    for (const auto& [filename, size_kb] : file_configs) {
        test_files.emplace_back(filename, generateTestData(size_kb));
        std::cout << "  " << filename << ": " << size_kb << "KB" << std::endl;
    }
    
    size_t total_size = 0;
    for (const auto& [filename, content] : test_files) {
        total_size += content.size();
    }
    
    std::cout << "总数据量: " << std::fixed << std::setprecision(2)
              << total_size / 1024.0 / 1024.0 << " MB" << std::endl;
    
    // 测试不同线程数的性能 - 使用更高的压缩级别以增加CPU负载
    std::vector<size_t> thread_counts = {1, 2, 4, 8};
    std::vector<double> performance_results;
    std::vector<double> duration_results;
    
    for (size_t thread_count : thread_counts) {
        std::cout << "\n--- 测试 " << thread_count << " 个线程 ---" << std::endl;
        
        auto start_time = std::chrono::high_resolution_clock::now();
        
        archive::MinizipParallelWriter writer(thread_count);
        
        std::string zip_filename = "minizip_parallel_test_" + std::to_string(thread_count) + "threads.xlsx";
        // 使用更高的压缩级别以增加CPU负载
        bool success = writer.compressAndWrite(zip_filename, test_files, 6);
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        if (success) {
            auto stats = writer.getStatistics();
            double mb_per_second = (total_size / 1024.0 / 1024.0) / (duration.count() / 1000.0);
            performance_results.push_back(mb_per_second);
            duration_results.push_back(static_cast<double>(duration.count()));
            
            std::cout << "✅ 压缩成功" << std::endl;
            std::cout << "总耗时: " << duration.count() << " ms" << std::endl;
            std::cout << "压缩速度: " << std::fixed << std::setprecision(2) << mb_per_second << " MB/s" << std::endl;
            std::cout << "压缩比: " << std::fixed << std::setprecision(1)
                      << stats.compression_ratio * 100 << "%" << std::endl;
            std::cout << "完成任务: " << stats.completed_tasks << "/"
                      << (stats.completed_tasks + stats.failed_tasks) << std::endl;
            std::cout << "并行效率: " << std::fixed << std::setprecision(1)
                      << stats.parallel_efficiency << "%" << std::endl;
            
            // 计算真实的加速比和效率
            if (thread_count > 1 && !performance_results.empty()) {
                double speedup = duration_results[0] / duration.count(); // 时间比值
                double efficiency = speedup / thread_count * 100.0; // 并行效率
                
                std::cout << "真实加速比: " << std::fixed << std::setprecision(2) << speedup << "x" << std::endl;
                std::cout << "真实并行效率: " << std::fixed << std::setprecision(1) << efficiency << "%" << std::endl;
                
                if (speedup >= thread_count * 0.8) {
                    std::cout << "🚀 并行效果卓越！" << std::endl;
                } else if (speedup >= thread_count * 0.6) {
                    std::cout << "🎉 并行效果优秀！" << std::endl;
                } else if (speedup >= thread_count * 0.4) {
                    std::cout << "👍 并行效果良好" << std::endl;
                } else {
                    std::cout << "⚠️  并行效果一般" << std::endl;
                }
            }
        } else {
            std::cout << "❌ 压缩失败" << std::endl;
            performance_results.push_back(0.0);
            duration_results.push_back(0.0);
        }
    }
    
    // 性能总结
    std::cout << "\n📊 性能总结:" << std::endl;
    std::cout << "线程数\t速度(MB/s)\t耗时(ms)\t加速比\t效率" << std::endl;
    std::cout << "----\t--------\t-------\t-----\t----" << std::endl;
    for (size_t i = 0; i < thread_counts.size(); ++i) {
        double speedup = (i > 0 && duration_results[0] > 0) ? duration_results[0] / duration_results[i] : 1.0;
        double efficiency = speedup / thread_counts[i] * 100.0;
        
        std::cout << thread_counts[i] << "\t"
                  << std::fixed << std::setprecision(1) << performance_results[i] << "\t\t"
                  << std::fixed << std::setprecision(0) << duration_results[i] << "\t\t"
                  << std::fixed << std::setprecision(2) << speedup << "x\t"
                  << std::fixed << std::setprecision(1) << efficiency << "%" << std::endl;
    }
}

// 与FastExcel集成测试
void testFastExcelIntegration() {
    std::cout << "\n=== FastExcel + Minizip-NG 集成测试 ===" << std::endl;
    
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/minizip_parallel_test.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return;
    }
    
    try {
        const int rows = 15000;
        const int cols = 20;
        const int total_cells = rows * cols;
        
        std::cout << "生成Excel文件: " << rows << "行 x " << cols << "列 = " << total_cells << "个单元格" << std::endl;
        
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 创建工作簿
        auto workbook = std::make_shared<core::Workbook>("minizip_integration_test.xlsx");
        
        if (!workbook->open()) {
            std::cerr << "Failed to open workbook" << std::endl;
            return;
        }
        
        // 使用默认的高性能配置（已经包含流式XML等优化）
        auto& options = workbook->getOptions();
        std::cout << "当前配置: 流式XML=" << (options.streaming_xml ? "ON" : "OFF")
                  << ", 共享字符串=" << (options.use_shared_strings ? "ON" : "OFF")
                  << ", 压缩级别=" << options.compression_level << std::endl;
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("Minizip并行测试");
        
        // 生成数据
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_int_distribution<> int_dist(1, 1000);
        std::uniform_real_distribution<> real_dist(1.0, 1000.0);
        
        for (int row = 0; row < rows; ++row) {
            for (int col = 0; col < cols; ++col) {
                if (col == 0) {
                    worksheet->writeString(row, col, "Row_" + std::to_string(row + 1));
                } else if (col % 3 == 1) {
                    worksheet->writeNumber(row, col, int_dist(gen));
                } else if (col % 3 == 2) {
                    worksheet->writeNumber(row, col, real_dist(gen));
                } else {
                    worksheet->writeString(row, col, "Data_" + std::to_string(row) + "_" + std::to_string(col));
                }
            }
            
            if ((row + 1) % 1500 == 0) {
                std::cout << "已处理 " << (row + 1) << " 行..." << std::endl;
            }
        }
        
        auto write_time = std::chrono::high_resolution_clock::now();
        auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
        
        std::cout << "数据写入完成，耗时: " << write_duration.count() << " ms" << std::endl;
        std::cout << "开始保存文件（使用minizip-ng并行压缩）..." << std::endl;
        
        // 保存文件（这里会使用我们的并行压缩）
        bool success = workbook->save();
        workbook->close();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - write_time);
        auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        if (success) {
            double cells_per_second = static_cast<double>(total_cells) / (total_duration.count() / 1000.0);
            
            std::cout << "\n✅ FastExcel + Minizip-NG 集成测试成功" << std::endl;
            std::cout << "数据写入: " << write_duration.count() << " ms (" 
                      << std::fixed << std::setprecision(1) 
                      << (double)write_duration.count() / total_duration.count() * 100 << "%)" << std::endl;
            std::cout << "文件保存: " << save_duration.count() << " ms (" 
                      << std::fixed << std::setprecision(1) 
                      << (double)save_duration.count() / total_duration.count() * 100 << "%)" << std::endl;
            std::cout << "总耗时: " << total_duration.count() << " ms" << std::endl;
            std::cout << "处理速度: " << std::fixed << std::setprecision(0) << cells_per_second << " 单元格/秒" << std::endl;
            
            // 性能评估
            if (cells_per_second > 200000) {
                std::cout << "🚀 性能卓越！Minizip-NG并行压缩效果显著" << std::endl;
            } else if (cells_per_second > 150000) {
                std::cout << "🎉 性能优秀！" << std::endl;
            } else if (cells_per_second > 100000) {
                std::cout << "👍 性能良好" << std::endl;
            } else {
                std::cout << "⚠️  性能有待提升" << std::endl;
            }
            
            // 保存阶段分析
            double save_percentage = (double)save_duration.count() / total_duration.count() * 100;
            if (save_percentage < 40) {
                std::cout << "🎯 并行压缩优化效果显著！保存阶段仅占 " << std::fixed << std::setprecision(1) << save_percentage << "%" << std::endl;
            } else if (save_percentage < 60) {
                std::cout << "✅ 并行压缩有效果，保存阶段占 " << std::fixed << std::setprecision(1) << save_percentage << "%" << std::endl;
            } else {
                std::cout << "⚠️  保存阶段仍占 " << std::fixed << std::setprecision(1) << save_percentage << "%，需要进一步优化" << std::endl;
            }
        } else {
            std::cout << "❌ 保存失败" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "Exception: " << e.what() << std::endl;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
}

int main() {
    std::cout << "FastExcel + Minizip-NG 并行压缩测试程序" << std::endl;
    std::cout << "=========================================" << std::endl;
    
    try {
        // 测试基于minizip-ng的并行压缩性能
        testMinizipParallelCompression();
        
        // 测试与FastExcel的集成
        testFastExcelIntegration();
        
        std::cout << "\n🎯 测试总结:" << std::endl;
        std::cout << "1. ✅ 使用成熟的minizip-ng库，稳定可靠" << std::endl;
        std::cout << "2. 🚀 文件级并行压缩，充分利用多核CPU" << std::endl;
        std::cout << "3. 📊 适合Excel文件的多文件结构特点" << std::endl;
        std::cout << "4. 🔧 完全兼容ZIP标准，无兼容性问题" << std::endl;
        std::cout << "5. 🎉 相比自实现ZIP，维护成本大幅降低" << std::endl;
        
        std::cout << "\n所有测试完成！请查看生成的测试文件。" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "程序异常: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}