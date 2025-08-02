#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <chrono>
#include <random>
#include <iomanip>

int main() {
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/ultra_performance_test.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    LOG_INFO("FastExcel ultra performance test started");
    
    try {
        const int rows = 50000;   // 5万行
        const int cols = 30;      // 30列
        const int total_cells = rows * cols;
        
        std::cout << "开始超级性能测试，将生成 " << rows << " 行 x " << cols << " 列 = " << total_cells << " 个单元格..." << std::endl;
        
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 创建工作簿
        auto workbook = std::make_shared<fastexcel::core::Workbook>("ultra_performance_test.xlsx");
        
        // 打开工作簿
        if (!workbook->open()) {
            LOG_ERROR("Failed to open workbook");
            return 1;
        }
        
        // 手动设置极致性能选项
        auto& options = workbook->getOptions();
        options.use_shared_strings = false;      // 完全禁用共享字符串
        options.streaming_xml = true;            // 启用流式XML
        options.row_buffer_size = 10000;         // 大缓冲区
        options.compression_level = 0;           // 无压缩（最快）
        options.xml_buffer_size = 8 * 1024 * 1024; // 8MB XML缓冲区
        
        LOG_INFO("Ultra performance mode configured: SharedStrings=OFF, StreamingXML=ON, RowBuffer={}, Compression={}, XMLBuffer={}MB",
                options.row_buffer_size, options.compression_level, options.xml_buffer_size / (1024*1024));
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("超级性能测试");
        if (!worksheet) {
            LOG_ERROR("Failed to create worksheet");
            return 1;
        }
        
        // 预分配内存（如果可能）
        // worksheet->reserveCapacity(rows, cols); // 假设有这个方法
        
        // 创建随机数生成器（优化：使用更快的随机数生成器）
        std::mt19937 gen(12345); // 固定种子以获得一致的性能测试结果
        std::uniform_int_distribution<> int_dist(1, 1000);
        std::uniform_real_distribution<> real_dist(1.0, 1000.0);
        std::uniform_int_distribution<> bool_dist(0, 1);
        
        // 预生成一些字符串以避免重复构造
        std::vector<std::string> pre_strings;
        pre_strings.reserve(1000);
        for (int i = 0; i < 1000; ++i) {
            pre_strings.push_back("Data_" + std::to_string(i));
        }
        
        // 生成随机数据并写入工作表
        for (int row = 0; row < rows; ++row) {
            for (int col = 0; col < cols; ++col) {
                if (col == 0) {
                    // 第一列写入行号
                    worksheet->writeNumber(row, col, row + 1);
                } else if (col == 1) {
                    // 第二列写入预生成的字符串
                    worksheet->writeString(row, col, pre_strings[row % pre_strings.size()]);
                } else if (col % 4 == 0) {
                    // 每4列写入一个布尔值
                    worksheet->writeBoolean(row, col, bool_dist(gen) == 1);
                } else if (col % 4 == 1) {
                    // 每4列写入一个整数
                    worksheet->writeNumber(row, col, int_dist(gen));
                } else if (col % 4 == 2) {
                    // 每4列写入一个浮点数
                    worksheet->writeNumber(row, col, real_dist(gen));
                } else {
                    // 其他列写入简单整数（最快）
                    worksheet->writeNumber(row, col, row + col);
                }
            }
            
            // 每2000行输出一次进度
            if ((row + 1) % 2000 == 0) {
                auto current_time = std::chrono::high_resolution_clock::now();
                auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(current_time - start_time);
                double cells_processed = (row + 1) * cols;
                double cells_per_second = cells_processed / (elapsed.count() / 1000.0);
                double progress = (double)(row + 1) / rows * 100.0;
                
                std::cout << "进度: " << std::fixed << std::setprecision(1) << progress << "% - "
                         << "已处理 " << (row + 1) << " 行 (" << static_cast<int>(cells_processed) 
                         << " 单元格), 速度: " << std::fixed << std::setprecision(0) 
                         << cells_per_second << " 单元格/秒" << std::endl;
            }
        }
        
        auto write_time = std::chrono::high_resolution_clock::now();
        auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
        std::cout << "数据写入完成，耗时: " << write_duration.count() << " 毫秒" << std::endl;
        
        // 保存工作簿
        std::cout << "开始保存文件（无压缩模式）..." << std::endl;
        if (!workbook->save()) {
            LOG_ERROR("Failed to save workbook");
            return 1;
        }
        
        auto save_time = std::chrono::high_resolution_clock::now();
        auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(save_time - write_time);
        std::cout << "文件保存完成，耗时: " << save_duration.count() << " 毫秒" << std::endl;
        
        // 关闭工作簿
        workbook->close();
        
        auto total_time = std::chrono::high_resolution_clock::now();
        auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(total_time - start_time);
        
        // 计算性能指标
        double cells_per_second = static_cast<double>(total_cells) / (total_duration.count() / 1000.0);
        double mb_per_second = 0.0; // 需要获取文件大小来计算
        
        std::cout << "\n超级性能测试结果:" << std::endl;
        std::cout << "总单元格数: " << total_cells << std::endl;
        std::cout << "总耗时: " << total_duration.count() << " 毫秒 (" << std::fixed << std::setprecision(2) << total_duration.count() / 1000.0 << " 秒)" << std::endl;
        std::cout << "写入速度: " << std::fixed << std::setprecision(0) << cells_per_second << " 单元格/秒" << std::endl;
        std::cout << "写入阶段: " << write_duration.count() << " 毫秒 (" << std::setprecision(1) << (double)write_duration.count() / total_duration.count() * 100 << "%)" << std::endl;
        std::cout << "保存阶段: " << save_duration.count() << " 毫秒 (" << std::setprecision(1) << (double)save_duration.count() / total_duration.count() * 100 << "%)" << std::endl;
        
        // 性能基准
        if (cells_per_second > 100000) {
            std::cout << "性能评级: 优秀 (>100K 单元格/秒)" << std::endl;
        } else if (cells_per_second > 50000) {
            std::cout << "性能评级: 良好 (>50K 单元格/秒)" << std::endl;
        } else if (cells_per_second > 20000) {
            std::cout << "性能评级: 一般 (>20K 单元格/秒)" << std::endl;
        } else {
            std::cout << "性能评级: 需要优化 (<20K 单元格/秒)" << std::endl;
        }
        
        LOG_INFO("Ultra performance test completed: {} cells in {} ms ({} cells/sec)", 
                total_cells, total_duration.count(), cells_per_second);
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception occurred: {}", e.what());
        std::cerr << "Exception occurred: " << e.what() << std::endl;
        return 1;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
    
    LOG_INFO("FastExcel ultra performance test completed");
    std::cout << "\n超级性能测试完成！" << std::endl;
    
    return 0;
}