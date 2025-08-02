#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <chrono>
#include <random>

int main() {
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/performance_test.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    LOG_INFO("FastExcel performance test started");
    
    try {
        const int rows = 100000;  // 1000行
        const int cols = 100;   // 20列
        const int total_cells = rows * cols;
        
        std::cout << "开始性能测试，将生成 " << rows << " 行 x " << cols << " 列 = " << total_cells << " 个单元格..." << std::endl;
        
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 创建工作簿
        auto workbook = std::make_shared<fastexcel::core::Workbook>("performance_test.xlsx");
        
        // 打开工作簿
        if (!workbook->open()) {
            LOG_ERROR("Failed to open workbook");
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("性能测试");
        if (!worksheet) {
            LOG_ERROR("Failed to create worksheet");
            return 1;
        }
        
        // 创建随机数生成器
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_int_distribution<> int_dist(1, 1000);
        std::uniform_real_distribution<> real_dist(1.0, 1000.0);
        std::uniform_int_distribution<> bool_dist(0, 1);
        
        // 生成随机数据并写入工作表
        for (int row = 0; row < rows; ++row) {
            for (int col = 0; col < cols; ++col) {
                if (col == 0) {
                    // 第一列写入行号
                    worksheet->writeNumber(row, col, row + 1);
                } else if (col == 1) {
                    // 第二列写入字符串
                    worksheet->writeString(row, col, "数据_" + std::to_string(row + 1));
                } else if (col % 3 == 0) {
                    // 每3列写入一个布尔值
                    worksheet->writeBoolean(row, col, bool_dist(gen) == 1);
                } else if (col % 3 == 1) {
                    // 每3列写入一个整数
                    worksheet->writeNumber(row, col, int_dist(gen));
                } else {
                    // 每3列写入一个浮点数
                    worksheet->writeNumber(row, col, real_dist(gen));
                }
            }
            
            // 每100行输出一次进度
            if ((row + 1) % 100 == 0) {
                std::cout << "已处理 " << (row + 1) << " 行..." << std::endl;
            }
        }
        
        auto write_time = std::chrono::high_resolution_clock::now();
        auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
        std::cout << "数据写入完成，耗时: " << write_duration.count() << " 毫秒" << std::endl;
        
        // 保存工作簿
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
        std::cout << "总耗时: " << total_duration.count() << " 毫秒" << std::endl;
        
        // 计算性能指标
        double cells_per_second = static_cast<double>(total_cells) / (total_duration.count() / 1000.0);
        double file_size_mb = 0.0; // 实际实现中需要获取文件大小
        
        std::cout << "\n性能测试结果:" << std::endl;
        std::cout << "总单元格数: " << total_cells << std::endl;
        std::cout << "总耗时: " << total_duration.count() << " 毫秒" << std::endl;
        std::cout << "写入速度: " << std::fixed << std::setprecision(2) << cells_per_second << " 单元格/秒" << std::endl;
        
        LOG_INFO("Performance test completed: {} cells in {} ms ({} cells/sec)", 
                total_cells, total_duration.count(), cells_per_second);
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception occurred: {}", e.what());
        std::cerr << "Exception occurred: " << e.what() << std::endl;
        return 1;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
    
    LOG_INFO("FastExcel performance test completed");
    std::cout << "\n性能测试完成，请查看生成的Excel文件和日志文件。" << std::endl;
    
    return 0;
}
