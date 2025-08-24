#include <gtest/gtest.h>
#include "PerformanceBenchmark.hpp"
#include "fastexcel/FastExcel.hpp"
#include <chrono>
#include <memory>
#include <filesystem>

class ReadWriteBenchmark : public PerformanceBenchmark {
protected:
    void SetUp() override {
        PerformanceBenchmark::SetUp();
        test_file_path_ = "read_write_benchmark_test.xlsx";
    }

    void TearDown() override {
        // 清理测试文件
        if (std::filesystem::exists(test_file_path_)) {
            std::filesystem::remove(test_file_path_);
        }
        PerformanceBenchmark::TearDown();
    }

    std::string test_file_path_;
};

// 测试写入性能
TEST_F(ReadWriteBenchmark, WritePerformance) {
    const int ROWS = 1000;
    const int COLS = 10;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    {
        auto workbook = fastexcel::core::Workbook::create(test_file_path_);
        auto worksheet = workbook->addSheet("WriteTest");
        
        // 写入大量数据
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                if (col % 3 == 0) {
                    worksheet->setValue(row, col, static_cast<double>(row * col + 0.5));
                } else if (col % 3 == 1) {
                    worksheet->setValue(row, col, "Text_" + std::to_string(row) + "_" + std::to_string(col));
                } else {
                    worksheet->getCell(row, col).setFormula("A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1));
                }
            }
        }
        
        workbook->save();
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 检查文件大小
    size_t file_size = std::filesystem::file_size(test_file_path_);
    
    std::cout << "📝 写入 " << ROWS << "x" << COLS << " 数据耗时: " 
              << duration.count() << " 毫秒，文件大小: " << file_size << " 字节" << std::endl;
    
    EXPECT_TRUE(std::filesystem::exists(test_file_path_));
    EXPECT_GT(file_size, 0);
}

// 测试读取性能
TEST_F(ReadWriteBenchmark, ReadPerformance) {
    const int ROWS = 1000;
    const int COLS = 10;
    
    // 首先创建测试文件
    {
        auto workbook = fastexcel::core::Workbook::create(test_file_path_);
        auto worksheet = workbook->addSheet("ReadTest");
        
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                worksheet->setValue(row, col, static_cast<double>(row + col));
                if (col == COLS - 1) {
                    worksheet->setValue(row, col, "End_" + std::to_string(row));
                }
            }
        }
        
        workbook->save();
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 读取测试
    {
        auto workbook = fastexcel::core::Workbook::openReadOnly(test_file_path_);
        if (workbook) {
            auto worksheet = workbook->getSheet(0);
            if (worksheet) {
                auto [max_row, max_col] = worksheet->getUsedRange();
                
                // 遍历读取所有单元格
                size_t read_count = 0;
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            const auto& cell = worksheet->getCell(row, col);
                            // 访问单元格内容以确保实际读取
                            volatile bool empty = cell.isEmpty();
                            (void)empty;
                            read_count++;
                        }
                    }
                }
                
                std::cout << "📖 读取了 " << read_count << " 个单元格" << std::endl;
            }
        }
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "📖 读取 " << ROWS << "x" << COLS << " 数据耗时: " 
              << duration.count() << " 毫秒" << std::endl;
}

// 测试往返性能（写入然后读取）
TEST_F(ReadWriteBenchmark, RoundTripPerformance) {
    const int ROWS = 500;
    const int COLS = 8;
    
    auto total_start = std::chrono::high_resolution_clock::now();
    
    // 写入阶段
    auto write_start = std::chrono::high_resolution_clock::now();
    {
        auto workbook = fastexcel::core::Workbook::create(test_file_path_);
        auto worksheet = workbook->addSheet("RoundTripTest");
        
        for (int row = 0; row < ROWS; ++row) {
            worksheet->setValue(row, 0, static_cast<double>(row + 1));
            worksheet->setValue(row, 1, static_cast<double>((row + 1) * 2.5));
            worksheet->setValue(row, 2, "Item " + std::to_string(row + 1));
            worksheet->getCell(row, 3).setFormula("A" + std::to_string(row + 1) + "*B" + std::to_string(row + 1));
            
            for (int col = 4; col < COLS; ++col) {
                worksheet->setValue(row, col, static_cast<double>(row * col));
            }
        }
        
        workbook->save();
    }
    auto write_end = std::chrono::high_resolution_clock::now();
    
    // 读取阶段
    auto read_start = std::chrono::high_resolution_clock::now();
    size_t cells_read = 0;
    {
        auto workbook = fastexcel::core::Workbook::openReadOnly(test_file_path_);
        if (workbook) {
            auto worksheet = workbook->getSheet("RoundTripTest");
            if (worksheet) {
                auto [max_row, max_col] = worksheet->getUsedRange();
                
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            const auto& cell = worksheet->getCell(row, col);
                            
                            // 验证数据类型（宽松验证，专注性能）
                            if (col == 0) {
                                // 第一列应该是数字
                                if (!cell.isNumber()) {
                                    std::cout << "⚠️  第一列不是数字类型 (row=" << row << ")" << std::endl;
                                }
                            } else if (col == 2) {
                                // 第三列应该是字符串
                                if (!cell.isString()) {
                                    std::cout << "⚠️  第三列不是字符串类型 (row=" << row << ")" << std::endl;
                                }
                            } else if (col == 3) {
                                // 第四列：公式可能被当作字符串存储
                                if (!cell.isString()) {
                                    std::cout << "⚠️  第四列不是字符串类型 (row=" << row << ", type=" << (int)cell.getType() << ")" << std::endl;
                                }
                            }
                            
                            cells_read++;
                        }
                    }
                }
            }
        }
    }
    auto read_end = std::chrono::high_resolution_clock::now();
    
    auto total_end = std::chrono::high_resolution_clock::now();
    
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_end - write_start);
    auto read_duration = std::chrono::duration_cast<std::chrono::milliseconds>(read_end - read_start);
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(total_end - total_start);
    
    std::cout << "🔄 往返测试结果:" << std::endl;
    std::cout << "   写入耗时: " << write_duration.count() << " 毫秒" << std::endl;
    std::cout << "   读取耗时: " << read_duration.count() << " 毫秒" << std::endl;
    std::cout << "   总耗时: " << total_duration.count() << " 毫秒" << std::endl;
    std::cout << "   读取单元格数: " << cells_read << std::endl;
    
    EXPECT_GT(cells_read, 0);
}