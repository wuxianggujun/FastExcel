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
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path(test_file_path_));
        workbook->open();
        auto worksheet = workbook->addWorksheet("WriteTest");
        
        // 写入大量数据
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                if (col % 3 == 0) {
                    worksheet->writeNumber(row, col, row * col + 0.5);
                } else if (col % 3 == 1) {
                    worksheet->writeString(row, col, "Text_" + std::to_string(row) + "_" + std::to_string(col));
                } else {
                    worksheet->writeFormula(row, col, "A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1));
                }
            }
        }
        
        workbook->save();
        workbook->close();
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
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path(test_file_path_));
        workbook->open();
        auto worksheet = workbook->addWorksheet("ReadTest");
        
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                worksheet->writeNumber(row, col, row + col);
                if (col == COLS - 1) {
                    worksheet->writeString(row, col, "End_" + std::to_string(row));
                }
            }
        }
        
        workbook->save();
        workbook->close();
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 读取测试
    {
        auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path(test_file_path_));
        if (workbook) {
            auto worksheet_names = workbook->getWorksheetNames();
            EXPECT_FALSE(worksheet_names.empty());
            
            auto worksheet = workbook->getWorksheet(worksheet_names[0]);
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
            
            workbook->close();
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
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path(test_file_path_));
        workbook->open();
        auto worksheet = workbook->addWorksheet("RoundTripTest");
        
        for (int row = 0; row < ROWS; ++row) {
            worksheet->writeNumber(row, 0, row + 1);
            worksheet->writeNumber(row, 1, (row + 1) * 2.5);
            worksheet->writeString(row, 2, "Item " + std::to_string(row + 1));
            worksheet->writeFormula(row, 3, "A" + std::to_string(row + 1) + "*B" + std::to_string(row + 1));
            
            for (int col = 4; col < COLS; ++col) {
                worksheet->writeNumber(row, col, row * col);
            }
        }
        
        workbook->save();
        workbook->close();
    }
    auto write_end = std::chrono::high_resolution_clock::now();
    
    // 读取阶段
    auto read_start = std::chrono::high_resolution_clock::now();
    size_t cells_read = 0;
    {
        auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path(test_file_path_));
        if (workbook) {
            auto worksheet = workbook->getWorksheet("RoundTripTest");
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
            workbook->close();
        }
    }
    auto read_end = std::chrono::high_resolution_clock::now();
    
    auto total_end = std::chrono::high_resolution_clock::now();
    
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_end - write_start);
    auto read_duration = std::chrono::duration_cast<std::chrono::milliseconds>(read_end - read_start);
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(total_end - total_start);
    
    size_t file_size = std::filesystem::file_size(test_file_path_);
    
    std::cout << "🔄 往返测试 " << ROWS << "x" << COLS << " 数据:" << std::endl;
    std::cout << "   写入耗时: " << write_duration.count() << " 毫秒" << std::endl;
    std::cout << "   读取耗时: " << read_duration.count() << " 毫秒，读取 " << cells_read << " 个单元格" << std::endl;
    std::cout << "   总耗时: " << total_duration.count() << " 毫秒，文件大小: " << file_size << " 字节" << std::endl;
    
    EXPECT_GT(cells_read, 0);
    EXPECT_GT(file_size, 0);
}

// 测试大文件写入性能
TEST_F(ReadWriteBenchmark, LargeFileWritePerformance) {
    const int ROWS = 10000;
    const int COLS = 5;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("large_" + test_file_path_));
        workbook->open();
        auto worksheet = workbook->addWorksheet("LargeData");
        
        // 写入大量数据
        for (int row = 0; row < ROWS; ++row) {
            for (int col = 0; col < COLS; ++col) {
                if (col == 0) {
                    worksheet->writeNumber(row, col, row + 1);
                } else if (col == 1) {
                    worksheet->writeNumber(row, col, (row + 1) * 1.5);
                } else {
                    worksheet->writeString(row, col, "Data_" + std::to_string(row) + "_" + std::to_string(col));
                }
            }
            
            // 每1000行输出一次进度
            if ((row + 1) % 1000 == 0) {
                std::cout << "✍️  已写入 " << (row + 1) << " 行数据..." << std::endl;
            }
        }
        
        workbook->save();
        workbook->close();
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::string large_file = "large_" + test_file_path_;
    size_t file_size = std::filesystem::file_size(large_file);
    
    std::cout << "📈 大文件写入 " << ROWS << "x" << COLS << " (" << (ROWS * COLS) << " 个单元格) 耗时: " 
              << duration.count() << " 毫秒，文件大小: " << (file_size / 1024) << " KB" << std::endl;
    
    // 清理大文件
    std::filesystem::remove(large_file);
    
    EXPECT_GT(file_size, 0);
}