#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <chrono>
#include <fstream>
#include <memory>

using namespace fastexcel;

// 演示XMLStreamWriter的回调模式功能
void demonstrateCallbackMode() {
    std::cout << "\n=== XMLStreamWriter 回调模式演示 ===" << std::endl;
    
    // 创建输出文件流
    std::ofstream output_file("streaming_output.xml", std::ios::binary);
    if (!output_file.is_open()) {
        std::cerr << "无法创建输出文件" << std::endl;
        return;
    }
    
    size_t total_bytes_written = 0;
    size_t chunk_count = 0;
    
    // 创建XMLStreamWriter并设置回调模式
    xml::XMLStreamWriter writer;
    writer.setCallbackMode([&output_file, &total_bytes_written, &chunk_count](const std::string& chunk) {
        // 直接将XML块写入文件，实现真正的流式写入
        output_file.write(chunk.c_str(), chunk.size());
        total_bytes_written += chunk.size();
        chunk_count++;
        
        // 每写入一定数量的块就输出进度
        if (chunk_count % 100 == 0) {
            std::cout << "已写入 " << chunk_count << " 个块，总计 " << total_bytes_written << " 字节" << std::endl;
        }
    }, true); // 启用自动刷新
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 生成大型XML文档
    writer.startDocument();
    writer.startElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    // 生成大量工作表数据
    const int num_sheets = 5;
    const int rows_per_sheet = 10000;
    const int cols_per_row = 10;
    
    for (int sheet = 0; sheet < num_sheets; ++sheet) {
        writer.startElement("worksheet");
        writer.writeAttribute("name", "Sheet" + std::to_string(sheet + 1));
        
        writer.startElement("sheetData");
        
        for (int row = 0; row < rows_per_sheet; ++row) {
            writer.startElement("row");
            writer.writeAttribute("r", std::to_string(row + 1));
            
            for (int col = 0; col < cols_per_row; ++col) {
                writer.startElement("c");
                writer.writeAttribute("r", std::string(1, 'A' + col) + std::to_string(row + 1));
                writer.writeAttribute("t", "inlineStr");
                
                writer.startElement("is");
                writer.startElement("t");
                writer.writeText("Cell(" + std::to_string(row) + "," + std::to_string(col) + ")");
                writer.endElement(); // t
                writer.endElement(); // is
                writer.endElement(); // c
            }
            
            writer.endElement(); // row
            
            // 每处理1000行就输出进度
            if ((row + 1) % 1000 == 0) {
                std::cout << "Sheet " << (sheet + 1) << ": 已处理 " << (row + 1) << " 行" << std::endl;
            }
        }
        
        writer.endElement(); // sheetData
        writer.endElement(); // worksheet
        
        std::cout << "完成 Sheet " << (sheet + 1) << std::endl;
    }
    
    writer.endElement(); // workbook
    writer.endDocument();
    
    // 最终刷新
    writer.flushBuffer();
    output_file.close();
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    
    std::cout << "\n流式XML写入完成:" << std::endl;
    std::cout << "- 总时间: " << duration.count() << " 毫秒" << std::endl;
    std::cout << "- 总字节数: " << total_bytes_written << " 字节" << std::endl;
    std::cout << "- 块数量: " << chunk_count << " 个" << std::endl;
    std::cout << "- 平均块大小: " << (chunk_count > 0 ? total_bytes_written / chunk_count : 0) << " 字节" << std::endl;
    std::cout << "- 写入速度: " << (duration.count() > 0 ? (total_bytes_written / 1024.0) / (duration.count() / 1000.0) : 0) << " KB/s" << std::endl;
}

// 比较缓冲模式和回调模式的性能
void comparePerformance() {
    std::cout << "\n=== 性能比较：缓冲模式 vs 回调模式 ===" << std::endl;
    
    const int test_rows = 5000;
    const int test_cols = 8;
    
    // 测试缓冲模式
    {
        std::cout << "\n测试缓冲模式..." << std::endl;
        auto start_time = std::chrono::high_resolution_clock::now();
        
        xml::XMLStreamWriter writer;
        writer.setBufferedMode();
        
        writer.startDocument();
        writer.startElement("data");
        
        for (int row = 0; row < test_rows; ++row) {
            writer.startElement("row");
            for (int col = 0; col < test_cols; ++col) {
                writer.startElement("cell");
                writer.writeText("Data_" + std::to_string(row) + "_" + std::to_string(col));
                writer.endElement();
            }
            writer.endElement();
        }
        
        writer.endElement();
        writer.endDocument();
        
        std::string result = writer.toString();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        std::cout << "缓冲模式结果:" << std::endl;
        std::cout << "- 时间: " << duration.count() << " 毫秒" << std::endl;
        std::cout << "- 输出大小: " << result.size() << " 字节" << std::endl;
        
        // 写入文件以便比较
        std::ofstream file("buffered_output.xml");
        file << result;
        file.close();
    }
    
    // 测试回调模式
    {
        std::cout << "\n测试回调模式..." << std::endl;
        auto start_time = std::chrono::high_resolution_clock::now();
        
        std::ofstream output_file("callback_output.xml");
        size_t total_size = 0;
        
        xml::XMLStreamWriter writer;
        writer.setCallbackMode([&output_file, &total_size](const std::string& chunk) {
            output_file.write(chunk.c_str(), chunk.size());
            total_size += chunk.size();
        }, true);
        
        writer.startDocument();
        writer.startElement("data");
        
        for (int row = 0; row < test_rows; ++row) {
            writer.startElement("row");
            for (int col = 0; col < test_cols; ++col) {
                writer.startElement("cell");
                writer.writeText("Data_" + std::to_string(row) + "_" + std::to_string(col));
                writer.endElement();
            }
            writer.endElement();
        }
        
        writer.endElement();
        writer.endDocument();
        writer.flushBuffer();
        
        output_file.close();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        std::cout << "回调模式结果:" << std::endl;
        std::cout << "- 时间: " << duration.count() << " 毫秒" << std::endl;
        std::cout << "- 输出大小: " << total_size << " 字节" << std::endl;
    }
}

// 演示高性能Excel文件生成
void demonstrateHighPerformanceExcel() {
    std::cout << "\n=== 高性能Excel文件生成演示 ===" << std::endl;
    
    try {
        auto workbook = core::Workbook::create("streaming_performance_test.xlsx");
        if (!workbook->open()) {
            std::cerr << "无法创建工作簿" << std::endl;
            return;
        }
        
        // 启用高性能模式
        workbook->setHighPerformanceMode(true);
        
        // 创建工作表
        auto worksheet = workbook->addWorksheet("PerformanceTest");
        if (!worksheet) {
            std::cerr << "无法创建工作表" << std::endl;
            return;
        }
        
        std::cout << "开始生成大量数据..." << std::endl;
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 生成大量数据
        const int num_rows = 50000;
        const int num_cols = 10;
        
        for (int row = 0; row < num_rows; ++row) {
            for (int col = 0; col < num_cols; ++col) {
                if (col == 0) {
                    // 第一列写入字符串
                    worksheet->writeString(row, col, "Row " + std::to_string(row + 1));
                } else if (col == 1) {
                    // 第二列写入数字
                    worksheet->writeNumber(row, col, row * col + 0.5);
                } else {
                    // 其他列写入公式或数据
                    worksheet->writeString(row, col, "Data_" + std::to_string(row) + "_" + std::to_string(col));
                }
            }
            
            // 每处理5000行输出进度
            if ((row + 1) % 5000 == 0) {
                std::cout << "已处理 " << (row + 1) << " 行" << std::endl;
            }
        }
        
        auto data_time = std::chrono::high_resolution_clock::now();
        auto data_duration = std::chrono::duration_cast<std::chrono::milliseconds>(data_time - start_time);
        
        std::cout << "数据写入完成，开始保存文件..." << std::endl;
        
        // 保存文件
        bool success = workbook->save();
        workbook->close();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - data_time);
        
        if (success) {
            std::cout << "\n高性能Excel文件生成完成:" << std::endl;
            std::cout << "- 数据行数: " << num_rows << std::endl;
            std::cout << "- 数据列数: " << num_cols << std::endl;
            std::cout << "- 总单元格数: " << (num_rows * num_cols) << std::endl;
            std::cout << "- 数据写入时间: " << data_duration.count() << " 毫秒" << std::endl;
            std::cout << "- 文件保存时间: " << save_duration.count() << " 毫秒" << std::endl;
            std::cout << "- 总时间: " << total_duration.count() << " 毫秒" << std::endl;
            std::cout << "- 处理速度: " << (total_duration.count() > 0 ? (num_rows * num_cols) / (total_duration.count() / 1000.0) : 0) << " 单元格/秒" << std::endl;
        } else {
            std::cerr << "保存文件失败" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cerr << "异常: " << e.what() << std::endl;
    }
}

int main() {
    std::cout << "FastExcel 流式XML写入演示程序" << std::endl;
    std::cout << "================================" << std::endl;
    
    // 设置日志级别
    utils::Logger::setLevel(utils::LogLevel::INFO);
    
    try {
        // 演示XMLStreamWriter的回调模式
        demonstrateCallbackMode();
        
        // 比较不同模式的性能
        comparePerformance();
        
        // 演示高性能Excel文件生成
        demonstrateHighPerformanceExcel();
        
        std::cout << "\n所有演示完成！" << std::endl;
        std::cout << "生成的文件:" << std::endl;
        std::cout << "- streaming_output.xml (流式XML输出)" << std::endl;
        std::cout << "- buffered_output.xml (缓冲模式输出)" << std::endl;
        std::cout << "- callback_output.xml (回调模式输出)" << std::endl;
        std::cout << "- streaming_performance_test.xlsx (高性能Excel文件)" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "程序异常: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}