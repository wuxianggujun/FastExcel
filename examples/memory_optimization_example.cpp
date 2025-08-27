/**
 * @file memory_optimization_example.cpp
 * @brief 演示内存管理优化的使用示例
 */

#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/memory/PoolManager.hpp"
#include "fastexcel/memory/PoolAllocator.hpp"
#include "fastexcel/utils/StringViewOptimized.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <chrono>
#include <vector>

using namespace fastexcel;

/**
 * @brief 展示内存池的使用
 */
void demonstrateMemoryPools() {
    std::cout << "=== Memory Pool Demonstration ===" << std::endl;
    
    // 获取Cell类型的内存池
    auto& cell_pool = memory::PoolManager::getInstance().getPool<core::Cell>();
    
    std::vector<core::Cell*> cells;
    
    // 批量分配Cell对象
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < 1000; ++i) {
        core::Cell* cell = cell_pool.allocate();
        cells.push_back(cell);
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "Allocated 1000 cells in " << duration.count() << " microseconds" << std::endl;
    std::cout << "Current pool usage: " << cell_pool.getCurrentUsage() << std::endl;
    std::cout << "Peak pool usage: " << cell_pool.getPeakUsage() << std::endl;
    
    // 批量释放
    start = std::chrono::high_resolution_clock::now();
    
    for (auto* cell : cells) {
        cell_pool.deallocate(cell);
    }
    
    end = std::chrono::high_resolution_clock::now();
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "Deallocated 1000 cells in " << duration.count() << " microseconds" << std::endl;
    std::cout << "Final pool usage: " << cell_pool.getCurrentUsage() << std::endl;
}

/**
 * @brief 展示string_view优化
 */
void demonstrateStringOptimizations() {
    std::cout << "\n=== String Optimization Demonstration ===" << std::endl;
    
    // StringJoiner使用示例
    utils::StringViewOptimized::StringJoiner joiner(", ");
    joiner.add("Apple")
          .add("Banana")
          .add("Cherry")
          .add("Date");
    
    std::string result = joiner.build();
    std::cout << "Joined string: " << result << std::endl;
    
    // StringBuilder使用示例
    utils::StringViewOptimized::StringBuilder builder(256);
    builder.append("Value: ")
           .append(42)
           .append(", Rate: ")
           .append(3.14159)
           .append("%");
    
    std::string built = builder.build();
    std::cout << "Built string: " << built << std::endl;
    
    // 字符串分割示例
    std::string_view text = "one,two,three,four,five";
    auto parts = utils::StringViewOptimized::split(text, ',');
    
    std::cout << "Split parts: ";
    for (const auto& part : parts) {
        std::cout << "[" << part << "] ";
    }
    std::cout << std::endl;
    
    // 格式化字符串示例
    std::string formatted = utils::StringViewOptimized::format(
        "Row: %d, Col: %d, Value: %.2f", 5, 3, 123.456);
    std::cout << "Formatted string: " << formatted << std::endl;
}

/**
 * @brief 展示优化的XML写入器
 */
void demonstrateOptimizedXMLWriter() {
    std::cout << "\n=== Optimized XML Writer Demonstration ===" << std::endl;
    
    // 创建内存模式的XML写入器
    xml::XMLStreamWriter xml_writer;
    
    // 写入XML内容
    xml_writer.startDocument();
    xml_writer.startElement("workbook");
    xml_writer.writeAttribute("version", "2.0");
    xml_writer.writeAttribute("optimized", true);
    
    // 批量写入属性
    xml_writer.startAttributeBatch();
    for (int i = 0; i < 10; ++i) {
        xml_writer.writeAttribute("attr" + std::to_string(i), "value" + std::to_string(i));
    }
    xml_writer.endAttributeBatch();
    
    xml_writer.startElement("worksheet");
    xml_writer.writeAttribute("name", "Sheet1");
    
    // 写入一些数据
    for (int row = 1; row <= 5; ++row) {
        for (int col = 1; col <= 5; ++col) {
            xml_writer.startElement("cell");
            xml_writer.writeAttribute("row", row);
            xml_writer.writeAttribute("col", col);
            xml_writer.writeText("Data_" + std::to_string(row) + "_" + std::to_string(col));
            xml_writer.endElement();
        }
    }
    
    xml_writer.endElement(); // worksheet
    xml_writer.endElement(); // workbook
    xml_writer.endDocument();
    
    // 获取结果
    std::string xml_content = xml_writer.toString();
    std::cout << "Generated XML size: " << xml_content.size() << " bytes" << std::endl;
    std::cout << "Bytes written: " << xml_writer.getBytesWritten() << std::endl;
    std::cout << "Flush count: " << xml_writer.getFlushCount() << std::endl;
    
    // 显示XML内容的前200字符
    std::cout << "XML content preview:" << std::endl;
    std::cout << xml_content.substr(0, 200) << "..." << std::endl;
}

/**
 * @brief 展示内存优化的Workbook
 */
void demonstrateOptimizedWorkbook() {
    std::cout << "\n=== Optimized Workbook Demonstration ===" << std::endl;
    
    try {
        // 创建内存优化的Workbook（使用统一内存管理）
        auto workbook = std::make_unique<core::Workbook>();
        
        // 创建工作表
        auto& worksheet = workbook->createWorksheet("OptimizedSheet");
        
        // 使用优化的字符串方法设置值
        std::vector<std::string_view> name_parts = {"John", "Q", "Public"};
        workbook->setCellComplexValue(1, 1, name_parts, " ");
        
        // 使用格式化方法设置值
        workbook->setCellFormattedValue(1, 2, "Score: %d/100", 85);
        
        // 批量设置优化字符串值
        for (int row = 2; row <= 10; ++row) {
            for (int col = 1; col <= 5; ++col) {
                std::string value = "R" + std::to_string(row) + "C" + std::to_string(col);
                workbook->setCellValueOptimized(row, col, value);
            }
        }
        
        // 获取内存统计
        auto stats = workbook->getMemoryStats();
        std::cout << "Memory Statistics:" << std::endl;
        std::cout << "  Cell allocations: " << stats.cell_allocations << std::endl;
        std::cout << "  Format allocations: " << stats.format_allocations << std::endl;
        std::cout << "  String optimizations: " << stats.string_optimizations << std::endl;
        std::cout << "  Cell pool usage: " << stats.cell_pool_usage << std::endl;
        std::cout << "  Format pool usage: " << stats.format_pool_usage << std::endl;
        std::cout << "  String pool size: " << stats.string_pool_size << std::endl;
        
        // 内存收缩
        workbook->shrinkMemory();
        std::cout << "Memory shrinking completed" << std::endl;
        
        std::cout << "Optimized workbook demonstration completed successfully!" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Error in optimized workbook demonstration: " << e.what() << std::endl;
    }
}

/**
 * @brief 性能比较测试
 */
void performanceComparison() {
    std::cout << "\n=== Performance Comparison ===" << std::endl;
    
    const int test_iterations = 10000;
    
    // 标准内存分配性能测试
    auto start = std::chrono::high_resolution_clock::now();
    
    std::vector<std::unique_ptr<core::Cell>> standard_cells;
    for (int i = 0; i < test_iterations; ++i) {
        standard_cells.push_back(std::make_unique<core::Cell>());
    }
    standard_cells.clear();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto standard_duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 内存池分配性能测试
    start = std::chrono::high_resolution_clock::now();
    
    auto& pool = memory::PoolManager::getInstance().getPool<core::Cell>();
    std::vector<core::Cell*> pool_cells;
    
    for (int i = 0; i < test_iterations; ++i) {
        pool_cells.push_back(pool.allocate());
    }
    
    for (auto* cell : pool_cells) {
        pool.deallocate(cell);
    }
    
    end = std::chrono::high_resolution_clock::now();
    auto pool_duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "Standard allocation: " << standard_duration.count() << " microseconds" << std::endl;
    std::cout << "Pool allocation: " << pool_duration.count() << " microseconds" << std::endl;
    std::cout << "Performance improvement: " 
              << (static_cast<double>(standard_duration.count()) / pool_duration.count()) 
              << "x faster" << std::endl;
}

int main() {
    try {
        // 设置日志级别
        FASTEXCEL_LOG_INFO("Starting FastExcel Memory Optimization Demonstration");
        
        demonstrateMemoryPools();
        demonstrateStringOptimizations();
        demonstrateOptimizedXMLWriter();
        demonstrateOptimizedWorkbook();
        performanceComparison();
        
        std::cout << "\n=== All demonstrations completed successfully! ===" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Fatal error: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}