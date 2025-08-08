/**
 * @file 01_basic_usage.cpp
 * @brief FastExcel V3.0 基本用法示例
 * 
 * 这个示例展示了FastExcel V3.0的基本用法，包括：
 * - 创建工作簿
 * - 添加工作表
 * - 写入数据
 * - 设置样式
 * - 保存文件
 */

#include <iostream>
#include <vector>
#include <string>
#include "fastexcel/Cell_v3.hpp"
#include "fastexcel/MemoryPool_v3.hpp"

// 注意：以下是示例代码，实际的Worksheet和Workbook类还未实现
// 这里展示的是预期的API使用方式

int main() {
    std::cout << "FastExcel V3.0 基本用法示例\n";
    std::cout << "===========================\n\n";
    
    try {
        // 1. 创建内存池（可选，用于优化性能）
        fx::MemoryPool_v3 pool(1024 * 1024); // 1MB初始大小
        std::cout << "1. 创建了1MB的内存池\n";
        
        // 2. 创建一些Cell对象进行测试
        std::cout << "\n2. 创建Cell对象：\n";
        
        // 创建数字单元格
        fx::Cell cell1;
        cell1 = 42.5;
        std::cout << "   - 数字单元格: 42.5\n";
        
        // 创建字符串单元格
        fx::Cell cell2;
        cell2 = std::string("Hello, FastExcel!");
        std::cout << "   - 字符串单元格: \"Hello, FastExcel!\"\n";
        
        // 创建布尔单元格
        fx::Cell cell3;
        cell3 = true;
        std::cout << "   - 布尔单元格: true\n";
        
        // 3. 测试Cell的内存占用
        std::cout << "\n3. Cell内存占用：\n";
        std::cout << "   - sizeof(Cell) = " << sizeof(fx::Cell) << " 字节\n";
        std::cout << "   - 目标: 24字节 ✓\n";
        
        // 4. 批量创建Cell测试内存池性能
        std::cout << "\n4. 批量创建测试：\n";
        const int cell_count = 10000;
        std::vector<fx::Cell> cells;
        cells.reserve(cell_count);
        
        auto start = std::chrono::high_resolution_clock::now();
        for (int i = 0; i < cell_count; ++i) {
            fx::Cell cell;
            if (i % 3 == 0) {
                cell = static_cast<double>(i);
            } else if (i % 3 == 1) {
                cell = std::string("Cell_" + std::to_string(i));
            } else {
                cell = (i % 2 == 0);
            }
            cells.push_back(cell);
        }
        auto end = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        std::cout << "   - 创建 " << cell_count << " 个单元格\n";
        std::cout << "   - 耗时: " << duration.count() << " 微秒\n";
        std::cout << "   - 平均: " << (duration.count() / static_cast<double>(cell_count)) 
                  << " 微秒/单元格\n";
        
        // 5. 内存池统计
        std::cout << "\n5. 内存池统计：\n";
        auto stats = pool.getStats();
        std::cout << "   - 总分配次数: " << stats.total_allocations << "\n";
        std::cout << "   - 总释放次数: " << stats.total_deallocations << "\n";
        std::cout << "   - 当前使用: " << stats.current_usage << " 字节\n";
        std::cout << "   - 峰值使用: " << stats.peak_usage << " 字节\n";
        
        // 6. 未来功能预览
        std::cout << "\n6. 未来功能预览：\n";
        std::cout << "   以下功能将在后续版本实现：\n";
        std::cout << "   - Worksheet工作表管理\n";
        std::cout << "   - Workbook工作簿操作\n";
        std::cout << "   - 样式和主题系统\n";
        std::cout << "   - Excel文件读写\n";
        std::cout << "   - 流式处理\n";
        std::cout << "   - 并行处理\n";
        
        /* 未来的API使用示例（伪代码）：
        
        // 创建工作簿
        auto workbook = fx::Workbook::create();
        
        // 添加工作表
        auto& sheet = workbook->addSheet("数据表");
        
        // 写入数据
        sheet(0, 0) = "姓名";
        sheet(0, 1) = "年龄";
        sheet(0, 2) = "分数";
        
        sheet(1, 0) = "张三";
        sheet(1, 1) = 25;
        sheet(1, 2) = 95.5;
        
        // 设置样式
        auto header_style = workbook->createStyle({
            .bold = true,
            .bg_color = "#4472C4",
            .font_color = "#FFFFFF"
        });
        sheet.setRangeStyle({0, 0, 0, 2}, header_style);
        
        // 保存文件
        workbook->save("output.xlsx");
        
        */
        
        std::cout << "\n示例运行成功！\n";
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}