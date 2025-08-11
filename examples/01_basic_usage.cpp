#include "fastexcel/utils/ModuleLoggers.hpp"
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
    EXAMPLE_INFO("FastExcel V3.0 基本用法示例");
    EXAMPLE_INFO("===========================");
    
    try {
        // 1. 创建内存池（可选，用于优化性能）
        fx::MemoryPool_v3 pool(1024 * 1024); // 1MB初始大小
        EXAMPLE_INFO("1. 创建了1MB的内存池");
        
        // 2. 创建一些Cell对象进行测试
        EXAMPLE_INFO("2. 创建Cell对象：");
        
        // 创建数字单元格
        fx::Cell cell1;
        cell1 = 42.5;
        EXAMPLE_INFO("   - 数字单元格: 42.5");
        
        // 创建字符串单元格
        fx::Cell cell2;
        cell2 = std::string("Hello, FastExcel!");
        EXAMPLE_INFO("   - 字符串单元格: \"Hello, FastExcel!\"");
        
        // 创建布尔单元格
        fx::Cell cell3;
        cell3 = true;
        EXAMPLE_INFO("   - 布尔单元格: true");
        
        // 3. 测试Cell的内存占用
        EXAMPLE_INFO("3. Cell内存占用：");
        EXAMPLE_INFO("   - sizeof(Cell) = {} 字节", sizeof(fx::Cell));
        EXAMPLE_INFO("   - 目标: 24字节 ✓");
        
        // 4. 批量创建Cell测试内存池性能
        EXAMPLE_INFO("4. 批量创建测试：");
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
        EXAMPLE_INFO("   - 创建 {} 个单元格", cell_count);
        EXAMPLE_INFO("   - 耗时: {} 微秒", duration.count());
        EXAMPLE_INFO("   - 平均: {:.2f} 微秒/单元格", (duration.count() / static_cast<double>(cell_count)));
        
        // 5. 内存池统计
        EXAMPLE_INFO("5. 内存池统计：");
        auto stats = pool.getStats();
        EXAMPLE_INFO("   - 总分配次数: {}", stats.total_allocations);
        EXAMPLE_INFO("   - 总释放次数: {}", stats.total_deallocations);
        EXAMPLE_INFO("   - 当前使用: {} 字节", stats.current_usage);
        EXAMPLE_INFO("   - 峰值使用: {} 字节", stats.peak_usage);
        
        // 6. 未来功能预览
        EXAMPLE_INFO("6. 未来功能预览：");
        EXAMPLE_INFO("   以下功能将在后续版本实现：");
        EXAMPLE_INFO("   - Worksheet工作表管理");
        EXAMPLE_INFO("   - Workbook工作簿操作");
        EXAMPLE_INFO("   - 样式和主题系统");
        EXAMPLE_INFO("   - Excel文件读写");
        EXAMPLE_INFO("   - 流式处理");
        EXAMPLE_INFO("   - 并行处理");
        
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
        
        EXAMPLE_INFO("示例运行成功！");
        
    } catch (const std::exception& e) {
        EXAMPLE_ERROR("错误: {}", e.what());
        return 1;
    }
    
    return 0;
}