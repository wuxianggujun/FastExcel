/**
 * @file high_performance_edit_example.cpp
 * @brief 高性能Excel编辑示例
 * 
 * 展示FastExcel的高性能特性：
 * - 双通道错误处理（零成本抽象）
 * - 流式处理大文件
 * - 内存池优化
 * - 批量操作
 * - 无异常模式（可选）
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/core/Expected.hpp"
#include <iostream>
#include <chrono>
#include <vector>
#include <random>

using namespace fastexcel;
using namespace fastexcel::core;

// 性能计时器
class Timer {
private:
    std::chrono::high_resolution_clock::time_point start_;
    std::string name_;

public:
    Timer(const std::string& name) : name_(name) {
        start_ = std::chrono::high_resolution_clock::now();
    }
    
    ~Timer() {
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start_);
        std::cout << "[" << name_ << "] 耗时: " << duration.count() << "ms" << std::endl;
    }
};

/**
 * @brief 示例1：零成本错误处理
 */
void demonstrateErrorHandling() {
    std::cout << "\n=== 示例1：双通道错误处理 ===" << std::endl;
    
    // 方式1：使用Expected（零成本）
    auto loadResult = [](const std::string& filename) -> Result<std::unique_ptr<Workbook>> {
        if (filename.empty()) {
            return makeError(ErrorCode::InvalidArgument, "文件名不能为空");
        }
        
        // 模拟加载
        auto workbook = Workbook::create(filename);
        if (!workbook) {
            return makeError(ErrorCode::FileNotFound, "无法创建工作簿");
        }
        
        return makeExpected(std::move(workbook));
    };
    
    // 零成本错误检查
    auto result = loadResult("test.xlsx");
    if (result.hasValue()) {
        std::cout << "✓ 工作簿加载成功（零成本模式）" << std::endl;
        auto& workbook = result.value();
        // 使用workbook...
    } else {
        std::cout << "✗ 加载失败: " << result.error().fullMessage() << std::endl;
    }

#if FASTEXCEL_USE_EXCEPTIONS
    // 方式2：异常模式（可选）
    try {
        auto result2 = loadResult("");
        auto& workbook = result2.valueOrThrow(); // 自动抛异常
        std::cout << "✓ 工作簿加载成功（异常模式）" << std::endl;
    } catch (const FastExcelException& e) {
        std::cout << "✗ 异常捕获: " << e.what() << " (错误码: " 
                  << static_cast<int>(e.code()) << ")" << std::endl;
    }
#endif
}

/**
 * @brief 示例2：高性能大文件编辑
 */
void demonstrateHighPerformanceEditing() {
    std::cout << "\n=== 示例2：高性能大文件编辑 ===" << std::endl;
    
    Timer timer("大文件编辑");
    
    // 创建工作簿并启用超高性能模式
    auto workbook = Workbook::create("large_file_edit.xlsx");
    if (!workbook->open()) {
        std::cout << "✗ 无法打开工作簿" << std::endl;
        return;
    }
    
    // 启用超高性能模式
    workbook->setHighPerformanceMode(true);
    std::cout << "✓ 已启用超高性能模式（无压缩、大缓冲区、流式XML）" << std::endl;
    
    auto worksheet = workbook->addWorksheet("大数据表");
    if (!worksheet) {
        std::cout << "✗ 无法创建工作表" << std::endl;
        return;
    }
    
    // 批量写入大量数据
    const int ROWS = 10000;
    const int COLS = 50;
    
    {
        Timer batch_timer("批量数据写入");
        
        // 使用批量操作API（如果实现了）
        std::vector<std::vector<Cell>> batch_data;
        batch_data.reserve(ROWS);
        
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> dis(1.0, 1000.0);
        
        for (int row = 0; row < ROWS; ++row) {
            std::vector<Cell> row_data;
            row_data.reserve(COLS);
            
            for (int col = 0; col < COLS; ++col) {
                if (col == 0) {
                    // 第一列：字符串
                    row_data.emplace_back("数据行" + std::to_string(row + 1));
                } else {
                    // 其他列：数字
                    row_data.emplace_back(dis(gen));
                }
            }
            
            batch_data.push_back(std::move(row_data));
        }
        
        // 批量设置数据（假设有这个API）
        // worksheet->setBatchData(0, 0, batch_data);
        
        // 当前使用逐个设置（演示）
        for (int row = 0; row < std::min(1000, ROWS); ++row) {
            for (int col = 0; col < std::min(10, COLS); ++col) {
                if (col == 0) {
                    worksheet->writeString(row, col, "数据" + std::to_string(row));
                } else {
                    worksheet->writeNumber(row, col, dis(gen));
                }
            }
        }
        
        std::cout << "✓ 批量写入完成: " << std::min(1000, ROWS) << " 行 x " 
                  << std::min(10, COLS) << " 列" << std::endl;
    }
    
    // 保存文件
    {
        Timer save_timer("文件保存");
        if (workbook->save()) {
            std::cout << "✓ 文件保存成功" << std::endl;
        } else {
            std::cout << "✗ 文件保存失败" << std::endl;
        }
    }
    
    workbook->close();
}

/**
 * @brief 示例3：内存优化编辑
 */
void demonstrateMemoryOptimizedEditing() {
    std::cout << "\n=== 示例3：内存优化编辑 ===" << std::endl;
    
    // 加载现有文件进行编辑
    auto workbook = Workbook::loadForEdit("test_data.xlsx");
    if (!workbook) {
        std::cout << "✗ 无法加载文件，创建新文件" << std::endl;
        workbook = Workbook::create("test_data.xlsx");
        workbook->open();
        
        // 创建测试数据
        auto ws = workbook->addWorksheet("测试数据");
        for (int i = 0; i < 100; ++i) {
            ws->writeString(i, 0, "测试" + std::to_string(i));
            ws->writeNumber(i, 1, i * 1.5);
            ws->writeBoolean(i, 2, i % 2 == 0);
        }
        workbook->save();
    }
    
    std::cout << "✓ 文件加载成功" << std::endl;
    
    // 获取工作簿统计信息
    auto stats = workbook->getStatistics();
    std::cout << "工作簿统计:" << std::endl;
    std::cout << "  - 工作表数量: " << stats.total_worksheets << std::endl;
    std::cout << "  - 总单元格数: " << stats.total_cells << std::endl;
    std::cout << "  - 格式数量: " << stats.total_formats << std::endl;
    std::cout << "  - 内存使用: " << stats.memory_usage / 1024 << " KB" << std::endl;
    
    // 批量编辑操作
    auto worksheet = workbook->getWorksheet(0);
    if (worksheet) {
        Timer edit_timer("批量编辑");
        
        // 全局查找替换
        int replacements = workbook->findAndReplaceAll("测试", "编辑后", {});
        std::cout << "✓ 全局替换完成: " << replacements << " 处" << std::endl;
        
        // 批量格式设置
        auto format = workbook->createFormat();
        format->setBold(true);
        format->setFontColor(Color::BLUE);
        
        // 设置标题行格式
        for (int col = 0; col < 3; ++col) {
            worksheet->setCellFormat(0, col, format);
        }
        
        std::cout << "✓ 格式设置完成" << std::endl;
    }
    
    // 保存修改
    if (workbook->save()) {
        std::cout << "✓ 修改保存成功" << std::endl;
    }
    
    workbook->close();
}

/**
 * @brief 示例4：流式处理超大文件
 */
void demonstrateStreamingProcessing() {
    std::cout << "\n=== 示例4：流式处理超大文件 ===" << std::endl;
    
    Timer timer("流式处理");
    
    auto workbook = Workbook::create("streaming_large.xlsx");
    workbook->open();
    
    // 确保启用流式模式
    auto options = workbook->getOptions();
    options.streaming_xml = true;
    options.use_shared_strings = false;  // 禁用共享字符串以获得最佳性能
    options.row_buffer_size = 10000;     // 大缓冲区
    workbook->setOptions(options);
    
    std::cout << "✓ 流式模式配置完成" << std::endl;
    
    auto worksheet = workbook->addWorksheet("流式数据");
    
    // 模拟流式写入大量数据
    const int TOTAL_ROWS = 50000;
    const int BATCH_SIZE = 1000;
    
    for (int batch = 0; batch < TOTAL_ROWS / BATCH_SIZE; ++batch) {
        Timer batch_timer("批次 " + std::to_string(batch + 1));
        
        int start_row = batch * BATCH_SIZE;
        int end_row = std::min(start_row + BATCH_SIZE, TOTAL_ROWS);
        
        for (int row = start_row; row < end_row; ++row) {
            // 写入不同类型的数据
            worksheet->writeString(row, 0, "流式数据行" + std::to_string(row + 1));
            worksheet->writeNumber(row, 1, row * 3.14159);
            worksheet->writeBoolean(row, 2, row % 3 == 0);
            worksheet->writeFormula(row, 3, "B" + std::to_string(row + 1) + "*2");
        }
        
        // 每个批次后可以选择性地刷新缓冲区
        // worksheet->flushBuffer(); // 如果实现了这个方法
        
        if ((batch + 1) % 10 == 0) {
            std::cout << "✓ 已处理 " << end_row << " 行数据" << std::endl;
        }
    }
    
    std::cout << "✓ 流式写入完成: " << TOTAL_ROWS << " 行" << std::endl;
    
    // 保存（流式模式下应该更快）
    {
        Timer save_timer("流式保存");
        if (workbook->save()) {
            std::cout << "✓ 流式保存成功" << std::endl;
        }
    }
    
    workbook->close();
}

/**
 * @brief 主函数
 */
int main() {
    std::cout << "FastExcel 高性能编辑示例" << std::endl;
    std::cout << "=========================" << std::endl;
    
#if FASTEXCEL_USE_EXCEPTIONS
    std::cout << "异常模式: 启用" << std::endl;
#else
    std::cout << "异常模式: 禁用（纯错误码模式）" << std::endl;
#endif
    
    try {
        // 运行所有示例
        demonstrateErrorHandling();
        demonstrateHighPerformanceEditing();
        demonstrateMemoryOptimizedEditing();
        demonstrateStreamingProcessing();
        
        std::cout << "\n=== 所有示例执行完成 ===" << std::endl;
        std::cout << "生成的文件:" << std::endl;
        std::cout << "  - large_file_edit.xlsx (高性能编辑)" << std::endl;
        std::cout << "  - test_data.xlsx (内存优化编辑)" << std::endl;
        std::cout << "  - streaming_large.xlsx (流式处理)" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "✗ 程序异常: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}