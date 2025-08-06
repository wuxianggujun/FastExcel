/**
 * @file optimization_demo.cpp
 * @brief FastExcel优化效果演示
 *
 * 展示新的时间工具类和设计模式优化的效果。
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include <iostream>
#include <chrono>
#include <thread>

void demonstrateTimeUtils() {
    std::cout << "\n=== 时间工具类演示 ===" << std::endl;
    
    // 获取当前时间
    auto current_time = fastexcel::utils::TimeUtils::getCurrentTime();
    std::cout << "当前时间: " << fastexcel::utils::TimeUtils::formatTimeISO8601(current_time) << std::endl;
    
    // 创建特定日期
    auto specific_date = fastexcel::utils::TimeUtils::createTime(2024, 8, 6, 9, 15, 30);
    std::cout << "特定日期: " << fastexcel::utils::TimeUtils::formatTime(specific_date, "%Y年%m月%d日 %H:%M:%S") << std::endl;
    
    // Excel序列号转换
    double excel_serial = fastexcel::utils::TimeUtils::toExcelSerialNumber(specific_date);
    std::cout << "Excel序列号: " << excel_serial << std::endl;
    
    // 性能计时器演示
    {
        fastexcel::utils::TimeUtils::PerformanceTimer timer("时间工具类测试");
        // 模拟一些工作
        std::this_thread::sleep_for(std::chrono::milliseconds(100));
        std::cout << "计时器已经运行了 " << timer.elapsedMs() << " 毫秒" << std::endl;
    } // 析构时自动输出总时间
}

void demonstrateWorkbookOptimization() {
    std::cout << "\n=== 工作簿优化演示 ===" << std::endl;
    
    try {
        // 初始化FastExcel库
        fastexcel::initialize();
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("optimization_demo.xlsx");
        if (!workbook->open()) {
            std::cerr << "无法打开工作簿" << std::endl;
            return;
        }
        
        // 添加一些测试数据
        auto worksheet = workbook->addWorksheet("优化演示");
        
        // 使用时间工具类设置文档属性
        auto creation_time = fastexcel::utils::TimeUtils::getCurrentTime();
        workbook->setCreatedTime(creation_time);
        workbook->setTitle("FastExcel优化演示");
        workbook->setAuthor("FastExcel优化版本");
        
        // 写入一些数据
        worksheet->writeString(0, 0, "优化项目");
        worksheet->writeString(0, 1, "状态");
        worksheet->writeString(0, 2, "完成时间");
        
        worksheet->writeString(1, 0, "时间工具类");
        worksheet->writeString(1, 1, "✅ 完成");
        worksheet->writeDateTime(1, 2, creation_time);
        
        worksheet->writeString(2, 0, "统一接口设计");
        worksheet->writeString(2, 1, "✅ 完成");
        worksheet->writeDateTime(2, 2, creation_time);
        
        worksheet->writeString(3, 0, "策略模式");
        worksheet->writeString(3, 1, "✅ 完成");
        worksheet->writeDateTime(3, 2, creation_time);
        
        // 演示不同的工作簿模式
        std::cout << "\n--- 工作簿模式演示 ---" << std::endl;
        
        // 批量模式
        workbook->setMode(fastexcel::core::WorkbookMode::BATCH);
        std::cout << "设置为批量模式: " << (workbook->getMode() == fastexcel::core::WorkbookMode::BATCH ? "成功" : "失败") << std::endl;
        
        // 流式模式
        workbook->setMode(fastexcel::core::WorkbookMode::STREAMING);
        std::cout << "设置为流式模式: " << (workbook->getMode() == fastexcel::core::WorkbookMode::STREAMING ? "成功" : "失败") << std::endl;
        
        // 自动模式
        workbook->setMode(fastexcel::core::WorkbookMode::AUTO);
        std::cout << "设置为自动模式: " << (workbook->getMode() == fastexcel::core::WorkbookMode::AUTO ? "成功" : "失败") << std::endl;
        
        // 保存文件
        auto start = std::chrono::high_resolution_clock::now();
        bool success = workbook->save();
        auto end = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
        std::cout << "文件保存" << (success ? "成功" : "失败") << "，耗时: " << duration.count() << "ms" << std::endl;
        
        // 获取统计信息
        auto stats = workbook->getStatistics();
        std::cout << "工作簿统计信息:" << std::endl;
        std::cout << "  - 工作表数量: " << stats.total_worksheets << std::endl;
        std::cout << "  - 总单元格数: " << stats.total_cells << std::endl;
        std::cout << "  - 格式数量: " << stats.total_formats << std::endl;
        std::cout << "  - 内存使用: " << stats.memory_usage << " 字节" << std::endl;
        
        std::cout << "\n优化效果总结:" << std::endl;
        std::cout << "✅ 时间处理统一 - 所有时间操作使用TimeUtils" << std::endl;
        std::cout << "✅ 智能模式选择 - 可以根据数据量自动选择最优模式" << std::endl;
        std::cout << "✅ 性能监控 - 支持实时统计信息" << std::endl;
        std::cout << "✅ 跨平台兼容 - 统一的时间处理API" << std::endl;
        
        // 清理资源
        workbook->close();
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
    }
}

void demonstrateDesignPatterns() {
    std::cout << "\n=== 设计模式演示 ===" << std::endl;
    
    std::cout << "1. 策略模式 (Strategy Pattern):" << std::endl;
    std::cout << "   - WorkbookMode 枚举定义不同的处理策略" << std::endl;
    std::cout << "   - BATCH/STREAMING/AUTO 模式可动态切换" << std::endl;
    std::cout << "   - 根据数据量自动选择最优策略" << std::endl;
    
    std::cout << "\n2. 工厂模式 (Factory Pattern):" << std::endl;
    std::cout << "   - Workbook::create() 使用工厂方法创建工作簿" << std::endl;
    std::cout << "   - Format 创建通过 createFormat() 工厂方法" << std::endl;
    std::cout << "   - 可以根据参数创建不同类型的对象" << std::endl;
    
    std::cout << "\n3. RAII模式 (Resource Acquisition Is Initialization):" << std::endl;
    std::cout << "   - TimeUtils::PerformanceTimer 自动管理计时资源" << std::endl;
    std::cout << "   - 智能指针管理内存资源" << std::endl;
    std::cout << "   - 工作簿自动管理文件句柄和资源" << std::endl;
    
    std::cout << "\n4. 单例模式 (Singleton Pattern):" << std::endl;
    std::cout << "   - FormatPool 管理全局格式资源" << std::endl;
    std::cout << "   - 避免重复创建相同的格式对象" << std::endl;
    
    std::cout << "\n5. 观察者模式 (Observer Pattern) - 计划中:" << std::endl;
    std::cout << "   - 进度通知系统" << std::endl;
    std::cout << "   - 事件驱动的状态更新" << std::endl;
}

int main() {
    std::cout << "FastExcel 代码优化演示程序" << std::endl;
    std::cout << "================================" << std::endl;
    
    // 演示时间工具类
    demonstrateTimeUtils();
    
    // 演示工作簿优化
    demonstrateWorkbookOptimization();
    
    // 演示设计模式
    demonstrateDesignPatterns();
    
    std::cout << "\n演示完成！生成的文件:" << std::endl;
    std::cout << "- optimization_demo.xlsx (优化演示文件)" << std::endl;
    
    return 0;
}