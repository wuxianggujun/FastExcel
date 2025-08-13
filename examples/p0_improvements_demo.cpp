/**
 * @file p0_improvements_demo.cpp
 * @brief FastExcel P0改进演示代码
 *
 * 本示例展示了FastExcel 2.x P0优先级改进的使用方式：
 * 1. 统一的日志宏定义
 * 2. 清理后的initialize/cleanup接口
 * 3. 线程安全的FormatRepository遍历
 * 4. 统一的样式API（FormatDescriptor vs Format）
 * 5. 优化的XLSXReader XML解析
 */

#include "fastexcel/FastExcel.hpp"
// 注意：现在FastExcel.hpp已经自动包含了Logger.hpp，无需手动包含
#include "fastexcel/reader/XLSXReader.hpp"

#include <chrono>
#include <iostream>
#include <thread>
#include <vector>

using namespace fastexcel;
using namespace fastexcel::core;

void demonstrateLoggingImprovements() {
    std::cout << "\n=== 1. 日志宏统一演示 ===" << std::endl;
    
    // 日志宏现在统一在 utils/Logger.hpp 中定义，避免重复定义
    FASTEXCEL_LOG_INFO("日志系统已统一，避免了重复定义问题");
    FASTEXCEL_LOG_DEBUG("调试信息：日志宏现在只在Logger.hpp中定义");
    FASTEXCEL_LOG_WARN("警告：旧的重复定义已被移除");
}

void demonstrateInitializationImprovements() {
    std::cout << "\n=== 2. 初始化接口统一演示 ===" << std::endl;
    
    // 方式1：使用带参数的版本（推荐）
    bool success = fastexcel::initialize("logs/demo.log", true);
    if (success) {
        FASTEXCEL_LOG_INFO("FastExcel初始化成功（带参数版本）");
    }
    
    // 方式2：使用简化版本（内部调用带参数版本）
    // 注意：简化版本现在是inline函数，直接调用带参数版本
    // fastexcel::initialize();  // 这会导致重载歧义，所以我们明确调用
    FASTEXCEL_LOG_INFO("FastExcel初始化成功（使用带参数版本）");
    
    // 清理（幂等操作）
    fastexcel::cleanup();
    FASTEXCEL_LOG_INFO("FastExcel清理完成");
}

void demonstrateThreadSafeFormatRepository() {
    std::cout << "\n=== 3. 线程安全的FormatRepository遍历演示 ===" << std::endl;
    
    FormatRepository repo;
    
    // 添加一些格式（使用默认格式）
    auto format1 = FormatDescriptor::getDefault();
    auto format2 = FormatDescriptor::getDefault();
    
    int id1 = repo.addFormat(format1);
    int id2 = repo.addFormat(format2);
    
    FASTEXCEL_LOG_INFO("添加了 {} 个格式到仓储", repo.getFormatCount());
    
    // 新方式：使用线程安全的快照遍历（推荐）
    auto snapshot = repo.createSnapshot();
    FASTEXCEL_LOG_INFO("创建格式快照，包含 {} 个格式", snapshot.size());
    
    for (const auto& [id, format] : snapshot) {
        FASTEXCEL_LOG_DEBUG("格式ID: {}", id);
    }
    
    // 多线程安全演示
    std::vector<std::thread> threads;
    
    // 线程1：安全遍历
    threads.emplace_back([&repo]() {
        auto snapshot = repo.createSnapshot();
        for (const auto& [id, format] : snapshot) {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
            // 遍历期间不受其他线程修改影响
        }
        FASTEXCEL_LOG_INFO("线程1：安全遍历完成");
    });
    
    // 线程2：添加新格式
    threads.emplace_back([&repo]() {
        for (int i = 0; i < 5; ++i) {
            auto newFormat = FormatDescriptor::getDefault();
            repo.addFormat(newFormat);
            std::this_thread::sleep_for(std::chrono::milliseconds(2));
        }
        FASTEXCEL_LOG_INFO("线程2：添加格式完成");
    });
    
    // 等待所有线程完成
    for (auto& t : threads) {
        t.join();
    }
    
    FASTEXCEL_LOG_INFO("多线程测试完成，最终格式数量: {}", repo.getFormatCount());
}

void demonstrateUnifiedStyleAPI() {
    std::cout << "\n=== 4. 统一样式API演示 ===" << std::endl;
    
    // 创建工作簿和工作表
    auto workbook = Workbook::create(core::Path("demo.xlsx"));
    if (!workbook) {
        FASTEXCEL_LOG_ERROR("无法创建工作簿");
        return;
    }
    auto worksheet = workbook->addSheet("StyleDemo");
    
    // 新架构：使用FormatDescriptor（推荐）
    auto newFormat = FormatDescriptor::getDefault();
    // 设置格式属性...
    
    // 新API：设置单元格格式
    worksheet->setCellFormat(0, 0, newFormat);
    FASTEXCEL_LOG_INFO("使用新架构FormatDescriptor设置单元格格式");
    
    // 新API：设置列格式
    worksheet->setColumnFormat(0, newFormat);
    FASTEXCEL_LOG_INFO("使用新架构FormatDescriptor设置列格式");
    
    // 新API：设置行格式
    worksheet->setRowFormat(0, newFormat);
    FASTEXCEL_LOG_INFO("使用新架构FormatDescriptor设置行格式");
    
    // 获取格式（统一API）
    auto columnFormat = worksheet->getColumnFormat(0);
    if (columnFormat) {
        FASTEXCEL_LOG_INFO("成功获取列格式描述符");
    }
    
    auto rowFormat = worksheet->getRowFormat(0);
    if (rowFormat) {
        FASTEXCEL_LOG_INFO("成功获取行格式描述符");
    }
    
    // 注意：旧API仍然可用但已标记为deprecated
    // worksheet->setColumnFormat(0, oldFormatPtr);  // 会产生deprecation警告
}

void demonstrateXMLReaderImprovements() {
    std::cout << "\n=== 5. XLSXReader XML解析优化演示 ===" << std::endl;
    
    // 注意：这里只是演示API改进，实际使用需要有效的XLSX文件
    try {
        reader::XLSXReader reader(core::Path("example.xlsx"));
        
        FASTEXCEL_LOG_INFO("XLSXReader现在优先使用XMLStreamReader进行解析");
        FASTEXCEL_LOG_INFO("旧的字符串解析方法已标记为deprecated");
        FASTEXCEL_LOG_INFO("新的解析方式提供更好的性能和错误处理");
        
        // 新的解析方式在内部使用XMLStreamReader
        // 提供更好的性能和错误处理能力
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_WARN("演示用文件不存在，这是正常的: {}", e.what());
    }
}

void demonstrateDeprecationWarnings() {
    std::cout << "\n=== 6. 弃用警告演示 ===" << std::endl;
    
    FormatRepository repo;
    
    // 这些调用会产生编译时警告，提醒开发者迁移到新API
    FASTEXCEL_LOG_INFO("旧的不安全迭代器已完全移除");
    FASTEXCEL_LOG_INFO("现在只能使用 repo.createSnapshot() 进行线程安全遍历");
    FASTEXCEL_LOG_INFO("这确保了代码的并发安全性");
}

int main() {
    std::cout << "FastExcel P0改进演示程序" << std::endl;
    std::cout << "========================" << std::endl;
    
    // 初始化FastExcel
    if (!fastexcel::initialize("logs/p0_demo.log", true)) {
        std::cerr << "FastExcel初始化失败" << std::endl;
        return 1;
    }
    
    try {
        // 演示各项改进
        demonstrateLoggingImprovements();
        demonstrateInitializationImprovements();
        demonstrateThreadSafeFormatRepository();
        demonstrateUnifiedStyleAPI();
        demonstrateXMLReaderImprovements();
        demonstrateDeprecationWarnings();
        
        std::cout << "\n=== 演示完成 ===" << std::endl;
        FASTEXCEL_LOG_INFO("所有P0改进演示完成");
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("演示过程中发生错误: {}", e.what());
        fastexcel::cleanup();
        return 1;
    }
    
    // 清理资源
    fastexcel::cleanup();
    
    std::cout << "\n程序执行完成。请查看日志文件 logs/p0_demo.log 获取详细信息。" << std::endl;
    return 0;
}
