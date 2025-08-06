/**
 * @file new_architecture_example.cpp
 * @brief 展示重构后FastExcel架构的示例代码
 * 
 * 这个示例展示了新架构的主要特性：
 * - 不可变格式对象
 * - Builder模式创建样式
 * - 线程安全的样式仓储
 * - 类型安全的API
 * - 自动样式去重
 */

#include "src/fastexcel/FastExcelNew.hpp"
#include <iostream>
#include <thread>
#include <vector>

using namespace fastexcel;

void demonstrateBasicUsage() {
    std::cout << "=== 基础用法演示 ===" << std::endl;
    
    // 创建工作簿
    auto workbook = createWorkbook("new_architecture_demo.xlsx");
    
    // 创建样式
    auto headerStyle = createStyle()
        .font("微软雅黑", 14, true)  // 字体：微软雅黑，14pt，粗体
        .fontColor(Color::WHITE)     // 白色字体
        .centerAlign()               // 水平居中
        .vcenterAlign()             // 垂直居中
        .fill(Color::BLUE)          // 蓝色背景
        .border(BorderStyle::Thin)   // 细边框
        .textWrap();                // 自动换行
    
    int headerStyleId = workbook->addStyle(headerStyle);
    
    auto dataStyle = createStyle()
        .font("Calibri", 11)
        .leftAlign()
        .border(BorderStyle::Thin, Color::GRAY);
    
    int dataStyleId = workbook->addStyle(dataStyle);
    
    // 货币样式
    auto moneyStyle = styles::money()
        .border(BorderStyle::Thin, Color::GRAY);
    int moneyStyleId = workbook->addStyle(moneyStyle);
    
    // 添加工作表
    auto sheet = workbook->addWorksheet("销售数据");
    
    // 写入表头
    sheet->writeString(0, 0, "姓名", headerStyleId);
    sheet->writeString(0, 1, "部门", headerStyleId);  
    sheet->writeString(0, 2, "销售额", headerStyleId);
    sheet->writeString(0, 3, "完成率", headerStyleId);
    
    // 写入数据
    std::vector<std::tuple<std::string, std::string, double, double>> data = {
        {"张三", "销售部", 15000.50, 0.85},
        {"李四", "市场部", 22300.75, 0.92}, 
        {"王五", "销售部", 18750.00, 0.78},
        {"赵六", "技术部", 25600.25, 1.05}
    };
    
    for (size_t i = 0; i < data.size(); ++i) {
        size_t row = i + 1;
        sheet->writeString(row, 0, std::get<0>(data[i]), dataStyleId);
        sheet->writeString(row, 1, std::get<1>(data[i]), dataStyleId);
        sheet->writeNumber(row, 2, std::get<2>(data[i]), moneyStyleId);
        sheet->writeNumber(row, 3, std::get<3>(data[i]), dataStyleId);
    }
    
    // 设置列宽
    sheet->setColumnWidth(0, 12);  // 姓名列
    sheet->setColumnWidth(1, 12);  // 部门列
    sheet->setColumnWidth(2, 15);  // 销售额列
    sheet->setColumnWidth(3, 12);  // 完成率列
    
    std::cout << "创建工作簿: " << workbook->getFilename() << std::endl;
    std::cout << "样式数量: " << workbook->getStyleCount() << std::endl;
    
    auto stats = workbook->getStyleStats();
    std::cout << "样式去重率: " << (stats.deduplication_ratio * 100) << "%" << std::endl;
    
    // 保存文件（这里只是演示，实际需要实现）
    // workbook->save();
    
    std::cout << "基础用法演示完成" << std::endl << std::endl;
}

void demonstrateStyleDeduplication() {
    std::cout << "=== 样式去重演示 ===" << std::endl;
    
    auto workbook = createWorkbook("deduplication_demo.xlsx");
    
    // 创建看似不同但实际相同的样式
    auto style1 = createStyle().bold().fontSize(12).fontColor(Color::BLACK);
    auto style2 = createStyle().fontSize(12).bold().fontColor(Color::BLACK);
    auto style3 = createStyle().fontColor(Color::BLACK).bold().fontSize(12);
    
    int id1 = workbook->addStyle(style1);
    int id2 = workbook->addStyle(style2);
    int id3 = workbook->addStyle(style3);
    
    std::cout << "样式ID1: " << id1 << std::endl;
    std::cout << "样式ID2: " << id2 << std::endl;
    std::cout << "样式ID3: " << id3 << std::endl;
    
    if (id1 == id2 && id2 == id3) {
        std::cout << "✅ 样式去重成功！三个相同样式被合并为一个" << std::endl;
    } else {
        std::cout << "❌ 样式去重失败" << std::endl;
    }
    
    auto stats = workbook->getStyleStats();
    std::cout << "总请求数: " << stats.total_requests << std::endl;
    std::cout << "唯一样式数: " << stats.unique_formats << std::endl;
    std::cout << "去重率: " << (stats.deduplication_ratio * 100) << "%" << std::endl;
    std::cout << std::endl;
}

void demonstrateThreadSafety() {
    std::cout << "=== 线程安全演示 ===" << std::endl;
    
    auto workbook = createWorkbook("thread_safe_demo.xlsx");
    
    // 多线程并发添加样式
    const int NUM_THREADS = 4;
    const int STYLES_PER_THREAD = 100;
    
    std::vector<std::thread> threads;
    std::vector<std::vector<int>> thread_style_ids(NUM_THREADS);
    
    // 启动多个线程并发添加样式
    for (int t = 0; t < NUM_THREADS; ++t) {
        threads.emplace_back([&, t]() {
            for (int i = 0; i < STYLES_PER_THREAD; ++i) {
                auto style = createStyle()
                    .fontSize(10 + (i % 10))  // 字体大小10-19
                    .fontColor(Color::fromRGB(0x000000 + i * t))  // 不同颜色
                    .bold(i % 2 == 0);  // 交替粗体
                
                int style_id = workbook->addStyle(style);
                thread_style_ids[t].push_back(style_id);
            }
        });
    }
    
    // 等待所有线程完成
    for (auto& thread : threads) {
        thread.join();
    }
    
    std::cout << "并发添加完成" << std::endl;
    std::cout << "总样式数: " << workbook->getStyleCount() << std::endl;
    
    auto stats = workbook->getStyleStats();
    std::cout << "去重效果: " << stats.total_requests << " -> " << stats.unique_formats 
              << " (去重率: " << (stats.deduplication_ratio * 100) << "%)" << std::endl;
    
    // 验证所有返回的样式ID都有效
    bool all_valid = true;
    for (const auto& ids : thread_style_ids) {
        for (int id : ids) {
            if (!workbook->isValidStyleId(id)) {
                all_valid = false;
                break;
            }
        }
    }
    
    std::cout << (all_valid ? "✅ 所有样式ID有效" : "❌ 存在无效样式ID") << std::endl;
    std::cout << std::endl;
}

void demonstrateStyleTransfer() {
    std::cout << "=== 样式传输演示 ===" << std::endl;
    
    // 创建源工作簿
    auto source_wb = createWorkbook("source.xlsx");
    
    auto style1 = styles::header().fill(Color::RED);
    auto style2 = styles::money().fontColor(Color::GREEN);
    auto style3 = createStyle().border(BorderStyle::Thick).fill(Color::YELLOW);
    
    int src_id1 = source_wb->addStyle(style1);
    int src_id2 = source_wb->addStyle(style2);
    int src_id3 = source_wb->addStyle(style3);
    
    std::cout << "源工作簿样式: " << src_id1 << ", " << src_id2 << ", " << src_id3 << std::endl;
    
    // 创建目标工作簿
    auto target_wb = createWorkbook("target.xlsx");
    
    // 先添加一些样式到目标工作簿
    auto existing_style = createStyle().italic().fontSize(16);
    int existing_id = target_wb->addStyle(existing_style);
    
    // 使用StyleTransferContext进行样式传输
    auto transfer_context = target_wb->copyStylesFrom(*source_wb);
    
    // 映射样式ID
    int target_id1 = transfer_context->mapStyleId(src_id1);
    int target_id2 = transfer_context->mapStyleId(src_id2);
    int target_id3 = transfer_context->mapStyleId(src_id3);
    
    std::cout << "映射后的目标样式: " << target_id1 << ", " << target_id2 << ", " << target_id3 << std::endl;
    
    // 检查样式内容是否一致
    auto src_format1 = source_wb->getStyle(src_id1);
    auto target_format1 = target_wb->getStyle(target_id1);
    
    if (*src_format1 == *target_format1) {
        std::cout << "✅ 样式传输成功，内容一致" << std::endl;
    } else {
        std::cout << "❌ 样式传输失败，内容不一致" << std::endl;
    }
    
    // 显示传输统计
    auto stats = transfer_context->getTransferStats();
    std::cout << "传输统计 - 源: " << stats.source_format_count 
              << ", 目标: " << stats.target_format_count
              << ", 传输: " << stats.transferred_count
              << ", 去重: " << stats.deduplicated_count << std::endl;
    std::cout << std::endl;
}

void demonstrateBuilderPattern() {
    std::cout << "=== Builder模式演示 ===" << std::endl;
    
    auto workbook = createWorkbook("builder_demo.xlsx");
    
    // 链式调用创建复杂样式
    auto complex_style = createStyle()
        .font("Arial", 14, true)                    // 字体设置
        .fontColor(Color::fromRGB(0x2E4057))       // 字体颜色
        .italic()                                   // 斜体
        .underline()                               // 下划线
        .centerAlign()                             // 水平居中
        .vcenterAlign()                           // 垂直居中
        .textWrap()                               // 自动换行  
        .fill(PatternType::LightGray, Color::fromRGB(0xF8F9FA))  // 浅灰填充
        .border(BorderStyle::Medium, Color::fromRGB(0x495057))   // 中等边框
        .numberFormat("#,##0.00")                 // 数字格式
        .rotation(45)                             // 旋转45度
        .indent(2);                               // 缩进2级
    
    int complex_id = workbook->addStyle(complex_style);
    
    // 基于现有样式创建变体
    auto variant_style = ui::StyleBuilder(*workbook->getStyle(complex_id))
        .fontColor(Color::RED)                    // 改变字体颜色
        .rotation(0);                             // 取消旋转
    
    int variant_id = workbook->addStyle(variant_style);
    
    std::cout << "复杂样式ID: " << complex_id << std::endl;
    std::cout << "变体样式ID: " << variant_id << std::endl;
    
    // 预定义样式快速创建
    int title_id = workbook->addStyle(styles::title().fill(Color::BLUE));
    int money_id = workbook->addStyle(styles::money().fontColor(Color::GREEN));
    int date_id = workbook->addStyle(styles::date().border(BorderStyle::Thin));
    
    std::cout << "预定义样式 - 标题: " << title_id 
              << ", 货币: " << money_id 
              << ", 日期: " << date_id << std::endl;
    
    std::cout << "Builder模式演示完成" << std::endl << std::endl;
}

int main() {
    std::cout << "FastExcel 新架构演示" << std::endl;
    std::cout << "版本: " << getVersion() << std::endl;
    std::cout << "========================================" << std::endl << std::endl;
    
    try {
        // 初始化库（在实际实现中）
        // initialize(Logger::Level::INFO, "fastexcel.log");
        
        // 运行各种演示
        demonstrateBasicUsage();
        demonstrateStyleDeduplication();
        demonstrateThreadSafety();
        demonstrateStyleTransfer(); 
        demonstrateBuilderPattern();
        
        std::cout << "========================================" << std::endl;
        std::cout << "所有演示完成！新架构的主要优势：" << std::endl;
        std::cout << "1. ✅ 线程安全 - 支持多线程并发操作" << std::endl;
        std::cout << "2. ✅ 自动去重 - 相同样式自动合并" << std::endl;
        std::cout << "3. ✅ 不可变性 - 样式对象创建后不可修改" << std::endl;
        std::cout << "4. ✅ 类型安全 - 编译期类型检查" << std::endl;
        std::cout << "5. ✅ 职责分离 - 清晰的分层架构" << std::endl;
        std::cout << "6. ✅ 易于扩展 - 支持多种输出格式" << std::endl;
        std::cout << "7. ✅ 样式传输 - 跨工作簿样式复制" << std::endl;
        std::cout << "8. ✅ 流畅API - Builder模式的链式调用" << std::endl;
        
        // cleanup();
    }
    catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}