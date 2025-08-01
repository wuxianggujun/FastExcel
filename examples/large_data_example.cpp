/**
 * @file large_data_example.cpp
 * @brief FastExcel大数据处理示例
 * 
 * 演示如何使用FastExcel高效处理大量数据
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <vector>
#include <random>
#include <chrono>
#include <iomanip>

// 模拟数据结构
struct SalesRecord {
    std::string product_name;
    std::string region;
    std::string salesperson;
    double quantity;
    double unit_price;
    double total_amount;
    std::string date;
    
    SalesRecord(const std::string& product, const std::string& reg, const std::string& sales,
                double qty, double price, const std::string& dt)
        : product_name(product), region(reg), salesperson(sales), 
          quantity(qty), unit_price(price), total_amount(qty * price), date(dt) {}
};

// 生成模拟数据
std::vector<SalesRecord> generateSalesData(size_t count) {
    std::vector<SalesRecord> data;
    data.reserve(count);
    
    // 随机数生成器
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> product_dist(0, 4);
    std::uniform_int_distribution<> region_dist(0, 3);
    std::uniform_int_distribution<> sales_dist(0, 9);
    std::uniform_real_distribution<> qty_dist(1.0, 1000.0);
    std::uniform_real_distribution<> price_dist(10.0, 5000.0);
    std::uniform_int_distribution<> day_dist(1, 28);
    std::uniform_int_distribution<> month_dist(1, 12);
    
    // 预定义数据
    std::vector<std::string> products = {"笔记本电脑", "智能手机", "平板电脑", "智能手表", "耳机"};
    std::vector<std::string> regions = {"华北", "华东", "华南", "西南"};
    std::vector<std::string> salespeople = {"张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十", "郑一", "王二"};
    
    for (size_t i = 0; i < count; ++i) {
        std::string product = products[product_dist(gen)];
        std::string region = regions[region_dist(gen)];
        std::string salesperson = salespeople[sales_dist(gen)];
        double quantity = std::round(qty_dist(gen) * 100) / 100;
        double unit_price = std::round(price_dist(gen) * 100) / 100;
        
        // 生成日期
        int month = month_dist(gen);
        int day = day_dist(gen);
        std::string date = "2024-" + (month < 10 ? "0" : "") + std::to_string(month) + 
                          "-" + (day < 10 ? "0" : "") + std::to_string(day);
        
        data.emplace_back(product, region, salesperson, quantity, unit_price, date);
    }
    
    return data;
}

int main() {
    try {
        // 初始化FastExcel库
        fastexcel::initialize();
        
        std::cout << "开始生成大数据Excel文件..." << std::endl;
        
        // 记录开始时间
        auto start_time = std::chrono::high_resolution_clock::now();
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("large_data_example.xlsx");
        
        // 设置常量内存模式以优化大数据处理
        workbook->setConstantMemory(true);
        
        if (!workbook->open()) {
            std::cerr << "无法打开工作簿" << std::endl;
            return -1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("销售数据");
        
        // ========== 创建格式 ==========
        
        // 表头格式
        auto header_format = workbook->createFormat();
        header_format->setBold(true);
        header_format->setBackgroundColor(fastexcel::core::COLOR_BLUE);
        header_format->setFontColor(fastexcel::core::COLOR_WHITE);
        header_format->setHorizontalAlign(fastexcel::core::HorizontalAlign::Center);
        header_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        // 数字格式
        auto number_format = workbook->createFormat();
        number_format->setNumberFormat("#,##0.00");
        
        // 货币格式
        auto currency_format = workbook->createFormat();
        currency_format->setNumberFormat("¥#,##0.00");
        
        // 文本格式
        auto text_format = workbook->createFormat();
        text_format->setHorizontalAlign(fastexcel::core::HorizontalAlign::Left);
        
        // ========== 写入表头 ==========
        
        std::vector<std::string> headers = {
            "产品名称", "销售区域", "销售员", "销售数量", "单价", "总金额", "销售日期"
        };
        
        for (size_t col = 0; col < headers.size(); ++col) {
            worksheet->writeString(0, static_cast<int>(col), headers[col], header_format);
        }
        
        // 设置列宽
        worksheet->setColumnWidth(0, 15);  // 产品名称
        worksheet->setColumnWidth(1, 10);  // 销售区域
        worksheet->setColumnWidth(2, 10);  // 销售员
        worksheet->setColumnWidth(3, 12);  // 销售数量
        worksheet->setColumnWidth(4, 12);  // 单价
        worksheet->setColumnWidth(5, 15);  // 总金额
        worksheet->setColumnWidth(6, 12);  // 销售日期
        
        // ========== 生成并写入大量数据 ==========
        
        const size_t TOTAL_RECORDS = 50000; // 5万条记录
        const size_t BATCH_SIZE = 1000;     // 批量处理大小
        
        std::cout << "生成 " << TOTAL_RECORDS << " 条销售记录..." << std::endl;
        
        size_t processed = 0;
        for (size_t batch_start = 0; batch_start < TOTAL_RECORDS; batch_start += BATCH_SIZE) {
            size_t batch_end = std::min(batch_start + BATCH_SIZE, TOTAL_RECORDS);
            size_t batch_size = batch_end - batch_start;
            
            // 生成批量数据
            auto batch_data = generateSalesData(batch_size);
            
            // 批量写入数据
            for (size_t i = 0; i < batch_data.size(); ++i) {
                int row = static_cast<int>(batch_start + i + 1); // +1 跳过表头
                const auto& record = batch_data[i];
                
                worksheet->writeString(row, 0, record.product_name, text_format);
                worksheet->writeString(row, 1, record.region, text_format);
                worksheet->writeString(row, 2, record.salesperson, text_format);
                worksheet->writeNumber(row, 3, record.quantity, number_format);
                worksheet->writeNumber(row, 4, record.unit_price, currency_format);
                worksheet->writeNumber(row, 5, record.total_amount, currency_format);
                worksheet->writeString(row, 6, record.date, text_format);
            }
            
            processed += batch_size;
            
            // 显示进度
            if (processed % 10000 == 0 || processed == TOTAL_RECORDS) {
                double progress = static_cast<double>(processed) / TOTAL_RECORDS * 100;
                std::cout << "进度: " << std::fixed << std::setprecision(1) 
                         << progress << "% (" << processed << "/" << TOTAL_RECORDS << ")" << std::endl;
            }
        }
        
        // ========== 添加汇总信息 ==========
        
        int summary_row = static_cast<int>(TOTAL_RECORDS + 2);
        
        // 汇总标题
        auto summary_format = workbook->createFormat();
        summary_format->setBold(true);
        summary_format->setBackgroundColor(fastexcel::core::COLOR_YELLOW);
        summary_format->setBorder(fastexcel::core::BorderStyle::Thin);
        
        worksheet->writeString(summary_row, 0, "汇总统计", summary_format);
        worksheet->writeString(summary_row + 1, 0, "总记录数:", text_format);
        worksheet->writeNumber(summary_row + 1, 1, static_cast<double>(TOTAL_RECORDS), number_format);
        
        worksheet->writeString(summary_row + 2, 0, "总销售额:", text_format);
        std::string total_formula = "SUM(F2:F" + std::to_string(TOTAL_RECORDS + 1) + ")";
        worksheet->writeFormula(summary_row + 2, 1, total_formula, currency_format);
        
        worksheet->writeString(summary_row + 3, 0, "平均单价:", text_format);
        std::string avg_formula = "AVERAGE(E2:E" + std::to_string(TOTAL_RECORDS + 1) + ")";
        worksheet->writeFormula(summary_row + 3, 1, avg_formula, currency_format);
        
        // ========== 设置工作表选项 ==========
        
        // 冻结表头
        worksheet->freezePanes(1, 0);
        
        // 自动筛选
        worksheet->setAutoFilter(0, 0, static_cast<int>(TOTAL_RECORDS), 6);
        
        // 设置打印选项
        worksheet->setPrintGridlines(true);
        worksheet->setLandscape(true);
        worksheet->setFitToPages(1, 0); // 适合一页宽度
        
        // ========== 设置文档属性 ==========
        
        workbook->setTitle("大数据销售报表");
        workbook->setAuthor("FastExcel大数据示例");
        workbook->setSubject("性能测试");
        workbook->setKeywords("Excel, 大数据, 性能, FastExcel");
        workbook->setComments("包含" + std::to_string(TOTAL_RECORDS) + "条销售记录的大数据报表");
        
        // 添加自定义属性
        workbook->setCustomProperty("记录数量", static_cast<double>(TOTAL_RECORDS));
        workbook->setCustomProperty("生成工具", "FastExcel");
        workbook->setCustomProperty("数据类型", "销售数据");
        
        std::cout << "开始保存文件..." << std::endl;
        
        // 保存文件
        if (!workbook->save()) {
            std::cerr << "保存文件失败" << std::endl;
            return -1;
        }
        
        // 计算耗时
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
        
        std::cout << "大数据Excel文件创建成功: large_data_example.xlsx" << std::endl;
        std::cout << "总记录数: " << TOTAL_RECORDS << std::endl;
        std::cout << "总耗时: " << duration.count() << " 毫秒" << std::endl;
        std::cout << "平均速度: " << std::fixed << std::setprecision(0) 
                 << (static_cast<double>(TOTAL_RECORDS) / duration.count() * 1000) 
                 << " 记录/秒" << std::endl;
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "发生错误: " << e.what() << std::endl;
        return -1;
    }
    
    return 0;
}