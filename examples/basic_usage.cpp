#include "fastexcel/FastExcel.h"
#include <iostream>

int main() {
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/basic_usage.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    LOG_INFO("FastExcel basic usage example started");
    
    try {
        // 创建工作簿
        auto workbook = std::make_shared<fastexcel::core::Workbook>("example_basic.xlsx");
        
        // 打开工作簿
        if (!workbook->open()) {
            LOG_ERROR("Failed to open workbook");
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("Sheet1");
        if (!worksheet) {
            LOG_ERROR("Failed to create worksheet");
            return 1;
        }
        
        // 写入表头
        worksheet->writeString(0, 0, "姓名");
        worksheet->writeString(0, 1, "年龄");
        worksheet->writeString(0, 2, "城市");
        worksheet->writeString(0, 3, "职业");
        
        // 写入数据
        worksheet->writeString(1, 0, "张三");
        worksheet->writeNumber(1, 1, 25);
        worksheet->writeString(1, 2, "北京");
        worksheet->writeString(1, 3, "工程师");
        
        worksheet->writeString(2, 0, "李四");
        worksheet->writeNumber(2, 1, 30);
        worksheet->writeString(2, 2, "上海");
        worksheet->writeString(2, 3, "设计师");
        
        worksheet->writeString(3, 0, "王五");
        worksheet->writeNumber(3, 1, 28);
        worksheet->writeString(3, 2, "广州");
        worksheet->writeString(3, 3, "产品经理");
        
        // 创建第二个工作表
        auto worksheet2 = workbook->addWorksheet("数据统计");
        
        // 写入统计信息
        worksheet2->writeString(0, 0, "统计项");
        worksheet2->writeString(0, 1, "数值");
        
        worksheet2->writeString(1, 0, "总人数");
        worksheet2->writeNumber(1, 1, 3);
        
        worksheet2->writeString(2, 0, "平均年龄");
        worksheet2->writeNumber(2, 1, 27.67);
        
        // 保存工作簿
        if (!workbook->save()) {
            LOG_ERROR("Failed to save workbook");
            return 1;
        }
        
        // 关闭工作簿
        workbook->close();
        
        LOG_INFO("Excel文件创建成功: example_basic.xlsx");
        std::cout << "Excel文件创建成功: example_basic.xlsx" << std::endl;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception occurred: {}", e.what());
        std::cerr << "Exception occurred: " << e.what() << std::endl;
        return 1;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
    
    LOG_INFO("FastExcel basic usage example completed");
    std::cout << "示例程序执行完成，请查看生成的Excel文件和日志文件。" << std::endl;
    
    return 0;
}