#include "fastexcel/utils/ModuleLoggers.hpp"
/**
 * @file 02_basic_usage.cpp
 * @brief FastExcel基本使用示例
 * 
 * 演示如何使用FastExcel创建Excel文件并写入数据
 */

#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <vector>

int main() {
    try {
        // 初始化FastExcel库
        fastexcel::initialize();
        
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create("basic_example.xlsx");
        if (!workbook) {
            EXAMPLE_ERROR("无法创建工作簿");
            return -1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("数据表");
        if (!worksheet) {
            EXAMPLE_ERROR("无法创建工作表");
            return -1;
        }
        
        // 写入基本数据
        worksheet->writeString(0, 0, "姓名");
        worksheet->writeString(0, 1, "年龄");
        worksheet->writeString(0, 2, "城市");
        worksheet->writeString(0, 3, "薪资");
        
        // 写入数据行
        worksheet->writeString(1, 0, "张三");
        worksheet->writeNumber(1, 1, 25);
        worksheet->writeString(1, 2, "北京");
        worksheet->writeNumber(1, 3, 8000.50);
        
        worksheet->writeString(2, 0, "李四");
        worksheet->writeNumber(2, 1, 30);
        worksheet->writeString(2, 2, "上海");
        worksheet->writeNumber(2, 3, 12000.00);
        
        worksheet->writeString(3, 0, "王五");
        worksheet->writeNumber(3, 1, 28);
        worksheet->writeString(3, 2, "广州");
        worksheet->writeNumber(3, 3, 9500.75);
        
        // 写入公式
        worksheet->writeString(4, 0, "平均薪资");
        worksheet->writeFormula(4, 3, "AVERAGE(D2:D4)");
        
        // 写入布尔值
        worksheet->writeString(5, 0, "数据完整");
        worksheet->writeBoolean(5, 1, true);
        
        // 设置文档属性
        workbook->setTitle("员工信息表");
        workbook->setAuthor("FastExcel示例");
        workbook->setSubject("基本使用演示");
        workbook->setKeywords("Excel, FastExcel, 示例");
        workbook->setComments("这是一个FastExcel基本使用示例");
        
        // 保存文件
        if (!workbook->save()) {
            EXAMPLE_ERROR("保存文件失败");
            return -1;
        }
        
        EXAMPLE_INFO("Excel文件创建成功: basic_example.xlsx");
        
        // 清理资源
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        EXAMPLE_ERROR("发生错误: {}", e.what());
        return -1;
    }
    
    return 0;
}