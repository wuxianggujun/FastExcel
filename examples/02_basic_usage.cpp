#include "fastexcel/utils/ModuleLoggers.hpp"
/**
 * @file 02_basic_usage.cpp
 * @brief FastExcel基本使用示例
 * 
 * 演示如何使用FastExcel创建Excel文件并写入数据
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Path.hpp"
#include <iostream>
#include <vector>

int main() {
    try {
        // 初始化FastExcel库（新API - 需要提供日志路径）
        if (!fastexcel::initialize("logs/basic_example.log", true)) {
            EXAMPLE_ERROR("无法初始化FastExcel库");
            return -1;
        }
        
        // 创建工作簿（新API - 使用Path对象）
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("basic_example.xlsx"));
        if (!workbook) {
            EXAMPLE_ERROR("无法创建工作簿");
            return -1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addSheet("数据表");
        if (!worksheet) {
            EXAMPLE_ERROR("无法创建工作表");
            return -1;
        }
        
        // 写入基本数据（新API - 使用模板化方法）
        worksheet->setValue(0, 0, std::string("姓名"));
        worksheet->setValue(0, 1, std::string("年龄"));
        worksheet->setValue(0, 2, std::string("城市"));
        worksheet->setValue(0, 3, std::string("薪资"));
        
        // 写入数据行（新API - 使用模板化方法）
        worksheet->setValue(1, 0, std::string("张三"));
        worksheet->setValue(1, 1, 25);
        worksheet->setValue(1, 2, std::string("北京"));
        worksheet->setValue(1, 3, 8000.50);
        
        worksheet->setValue(2, 0, std::string("李四"));
        worksheet->setValue(2, 1, 30);
        worksheet->setValue(2, 2, std::string("上海"));
        worksheet->setValue(2, 3, 12000.00);
        
        worksheet->setValue(3, 0, std::string("王五"));
        worksheet->setValue(3, 1, 28);
        worksheet->setValue(3, 2, std::string("广州"));
        worksheet->setValue(3, 3, 9500.75);
        
        // 写入公式（新API - 使用Excel地址格式）
        worksheet->setValue("A5", std::string("平均薪资"));
        worksheet->getCell(4, 3).setFormula("AVERAGE(D2:D4)");
        
        // 写入布尔值（新API - 使用模板化方法）
        worksheet->setValue(5, 0, std::string("数据完整"));
        worksheet->setValue(5, 1, true);
        
        // 设置文档属性（新API - 批量设置）
        workbook->setDocumentProperties(
            "员工信息表",           // title
            "基本使用演示",         // subject
            "FastExcel示例",        // author
            "FastExcel公司",        // company
            "这是一个FastExcel基本使用示例"  // comments
        );
        workbook->setKeywords("Excel, FastExcel, 示例");
        
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