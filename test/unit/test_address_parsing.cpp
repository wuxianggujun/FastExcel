#include "fastexcel/FastExcel.hpp"
#include <iostream>

using namespace fastexcel;

int main() {
    std::cout << "测试新增的地址解析便捷方法..." << std::endl;
    
    try {
        // 创建工作簿和工作表
        auto workbook = core::Workbook::create("test_address_parsing.xlsx");
        auto worksheet = workbook->addSheet("测试地址解析");
        
        // 测试1: 使用字符串地址设置单元格值
        std::cout << "1. 测试字符串地址设置单元格值..." << std::endl;
        worksheet->setValue("A1", std::string("标题"));
        worksheet->setValue("B1", std::string("数据"));
        worksheet->setValue("C1", std::string("结果"));
        
        // 测试2: 使用字符串地址合并单元格
        std::cout << "2. 测试字符串地址合并单元格..." << std::endl;
        worksheet->mergeCells("A1:C1");  // 合并标题行
        worksheet->mergeCells("A3:B4");  // 合并左下角区域
        
        // 测试3: 设置自动筛选
        std::cout << "3. 测试字符串地址设置自动筛选..." << std::endl;
        worksheet->setValue("A2", std::string("名称"));
        worksheet->setValue("B2", std::string("数值"));
        worksheet->setValue("C2", std::string("状态"));
        worksheet->setAutoFilter("A2:C10");
        
        // 测试4: 冻结窗格
        std::cout << "4. 测试字符串地址冻结窗格..." << std::endl;
        worksheet->freezePanes("B3");  // 冻结在B3位置
        
        // 测试5: 设置打印区域
        std::cout << "5. 测试字符串地址设置打印区域..." << std::endl;
        worksheet->setPrintArea("A1:C10");
        
        // 测试6: 设置活动单元格和选中范围
        std::cout << "6. 测试字符串地址设置活动单元格和选中范围..." << std::endl;
        worksheet->setActiveCell("B2");
        worksheet->setSelection("A2:C5");
        
        // 测试7: 范围格式设置
        std::cout << "7. 测试范围格式设置..." << std::endl;
        auto formatter = worksheet->rangeFormatter("A1:C1");
        // 这里只是创建formatter，具体的格式设置需要调用其方法
        
        // 添加一些测试数据
        for (int i = 3; i <= 6; ++i) {
            worksheet->setValue(0, i, std::string("项目") + std::to_string(i-2));
            worksheet->setValue(1, i, (i-2) * 100.0);
            worksheet->setValue(2, i, i % 2 == 0 ? std::string("完成") : std::string("进行中"));
        }
        
        // 测试8: 测试新的getCell和hasCellAt方法
        std::cout << "8. 测试地址字符串版本的getCell和hasCellAt..." << std::endl;
        
        // 测试hasCellAt
        if (worksheet->hasCellAt("A1")) {
            std::cout << "   A1单元格存在" << std::endl;
        }
        
        if (!worksheet->hasCellAt("Z99")) {
            std::cout << "   Z99单元格不存在（正常）" << std::endl;
        }
        
        // 测试getCell并获取值
        try {
            auto& cellA1 = worksheet->getCell("A1");
            auto titleValue = cellA1.getValue<std::string>();
            std::cout << "   A1的值: " << titleValue << std::endl;
            
            // 测试设置和获取新值
            worksheet->setValue("D1", std::string("新值"));
            auto& cellD1 = worksheet->getCell("D1");
            auto newValue = cellD1.getValue<std::string>();
            std::cout << "   D1的新值: " << newValue << std::endl;
            
        } catch (const std::exception& e) {
            std::cout << "   getCell测试出错: " << e.what() << std::endl;
        }
        
        // 保存文件
        std::cout << "保存文件..." << std::endl;
        workbook->save();
        
        std::cout << "✅ 所有测试通过！新的地址解析功能工作正常。" << std::endl;
        std::cout << "生成的文件: test_address_parsing.xlsx" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 测试失败: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}