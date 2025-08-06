/**
 * @file zip_compatibility_demo.cpp
 * @brief ZIP兼容性验证示例程序
 * 
 * 验证生成的XLSX文件与Excel的兼容性
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include <iostream>
#include <filesystem>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    std::cout << "ZIP兼容性验证程序" << std::endl;
    std::cout << "==============================" << std::endl;
    
    try {
        // 初始化FastExcel
        void (*init_func)() = fastexcel::initialize;
        init_func();
        
        std::string filename = "zip_compatibility_test.xlsx";
        
        // 创建工作簿
        auto workbook = Workbook::create(filename);
        if (!workbook->open()) {
            throw std::runtime_error("无法打开工作簿");
        }
        
        std::cout << "✓ 工作簿创建成功" << std::endl;
        
        // 添加工作表并写入测试数据
        auto worksheet = workbook->addWorksheet("CompatibilityTest");
        
        // 写入各种类型的数据
        worksheet->writeString(0, 0, "ZIP兼容性测试");
        worksheet->writeNumber(0, 1, 2025);
        worksheet->writeString(1, 0, "当前时间");
        
        // 使用封装的TimeUtils获取时间
        auto current_time = utils::TimeUtils::getCurrentTime();
        std::string time_str = utils::TimeUtils::formatTime(current_time, "%Y-%m-%d %H:%M:%S");
        worksheet->writeString(1, 1, time_str);
        
        // 添加格式化数据
        auto bold_format = workbook->createFormat();
        bold_format->setBold(true);
        bold_format->setFontSize(14);
        worksheet->writeString(3, 0, "粗体文字", bold_format);
        
        auto colored_format = workbook->createFormat();
        colored_format->setFontColor(Color::blue());
        colored_format->setItalic(true);
        worksheet->writeString(4, 0, "蓝色斜体", colored_format);
        
        // 写入数字和计算数据
        worksheet->writeNumber(6, 0, 123.456);
        worksheet->writeNumber(6, 1, 789.012);
        worksheet->writeFormula(6, 2, "A7+B7"); // Excel style (1-based)
        
        // 设置文档属性
        workbook->setTitle("ZIP兼容性测试文档");
        workbook->setAuthor("FastExcel");
        workbook->setSubject("验证XLSX文件格式兼容性");
        workbook->setComments("此文件用于验证与Microsoft Excel的兼容性");
        
        std::cout << "✓ 测试数据写入完成" << std::endl;
        
        // 保存文件
        if (!workbook->save()) {
            throw std::runtime_error("保存文件失败");
        }
        
        workbook->close();
        std::cout << "✓ 文件保存成功: " << filename << std::endl;
        
        // 验证文件存在和大小
        if (std::filesystem::exists(filename)) {
            auto file_size = std::filesystem::file_size(filename);
            std::cout << "文件大小: " << file_size << " 字节" << std::endl;
            
            if (file_size > 0) {
                std::cout << "✅ ZIP兼容性验证成功！" << std::endl;
                std::cout << "请用Microsoft Excel打开 '" << filename << "' 验证兼容性。" << std::endl;
            } else {
                std::cout << "❌ 文件大小为0，可能存在问题" << std::endl;
            }
        } else {
            std::cout << "❌ 文件未创建" << std::endl;
        }
        
        // 清理
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 错误: " << e.what() << std::endl;
        fastexcel::cleanup();
        return 1;
    }
    
    return 0;
}