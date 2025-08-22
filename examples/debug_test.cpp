#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/Path.hpp"
#include <iostream>

using namespace fastexcel;

int main() {
    try {
        // 初始化FastExcel库
        if (!fastexcel::initialize("debug.log", true)) {
            std::cerr << "无法初始化FastExcel库" << std::endl;
            return -1;
        }
        
        std::cout << "创建工作簿..." << std::endl;
        
        // 创建新工作簿
        auto workbook = core::Workbook::create(core::Path("debug_sample.xlsx"));
        if (!workbook) {
            std::cerr << "无法创建工作簿" << std::endl;
            return -1;
        }
        
        std::cout << "添加工作表..." << std::endl;
        auto worksheet = workbook->addSheet("测试数据");
        
        std::cout << "写入数据..." << std::endl;
        // 写入简单数据
        worksheet->setValue(0, 0, std::string("名称"));
        worksheet->setValue(0, 1, std::string("数值"));
        worksheet->setValue(1, 0, std::string("测试"));
        worksheet->setValue(1, 1, 123.0);
        
        std::cout << "保存文件..." << std::endl;
        if (workbook->save()) {
            std::cout << "✓ 成功保存: debug_sample.xlsx" << std::endl;
        } else {
            std::cout << "保存失败" << std::endl;
        }
        
        workbook->close();
        
        std::cout << "尝试读取文件..." << std::endl;
        
        // 尝试读取
        auto read_workbook = core::Workbook::openForReading(core::Path("debug_sample.xlsx"));
        if (!read_workbook) {
            std::cerr << "无法读取文件" << std::endl;
            return -1;
        }
        
        std::cout << "成功读取文件" << std::endl;
        read_workbook->close();
        
        // 清理
        fastexcel::cleanup();
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return -1;
    }
    
    return 0;
}