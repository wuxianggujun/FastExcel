#include "fastexcel/FastExcel.h"
#include <iostream>

int main() {
    // 初始化FastExcel库
    if (!fastexcel::initialize("logs/formatting.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }
    
    LOG_INFO("FastExcel formatting example started");
    
    try {
        // 创建工作簿
        auto workbook = std::make_shared<fastexcel::core::Workbook>("example_formatting.xlsx");
        
        // 打开工作簿
        if (!workbook->open()) {
            LOG_ERROR("Failed to open workbook");
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addWorksheet("格式化示例");
        if (!worksheet) {
            LOG_ERROR("Failed to create worksheet");
            return 1;
        }
        
        // 创建格式
        auto header_format = workbook->createFormat();
        header_format->setBold(true);
        header_format->setFontSize(12);
        header_format->setFontColor(0xFFFFFF); // 白色
        header_format->setBackgroundColor(0x4F81BD); // 蓝色
        header_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Center);
        header_format->setVerticalAlignment(fastexcel::core::VerticalAlignment::Center);
        
        auto number_format = workbook->createFormat();
        number_format->setNumberFormat("#,##0.00");
        number_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Right);
        
        auto date_format = workbook->createFormat();
        date_format->setNumberFormat("yyyy-mm-dd");
        date_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Center);
        
        auto currency_format = workbook->createFormat();
        currency_format->setNumberFormat("\"¥\"#,##0.00");
        currency_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Right);
        currency_format->setFontColor(0x00B050); // 绿色
        
        auto border_format = workbook->createFormat();
        border_format->setBorderStyle("thin", 0x000000); // 黑色细边框
        border_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Center);
        border_format->setVerticalAlignment(fastexcel::core::VerticalAlignment::Center);
        
        // 写入表头
        worksheet->writeString(0, 0, "产品名称", header_format);
        worksheet->writeString(0, 1, "单价", header_format);
        worksheet->writeString(0, 2, "数量", header_format);
        worksheet->writeString(0, 3, "总价", header_format);
        worksheet->writeString(0, 4, "日期", header_format);
        
        // 写入数据
        worksheet->writeString(1, 0, "笔记本电脑");
        worksheet->writeNumber(1, 1, 5999.00, currency_format);
        worksheet->writeNumber(1, 2, 2, number_format);
        worksheet->writeNumber(1, 3, 11998.00, currency_format);
        worksheet->writeString(1, 4, "2023-10-01", date_format);
        
        worksheet->writeString(2, 0, "智能手机");
        worksheet->writeNumber(2, 1, 3999.00, currency_format);
        worksheet->writeNumber(2, 2, 5, number_format);
        worksheet->writeNumber(2, 3, 19995.00, currency_format);
        worksheet->writeString(2, 4, "2023-10-02", date_format);
        
        worksheet->writeString(3, 0, "平板电脑");
        worksheet->writeNumber(3, 1, 2999.00, currency_format);
        worksheet->writeNumber(3, 2, 3, number_format);
        worksheet->writeNumber(3, 3, 8997.00, currency_format);
        worksheet->writeString(3, 4, "2023-10-03", date_format);
        
        // 添加汇总行
        auto total_format = workbook->createFormat();
        total_format->setBold(true);
        total_format->setBackgroundColor(0xFFC000); // 黄色
        total_format->setBorderStyle("medium", 0x000000); // 黑色中等边框
        total_format->setHorizontalAlignment(fastexcel::core::HorizontalAlignment::Right);
        
        worksheet->writeString(4, 0, "总计", total_format);
        worksheet->writeNumber(4, 3, 40990.00, currency_format);
        
        // 创建第二个工作表，展示更多格式选项
        auto worksheet2 = workbook->addWorksheet("更多格式");
        
        // 创建更多格式
        auto bold_format = workbook->createFormat();
        bold_format->setBold(true);
        
        auto italic_format = workbook->createFormat();
        italic_format->setItalic(true);
        
        auto underline_format = workbook->createFormat();
        underline_format->setUnderline(fastexcel::core::FontUnderline::Single);
        
        auto wrap_format = workbook->createFormat();
        wrap_format->setWrapText(true);
        
        // 写入示例数据
        worksheet2->writeString(0, 0, "格式类型", bold_format);
        worksheet2->writeString(0, 1, "示例", bold_format);
        
        worksheet2->writeString(1, 0, "粗体");
        worksheet2->writeString(1, 1, "这是粗体文本", bold_format);
        
        worksheet2->writeString(2, 0, "斜体");
        worksheet2->writeString(2, 1, "这是斜体文本", italic_format);
        
        worksheet2->writeString(3, 0, "下划线");
        worksheet2->writeString(3, 1, "这是带下划线的文本", underline_format);
        
        worksheet2->writeString(4, 0, "自动换行");
        worksheet2->writeString(4, 1, "这是一段很长的文本，它会自动换行显示，以便适应单元格的宽度。", wrap_format);
        
        // 调整行高以适应自动换行
        // 注意：在实际实现中，可能需要添加设置行高的功能
        
        // 保存工作簿
        if (!workbook->save()) {
            LOG_ERROR("Failed to save workbook");
            return 1;
        }
        
        // 关闭工作簿
        workbook->close();
        
        LOG_INFO("Excel文件创建成功: example_formatting.xlsx");
        std::cout << "Excel文件创建成功: example_formatting.xlsx" << std::endl;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception occurred: {}", e.what());
        std::cerr << "Exception occurred: " << e.what() << std::endl;
        return 1;
    }
    
    // 清理FastExcel库
    fastexcel::cleanup();
    
    LOG_INFO("FastExcel formatting example completed");
    std::cout << "格式化示例程序执行完成，请查看生成的Excel文件和日志文件。" << std::endl;
    
    return 0;
}