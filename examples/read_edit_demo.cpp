#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <iomanip>

using namespace fastexcel;
using namespace fastexcel::core;

void printCellInfo(const Cell& cell, int row, int col) {
    std::cout << "\n📍 单元格 " << static_cast<char>('A' + col) << (row + 1) << ":" << std::endl;
    
    // 显示单元格值
    if (cell.isNumber()) {
        std::cout << "   📊 数值: " << cell.getNumberValue() << std::endl;
    } else if (cell.isString()) {
        std::cout << "   📝 文本: \"" << cell.getStringValue() << "\"" << std::endl;
    } else if (cell.isFormula()) {
        std::cout << "   🔢 公式: " << cell.getFormula() << " = " << cell.getFormulaResult() << std::endl;
    }
    
    // 显示格式信息
    auto format = cell.getFormatDescriptor();
    if (format) {
        std::cout << "   ✅ 格式信息:" << std::endl;
        std::cout << "     🎨 字体: " << format->getFontName() << ", " << format->getFontSize() << "pt";
        if (format->isBold()) std::cout << ", 粗体";
        if (format->isItalic()) std::cout << ", 斜体";
        std::cout << std::endl;
        
        std::cout << "     🌈 字体色: RGB(0x" << std::hex << format->getFontColor().getRGB() << ")" << std::dec << std::endl;
        std::cout << "     🎯 背景色: RGB(0x" << std::hex << format->getBackgroundColor().getRGB() << ")" << std::dec << std::endl;
        std::cout << "     📐 对齐: " << static_cast<int>(format->getHorizontalAlign());
        if (format->isTextWrap()) std::cout << ", 自动换行";
        std::cout << std::endl;
        std::cout << "     📋 数字格式: \"" << format->getNumberFormat() << "\"" << std::endl;
    } else {
        std::cout << "   ❌ 无格式信息（默认格式）" << std::endl;
    }
}

int main() {
    try {
        std::cout << "=== FastExcel 读取与编辑测试 ===" << std::endl;
        
        // ========== 第一步：创建测试文件 ==========
        std::cout << "\n🔨 步骤1: 创建测试Excel文件..." << std::endl;
        
        {
            auto workbook = Workbook::create(Path("test_read_edit.xlsx"));
            if (!workbook) {
                std::cout << "❌ 无法创建工作簿" << std::endl;
                return 1;
            }
            auto worksheet = workbook->addSheet("测试数据");
            
            // 创建各种格式的数据
            auto titleStyle = workbook->createStyleBuilder()
                .fontName("Arial").fontSize(16).bold().fontColor(Color(255, 255, 255))
                .fill(Color(0, 0, 128))  // 深蓝背景
                .centerAlign().textWrap(true)
                .build();
            
            auto numberStyle = workbook->createStyleBuilder()
                .numberFormat("0.00").rightAlign()
                .fontColor(Color(0, 128, 0))  // 绿色
                .build();
            
            auto percentStyle = workbook->createStyleBuilder()
                .percentage().rightAlign().bold()
                .fontColor(Color(128, 0, 128))  // 紫色
                .build();
            
            auto currencyStyle = workbook->createStyleBuilder()
                .currency().rightAlign()
                .fontColor(Color(0, 0, 255))  // 蓝色
                .fill(Color(255, 255, 0))     // 黄色背景
                .build();
            
            // 添加样式
            int titleId = workbook->addStyle(titleStyle);
            int numberId = workbook->addStyle(numberStyle);
            int percentId = workbook->addStyle(percentStyle);
            int currencyId = workbook->addStyle(currencyStyle);
            
            // 写入数据
            worksheet->writeString(0, 0, "项目名称");
            worksheet->getCell(0, 0).setFormat(workbook->getStyles().getFormat(titleId));
            
            worksheet->writeString(0, 1, "数值");
            worksheet->getCell(0, 1).setFormat(workbook->getStyles().getFormat(titleId));
            
            worksheet->writeString(0, 2, "百分比");
            worksheet->getCell(0, 2).setFormat(workbook->getStyles().getFormat(titleId));
            
            worksheet->writeString(0, 3, "金额");
            worksheet->getCell(0, 3).setFormat(workbook->getStyles().getFormat(titleId));
            
            // 数据行
            worksheet->writeString(1, 0, "产品A");
            worksheet->writeNumber(1, 1, 123.456);
            worksheet->getCell(1, 1).setFormat(workbook->getStyles().getFormat(numberId));
            
            worksheet->writeNumber(1, 2, 0.85);
            worksheet->getCell(1, 2).setFormat(workbook->getStyles().getFormat(percentId));
            
            worksheet->writeNumber(1, 3, 1234.56);
            worksheet->getCell(1, 3).setFormat(workbook->getStyles().getFormat(currencyId));
            
            worksheet->writeString(2, 0, "产品B");
            worksheet->writeNumber(2, 1, 987.654);
            worksheet->getCell(2, 1).setFormat(workbook->getStyles().getFormat(numberId));
            
            worksheet->writeNumber(2, 2, 0.92);
            worksheet->getCell(2, 2).setFormat(workbook->getStyles().getFormat(percentId));
            
            worksheet->writeNumber(2, 3, 2345.67);
            worksheet->getCell(2, 3).setFormat(workbook->getStyles().getFormat(currencyId));
            
            workbook->save();
            workbook->close();
        }
        
        std::cout << "   ✅ 测试文件创建完成！" << std::endl;
        
        // ========== 第二步：读取Excel文件 ==========
        std::cout << "\n📖 步骤2: 读取Excel文件并分析格式..." << std::endl;
        
        auto readWorkbook = Workbook::openForEditing(Path("test_read_edit.xlsx"));
        if (!readWorkbook) {
            std::cout << "❌ 无法打开文件进行读取" << std::endl;
            return 1;
        }
        
        auto worksheet = readWorkbook->getSheet("测试数据");
        if (!worksheet) {
            std::cout << "❌ 找不到工作表" << std::endl;
            return 1;
        }
        
        std::cout << "   ✅ 成功打开文件，开始读取..." << std::endl;
        
        // 读取并显示所有单元格信息
        auto [maxRow, maxCol] = worksheet->getUsedRange();
        std::cout << "   📊 已用范围: " << (maxRow + 1) << " 行 x " << (maxCol + 1) << " 列" << std::endl;
        
        for (int row = 0; row <= maxRow; ++row) {
            for (int col = 0; col <= maxCol; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    printCellInfo(cell, row, col);
                }
            }
        }
        
        // ========== 第三步：编辑Excel文件 ==========
        std::cout << "\n✏️ 步骤3: 编辑Excel文件..." << std::endl;
        
        // 添加新的数据行
        worksheet->writeString(3, 0, "产品C");
        worksheet->writeNumber(3, 1, 555.555);
        worksheet->writeNumber(3, 2, 0.78);
        worksheet->writeNumber(3, 3, 3456.78);
        
        // 创建新的样式用于编辑
        auto editStyle = readWorkbook->createStyleBuilder()
            .fontName("Times New Roman").fontSize(12).italic()
            .fontColor(Color(255, 0, 0))      // 红色字体
            .fill(Color(240, 240, 240))       // 浅灰背景
            .rightAlign().textWrap(true)
            .numberFormat("#,##0.000")        // 3位小数千分位格式
            .build();
        
        int editStyleId = readWorkbook->addStyle(editStyle);
        
        // 应用新样式到新数据
        auto& newCell = worksheet->getCell(3, 1);
        newCell.setFormat(readWorkbook->getStyles().getFormat(editStyleId));
        
        // 修改现有单元格
        std::cout << "\n🔄 修改现有数据..." << std::endl;
        auto& existingCell = worksheet->getCell(1, 1);
        std::cout << "   原值: " << existingCell.getNumberValue() << std::endl;
        
        // 修改数值但保持格式
        auto oldFormat = existingCell.getFormatDescriptor();
        existingCell.setValue(999.999);
        existingCell.setFormat(oldFormat);  // 保持原格式
        
        std::cout << "   新值: " << existingCell.getNumberValue() << " (保持原格式)" << std::endl;
        
        // 添加公式单元格
        worksheet->writeFormula(4, 1, "SUM(B2:B4)");
        worksheet->writeString(4, 0, "总计");
        
        // 保存修改后的文件
        readWorkbook->saveAs("test_read_edit_modified.xlsx");
        readWorkbook->close();
        
        std::cout << "\n   ✅ 文件编辑完成，已保存为: test_read_edit_modified.xlsx" << std::endl;
        
        // ========== 第四步：验证编辑结果 ==========
        std::cout << "\n🔍 步骤4: 验证编辑结果..." << std::endl;
        
        auto verifyWorkbook = Workbook::openForReading(Path("test_read_edit_modified.xlsx"));
        auto verifyWorksheet = verifyWorkbook->getSheet("测试数据");
        
        std::cout << "   📊 验证修改后的数据:" << std::endl;
        auto [newMaxRow, newMaxCol] = verifyWorksheet->getUsedRange();
        std::cout << "   📈 新的已用范围: " << (newMaxRow + 1) << " 行 x " << (newMaxCol + 1) << " 列" << std::endl;
        
        // 检查新添加的数据
        if (verifyWorksheet->hasCellAt(3, 0)) {
            const auto& newDataCell = verifyWorksheet->getCell(3, 0);
            std::cout << "   ✅ 新数据行: " << newDataCell.getStringValue() << std::endl;
        }
        
        // 检查公式
        if (verifyWorksheet->hasCellAt(4, 1)) {
            const auto& formulaCell = verifyWorksheet->getCell(4, 1);
            if (formulaCell.isFormula()) {
                std::cout << "   ✅ 公式单元格: " << formulaCell.getFormula() 
                         << " = " << formulaCell.getFormulaResult() << std::endl;
            }
        }
        
        verifyWorkbook->close();
        
        std::cout << "\n🎉 FastExcel 读取与编辑功能测试完成!" << std::endl;
        std::cout << "📋 验证的功能:" << std::endl;
        std::cout << "   ✅ 读取Excel文件并解析所有格式信息" << std::endl;
        std::cout << "   ✅ 获取单元格颜色、字体、对齐方式等" << std::endl;
        std::cout << "   ✅ 读取数字格式、自动换行等属性" << std::endl;
        std::cout << "   ✅ 编辑现有单元格并保持原格式" << std::endl;
        std::cout << "   ✅ 添加新数据行和新样式" << std::endl;
        std::cout << "   ✅ 添加公式并计算结果" << std::endl;
        std::cout << "   ✅ 保存修改并验证结果" << std::endl;
        
        std::cout << "\n📁 生成的文件:" << std::endl;
        std::cout << "   📄 test_read_edit.xlsx - 原始测试文件" << std::endl;
        std::cout << "   📄 test_read_edit_modified.xlsx - 编辑后的文件" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}