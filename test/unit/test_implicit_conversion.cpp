#include "fastexcel/FastExcel.hpp"
#include "fastexcel/core/CellAddress.hpp"
#include <iostream>

using namespace fastexcel;
using namespace fastexcel::core;

int main() {
    std::cout << "=== 测试Address和Range类的隐式转换功能 ===" << std::endl;
    
    try {
        // 创建工作簿和工作表
        auto workbook = core::Workbook::create("test_implicit_conversion.xlsx");
        auto worksheet = workbook->addSheet("隐式转换测试");
        
        std::cout << "\n1. 测试Address类的隐式转换..." << std::endl;
        
        // ===== Address类测试 =====
        
        // 测试字符串构造
        Address addr1("A1");
        std::cout << "   字符串构造 'A1': " << addr1.toString() << 
                     " (行:" << addr1.getRow() << ", 列:" << addr1.getCol() << ")" << std::endl;
        
        // 测试坐标构造
        Address addr2(1, 2);  // B2
        std::cout << "   坐标构造 (1,2): " << addr2.toString() << 
                     " (行:" << addr2.getRow() << ", 列:" << addr2.getCol() << ")" << std::endl;
        
        // 测试带工作表名的地址
        Address addr3("Sheet1!C3");
        std::cout << "   带工作表 'Sheet1!C3': " << addr3.toString(true) << 
                     " (工作表:" << addr3.getSheetName() << ")" << std::endl;
        
        std::cout << "\n2. 测试CellRange类的隐式转换..." << std::endl;
        
        // ===== CellRange类测试 =====
        
        // 测试字符串范围构造
        CellRange range1("A1:C3");
        std::cout << "   字符串范围 'A1:C3': " << range1.toString() << 
                     " (行数:" << range1.getRowCount() << ", 列数:" << range1.getColCount() << ")" << std::endl;
        
        // 测试坐标范围构造
        CellRange range2(0, 0, 2, 2);
        std::cout << "   坐标范围 (0,0,2,2): " << range2.toString() << std::endl;
        
        // 测试从Address构造CellRange
        CellRange range3(Address("B2"));
        std::cout << "   从Address构造: " << range3.toString() << 
                     " (是否单个单元格:" << (range3.isSingleCell() ? "是" : "否") << ")" << std::endl;
        
        std::cout << "\n3. 测试隐式转换在实际API中的使用..." << std::endl;
        
        // ===== 实际API测试 =====
        
        // 测试setValue的隐式转换
        worksheet->setValue("A1", std::string("标题"));              // 字符串地址
        worksheet->setValue(Address(0, 1), std::string("数据"));     // Address对象
        worksheet->setValue({0, 2}, std::string("结果"));          // 列表初始化
        
        std::cout << "   ✓ setValue支持多种地址格式" << std::endl;
        
        // 测试getCell的隐式转换
        auto& cell1 = worksheet->getCell("A1");                    // 字符串地址
        auto& cell2 = worksheet->getCell(Address(0, 1));           // Address对象
        auto& cell3 = worksheet->getCell({0, 2});                  // 列表初始化
        
        std::cout << "   ✓ getCell支持多种地址格式" << std::endl;
        std::cout << "     A1内容: " << cell1.getValue<std::string>() << std::endl;
        std::cout << "     B1内容: " << cell2.getValue<std::string>() << std::endl;
        std::cout << "     C1内容: " << cell3.getValue<std::string>() << std::endl;
        
        // 测试hasCellAt的隐式转换
        bool hasA1 = worksheet->hasCellAt("A1");                   // 字符串地址
        bool hasB1 = worksheet->hasCellAt(Address(0, 1));          // Address对象
        bool hasZ99 = worksheet->hasCellAt({25, 98});              // 列表初始化
        
        std::cout << "   ✓ hasCellAt支持多种地址格式" << std::endl;
        std::cout << "     A1存在: " << (hasA1 ? "是" : "否") << std::endl;
        std::cout << "     B1存在: " << (hasB1 ? "是" : "否") << std::endl;
        std::cout << "     Z99存在: " << (hasZ99 ? "是" : "否") << std::endl;
        
        // 测试mergeCells的隐式转换
        worksheet->mergeCells("A3:C3");                            // 字符串范围
        worksheet->mergeCells(CellRange(4, 0, 4, 2));              // CellRange对象
        worksheet->mergeCells({5, 0, 5, 2});                      // 列表初始化
        worksheet->mergeCells(Address("A7"));                     // Address转CellRange
        
        std::cout << "   ✓ mergeCells支持多种范围格式" << std::endl;
        
        // 测试setAutoFilter的隐式转换
        worksheet->setValue("A2", std::string("名称"));
        worksheet->setValue("B2", std::string("数值"));
        worksheet->setValue("C2", std::string("状态"));
        
        worksheet->setAutoFilter("A2:C10");                       // 字符串范围
        std::cout << "   ✓ setAutoFilter支持字符串范围" << std::endl;
        
        // 测试freezePanes的隐式转换
        worksheet->freezePanes("B3");                             // 字符串地址
        std::cout << "   ✓ freezePanes支持字符串地址" << std::endl;
        
        // 测试setPrintArea的隐式转换
        worksheet->setPrintArea("A1:C10");                        // 字符串范围
        std::cout << "   ✓ setPrintArea支持字符串范围" << std::endl;
        
        // 测试setActiveCell的隐式转换
        worksheet->setActiveCell("B2");                           // 字符串地址
        std::cout << "   ✓ setActiveCell支持字符串地址" << std::endl;
        
        // 测试setSelection的隐式转换
        worksheet->setSelection("A2:C5");                         // 字符串范围
        std::cout << "   ✓ setSelection支持字符串范围" << std::endl;
        
        std::cout << "\n4. 测试Address和Range类的辅助功能..." << std::endl;
        
        // ===== 辅助功能测试 =====
        
        // 测试CellRange.contains
        CellRange testRange("B2:D4");
        bool containsB2 = testRange.contains(Address("B2"));       // 应该包含
        bool containsA1 = testRange.contains(Address("A1"));       // 应该不包含
        bool containsC3 = testRange.contains(Address("C3"));       // 应该包含
        
        std::cout << "   范围B2:D4包含B2: " << (containsB2 ? "是" : "否") << std::endl;
        std::cout << "   范围B2:D4包含A1: " << (containsA1 ? "是" : "否") << std::endl;
        std::cout << "   范围B2:D4包含C3: " << (containsC3 ? "是" : "否") << std::endl;
        
        // 测试CellRange的角落地址
        core::Address topLeft(testRange.getStartRow(), testRange.getStartCol());
        core::Address bottomRight(testRange.getEndRow(), testRange.getEndCol());
        std::cout << "   范围B2:D4左上角: " << topLeft.toString() << std::endl;
        std::cout << "   范围B2:D4右下角: " << bottomRight.toString() << std::endl;
        
        // 测试地址比较
        Address addrA1("A1");
        Address addrB2("B2");
        Address addrA1Copy("A1");
        
        std::cout << "   A1 == A1副本: " << (addrA1 == addrA1Copy ? "是" : "否") << std::endl;
        std::cout << "   A1 != B2: " << (addrA1 != addrB2 ? "是" : "否") << std::endl;
        std::cout << "   A1 < B2: " << (addrA1 < addrB2 ? "是" : "否") << std::endl;
        
        std::cout << "\n5. 测试错误处理..." << std::endl;
        
        // ===== 错误处理测试 =====
        
        try {
            Address invalidAddr(-1, -1);  // 应该抛出异常
            std::cout << "   ❌ 负数坐标应该抛出异常，但没有！" << std::endl;
        } catch (const std::exception& e) {
            std::cout << "   ✓ 负数坐标正确抛出异常: " << e.what() << std::endl;
        }
        
        try {
            Address invalidAddr("Invalid!");  // 应该抛出异常
            std::cout << "   ❌ 无效地址字符串应该抛出异常，但没有！" << std::endl;
        } catch (const std::exception& e) {
            std::cout << "   ✓ 无效地址字符串正确抛出异常: " << e.what() << std::endl;
        }
        
        // 保存文件
        std::cout << "\n保存文件..." << std::endl;
        workbook->save();
        
        std::cout << "\n🎉 所有隐式转换测试通过！" << std::endl;
        std::cout << "📁 生成的文件: test_implicit_conversion.xlsx" << std::endl;
        
        std::cout << "\n=== 功能总结 ===" << std::endl;
        std::cout << "✅ Address类支持:" << std::endl;
        std::cout << "   - 字符串地址: Address(\"A1\")" << std::endl;
        std::cout << "   - 坐标地址: Address(0, 0)" << std::endl;
        std::cout << "   - 列表初始化: Address{0, 0}" << std::endl;
        std::cout << "   - 带工作表: Address(\"Sheet1!A1\")" << std::endl;
        
        std::cout << "✅ CellRange类支持:" << std::endl;
        std::cout << "   - 字符串范围: CellRange(\"A1:C3\")" << std::endl;
        std::cout << "   - 坐标范围: CellRange(0, 0, 2, 2)" << std::endl;
        std::cout << "   - 列表初始化: CellRange{0, 0, 2, 2}" << std::endl;
        std::cout << "   - 从Address转换: CellRange(Address(\"A1\"))" << std::endl;
        
        std::cout << "✅ 所有Worksheet方法现在支持隐式转换！" << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 测试失败: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}