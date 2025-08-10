#include "fastexcel/FastExcel.hpp"
#include <iostream>

int main() {
    try {
        std::cout << "=== 共享公式读写完整性测试 ===" << std::endl;
        
        // ========== 第一步：创建并保存带共享公式的文件 ==========
        std::cout << "\n1. 创建带共享公式的Excel文件..." << std::endl;
        {
            auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("roundtrip_test.xlsx"));
            if (!workbook->open()) {
                std::cout << "无法创建工作簿" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook->addWorksheet("SharedFormulaRoundTrip");
            if (!worksheet) {
                std::cout << "无法添加工作表" << std::endl;
                return 1;
            }
            
            // 写入基础数据
            for (int row = 0; row < 8; ++row) {
                worksheet->writeNumber(row, 0, row + 1);      // A列：1-8
                worksheet->writeNumber(row, 1, (row + 1) * 3);// B列：3,6,9,12...
            }
            
            // 创建第一个共享公式：C1:C5 = A+B
            int si1 = worksheet->createSharedFormula(0, 2, 4, 2, "A1+B1");
            std::cout << "创建共享公式1: si=" << si1 << ", C1:C5 = A+B" << std::endl;
            
            // 创建第二个共享公式：E1:E8 = A*2
            int si2 = worksheet->createSharedFormula(0, 4, 7, 4, "A1*2");
            std::cout << "创建共享公式2: si=" << si2 << ", E1:E8 = A*2" << std::endl;
            
            // 创建第三个共享公式：F6:F8 = A6+B6+10
            int si3 = worksheet->createSharedFormula(5, 5, 7, 5, "A6+B6+10");
            std::cout << "创建共享公式3: si=" << si3 << ", F6:F8 = A+B+10" << std::endl;
            
            // 打印统计信息
            auto* manager = worksheet->getSharedFormulaManager();
            if (manager) {
                auto stats = manager->getStatistics();
                std::cout << "创建后统计信息:" << std::endl;
                std::cout << "  共享公式总数: " << stats.total_shared_formulas << std::endl;
                std::cout << "  受影响单元格: " << stats.total_affected_cells << std::endl;
                std::cout << "  内存节省: " << stats.memory_saved << " 字节" << std::endl;
                std::cout << "  压缩比: " << stats.average_compression_ratio << std::endl;
            }
            
            if (!workbook->save()) {
                std::cout << "保存失败" << std::endl;
                return 1;
            }
            workbook->close();
            std::cout << "文件保存成功！" << std::endl;
        }
        
        // ========== 第二步：读取文件并验证共享公式 ==========
        std::cout << "\n2. 读取Excel文件并验证共享公式..." << std::endl;
        {
            auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path("roundtrip_test.xlsx"));
            if (!workbook) {
                std::cout << "无法打开文件进行读取" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook->getWorksheet("SharedFormulaRoundTrip");
            if (!worksheet) {
                std::cout << "无法获取工作表" << std::endl;
                return 1;
            }
            
            // 验证共享公式管理器
            auto* manager = worksheet->getSharedFormulaManager();
            if (!manager) {
                std::cout << "❌ 警告：读取后没有共享公式管理器！" << std::endl;
            } else {
                auto stats = manager->getStatistics();
                std::cout << "读取后统计信息:" << std::endl;
                std::cout << "  共享公式总数: " << stats.total_shared_formulas << std::endl;
                std::cout << "  受影响单元格: " << stats.total_affected_cells << std::endl;
                std::cout << "  内存节省: " << stats.memory_saved << " 字节" << std::endl;
                std::cout << "  压缩比: " << stats.average_compression_ratio << std::endl;
                
                // 打印调试信息
                std::cout << "\n共享公式详细信息:" << std::endl;
                manager->debugPrint();
            }
            
            // 验证特定单元格的公式
            std::cout << "\n验证单元格公式:" << std::endl;
            
            // 验证C列（共享公式1）
            for (int row = 0; row < 5; ++row) {
                if (worksheet->hasCellAt(row, 2)) {
                    const auto& cell = worksheet->getCell(row, 2);
                    if (cell.isFormula()) {
                        std::cout << "  C" << (row + 1) << ": " << cell.getFormula() 
                                  << " (共享: " << (cell.isSharedFormula() ? "是" : "否") << ")";
                        if (cell.isSharedFormula()) {
                            std::cout << " [si=" << cell.getSharedFormulaIndex() << "]";
                        }
                        std::cout << std::endl;
                    }
                }
            }
            
            // 验证E列（共享公式2）
            for (int row = 0; row < 8; ++row) {
                if (worksheet->hasCellAt(row, 4)) {
                    const auto& cell = worksheet->getCell(row, 4);
                    if (cell.isFormula()) {
                        std::cout << "  E" << (row + 1) << ": " << cell.getFormula() 
                                  << " (共享: " << (cell.isSharedFormula() ? "是" : "否") << ")";
                        if (cell.isSharedFormula()) {
                            std::cout << " [si=" << cell.getSharedFormulaIndex() << "]";
                        }
                        std::cout << std::endl;
                    }
                }
            }
            
            workbook->close();
        }
        
        // ========== 第三步：修改并重新保存 ==========
        std::cout << "\n3. 修改文件并重新保存..." << std::endl;
        {
            auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path("roundtrip_test.xlsx"));
            if (!workbook) {
                std::cout << "无法打开文件进行修改" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook->getWorksheet("SharedFormulaRoundTrip");
            if (!worksheet) {
                std::cout << "无法获取工作表进行修改" << std::endl;
                return 1;
            }
            
            // 添加一个新的共享公式：G1:G3 = A1+5
            int si4 = worksheet->createSharedFormula(0, 6, 2, 6, "A1+5");
            std::cout << "添加新共享公式: si=" << si4 << ", G1:G3 = A+5" << std::endl;
            
            // 修改一些数据
            worksheet->writeNumber(0, 0, 100);  // 改变A1，应该影响多个共享公式
            
            // 保存修改
            if (!workbook->saveAs("roundtrip_modified.xlsx")) {
                std::cout << "保存修改失败" << std::endl;
                return 1;
            }
            workbook->close();
            std::cout << "修改保存成功！" << std::endl;
        }
        
        std::cout << "\n=== 测试完成 ===" << std::endl;
        std::cout << "请检查生成的文件:" << std::endl;
        std::cout << "  - roundtrip_test.xlsx (原始文件)" << std::endl;
        std::cout << "  - roundtrip_modified.xlsx (修改后文件)" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "错误: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}
