#include "fastexcel/FastExcel.hpp"
#include <iostream>

int main() {
    try {
        // 创建工作簿
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("test_shared_formula.xlsx"));
        if (!workbook) {
            std::cout << "Failed to create workbook" << std::endl;
            return 1;
        }
        
        // 添加工作表
        auto worksheet = workbook->addSheet("SharedFormulaTest");
        if (!worksheet) {
            std::cout << "Failed to add worksheet" << std::endl;
            return 1;
        }
        
        // 写入一些数据作为基础
        std::cout << "Writing base data..." << std::endl;
        for (int row = 0; row < 5; ++row) {
            worksheet->setValue(row, 0, row + 1);  // A列：1, 2, 3, 4, 5
            worksheet->setValue(row, 1, (row + 1) * 2);  // B列：2, 4, 6, 8, 10
        }
        
        // 创建共享公式：C列 = A列 + B列
        std::cout << "Creating shared formula..." << std::endl;
        int shared_index = worksheet->createSharedFormula(0, 2, 4, 2, "A1+B1");
        
        if (shared_index >= 0) {
            std::cout << "Shared formula created successfully with index: " << shared_index << std::endl;
            
            // 获取共享公式管理器并打印统计信息
            auto* manager = worksheet->getSharedFormulaManager();
            if (manager) {
                auto stats = manager->getStatistics();
                std::cout << "Shared formula statistics:" << std::endl;
                std::cout << "  Total shared formulas: " << stats.total_shared_formulas << std::endl;
                std::cout << "  Total affected cells: " << stats.total_affected_cells << std::endl;
                std::cout << "  Memory saved: " << stats.memory_saved << " bytes" << std::endl;
                std::cout << "  Average compression ratio: " << stats.average_compression_ratio << std::endl;
                
                // 打印调试信息
                manager->debugPrint();
                
                // 测试公式展开
                std::cout << "\nTesting formula expansion:" << std::endl;
                for (int row = 0; row < 5; ++row) {
                    std::string expanded = manager->getExpandedFormula(row, 2);
                    std::cout << "  Cell C" << (row + 1) << ": " << expanded << std::endl;
                }
            }
        } else {
            std::cout << "Failed to create shared formula" << std::endl;
        }
        
        // 测试单个公式写入
        std::cout << "\nWriting individual formulas for comparison..." << std::endl;
        for (int row = 0; row < 5; ++row) {
            std::string formula = "A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1);
            worksheet->getCell(row, 3).setFormula(formula);  // D列使用单独的公式
        }
        
        // 保存工作簿
        std::cout << "\nSaving workbook..." << std::endl;
        if (workbook->save()) {
            std::cout << "Workbook saved successfully!" << std::endl;
        } else {
            std::cout << "Failed to save workbook" << std::endl;
        }
        
        workbook->close();
        std::cout << "Test completed." << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "Error: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}