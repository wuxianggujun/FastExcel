#include "fastexcel/FastExcel.hpp"
#include <iostream>
#include <iomanip>
#include <map>
#include <vector>

/**
 * @brief 高级公式优化分析器演示程序
 * 
 * 展示如何使用FastExcel的共享公式系统进行：
 * 1. 自动检测可优化的公式模式
 * 2. 分析优化潜力和收益
 * 3. 提供优化建议
 * 4. 执行自动优化
 */

class FormulaOptimizationAnalyzer {
public:
    struct OptimizationReport {
        size_t total_formulas = 0;
        size_t optimizable_formulas = 0;
        size_t memory_savings_bytes = 0;
        double optimization_ratio = 0.0;
        std::vector<std::string> recommendations;
    };

    static OptimizationReport analyzeWorksheet(std::shared_ptr<fastexcel::core::Worksheet> worksheet) {
        OptimizationReport report;
        
        if (!worksheet) {
            return report;
        }

        // 收集所有公式
        std::map<std::pair<int, int>, std::string> formulas;
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    if (cell.isFormula()) {
                        formulas[{row, col}] = cell.getFormula();
                    }
                }
            }
        }

        report.total_formulas = formulas.size();

        if (formulas.empty()) {
            report.recommendations.push_back("📊 工作表中未发现公式，无需优化");
            return report;
        }

        // 使用SharedFormulaManager检测优化模式
        auto* manager = worksheet->getSharedFormulaManager();
        if (manager) {
            auto patterns = manager->detectSharedFormulaPatterns(formulas);
            
            // 分析优化潜力
            size_t optimizable_count = 0;
            size_t estimated_savings = 0;
            
            for (const auto& pattern : patterns) {
                optimizable_count += pattern.matching_cells.size();
                estimated_savings += pattern.estimated_savings;
            }

            report.optimizable_formulas = optimizable_count;
            report.memory_savings_bytes = estimated_savings;
            
            if (report.total_formulas > 0) {
                report.optimization_ratio = static_cast<double>(optimizable_count) / report.total_formulas * 100.0;
            }

            // 生成优化建议
            generateRecommendations(report, patterns, formulas);
        }

        return report;
    }

private:
    static void generateRecommendations(OptimizationReport& report, 
                                       const std::vector<fastexcel::core::SharedFormulaManager::FormulaPattern>& patterns,
                                       const std::map<std::pair<int, int>, std::string>& formulas) {
        
        if (patterns.empty()) {
            report.recommendations.push_back("✅ 未发现可优化的公式模式");
            report.recommendations.push_back("💡 建议：考虑使用更多相似的公式来获得优化效果");
            return;
        }

        report.recommendations.push_back("🎯 发现优化机会：");
        
        for (size_t i = 0; i < std::min(patterns.size(), size_t(5)); ++i) {
            const auto& pattern = patterns[i];
            
            std::string recommendation = "  📈 模式 " + std::to_string(i + 1) + ": " +
                std::to_string(pattern.matching_cells.size()) + " 个相似公式，" +
                "预估节省 " + std::to_string(pattern.estimated_savings) + " 字节";
            
            report.recommendations.push_back(recommendation);
            
            // 显示具体的公式示例
            if (!pattern.matching_cells.empty()) {
                auto first_pos = pattern.matching_cells[0];
                auto formula_it = formulas.find(first_pos);
                if (formula_it != formulas.end()) {
                    std::string cell_ref = fastexcel::utils::CommonUtils::cellReference(first_pos.first, first_pos.second);
                    report.recommendations.push_back("     📝 示例: " + cell_ref + " = " + formula_it->second);
                }
            }
        }

        // 总体建议
        if (report.optimization_ratio > 50.0) {
            report.recommendations.push_back("🚀 高优化潜力：建议立即执行自动优化");
        } else if (report.optimization_ratio > 20.0) {
            report.recommendations.push_back("📊 中等优化潜力：建议考虑执行优化");
        } else {
            report.recommendations.push_back("💭 低优化潜力：可选择性执行优化");
        }

        // 具体操作建议
        report.recommendations.push_back("🛠️ 执行方法：调用 worksheet->optimizeFormulas() 自动优化");
    }
};

int main() {
    try {
        std::cout << "=== 公式优化分析器演示 ===\n" << std::endl;

        // ========== 第一步：创建测试工作簿 ==========
        std::cout << "1. 创建包含各种公式模式的测试工作簿..." << std::endl;
        
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("formula_optimization_test.xlsx"));
        if (!workbook->open()) {
            std::cout << "❌ 无法创建工作簿" << std::endl;
            return 1;
        }

        auto worksheet = workbook->addWorksheet("OptimizationTest");
        if (!worksheet) {
            std::cout << "❌ 无法添加工作表" << std::endl;
            return 1;
        }

        // 创建基础数据
        for (int row = 0; row < 20; ++row) {
            worksheet->writeNumber(row, 0, row + 1);        // A列：1-20
            worksheet->writeNumber(row, 1, (row + 1) * 2);  // B列：2,4,6,8...
            worksheet->writeNumber(row, 2, (row + 1) * 3);  // C列：3,6,9,12...
        }

        // 模式1：简单加法公式（A+B）- 高度相似
        for (int row = 0; row < 10; ++row) {
            std::string formula = "A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1);
            worksheet->writeFormula(row, 3, formula); // D列
        }

        // 模式2：复杂计算公式（A*B+C）- 高度相似
        for (int row = 0; row < 8; ++row) {
            std::string formula = "A" + std::to_string(row + 1) + "*B" + std::to_string(row + 1) + "+C" + std::to_string(row + 1);
            worksheet->writeFormula(row, 4, formula); // E列
        }

        // 模式3：求和公式（SUM）- 中等相似
        for (int row = 2; row < 12; ++row) {
            std::string formula = "SUM(A1:A" + std::to_string(row + 1) + ")";
            worksheet->writeFormula(row, 5, formula); // F列
        }

        // 模式4：条件公式（IF）- 低相似度
        for (int row = 0; row < 5; ++row) {
            std::string formula = "IF(A" + std::to_string(row + 1) + ">10,\"大\",\"小\")";
            worksheet->writeFormula(row, 6, formula); // G列
        }

        // 模式5：独立公式 - 无相似
        worksheet->writeFormula(0, 7, "AVERAGE(A1:A20)");
        worksheet->writeFormula(1, 7, "MAX(B1:B20)");
        worksheet->writeFormula(2, 7, "MIN(C1:C20)");

        std::cout << "✅ 测试数据创建完成" << std::endl;

        // ========== 第二步：执行优化分析 ==========
        std::cout << "\n2. 执行公式优化分析..." << std::endl;
        
        auto report = FormulaOptimizationAnalyzer::analyzeWorksheet(worksheet);

        // 显示分析报告
        std::cout << "\n📊 === 优化分析报告 ===" << std::endl;
        std::cout << "总公式数量: " << report.total_formulas << std::endl;
        std::cout << "可优化公式数量: " << report.optimizable_formulas << std::endl;
        std::cout << "预估内存节省: " << report.memory_savings_bytes << " 字节" << std::endl;
        std::cout << "优化潜力: " << std::fixed << std::setprecision(1) << report.optimization_ratio << "%" << std::endl;
        
        std::cout << "\n📋 优化建议:" << std::endl;
        for (const auto& recommendation : report.recommendations) {
            std::cout << recommendation << std::endl;
        }

        // ========== 第三步：执行自动优化 ==========
        std::cout << "\n3. 执行自动优化..." << std::endl;
        
        auto* manager = worksheet->getSharedFormulaManager();
        if (manager) {
            // 收集现有公式进行优化
            std::map<std::pair<int, int>, std::string> formulas;
            auto [max_row, max_col] = worksheet->getUsedRange();
            
            for (int row = 0; row <= max_row; ++row) {
                for (int col = 0; col <= max_col; ++col) {
                    if (worksheet->hasCellAt(row, col)) {
                        const auto& cell = worksheet->getCell(row, col);
                        if (cell.isFormula() && !cell.isSharedFormula()) { // 只优化非共享公式
                            formulas[{row, col}] = cell.getFormula();
                        }
                    }
                }
            }

            int optimized_count = manager->optimizeFormulas(formulas, 3); // 至少3个相似公式才优化
            
            if (optimized_count > 0) {
                std::cout << "✅ 成功优化 " << optimized_count << " 个公式为共享公式" << std::endl;
                
                // 显示优化后的统计信息
                auto stats = manager->getStatistics();
                std::cout << "\n📈 优化后统计信息:" << std::endl;
                std::cout << "  共享公式总数: " << stats.total_shared_formulas << std::endl;
                std::cout << "  受影响单元格: " << stats.total_affected_cells << std::endl;
                std::cout << "  内存节省: " << stats.memory_saved << " 字节" << std::endl;
                std::cout << "  平均压缩比: " << std::fixed << std::setprecision(2) << stats.average_compression_ratio << std::endl;
            } else {
                std::cout << "ℹ️ 未找到足够的相似公式进行优化（需要至少3个相似公式）" << std::endl;
            }
        }

        // ========== 第四步：保存并再次分析 ==========
        std::cout << "\n4. 保存文件并验证优化效果..." << std::endl;
        
        if (!workbook->save()) {
            std::cout << "❌ 保存失败" << std::endl;
            return 1;
        }
        workbook->close();

        // 重新打开文件验证
        auto verification_workbook = fastexcel::core::Workbook::open(fastexcel::core::Path("formula_optimization_test.xlsx"));
        if (verification_workbook) {
            auto verification_worksheet = verification_workbook->getWorksheet("OptimizationTest");
            if (verification_worksheet) {
                auto final_report = FormulaOptimizationAnalyzer::analyzeWorksheet(verification_worksheet);
                
                std::cout << "\n📋 验证结果:" << std::endl;
                std::cout << "  原始公式数量: " << report.total_formulas << " → " << final_report.total_formulas << std::endl;
                std::cout << "  优化潜力: " << std::fixed << std::setprecision(1) 
                         << report.optimization_ratio << "% → " << final_report.optimization_ratio << "%" << std::endl;
                
                if (final_report.optimization_ratio < report.optimization_ratio) {
                    std::cout << "✅ 优化效果显著！优化潜力降低了 " 
                             << (report.optimization_ratio - final_report.optimization_ratio) << "%" << std::endl;
                } else {
                    std::cout << "ℹ️ 优化效果有限，可能需要调整优化策略" << std::endl;
                }
            }
            verification_workbook->close();
        }

        std::cout << "\n=== 分析完成 ===" << std::endl;
        std::cout << "生成文件: formula_optimization_test.xlsx" << std::endl;
        std::cout << "\n💡 使用建议:" << std::endl;
        std::cout << "1. 在实际项目中定期运行公式优化分析" << std::endl;
        std::cout << "2. 对于大型工作表，优化效果更加明显" << std::endl;
        std::cout << "3. 建议在保存前执行自动优化以减少文件大小" << std::endl;

    } catch (const std::exception& e) {
        std::cout << "❌ 错误: " << e.what() << std::endl;
        return 1;
    }

    return 0;
}