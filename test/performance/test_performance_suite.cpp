#include "PerformanceTestSuite.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <string>

using namespace fastexcel::test;

void printUsage() {
    std::cout << "FastExcel Performance Test Suite\n";
    std::cout << "使用方法:\n";
    std::cout << "  fastexcel_performance_tests [选项]\n\n";
    std::cout << "选项:\n";
    std::cout << "  --all               运行所有性能测试（默认）\n";
    std::cout << "  --read              只运行读取性能测试\n";
    std::cout << "  --write             只运行写入性能测试\n";
    std::cout << "  --parsing           只运行解析性能测试\n";
    std::cout << "  --shared-formula    只运行共享公式性能测试\n";
    std::cout << "  --output <dir>      指定输出目录（默认: performance_results）\n";
    std::cout << "  --help              显示此帮助信息\n";
}

int main(int argc, char* argv[]) {
    // 初始化日志
    fastexcel::utils::Logger::getInstance().setLevel(fastexcel::utils::LogLevel::INFO);
    
    std::string test_type = "all";
    std::string output_dir = "performance_results";
    
    // 解析命令行参数
    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        
        if (arg == "--help" || arg == "-h") {
            printUsage();
            return 0;
        } else if (arg == "--all") {
            test_type = "all";
        } else if (arg == "--read") {
            test_type = "read";
        } else if (arg == "--write") {
            test_type = "write";
        } else if (arg == "--parsing") {
            test_type = "parsing";
        } else if (arg == "--shared-formula") {
            test_type = "shared-formula";
        } else if (arg == "--output" && i + 1 < argc) {
            output_dir = argv[++i];
        } else {
            std::cerr << "未知参数: " << arg << std::endl;
            printUsage();
            return 1;
        }
    }
    
    std::cout << "🚀 FastExcel 性能测试套件启动..." << std::endl;
    std::cout << "测试类型: " << test_type << std::endl;
    std::cout << "输出目录: " << output_dir << std::endl;
    std::cout << "===============================================" << std::endl;
    
    try {
        ComprehensivePerformanceTestSuite suite(output_dir);
        
        if (test_type == "all") {
            suite.runAllTests();
        } else if (test_type == "read") {
            suite.runReadTests();
        } else if (test_type == "write") {
            suite.runWriteTests();
        } else if (test_type == "parsing") {
            suite.runParsingTests();
        } else if (test_type == "shared-formula") {
            suite.runSharedFormulaTests();
        } else {
            std::cerr << "无效的测试类型: " << test_type << std::endl;
            return 1;
        }
        
        std::cout << "\n🎉 所有性能测试完成！" << std::endl;
        std::cout << "📊 测试报告已保存到: " << output_dir << std::endl;
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "❌ 性能测试执行失败: " << e.what() << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "❌ 性能测试执行失败: 未知错误" << std::endl;
        return 1;
    }
}