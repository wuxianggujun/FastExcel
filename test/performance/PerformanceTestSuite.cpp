#include "PerformanceTestSuite.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <fstream>
#include <iomanip>
#include <sstream>
#include <algorithm>
#include <random>

#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#else
#include <sys/resource.h>
#include <unistd.h>
#endif

namespace fastexcel {
namespace test {

// ========== MemoryMonitor 实现 ==========

MemoryMonitor::MemoryMonitor() : initial_memory_(0), peak_memory_(0), monitoring_(false) {}

MemoryMonitor::~MemoryMonitor() {
    if (monitoring_) {
        stopMonitoring();
    }
}

void MemoryMonitor::startMonitoring() {
    initial_memory_ = getProcessMemoryUsage();
    peak_memory_ = initial_memory_;
    monitoring_ = true;
    memory_snapshots_.clear();
}

void MemoryMonitor::stopMonitoring() {
    monitoring_ = false;
}

size_t MemoryMonitor::getCurrentMemoryUsage() {
    return getProcessMemoryUsage();
}

size_t MemoryMonitor::getPeakMemoryUsage() {
    if (monitoring_) {
        size_t current = getProcessMemoryUsage();
        peak_memory_ = std::max(peak_memory_, current);
    }
    return peak_memory_;
}

void MemoryMonitor::recordMemorySnapshot(const std::string& checkpoint) {
    if (monitoring_) {
        size_t current = getProcessMemoryUsage();
        memory_snapshots_.emplace_back(checkpoint, current);
        peak_memory_ = std::max(peak_memory_, current);
    }
}

size_t MemoryMonitor::getProcessMemoryUsage() {
#ifdef _WIN32
    PROCESS_MEMORY_COUNTERS_EX pmc;
    if (GetProcessMemoryInfo(GetCurrentProcess(), (PROCESS_MEMORY_COUNTERS*)&pmc, sizeof(pmc))) {
        return pmc.WorkingSetSize / 1024; // 返回KB
    }
    return 0;
#else
    struct rusage usage;
    if (getrusage(RUSAGE_SELF, &usage) == 0) {
        return usage.ru_maxrss; // Linux上已经是KB，macOS上是字节
    }
    return 0;
#endif
}

// ========== PerformanceTestBase 实现 ==========

PerformanceTestBase::PerformanceTestBase(const std::string& suite_name)
    : test_suite_name_(suite_name), memory_monitor_(std::make_unique<MemoryMonitor>()) {
    setupTest();
}

// 显式实例化模板方法
template PerformanceResult PerformanceTestBase::measurePerformance<std::function<void()>>(
    const std::string& test_name, size_t operations_count, std::function<void()>&& test_function);

template<typename Func>
PerformanceResult PerformanceTestBase::measurePerformance(const std::string& test_name,
                                                         size_t operations_count,
                                                         Func&& test_function) {
    PerformanceResult result;
    result.test_name = test_name;
    result.operations_count = operations_count;

    // 开始监控内存
    memory_monitor_->startMonitoring();
    
    // 记录开始时间
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 执行测试函数
    try {
        test_function();
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Performance test '{}' failed: {}", test_name, e.what());
        result.execution_time_ms = -1; // 标记为失败
        return result;
    }
    
    // 记录结束时间
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    result.execution_time_ms = duration.count() / 1000.0;
    result.memory_usage_kb = memory_monitor_->getCurrentMemoryUsage();
    result.peak_memory_kb = memory_monitor_->getPeakMemoryUsage();
    
    memory_monitor_->stopMonitoring();
    
    result.calculateOperationRate();
    
    // 记录结果
    results_.push_back(result);
    logResult(result);
    
    return result;
}

void PerformanceTestBase::generateReport(const std::string& output_file) {
    std::ostream* out;
    std::ofstream file_stream;
    
    if (!output_file.empty()) {
        file_stream.open(output_file);
        out = &file_stream;
    } else {
        out = &std::cout;
    }
    
    *out << "\\n========== Performance Test Report: " << test_suite_name_ << " ==========\\n";
    *out << std::left << std::setw(30) << "Test Name"
         << std::setw(15) << "Time (ms)"
         << std::setw(15) << "Memory (KB)"
         << std::setw(15) << "Peak Mem (KB)"
         << std::setw(12) << "Operations"
         << std::setw(15) << "Ops/Second" << "\\n";
    *out << std::string(102, '-') << "\\n";
    
    for (const auto& result : results_) {
        if (result.execution_time_ms >= 0) { // 只显示成功的测试
            *out << std::left << std::setw(30) << result.test_name
                 << std::setw(15) << std::fixed << std::setprecision(2) << result.execution_time_ms
                 << std::setw(15) << result.memory_usage_kb
                 << std::setw(15) << result.peak_memory_kb
                 << std::setw(12) << result.operations_count
                 << std::setw(15) << std::fixed << std::setprecision(0) << result.operations_per_second << "\\n";
        }
    }
    *out << std::string(102, '=') << "\\n\\n";
}

void PerformanceTestBase::exportToCSV(const std::string& csv_file) {
    std::ofstream file(csv_file);
    if (!file.is_open()) {
        FASTEXCEL_LOG_ERROR("无法打开CSV文件进行写入: {}", csv_file);
        return;
    }
    
    // CSV标题
    file << "TestSuite,TestName,ExecutionTime(ms),MemoryUsage(KB),PeakMemory(KB),Operations,OperationsPerSecond,FileSize(bytes)\\n";
    
    for (const auto& result : results_) {
        if (result.execution_time_ms >= 0) {
            file << test_suite_name_ << ","
                 << result.test_name << ","
                 << result.execution_time_ms << ","
                 << result.memory_usage_kb << ","
                 << result.peak_memory_kb << ","
                 << result.operations_count << ","
                 << result.operations_per_second << ","
                 << result.file_size_bytes << "\\n";
        }
    }
    
    file.close();
    FASTEXCEL_LOG_DEBUG("性能测试结果已导出到CSV文件: {}", csv_file);
}

void PerformanceTestBase::setupTest() {
    FASTEXCEL_LOG_DEBUG("设置性能测试环境: {}", test_suite_name_);
}

void PerformanceTestBase::teardownTest() {
    FASTEXCEL_LOG_DEBUG("清理性能测试环境: {}", test_suite_name_);
}

void PerformanceTestBase::logResult(const PerformanceResult& result) {
    if (result.execution_time_ms >= 0) {
        FASTEXCEL_LOG_DEBUG("性能测试 '{}' 完成: {:.2f}ms, {:.0f} ops/sec, 内存使用 {}KB", 
                 result.test_name, result.execution_time_ms, result.operations_per_second, result.memory_usage_kb);
    } else {
        FASTEXCEL_LOG_ERROR("性能测试 '{}' 失败", result.test_name);
    }
}

// ========== ReadPerformanceTest 实现 ==========

void ReadPerformanceTest::testBasicFileReading() {
    createTestFiles();
    
    for (const auto& file_path : test_files_) {
        std::string test_name = "BasicRead_" + file_path;
        measurePerformance(test_name, 1, [&]() {
            auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path(file_path));
            if (workbook) {
                auto sheets = workbook->getWorksheetNames();
                for (const auto& sheet_name : sheets) {
                    auto worksheet = workbook->getWorksheet(sheet_name);
                    if (worksheet) {
                        auto [max_row, max_col] = worksheet->getUsedRange();
                        // 模拟实际使用：访问一些单元格
                        for (int row = 0; row <= std::min(max_row, 100); ++row) {
                            for (int col = 0; col <= std::min(max_col, 10); ++col) {
                                if (worksheet->hasCellAt(row, col)) {
                                    const auto& cell = worksheet->getCell(row, col);
                                    // 简单访问以触发实际读取
                                    volatile auto value = cell.isEmpty();
                                    (void)value;
                                }
                            }
                        }
                    }
                }
                workbook->close();
            }
        });
    }
}

void ReadPerformanceTest::testSharedFormulaReading() {
    // 创建包含共享公式的测试文件
    std::string test_file = "shared_formula_test.xlsx";
    
    // 创建测试文件
    {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path(test_file));
        workbook->open();
        auto worksheet = workbook->addWorksheet("SharedFormulaTest");
        
        // 添加基础数据
        for (int row = 0; row < 1000; ++row) {
            worksheet->writeNumber(row, 0, row + 1);
            worksheet->writeNumber(row, 1, (row + 1) * 2);
        }
        
        // 创建共享公式
        worksheet->createSharedFormula(0, 2, 999, 2, "A1+B1");
        worksheet->createSharedFormula(0, 3, 999, 3, "A1*B1");
        
        workbook->save();
        workbook->close();
    }
    
    // 测试读取性能
    measurePerformance("SharedFormulaReading", 2000, [&]() {
        auto workbook = fastexcel::core::Workbook::open(fastexcel::core::Path(test_file));
        if (workbook) {
            auto worksheet = workbook->getWorksheet("SharedFormulaTest");
            if (worksheet) {
                auto* manager = worksheet->getSharedFormulaManager();
                if (manager) {
                    auto stats = manager->getStatistics();
                    // 验证共享公式是否正确读取
                    assert(stats.total_shared_formulas > 0);
                }
            }
            workbook->close();
        }
    });
}

void ReadPerformanceTest::testReadingByFileSize() {
    std::vector<std::pair<std::string, size_t>> file_sizes = {
        {"small_file.xlsx", 100},     // 100行
        {"medium_file.xlsx", 1000},   // 1000行
        {"large_file.xlsx", 10000}    // 10000行
    };
    
    // 创建不同大小的测试文件
    for (const auto& [filename, row_count] : file_sizes) {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path(filename));
        workbook->open();
        auto worksheet = workbook->addWorksheet("TestData");
        
        for (size_t row = 0; row < row_count; ++row) {
            worksheet->writeNumber(row, 0, row);
            worksheet->writeString(row, 1, "Data " + std::to_string(row));
            worksheet->writeNumber(row, 2, row * 1.5);
        }
        
        workbook->save();
        workbook->close();
        
        // 测试读取性能
        std::string test_name = "ReadFileSize_" + std::to_string(row_count);
        measurePerformance(test_name, row_count, [&]() {
            auto test_workbook = fastexcel::core::Workbook::open(fastexcel::core::Path(filename));
            if (test_workbook) {
                auto test_worksheet = test_workbook->getWorksheet("TestData");
                if (test_worksheet) {
                    auto [max_row, max_col] = test_worksheet->getUsedRange();
                    size_t cell_count = 0;
                    for (int row = 0; row <= max_row; ++row) {
                        for (int col = 0; col <= max_col; ++col) {
                            if (test_worksheet->hasCellAt(row, col)) {
                                const auto& cell = test_worksheet->getCell(row, col);
                                volatile auto empty = cell.isEmpty();
                                (void)empty;
                                cell_count++;
                            }
                        }
                    }
                }
                test_workbook->close();
            }
        });
    }
}

void ReadPerformanceTest::runAllTests() {
    std::cout << "\\n🚀 开始读取性能测试..." << std::endl;
    
    testBasicFileReading();
    testSharedFormulaReading();
    testReadingByFileSize();
    
    std::cout << "✅ 读取性能测试完成!" << std::endl;
    generateReport();
}

void ReadPerformanceTest::createTestFiles() {
    test_files_ = {"basic_test.xlsx"};
    
    // 创建基本测试文件
    auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("basic_test.xlsx"));
    workbook->open();
    auto worksheet = workbook->addWorksheet("BasicTest");
    
    for (int row = 0; row < 100; ++row) {
        worksheet->writeString(row, 0, "Cell " + std::to_string(row));
        worksheet->writeNumber(row, 1, row * 2.5);
    }
    
    workbook->save();
    workbook->close();
}

// ========== WritePerformanceTest 实现 ==========

void WritePerformanceTest::testBasicFileWriting() {
    measurePerformance("BasicFileWriting", 100, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("write_test_basic.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("WriteTest");
        
        for (int row = 0; row < 100; ++row) {
            worksheet->writeString(row, 0, "Row " + std::to_string(row));
            worksheet->writeNumber(row, 1, row);
            worksheet->writeNumber(row, 2, row * 1.5);
        }
        
        workbook->save();
        workbook->close();
    });
}

void WritePerformanceTest::testSharedFormulaWriting() {
    measurePerformance("SharedFormulaWriting", 2000, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("shared_formula_write_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("SharedFormulaWrite");
        
        // 基础数据
        for (int row = 0; row < 1000; ++row) {
            worksheet->writeNumber(row, 0, row + 1);
            worksheet->writeNumber(row, 1, (row + 1) * 2);
        }
        
        // 创建共享公式
        worksheet->createSharedFormula(0, 2, 999, 2, "A1+B1");
        worksheet->createSharedFormula(0, 3, 999, 3, "A1*B1*2");
        
        workbook->save();
        workbook->close();
    });
}

void WritePerformanceTest::testBatchVsStreamingMode() {
    std::vector<std::pair<std::string, fastexcel::core::WorkbookMode>> modes = {
        {"BatchMode", fastexcel::core::WorkbookMode::BATCH},
        {"StreamingMode", fastexcel::core::WorkbookMode::STREAMING}
    };
    
    for (const auto& [mode_name, mode] : modes) {
        measurePerformance("Write_" + mode_name, 10000, [&]() {
            auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("mode_test_" + mode_name + ".xlsx"));
            
            // 设置工作模式
            fastexcel::core::WorkbookOptions options;
            options.mode = mode;
            workbook->setOptions(options);
            
            workbook->open();
            auto worksheet = workbook->addWorksheet("ModeTest");
            
            for (int row = 0; row < 5000; ++row) {
                worksheet->writeString(row, 0, "Data " + std::to_string(row));
                worksheet->writeNumber(row, 1, row);
                worksheet->writeFormula(row, 2, "B" + std::to_string(row + 1) + "*2");
            }
            
            workbook->save();
            workbook->close();
        });
    }
}

void WritePerformanceTest::runAllTests() {
    std::cout << "\\n📝 开始写入性能测试..." << std::endl;
    
    testBasicFileWriting();
    testSharedFormulaWriting();
    testBatchVsStreamingMode();
    
    std::cout << "✅ 写入性能测试完成!" << std::endl;
    generateReport();
}

void WritePerformanceTest::generateTestData(size_t rows, size_t cols, double formula_ratio) {
    test_workbook_ = fastexcel::core::Workbook::create(fastexcel::core::Path("test_data.xlsx"));
    test_workbook_->open();
    auto worksheet = test_workbook_->addWorksheet("GeneratedData");
    
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_real_distribution<> dis(0.0, 1.0);
    
    for (size_t row = 0; row < rows; ++row) {
        for (size_t col = 0; col < cols; ++col) {
            if (dis(gen) < formula_ratio) {
                // 生成公式
                worksheet->writeFormula(row, col, "A" + std::to_string(row + 1) + "*" + std::to_string(col + 1));
            } else {
                // 生成数据
                worksheet->writeNumber(row, col, row * col + dis(gen) * 100);
            }
        }
    }
}

// ========== ComprehensivePerformanceTestSuite 实现 ==========

ComprehensivePerformanceTestSuite::ComprehensivePerformanceTestSuite(const std::string& output_dir)
    : output_directory_(output_dir) {
    setupOutputDirectory();
    
    read_test_ = std::make_unique<ReadPerformanceTest>();
    write_test_ = std::make_unique<WritePerformanceTest>();
    parsing_test_ = std::make_unique<ParsingPerformanceTest>();
    shared_formula_test_ = std::make_unique<SharedFormulaPerformanceTest>();
}

void ComprehensivePerformanceTestSuite::runAllTests() {
    std::cout << "\\n🎯 开始综合性能测试套件..." << std::endl;
    
    runReadTests();
    runWriteTests();
    runParsingTests();
    runSharedFormulaTests();
    
    generateComprehensiveReport();
    
    std::cout << "\\n🎉 所有性能测试完成!" << std::endl;
}

void ComprehensivePerformanceTestSuite::runReadTests() {
    std::cout << "\\n📖 执行读取测试..." << std::endl;
    read_test_->runAllTests();
}

void ComprehensivePerformanceTestSuite::runWriteTests() {
    std::cout << "\\n📝 执行写入测试..." << std::endl;
    write_test_->runAllTests();
}

void ComprehensivePerformanceTestSuite::generateComprehensiveReport() {
    std::string report_file = output_directory_ + "/comprehensive_performance_report.txt";
    std::ofstream report(report_file);
    
    report << "FastExcel 综合性能测试报告\\n";
    report << "生成时间: " << std::time(nullptr) << "\\n";
    report << "=====================================\\n\\n";
    
    // 汇总各个测试套件的结果
    const auto& read_results = read_test_->getResults();
    const auto& write_results = write_test_->getResults();
    
    report << "📊 测试总览:\\n";
    report << "读取测试: " << read_results.size() << " 项\\n";
    report << "写入测试: " << write_results.size() << " 项\\n";
    
    report.close();
    
    std::cout << "📊 综合报告已生成: " << report_file << std::endl;
}

void ComprehensivePerformanceTestSuite::runParsingTests() {
    std::cout << "\n🔍 执行解析测试..." << std::endl;
    parsing_test_->runAllTests();
}

void ComprehensivePerformanceTestSuite::runSharedFormulaTests() {
    std::cout << "\n📊 执行共享公式测试..." << std::endl;
    shared_formula_test_->runAllTests();
}

// ========== ParsingPerformanceTest 实现 ==========

void ParsingPerformanceTest::testXMLParsingSpeed() {
    prepareTestXMLData();
    
    for (const auto& [xml_type, xml_content] : test_xml_data_) {
        std::string test_name = "XMLParsing_" + xml_type;
        measurePerformance(test_name, 1, [&]() {
            // 模拟XML解析
            size_t element_count = 0;
            size_t pos = 0;
            while ((pos = xml_content.find('<', pos)) != std::string::npos) {
                element_count++;
                pos++;
            }
            volatile auto count = element_count;
            (void)count;
        });
    }
}

void ParsingPerformanceTest::testLargeXMLParsing() {
    // 创建大型XML测试数据
    std::string large_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet>\n";
    for (int i = 0; i < 10000; ++i) {
        large_xml += "<row r=\"" + std::to_string(i + 1) + "\">\n";
        for (int j = 0; j < 10; ++j) {
            std::string cell_ref = utils::CommonUtils::cellReference(i, j);
            large_xml += "<c r=\"" + cell_ref + "\"><v>" + std::to_string(i * j) + "</v></c>\n";
        }
        large_xml += "</row>\n";
    }
    large_xml += "</worksheet>";
    
    measurePerformance("LargeXMLParsing", 100000, [&]() {
        // 模拟大型XML解析
        size_t cell_count = 0;
        size_t pos = 0;
        while ((pos = large_xml.find("<c r=", pos)) != std::string::npos) {
            cell_count++;
            pos += 5;
        }
        volatile auto count = cell_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testStylesParsing() {
    // 创建样式XML测试数据
    std::string styles_xml = "<?xml version=\"1.0\"?>\n<styleSheet>\n<fills>\n";
    for (int i = 0; i < 1000; ++i) {
        styles_xml += "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF" + 
                     std::to_string(100000 + i) + "\"/></patternFill></fill>\n";
    }
    styles_xml += "</fills>\n</styleSheet>";
    
    measurePerformance("StylesParsing", 1000, [&]() {
        size_t fill_count = 0;
        size_t pos = 0;
        while ((pos = styles_xml.find("<fill>", pos)) != std::string::npos) {
            fill_count++;
            pos += 6;
        }
        volatile auto count = fill_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testComplexStylesParsing() {
    // 创建复杂样式XML
    std::string complex_styles = "<?xml version=\"1.0\"?>\n<styleSheet>\n";
    complex_styles += "<numFmts>\n";
    for (int i = 0; i < 100; ++i) {
        complex_styles += "<numFmt numFmtId=\"" + std::to_string(164 + i) + "\" formatCode=\"General\"/>\n";
    }
    complex_styles += "</numFmts>\n<fonts>\n";
    for (int i = 0; i < 500; ++i) {
        complex_styles += "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>\n";
    }
    complex_styles += "</fonts>\n</styleSheet>";
    
    measurePerformance("ComplexStylesParsing", 600, [&]() {
        size_t format_count = 0;
        size_t pos = 0;
        while ((pos = complex_styles.find("<numFmt", pos)) != std::string::npos) {
            format_count++;
            pos += 7;
        }
        while ((pos = complex_styles.find("<font>", pos)) != std::string::npos) {
            format_count++;
            pos += 6;
        }
        volatile auto count = format_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testSharedStringsParsing() {
    // 创建共享字符串XML
    std::string shared_strings = "<?xml version=\"1.0\"?>\n<sst>\n";
    for (int i = 0; i < 5000; ++i) {
        shared_strings += "<si><t>String " + std::to_string(i) + "</t></si>\n";
    }
    shared_strings += "</sst>";
    
    measurePerformance("SharedStringsParsing", 5000, [&]() {
        size_t string_count = 0;
        size_t pos = 0;
        while ((pos = shared_strings.find("<si>", pos)) != std::string::npos) {
            string_count++;
            pos += 4;
        }
        volatile auto count = string_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testWorksheetParsing() {
    // 创建工作表XML
    std::string worksheet_xml = "<?xml version=\"1.0\"?>\n<worksheet>\n<sheetData>\n";
    for (int row = 0; row < 1000; ++row) {
        worksheet_xml += "<row r=\"" + std::to_string(row + 1) + "\">\n";
        for (int col = 0; col < 5; ++col) {
            std::string cell_ref = utils::CommonUtils::cellReference(row, col);
            worksheet_xml += "<c r=\"" + cell_ref + "\"><v>" + std::to_string(row * col) + "</v></c>\n";
        }
        worksheet_xml += "</row>\n";
    }
    worksheet_xml += "</sheetData>\n</worksheet>";
    
    measurePerformance("WorksheetParsing", 5000, [&]() {
        size_t cell_count = 0;
        size_t pos = 0;
        while ((pos = worksheet_xml.find("<c r=", pos)) != std::string::npos) {
            cell_count++;
            pos += 5;
        }
        volatile auto count = cell_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testMultipleWorksheetsParsing() {
    // 创建多个工作表数据
    std::vector<std::string> worksheets;
    for (int sheet = 0; sheet < 5; ++sheet) {
        std::string xml = "<?xml version=\"1.0\"?>\n<worksheet>\n<sheetData>\n";
        for (int row = 0; row < 200; ++row) {
            xml += "<row r=\"" + std::to_string(row + 1) + "\">\n";
            for (int col = 0; col < 3; ++col) {
                std::string cell_ref = utils::CommonUtils::cellReference(row, col);
                xml += "<c r=\"" + cell_ref + "\"><v>Sheet" + std::to_string(sheet) + "_" + std::to_string(row * col) + "</v></c>\n";
            }
            xml += "</row>\n";
        }
        xml += "</sheetData>\n</worksheet>";
        worksheets.push_back(xml);
    }
    
    measurePerformance("MultipleWorksheetsParsing", 3000, [&]() {
        size_t total_cells = 0;
        for (const auto& worksheet : worksheets) {
            size_t pos = 0;
            while ((pos = worksheet.find("<c r=", pos)) != std::string::npos) {
                total_cells++;
                pos += 5;
            }
        }
        volatile auto count = total_cells;
        (void)count;
    });
}

void ParsingPerformanceTest::testFormulaParsingSpeed() {
    // 创建包含公式的工作表
    std::string formula_xml = "<?xml version=\"1.0\"?>\n<worksheet>\n<sheetData>\n";
    for (int row = 0; row < 500; ++row) {
        formula_xml += "<row r=\"" + std::to_string(row + 1) + "\">\n";
        std::string cell_ref = utils::CommonUtils::cellReference(row, 0);
        formula_xml += "<c r=\"" + cell_ref + "\"><f>SUM(A1:A" + std::to_string(row + 1) + ")</f></c>\n";
        formula_xml += "</row>\n";
    }
    formula_xml += "</sheetData>\n</worksheet>";
    
    measurePerformance("FormulaParsingSpeed", 500, [&]() {
        size_t formula_count = 0;
        size_t pos = 0;
        while ((pos = formula_xml.find("<f>", pos)) != std::string::npos) {
            formula_count++;
            pos += 3;
        }
        volatile auto count = formula_count;
        (void)count;
    });
}

void ParsingPerformanceTest::testComplexFormulaParsing() {
    // 创建复杂公式测试数据
    std::vector<std::string> complex_formulas = {
        "IF(AND(A1>0,B1<100),SUM(C1:C10)*0.1,AVERAGE(D1:D20))",
        "VLOOKUP(E1,Sheet2!A:B,2,FALSE)",
        "INDEX(MATCH(F1,G:G,0),MATCH(H1,1:1,0))",
        "SUMIFS(I:I,J:J,\">\"&K1,L:L,\"<\"&M1)",
        "=CONCATENATE(\"Hello \",N1,\" World \",O1)"
    };
    
    measurePerformance("ComplexFormulaParsing", 2500, [&]() {
        for (const auto& formula : complex_formulas) {
            // 模拟复杂公式解析
            size_t func_count = 0;
            for (const std::string& func : {"IF", "AND", "SUM", "VLOOKUP", "INDEX", "MATCH", "SUMIFS", "CONCATENATE"}) {
                if (formula.find(func) != std::string::npos) {
                    func_count++;
                }
            }
            volatile auto count = func_count;
            (void)count;
        }
    });
}

void ParsingPerformanceTest::testSharedFormulaParsing() {
    // 创建共享公式XML
    std::string shared_formula_xml = "<?xml version=\"1.0\"?>\n<worksheet>\n<sheetData>\n";
    for (int row = 0; row < 100; ++row) {
        shared_formula_xml += "<row r=\"" + std::to_string(row + 1) + "\">\n";
        std::string cell_ref = utils::CommonUtils::cellReference(row, 0);
        if (row == 0) {
            shared_formula_xml += "<c r=\"" + cell_ref + "\"><f t=\"shared\" si=\"0\" ref=\"A1:A100\">A1+B1</f></c>\n";
        } else {
            shared_formula_xml += "<c r=\"" + cell_ref + "\"><f t=\"shared\" si=\"0\"/></c>\n";
        }
        shared_formula_xml += "</row>\n";
    }
    shared_formula_xml += "</sheetData>\n</worksheet>";
    
    measurePerformance("SharedFormulaParsing", 100, [&]() {
        size_t shared_count = 0;
        size_t pos = 0;
        while ((pos = shared_formula_xml.find("t=\"shared\"", pos)) != std::string::npos) {
            shared_count++;
            pos += 10;
        }
        volatile auto count = shared_count;
        (void)count;
    });
}

void ParsingPerformanceTest::runAllTests() {
    std::cout << "\n🔍 开始解析性能测试..." << std::endl;
    
    testXMLParsingSpeed();
    testLargeXMLParsing();
    testStylesParsing();
    testComplexStylesParsing();
    testSharedStringsParsing();
    testWorksheetParsing();
    testMultipleWorksheetsParsing();
    testFormulaParsingSpeed();
    testComplexFormulaParsing();
    testSharedFormulaParsing();
    
    std::cout << "✅ 解析性能测试完成!" << std::endl;
    generateReport();
}

void ParsingPerformanceTest::prepareTestXMLData() {
    test_xml_data_["SimpleWorksheet"] = "<?xml version=\"1.0\"?><worksheet><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>";
    test_xml_data_["BasicStyles"] = "<?xml version=\"1.0\"?><styleSheet><fonts><font><sz val=\"11\"/></font></fonts></styleSheet>";
    test_xml_data_["SharedStrings"] = "<?xml version=\"1.0\"?><sst><si><t>Test</t></si></sst>";
}

// ========== SharedFormulaPerformanceTest 实现 ==========

void SharedFormulaPerformanceTest::testSharedFormulaCreation() {
    measurePerformance("SharedFormulaCreation", 100, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("shared_formula_creation_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("CreationTest");
        
        // 创建100个共享公式
        for (int i = 0; i < 10; ++i) {
            for (int j = 0; j < 10; ++j) {
                int start_row = i * 10;
                int start_col = j * 2;
                worksheet->createSharedFormula(start_row, start_col, start_row + 9, start_col, 
                                             "A" + std::to_string(start_row + 1) + "+B" + std::to_string(start_row + 1));
            }
        }
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::testLargeSharedFormulaCreation() {
    measurePerformance("LargeSharedFormulaCreation", 10000, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("large_shared_formula_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("LargeTest");
        
        // 创建大型共享公式（10000个单元格）
        worksheet->createSharedFormula(0, 0, 99, 99, "A1+B1");
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::testFormulaOptimizationSpeed() {
    createFormulaTestData(1000);
    
    measurePerformance("FormulaOptimizationSpeed", 1000, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("optimization_speed_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("OptimizationTest");
        
        // 添加相似公式
        for (const auto& [pos, formula] : test_formulas_) {
            worksheet->writeFormula(pos.first, pos.second, formula);
        }
        
        // 执行优化
        int optimized = worksheet->optimizeFormulas(3);
        volatile auto count = optimized;
        (void)count;
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::testBatchOptimization() {
    measurePerformance("BatchOptimization", 5000, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("batch_optimization_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("BatchTest");
        
        // 批量添加公式
        for (int row = 0; row < 100; ++row) {
            for (int col = 0; col < 5; ++col) {
                std::string formula = "A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1);
                worksheet->writeFormula(row, col + 2, formula);
            }
        }
        
        // 批量优化
        int optimized = worksheet->optimizeFormulas(3);
        volatile auto count = optimized;
        (void)count;
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::testPatternDetectionSpeed() {
    createFormulaTestData(2000);
    
    measurePerformance("PatternDetectionSpeed", 2000, [&]() {
        fastexcel::core::SharedFormulaManager manager;
        
        // 检测模式
        auto patterns = manager.detectSharedFormulaPatterns(test_formulas_);
        volatile auto count = patterns.size();
        (void)count;
    });
}

void SharedFormulaPerformanceTest::testComplexPatternDetection() {
    // 创建复杂模式的公式
    std::map<std::pair<int, int>, std::string> complex_formulas;
    
    // 模式1: SUM函数
    for (int i = 0; i < 50; ++i) {
        complex_formulas[{i, 0}] = "SUM(A" + std::to_string(i + 1) + ":A" + std::to_string(i + 10) + ")";
    }
    
    // 模式2: IF条件
    for (int i = 0; i < 30; ++i) {
        complex_formulas[{i, 1}] = "IF(B" + std::to_string(i + 1) + ">0,B" + std::to_string(i + 1) + "*2,0)";
    }
    
    // 模式3: VLOOKUP
    for (int i = 0; i < 20; ++i) {
        complex_formulas[{i, 2}] = "VLOOKUP(C" + std::to_string(i + 1) + ",Table1,2,FALSE)";
    }
    
    measurePerformance("ComplexPatternDetection", 100, [&]() {
        fastexcel::core::SharedFormulaManager manager;
        auto patterns = manager.detectSharedFormulaPatterns(complex_formulas);
        volatile auto count = patterns.size();
        (void)count;
    });
}

void SharedFormulaPerformanceTest::testMemoryUsageComparison() {
    measurePerformance("MemoryUsageComparison", 1000, [&]() {
        // 创建工作簿测试内存使用
        auto workbook1 = fastexcel::core::Workbook::create(fastexcel::core::Path("memory_test_normal.xlsx"));
        workbook1->open();
        auto worksheet1 = workbook1->addWorksheet("Normal");
        
        // 普通公式（不优化）
        for (int i = 0; i < 100; ++i) {
            worksheet1->writeFormula(i, 0, "A" + std::to_string(i + 1) + "+B" + std::to_string(i + 1));
        }
        
        auto workbook2 = fastexcel::core::Workbook::create(fastexcel::core::Path("memory_test_shared.xlsx"));
        workbook2->open();
        auto worksheet2 = workbook2->addWorksheet("Shared");
        
        // 共享公式
        worksheet2->createSharedFormula(0, 0, 99, 0, "A1+B1");
        
        workbook1->close();
        workbook2->close();
    });
}

void SharedFormulaPerformanceTest::testSharedFormulaXMLGeneration() {
    measurePerformance("SharedFormulaXMLGeneration", 1000, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("xml_generation_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("XMLTest");
        
        // 创建共享公式
        worksheet->createSharedFormula(0, 0, 99, 0, "A1*2");
        worksheet->createSharedFormula(0, 1, 99, 1, "B1+10");
        
        // 生成XML（模拟）
        std::string xml_output;
        worksheet->generateXML([&](const char* data, size_t size) {
            xml_output.append(data, size);
        });
        
        volatile auto size = xml_output.size();
        (void)size;
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::testFullOptimizationWorkflow() {
    measurePerformance("FullOptimizationWorkflow", 500, [&]() {
        auto workbook = fastexcel::core::Workbook::create(fastexcel::core::Path("full_workflow_test.xlsx"));
        workbook->open();
        auto worksheet = workbook->addWorksheet("WorkflowTest");
        
        // 1. 添加数据
        for (int i = 0; i < 50; ++i) {
            worksheet->writeNumber(i, 0, i + 1);
            worksheet->writeNumber(i, 1, (i + 1) * 2);
        }
        
        // 2. 添加普通公式
        for (int i = 0; i < 50; ++i) {
            worksheet->writeFormula(i, 2, "A" + std::to_string(i + 1) + "+B" + std::to_string(i + 1));
            worksheet->writeFormula(i, 3, "A" + std::to_string(i + 1) + "*B" + std::to_string(i + 1));
        }
        
        // 3. 分析优化潜力
        auto report = worksheet->analyzeFormulaOptimization();
        
        // 4. 执行优化
        int optimized = worksheet->optimizeFormulas(3);
        
        // 5. 保存文件
        workbook->save();
        
        volatile auto count = optimized + static_cast<int>(report.total_formulas);
        (void)count;
        
        workbook->close();
    });
}

void SharedFormulaPerformanceTest::runAllTests() {
    std::cout << "\n📊 开始共享公式性能测试..." << std::endl;
    
    testSharedFormulaCreation();
    testLargeSharedFormulaCreation();
    testFormulaOptimizationSpeed();
    testBatchOptimization();
    testPatternDetectionSpeed();
    testComplexPatternDetection();
    testMemoryUsageComparison();
    testSharedFormulaXMLGeneration();
    testFullOptimizationWorkflow();
    
    std::cout << "✅ 共享公式性能测试完成!" << std::endl;
    generateReport();
}

void SharedFormulaPerformanceTest::createFormulaTestData(size_t formula_count) {
    test_formulas_.clear();
    
    // 创建相似公式模式
    for (size_t i = 0; i < formula_count; ++i) {
        int row = static_cast<int>(i / 10);
        int col = static_cast<int>(i % 10);
        
        // 创建不同类型的公式模式
        if (i % 3 == 0) {
            test_formulas_[{row, col}] = "A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1);
        } else if (i % 3 == 1) {
            test_formulas_[{row, col}] = "C" + std::to_string(row + 1) + "*D" + std::to_string(row + 1);
        } else {
            test_formulas_[{row, col}] = "SUM(E" + std::to_string(row + 1) + ":E" + std::to_string(row + 5) + ")";
        }
    }
}

void ComprehensivePerformanceTestSuite::setupOutputDirectory() {
    // 创建输出目录的实现
    FASTEXCEL_LOG_DEBUG("设置性能测试输出目录: {}", output_directory_);
}

} // namespace test
} // namespace fastexcel