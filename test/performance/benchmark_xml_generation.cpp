#include <gtest/gtest.h>
#include "PerformanceBenchmark.hpp"
#include "fastexcel/FastExcel.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"
#include <chrono>
#include <memory>
#include <sstream>

class XMLGenerationBenchmark : public PerformanceBenchmark {
protected:
    void SetUp() override {
        PerformanceBenchmark::SetUp();
        workbook_ = fastexcel::core::Workbook::create("xml_benchmark_test.xlsx");
        worksheet_ = workbook_->addSheet("XMLBenchmark");
    }

    void TearDown() override {
        if (workbook_) {
            workbook_->save();
        }
        PerformanceBenchmark::TearDown();
    }

    std::shared_ptr<fastexcel::core::Workbook> workbook_;
    std::shared_ptr<fastexcel::core::Worksheet> worksheet_;
};

// 测试XML生成性能
TEST_F(XMLGenerationBenchmark, XMLGenerationPerformance) {
    const int ROWS = 1000;
    const int COLS = 10;
    
    // 填充测试数据
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            if (col % 2 == 0) {
                worksheet_->setValue(row, col, static_cast<double>(row * col + 1.5));
            } else {
                worksheet_->setValue(row, col, "Cell_" + std::to_string(row) + "_" + std::to_string(col));
            }
        }
    }
    
    std::string xml_output;
    size_t total_size = 0;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 使用标准XML生成
    worksheet_->generateXML([&](const char* data, size_t size) {
        xml_output.append(data, size);
        total_size += size;
    });
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "⚡ 生成 " << ROWS << "x" << COLS << " 工作表XML耗时: " 
              << duration.count() << " 毫秒，生成 " << total_size << " 字节" << std::endl;
    
    EXPECT_GT(total_size, 0);
    EXPECT_FALSE(xml_output.empty());
}

// 测试大数据XML生成性能
TEST_F(XMLGenerationBenchmark, LargeDataXMLPerformance) {
    const int ROWS = 2000;
    const int COLS = 10;
    
    // 填充测试数据
    for (int row = 0; row < ROWS; ++row) {
        for (int col = 0; col < COLS; ++col) {
            if (col % 3 == 0) {
                worksheet_->setValue(row, col, static_cast<double>(row + col));
            } else if (col % 3 == 1) {
                worksheet_->setValue(row, col, "Data" + std::to_string(row));
            } else {
                worksheet_->getCell(row, col).setFormula("A" + std::to_string(row + 1) + "+B" + std::to_string(row + 1));
            }
        }
    }
    
    std::string xml_output;
    size_t total_size = 0;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 使用标准XML生成
    worksheet_->generateXML([&](const char* data, size_t size) {
        xml_output.append(data, size);
        total_size += size;
    });
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "📦 生成大数据 " << ROWS << "x" << COLS << " 工作表XML耗时: " 
              << duration.count() << " 毫秒，生成 " << total_size << " 字节" << std::endl;
    
    EXPECT_GT(total_size, 0);
    EXPECT_FALSE(xml_output.empty());
}

// 测试带共享公式的XML生成性能
TEST_F(XMLGenerationBenchmark, SharedFormulaXMLPerformance) {
    const int ROWS = 500;
    
    // 添加基础数据
    for (int i = 0; i < ROWS; ++i) {
        worksheet_->setValue(i, 0, static_cast<double>(i + 1));
        worksheet_->setValue(i, 1, static_cast<double>((i + 1) * 2));
    }
    
    // 创建共享公式
    worksheet_->createSharedFormula(0, 2, ROWS - 1, 2, "A1+B1");
    worksheet_->createSharedFormula(0, 3, ROWS - 1, 3, "A1*B1");
    
    std::string xml_output;
    size_t total_size = 0;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 生成包含共享公式的XML
    worksheet_->generateXML([&](const char* data, size_t size) {
        xml_output.append(data, size);
        total_size += size;
    });
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "🔗 生成包含共享公式的XML耗时: " << duration.count() 
              << " 毫秒，生成 " << total_size << " 字节" << std::endl;
    
    // 验证XML包含共享公式标记
    EXPECT_NE(xml_output.find("t=\"shared\""), std::string::npos);
    EXPECT_GT(total_size, 0);
}

// 测试XMLStreamWriter性能
TEST_F(XMLGenerationBenchmark, XMLStreamWriterPerformance) {
    const int ELEMENT_COUNT = 10000;
    
    std::ostringstream output_stream;
    auto start = std::chrono::high_resolution_clock::now();
    
    {
        fastexcel::xml::XMLStreamWriter writer([&](const char* data, size_t size) {
            output_stream.write(data, size);
        });
        
        writer.startDocument();
        writer.startElement("worksheet");
        writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        
        writer.startElement("sheetData");
        
        // 写入大量XML元素
        for (int i = 0; i < ELEMENT_COUNT; ++i) {
            writer.startElement("row");
            writer.writeAttribute("r", std::to_string(i + 1).c_str());
            
            writer.startElement("c");
            writer.writeAttribute("r", ("A" + std::to_string(i + 1)).c_str());
            writer.startElement("v");
            writer.writeText(std::to_string(i).c_str());
            writer.endElement(); // v
            writer.endElement(); // c
            
            writer.endElement(); // row
        }
        
        writer.endElement(); // sheetData
        writer.endElement(); // worksheet
        writer.endDocument();
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::string result = output_stream.str();
    
    std::cout << "🚀 XMLStreamWriter 写入 " << ELEMENT_COUNT << " 个元素耗时: " 
              << duration.count() << " 微秒，生成 " << result.size() << " 字节" << std::endl;
    
    EXPECT_GT(result.size(), 0);
    EXPECT_NE(result.find("<worksheet"), std::string::npos);
    EXPECT_NE(result.find("</worksheet>"), std::string::npos);
}