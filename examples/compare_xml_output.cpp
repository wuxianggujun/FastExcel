#include "fastexcel/fastexcel.hpp"
#include <iostream>
#include <fstream>
#include <filesystem>

using namespace fastexcel;

void createTestData(std::shared_ptr<Worksheet> worksheet) {
    // 创建相同的测试数据
    worksheet->writeString(0, 0, "Name");
    worksheet->writeString(0, 1, "Age");
    worksheet->writeString(0, 2, "City");
    
    worksheet->writeString(1, 0, "Alice");
    worksheet->writeNumber(1, 1, 25);
    worksheet->writeString(1, 2, "New York");
    
    worksheet->writeString(2, 0, "Bob");
    worksheet->writeNumber(2, 1, 30);
    worksheet->writeString(2, 2, "London");
    
    worksheet->writeString(3, 0, "Charlie");
    worksheet->writeNumber(3, 1, 35);
    worksheet->writeString(3, 2, "Tokyo");
}

bool extractWorksheetXML(const std::string& xlsxFile, const std::string& outputXmlFile) {
    try {
        // 使用系统命令解压ZIP文件中的worksheet XML
        std::string tempDir = "temp_extract_" + std::to_string(std::time(nullptr));
        std::filesystem::create_directory(tempDir);
        
        // 解压XLSX文件
        std::string unzipCmd = "cd " + tempDir + " && unzip -q ../" + xlsxFile;
        int result = system(unzipCmd.c_str());
        
        if (result != 0) {
            std::cout << "Failed to extract " << xlsxFile << std::endl;
            return false;
        }
        
        // 复制worksheet XML文件
        std::string worksheetPath = tempDir + "/xl/worksheets/sheet1.xml";
        if (std::filesystem::exists(worksheetPath)) {
            std::filesystem::copy_file(worksheetPath, outputXmlFile);
            
            // 清理临时目录
            std::filesystem::remove_all(tempDir);
            return true;
        } else {
            std::cout << "Worksheet XML not found in " << xlsxFile << std::endl;
            std::filesystem::remove_all(tempDir);
            return false;
        }
        
    } catch (const std::exception& e) {
        std::cout << "Exception extracting XML: " << e.what() << std::endl;
        return false;
    }
}

void compareFiles(const std::string& file1, const std::string& file2) {
    std::ifstream f1(file1);
    std::ifstream f2(file2);
    
    if (!f1.is_open() || !f2.is_open()) {
        std::cout << "Failed to open files for comparison" << std::endl;
        return;
    }
    
    std::string content1((std::istreambuf_iterator<char>(f1)), std::istreambuf_iterator<char>());
    std::string content2((std::istreambuf_iterator<char>(f2)), std::istreambuf_iterator<char>());
    
    f1.close();
    f2.close();
    
    std::cout << "\n=== XML Content Comparison ===" << std::endl;
    std::cout << "Batch mode XML size: " << content1.size() << " bytes" << std::endl;
    std::cout << "Streaming mode XML size: " << content2.size() << " bytes" << std::endl;
    
    if (content1 == content2) {
        std::cout << "✓ XML contents are IDENTICAL" << std::endl;
    } else {
        std::cout << "✗ XML contents are DIFFERENT" << std::endl;
        
        // 找出第一个不同的位置
        size_t diffPos = 0;
        size_t minSize = std::min(content1.size(), content2.size());
        
        for (size_t i = 0; i < minSize; ++i) {
            if (content1[i] != content2[i]) {
                diffPos = i;
                break;
            }
        }
        
        if (diffPos < minSize) {
            std::cout << "First difference at position " << diffPos << ":" << std::endl;
            
            // 显示差异上下文
            size_t start = (diffPos > 50) ? diffPos - 50 : 0;
            size_t end1 = std::min(diffPos + 50, content1.size());
            size_t end2 = std::min(diffPos + 50, content2.size());
            
            std::cout << "Batch mode:    \"" << content1.substr(start, end1 - start) << "\"" << std::endl;
            std::cout << "Streaming mode:\"" << content2.substr(start, end2 - start) << "\"" << std::endl;
        } else {
            std::cout << "Files have different lengths" << std::endl;
        }
    }
}

int main() {
    std::cout << "FastExcel XML Content Comparison Test" << std::endl;
    std::cout << "=====================================" << std::endl;
    
    try {
        // 创建批量模式文件
        std::cout << "\nCreating BATCH mode file..." << std::endl;
        {
            Workbook workbook("test_batch_xml.xlsx");
            workbook.setMode(WorkbookMode::BATCH);
            
            if (!workbook.open()) {
                std::cout << "Failed to open batch workbook" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook.addWorksheet("TestSheet");
            createTestData(worksheet);
            
            if (!workbook.save()) {
                std::cout << "Failed to save batch workbook" << std::endl;
                return 1;
            }
            
            workbook.close();
        }
        
        // 创建流模式文件
        std::cout << "Creating STREAMING mode file..." << std::endl;
        {
            Workbook workbook("test_streaming_xml.xlsx");
            workbook.setMode(WorkbookMode::STREAMING);
            
            if (!workbook.open()) {
                std::cout << "Failed to open streaming workbook" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook.addWorksheet("TestSheet");
            createTestData(worksheet);
            
            if (!workbook.save()) {
                std::cout << "Failed to save streaming workbook" << std::endl;
                return 1;
            }
            
            workbook.close();
        }
        
        std::cout << "Both files created successfully" << std::endl;
        
        // 提取XML内容进行比较
        std::cout << "\nExtracting XML content for comparison..." << std::endl;
        
        if (extractWorksheetXML("test_batch_xml.xlsx", "batch_worksheet.xml") &&
            extractWorksheetXML("test_streaming_xml.xlsx", "streaming_worksheet.xml")) {
            
            compareFiles("batch_worksheet.xml", "streaming_worksheet.xml");
            
        } else {
            std::cout << "Failed to extract XML files for comparison" << std::endl;
            std::cout << "Please manually extract and compare the worksheet XML files:" << std::endl;
            std::cout << "- test_batch_xml.xlsx -> xl/worksheets/sheet1.xml" << std::endl;
            std::cout << "- test_streaming_xml.xlsx -> xl/worksheets/sheet1.xml" << std::endl;
        }
        
    } catch (const std::exception& e) {
        std::cout << "Exception: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}