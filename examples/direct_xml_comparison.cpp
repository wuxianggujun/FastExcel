#include "fastexcel/fastexcel.hpp"
#include <iostream>
#include <sstream>

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

std::string captureWorksheetXML(std::shared_ptr<Worksheet> worksheet) {
    std::string xmlContent;
    
    // 直接调用generateXML方法捕获XML内容
    worksheet->generateXML([&xmlContent](const char* data, size_t size) {
        xmlContent.append(data, size);
    });
    
    return xmlContent;
}

void compareXMLStrings(const std::string& batchXML, const std::string& streamingXML) {
    std::cout << "\n=== Direct XML Content Comparison ===" << std::endl;
    std::cout << "Batch mode XML size: " << batchXML.size() << " bytes" << std::endl;
    std::cout << "Streaming mode XML size: " << streamingXML.size() << " bytes" << std::endl;
    
    if (batchXML == streamingXML) {
        std::cout << "✓ XML contents are IDENTICAL" << std::endl;
        std::cout << "The XML generation logic is consistent between modes." << std::endl;
    } else {
        std::cout << "✗ XML contents are DIFFERENT" << std::endl;
        
        // 找出第一个不同的位置
        size_t diffPos = 0;
        size_t minSize = std::min(batchXML.size(), streamingXML.size());
        
        for (size_t i = 0; i < minSize; ++i) {
            if (batchXML[i] != streamingXML[i]) {
                diffPos = i;
                break;
            }
        }
        
        if (diffPos < minSize) {
            std::cout << "First difference at position " << diffPos << ":" << std::endl;
            
            // 显示差异上下文
            size_t start = (diffPos > 100) ? diffPos - 100 : 0;
            size_t end1 = std::min(diffPos + 100, batchXML.size());
            size_t end2 = std::min(diffPos + 100, streamingXML.size());
            
            std::cout << "\nBatch mode context:" << std::endl;
            std::cout << "\"" << batchXML.substr(start, end1 - start) << "\"" << std::endl;
            
            std::cout << "\nStreaming mode context:" << std::endl;
            std::cout << "\"" << streamingXML.substr(start, end2 - start) << "\"" << std::endl;
            
            // 显示具体的字符差异
            std::cout << "\nCharacter difference:" << std::endl;
            std::cout << "Batch:    '" << (char)batchXML[diffPos] << "' (0x" << std::hex << (int)(unsigned char)batchXML[diffPos] << ")" << std::endl;
            std::cout << "Streaming:'" << (char)streamingXML[diffPos] << "' (0x" << std::hex << (int)(unsigned char)streamingXML[diffPos] << ")" << std::dec << std::endl;
            
        } else {
            std::cout << "Files have different lengths" << std::endl;
            if (batchXML.size() > streamingXML.size()) {
                std::cout << "Batch mode has extra content: \"" << batchXML.substr(streamingXML.size()) << "\"" << std::endl;
            } else {
                std::cout << "Streaming mode has extra content: \"" << streamingXML.substr(batchXML.size()) << "\"" << std::endl;
            }
        }
    }
}

void analyzeXMLStructure(const std::string& xml, const std::string& modeName) {
    std::cout << "\n=== " << modeName << " XML Structure Analysis ===" << std::endl;
    
    // 检查XML声明
    if (xml.find("<?xml") == 0) {
        size_t declEnd = xml.find("?>");
        if (declEnd != std::string::npos) {
            std::string declaration = xml.substr(0, declEnd + 2);
            std::cout << "XML Declaration: " << declaration << std::endl;
        }
    }
    
    // 检查根元素
    size_t worksheetStart = xml.find("<worksheet");
    if (worksheetStart != std::string::npos) {
        size_t worksheetEnd = xml.find(">", worksheetStart);
        if (worksheetEnd != std::string::npos) {
            std::string rootElement = xml.substr(worksheetStart, worksheetEnd - worksheetStart + 1);
            std::cout << "Root element: " << rootElement << std::endl;
        }
    }
    
    // 统计主要元素
    size_t sheetDataPos = xml.find("<sheetData");
    size_t sheetDataEndPos = xml.find("</sheetData>");
    
    if (sheetDataPos != std::string::npos && sheetDataEndPos != std::string::npos) {
        std::string sheetDataContent = xml.substr(sheetDataPos, sheetDataEndPos - sheetDataPos + 12);
        
        // 统计行和单元格数量
        size_t rowCount = 0;
        size_t cellCount = 0;
        
        size_t pos = 0;
        while ((pos = sheetDataContent.find("<row", pos)) != std::string::npos) {
            rowCount++;
            pos++;
        }
        
        pos = 0;
        while ((pos = sheetDataContent.find("<c ", pos)) != std::string::npos) {
            cellCount++;
            pos++;
        }
        
        std::cout << "Rows: " << rowCount << ", Cells: " << cellCount << std::endl;
    }
}

int main() {
    std::cout << "FastExcel Direct XML Comparison Test" << std::endl;
    std::cout << "====================================" << std::endl;
    
    try {
        std::string batchXML, streamingXML;
        
        // 测试批量模式XML生成
        std::cout << "\nTesting BATCH mode XML generation..." << std::endl;
        {
            Workbook workbook("temp_batch.xlsx");
            workbook.setMode(WorkbookMode::BATCH);
            
            if (!workbook.open()) {
                std::cout << "Failed to open batch workbook" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook.addWorksheet("TestSheet");
            createTestData(worksheet);
            
            // 直接捕获XML内容
            batchXML = captureWorksheetXML(worksheet);
            
            workbook.close();
        }
        
        // 测试流模式XML生成
        std::cout << "Testing STREAMING mode XML generation..." << std::endl;
        {
            Workbook workbook("temp_streaming.xlsx");
            workbook.setMode(WorkbookMode::STREAMING);
            
            if (!workbook.open()) {
                std::cout << "Failed to open streaming workbook" << std::endl;
                return 1;
            }
            
            auto worksheet = workbook.addWorksheet("TestSheet");
            createTestData(worksheet);
            
            // 直接捕获XML内容
            streamingXML = captureWorksheetXML(worksheet);
            
            workbook.close();
        }
        
        // 分析XML结构
        analyzeXMLStructure(batchXML, "BATCH");
        analyzeXMLStructure(streamingXML, "STREAMING");
        
        // 比较XML内容
        compareXMLStrings(batchXML, streamingXML);
        
        // 保存XML内容到文件以便进一步分析
        std::ofstream batchFile("batch_direct.xml");
        batchFile << batchXML;
        batchFile.close();
        
        std::ofstream streamingFile("streaming_direct.xml");
        streamingFile << streamingXML;
        streamingFile.close();
        
        std::cout << "\nXML content saved to:" << std::endl;
        std::cout << "- batch_direct.xml" << std::endl;
        std::cout << "- streaming_direct.xml" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "Exception: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}