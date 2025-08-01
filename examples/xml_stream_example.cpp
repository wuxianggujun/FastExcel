#include "fastexcel/xml/XMLStreamWriter.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <iostream>
#include <sstream>
#include <chrono>

using namespace fastexcel::xml;

/**
 * @brief 演示XMLStreamWriter和XMLStreamReader的配合使用
 * 
 * 这个示例展示了如何：
 * 1. 使用XMLStreamWriter创建Excel格式的XML文档
 * 2. 使用XMLStreamReader解析生成的XML文档
 * 3. 流式处理大量数据而不占用过多内存
 */

void demonstrateBasicUsage() {
    std::cout << "\n=== 基本用法演示 ===\n";
    
    // 1. 使用XMLStreamWriter创建XML
    XMLStreamWriter writer;
    writer.setBufferedMode();
    
    writer.startDocument();
    writer.startElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    writer.startElement("sheets");
    
    // 创建多个工作表
    for (int i = 1; i <= 3; ++i) {
        writer.writeEmptyElement("sheet");
        writer.writeAttribute("name", ("Sheet" + std::to_string(i)).c_str());
        writer.writeAttribute("sheetId", i);
        writer.writeAttribute("r:id", ("rId" + std::to_string(i)).c_str());
    }
    
    writer.endElement(); // sheets
    writer.endElement(); // workbook
    writer.endDocument();
    
    std::string xml_content = writer.toString();
    std::cout << "生成的XML内容:\n" << xml_content << "\n\n";
    
    // 2. 使用XMLStreamReader解析XML
    XMLStreamReader reader;
    
    std::vector<std::pair<std::string, int>> sheets;
    
    reader.setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        if (name == "sheet") {
            std::string sheet_name;
            int sheet_id = 0;
            
            for (const auto& attr : attributes) {
                if (attr.name == "name") {
                    sheet_name = attr.value;
                } else if (attr.name == "sheetId") {
                    sheet_id = std::stoi(attr.value);
                }
            }
            
            if (!sheet_name.empty() && sheet_id > 0) {
                sheets.emplace_back(sheet_name, sheet_id);
                std::cout << "解析到工作表: " << sheet_name << " (ID: " << sheet_id << ")\n";
            }
        }
    });
    
    XMLParseError result = reader.parseFromString(xml_content);
    if (result == XMLParseError::Ok) {
        std::cout << "成功解析XML，共找到 " << sheets.size() << " 个工作表\n";
    } else {
        std::cout << "解析失败: " << reader.getLastErrorMessage() << "\n";
    }
}

void demonstrateStreamProcessing() {
    std::cout << "\n=== 流式处理演示 ===\n";
    
    // 模拟处理大型Excel工作表
    const int ROWS = 1000;
    const int COLS = 10;
    
    std::cout << "生成包含 " << ROWS << " 行 " << COLS << " 列的工作表数据...\n";
    
    // 1. 流式写入大量数据
    XMLStreamWriter writer;
    writer.setBufferedMode();
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    writer.startDocument();
    writer.startElement("worksheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    
    writer.startElement("sheetData");
    
    for (int row = 1; row <= ROWS; ++row) {
        writer.startElement("row");
        writer.writeAttribute("r", row);
        
        for (int col = 1; col <= COLS; ++col) {
            writer.startElement("c");
            
            // 生成单元格引用 (A1, B1, C1, ...)
            std::string cell_ref;
            cell_ref += static_cast<char>('A' + col - 1);
            cell_ref += std::to_string(row);
            writer.writeAttribute("r", cell_ref.c_str());
            writer.writeAttribute("t", "inlineStr");
            
            writer.startElement("is");
            writer.startElement("t");
            
            std::string cell_value = "Cell " + std::to_string(row) + "," + std::to_string(col);
            writer.writeText(cell_value.c_str());
            
            writer.endElement(); // t
            writer.endElement(); // is
            writer.endElement(); // c
        }
        
        writer.endElement(); // row
    }
    
    writer.endElement(); // sheetData
    writer.endElement(); // worksheet
    writer.endDocument();
    
    auto write_end_time = std::chrono::high_resolution_clock::now();
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_end_time - start_time);
    
    std::string xml_content = writer.toString();
    std::cout << "XML生成完成，大小: " << xml_content.size() << " 字节，耗时: " << write_duration.count() << " ms\n";
    
    // 2. 流式解析大量数据
    XMLStreamReader reader;
    
    int cell_count = 0;
    int row_count = 0;
    
    reader.setStartElementCallback([&](const std::string& name, const std::vector<XMLAttribute>& attributes, int depth) {
        if (name == "row") {
            row_count++;
        } else if (name == "c") {
            cell_count++;
        }
    });
    
    auto parse_start_time = std::chrono::high_resolution_clock::now();
    
    XMLParseError result = reader.parseFromString(xml_content);
    
    auto parse_end_time = std::chrono::high_resolution_clock::now();
    auto parse_duration = std::chrono::duration_cast<std::chrono::milliseconds>(parse_end_time - parse_start_time);
    
    if (result == XMLParseError::Ok) {
        std::cout << "解析完成，找到 " << row_count << " 行，" << cell_count << " 个单元格，耗时: " << parse_duration.count() << " ms\n";
        std::cout << "解析速度: " << (xml_content.size() / 1024.0 / 1024.0) / (parse_duration.count() / 1000.0) << " MB/s\n";
    } else {
        std::cout << "解析失败: " << reader.getLastErrorMessage() << "\n";
    }
}

void demonstrateDOMParsing() {
    std::cout << "\n=== DOM解析演示 ===\n";
    
    // 创建一个复杂的XML结构
    std::string complex_xml = R"(<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="24816"/>
    <workbookPr defaultThemeVersion="166925"/>
    <sheets>
        <sheet name="Sales Data" sheetId="1" r:id="rId1"/>
        <sheet name="Summary" sheetId="2" r:id="rId2"/>
        <sheet name="Charts" sheetId="3" r:id="rId3"/>
    </sheets>
    <definedNames>
        <definedName name="Print_Area" localSheetId="0">'Sales Data'!$A$1:$H$100</definedName>
        <definedName name="Database" localSheetId="0">'Sales Data'!$A$1:$H$1000</definedName>
    </definedNames>
    <calcPr calcId="191029"/>
</workbook>)";
    
    XMLStreamReader reader;
    auto root = reader.parseToDOM(complex_xml);
    
    if (root) {
        std::cout << "成功解析为DOM结构\n";
        std::cout << "根元素: " << root->name << "\n";
        
        // 查找sheets元素
        auto sheets = root->findChild("sheets");
        if (sheets) {
            std::cout << "找到 " << sheets->children.size() << " 个工作表:\n";
            
            for (const auto& sheet : sheets->children) {
                if (sheet->name == "sheet") {
                    std::string name = sheet->getAttribute("name");
                    std::string id = sheet->getAttribute("sheetId");
                    std::cout << "  - " << name << " (ID: " << id << ")\n";
                }
            }
        }
        
        // 查找定义的名称
        auto definedNames = root->findChild("definedNames");
        if (definedNames) {
            std::cout << "找到 " << definedNames->children.size() << " 个定义的名称:\n";
            
            for (const auto& definedName : definedNames->children) {
                if (definedName->name == "definedName") {
                    std::string name = definedName->getAttribute("name");
                    std::string range = definedName->getTextContent();
                    std::cout << "  - " << name << ": " << range << "\n";
                }
            }
        }
    } else {
        std::cout << "DOM解析失败: " << reader.getLastErrorMessage() << "\n";
    }
}

void demonstrateErrorHandling() {
    std::cout << "\n=== 错误处理演示 ===\n";
    
    // 故意创建一个有错误的XML
    std::string invalid_xml = R"(<?xml version="1.0"?>
<workbook>
    <sheets>
        <sheet name="Sheet1" sheetId="1">
        <!-- 注意：这个sheet标签没有正确关闭 -->
    </sheets>
</workbook>)";
    
    XMLStreamReader reader;
    
    bool error_occurred = false;
    reader.setErrorCallback([&](XMLParseError error, const std::string& message, int line, int column) {
        error_occurred = true;
        std::cout << "解析错误 (第" << line << "行，第" << column << "列): " << message << "\n";
    });
    
    XMLParseError result = reader.parseFromString(invalid_xml);
    
    if (result != XMLParseError::Ok) {
        std::cout << "解析失败，错误代码: " << static_cast<int>(result) << "\n";
        std::cout << "错误信息: " << reader.getLastErrorMessage() << "\n";
    }
    
    if (error_occurred) {
        std::cout << "错误回调被正确调用\n";
    }
}

int main() {
    // 初始化日志系统
    fastexcel::Logger::getInstance().initialize("logs/xml_stream_example.log", 
                                               fastexcel::Logger::Level::INFO, 
                                               true);
    
    std::cout << "FastExcel XML流式处理示例\n";
    std::cout << "========================\n";
    
    try {
        demonstrateBasicUsage();
        demonstrateStreamProcessing();
        demonstrateDOMParsing();
        demonstrateErrorHandling();
        
        std::cout << "\n所有演示完成！\n";
        
    } catch (const std::exception& e) {
        std::cout << "发生异常: " << e.what() << "\n";
        return 1;
    }
    
    // 关闭日志系统
    fastexcel::Logger::getInstance().shutdown();
    
    return 0;
}