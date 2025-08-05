/**
 * @file zip_streaming_test.cpp
 * @brief ZIP流式写入修复测试
 *
 * 专门测试openEntry()/writeChunk()路径，验证Excel兼容性修复
 */

#include "fastexcel/FastExcel.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <filesystem>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

using namespace fastexcel::archive;

// 生成大量测试数据
std::string generateLargeContent(size_t size) {
    std::string content;
    content.reserve(size);
    
    // 生成一些有意义的内容，而不是随机字符
    const std::string pattern = "This is a test line for streaming write functionality. ";
    while (content.size() < size) {
        content += pattern;
        if (content.size() + pattern.size() > size) {
            content += pattern.substr(0, size - content.size());
            break;
        }
    }
    
    return content;
}

// 测试流式写入功能
bool testStreamingWrite() {
    std::cout << "=== 测试流式写入功能 ===" << std::endl;
    
    std::string filename = "zip_streaming_test.xlsx";
    ZipArchive archive(filename);
    
    // 打开文件进行写入
    if (!archive.open(true)) {
        std::cerr << "无法打开ZIP文件进行写入" << std::endl;
        return false;
    }
    
    // 测试1：添加一个小文件（使用addFile路径）
    std::string smallContent = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>)";
    
    auto result = archive.addFile("[Content_Types].xml", smallContent);
    if (result != ZipError::Ok) {
        std::cerr << "添加小文件失败" << std::endl;
        return false;
    }
    std::cout << "✓ 成功添加小文件 ([Content_Types].xml)" << std::endl;
    
    // 测试2：使用流式写入添加一个大文件
    std::cout << "开始流式写入大文件..." << std::endl;
    
    auto streamResult = archive.openEntry("xl/sharedStrings.xml");
    if (streamResult != ZipError::Ok) {
        std::cerr << "打开流式条目失败" << std::endl;
        return false;
    }
    std::cout << "✓ 成功打开流式条目" << std::endl;
    
    // 生成并分块写入大内容（2MB）
    const size_t totalSize = 2 * 1024 * 1024;  // 2MB
    const size_t chunkSize = 64 * 1024;        // 64KB chunks
    
    std::string xmlHeader = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1000" uniqueCount="1000">
)";
    
    // 写入XML头部
    auto writeResult = archive.writeChunk(xmlHeader.data(), xmlHeader.size());
    if (writeResult != ZipError::Ok) {
        std::cerr << "写入XML头部失败" << std::endl;
        return false;
    }
    
    size_t writtenSize = xmlHeader.size();
    
    // 分块写入大量数据
    for (size_t i = 0; writtenSize < totalSize - 1000; ++i) {  // 留一些空间给结尾
        std::ostringstream oss;
        oss << "<si><t>String entry " << i << " with some additional content to make it larger</t></si>\n";
        std::string chunk = oss.str();
        
        // 确保不超过总大小
        if (writtenSize + chunk.size() > totalSize - 1000) {
            chunk = chunk.substr(0, totalSize - 1000 - writtenSize);
        }
        
        writeResult = archive.writeChunk(chunk.data(), chunk.size());
        if (writeResult != ZipError::Ok) {
            std::cerr << "写入数据块失败，块 " << i << std::endl;
            return false;
        }
        
        writtenSize += chunk.size();
        
        // 每写入1MB显示进度
        if (i % 1000 == 0) {
            std::cout << "已写入: " << writtenSize / 1024 << " KB" << std::endl;
        }
    }
    
    // 写入XML结尾
    std::string xmlFooter = "</sst>";
    writeResult = archive.writeChunk(xmlFooter.data(), xmlFooter.size());
    if (writeResult != ZipError::Ok) {
        std::cerr << "写入XML结尾失败" << std::endl;
        return false;
    }
    
    writtenSize += xmlFooter.size();
    std::cout << "✓ 流式写入完成，总大小: " << writtenSize / 1024 << " KB" << std::endl;
    
    // 关闭流式条目 - 这是关键步骤
    auto closeResult = archive.closeEntry();
    if (closeResult != ZipError::Ok) {
        std::cerr << "关闭流式条目失败！这是致命错误。" << std::endl;
        return false;
    }
    std::cout << "✓ 成功关闭流式条目" << std::endl;
    
    // 测试3：再添加一个小文件，确保流式写入后还能正常工作
    std::string relContent = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
    
    result = archive.addFile("_rels/.rels", relContent);
    if (result != ZipError::Ok) {
        std::cerr << "流式写入后添加文件失败" << std::endl;
        return false;
    }
    std::cout << "✓ 流式写入后成功添加文件" << std::endl;
    
    // 关键：显式关闭ZIP文件并检查返回值
    std::cout << "正在关闭ZIP文件..." << std::endl;
    if (!archive.close()) {
        std::cerr << "关闭ZIP文件失败！这意味着中央目录没有正确写入。" << std::endl;
        return false;
    }
    
    std::cout << "✓ ZIP文件成功关闭" << std::endl;
    return true;
}

int main() {
    
    if (!fastexcel::initialize("logs/zip_streaming_test.log", true)) {
        std::cerr << "Failed to initialize FastExcel library" << std::endl;
        return 1;
    }

    
    try {
        std::cout << "=== ZIP流式写入修复测试 ===" << std::endl;
        std::cout << "这个测试专门验证openEntry()/writeChunk()路径" << std::endl;
        std::cout << "目的：确保生成的ZIP文件能被Excel正常打开，无需修复" << std::endl;
        std::cout << std::endl;
        
        if (!testStreamingWrite()) {
            std::cerr << "流式写入测试失败！" << std::endl;
            return 1;
        }
        
        std::cout << std::endl;
        std::cout << "=== 测试成功完成 ===" << std::endl;
        std::cout << "生成的文件: zip_streaming_test.xlsx" << std::endl;
        std::cout << std::endl;
        std::cout << "=== 验证建议 ===" << std::endl;
        std::cout << "1. 用010 Editor打开文件，搜索十六进制 '504B0506' (EOCD签名)" << std::endl;
        std::cout << "2. 运行命令：unzip -t zip_streaming_test.xlsx" << std::endl;
        std::cout << "3. 运行命令：zip -T zip_streaming_test.xlsx" << std::endl;
        std::cout << "4. 用Excel打开文件，看是否还有修复提示" << std::endl;
        std::cout << "5. 检查文件是否包含大的sharedStrings.xml（约2MB）" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "发生异常: " << e.what() << std::endl;
        return 1;
    }
}
