/**
 * @file zip_directory_test.cpp
 * @brief ZIP目录打包测试
 * 
 * 测试addFile()路径，将完整的Excel目录结构打包为ZIP文件
 */

#include "fastexcel/archive/ZipArchive.hpp"
#include <iostream>
#include <string>
#include <filesystem>
#include <fstream>

using namespace fastexcel::archive;
namespace fs = std::filesystem;

// 递归读取目录并添加到ZIP
bool addDirectoryToZip(ZipArchive& archive, const fs::path& dirPath, const std::string& zipPrefix = "") {
    try {
        for (const auto& entry : fs::recursive_directory_iterator(dirPath)) {
            if (entry.is_regular_file()) {
                // 构建ZIP内部路径
                auto relativePath = fs::relative(entry.path(), dirPath);
                std::string zipPath = zipPrefix + relativePath.string();
                
                // 将反斜杠替换为正斜杠（ZIP标准）
                std::replace(zipPath.begin(), zipPath.end(), '\\', '/');
                
                // 读取文件内容
                std::ifstream file(entry.path(), std::ios::binary);
                if (!file) {
                    std::cerr << "无法读取文件: " << entry.path() << std::endl;
                    return false;
                }
                
                std::string content((std::istreambuf_iterator<char>(file)),
                                   std::istreambuf_iterator<char>());
                file.close();
                
                // 添加到ZIP
                auto result = archive.addFile(zipPath, content);
                if (result != ZipError::Ok) {
                    std::cerr << "添加文件失败: " << zipPath << std::endl;
                    return false;
                }
                
                std::cout << "已添加: " << zipPath << " (" << content.size() << " bytes)" << std::endl;
            }
        }
        return true;
    } catch (const std::exception& e) {
        std::cerr << "处理目录时出错: " << e.what() << std::endl;
        return false;
    }
}

int main() {
    try {
        std::cout << "=== ZIP目录打包测试 ===" << std::endl;
        std::cout << "这个测试验证addFile()路径（批量写入）" << std::endl;
        std::cout << std::endl;
        
        // 源目录和目标文件
        std::string sourceDir = "cmake-build-debug/bin/examples/simple_test";
        std::string outputFile = "zip_directory_test.xlsx";
        
        // 检查源目录是否存在
        if (!fs::exists(sourceDir)) {
            std::cerr << "源目录不存在: " << sourceDir << std::endl;
            std::cerr << "请先运行simple_test示例生成测试数据" << std::endl;
            return 1;
        }
        
        std::cout << "源目录: " << sourceDir << std::endl;
        std::cout << "输出文件: " << outputFile << std::endl;
        std::cout << std::endl;
        
        // 创建ZIP文件
        ZipArchive archive(outputFile);
        
        // 打开文件进行写入
        if (!archive.open(true)) {
            std::cerr << "无法打开ZIP文件进行写入" << std::endl;
            return 1;
        }
        
        std::cout << "开始添加文件..." << std::endl;
        
        // 添加整个目录
        if (!addDirectoryToZip(archive, sourceDir)) {
            std::cerr << "添加目录失败" << std::endl;
            return 1;
        }
        
        std::cout << "所有文件添加完成，正在关闭ZIP..." << std::endl;
        
        // 关键：显式关闭ZIP文件并检查返回值
        if (!archive.close()) {
            std::cerr << "关闭ZIP文件失败！这意味着中央目录没有正确写入。" << std::endl;
            std::cerr << "请检查日志以获取详细错误信息。" << std::endl;
            return 1;
        }
        
        std::cout << "✓ ZIP文件创建成功：" << outputFile << std::endl;
        
        // 验证文件大小
        if (fs::exists(outputFile)) {
            auto fileSize = fs::file_size(outputFile);
            std::cout << "文件大小: " << fileSize << " bytes" << std::endl;
            
            if (fileSize < 100) {
                std::cerr << "警告：文件大小异常小，可能创建失败" << std::endl;
                return 1;
            }
        }
        
        std::cout << std::endl;
        std::cout << "=== 验证建议 ===" << std::endl;
        std::cout << "1. 用010 Editor打开文件，搜索十六进制 '504B0506' (EOCD签名)" << std::endl;
        std::cout << "2. 运行命令：unzip -t " << outputFile << std::endl;
        std::cout << "3. 运行命令：zip -T " << outputFile << std::endl;
        std::cout << "4. 用Excel打开文件，看是否还有修复提示" << std::endl;
        std::cout << "5. 检查文件末尾是否有完整的中央目录结构" << std::endl;
        
        return 0;
        
    } catch (const std::exception& e) {
        std::cerr << "发生异常: " << e.what() << std::endl;
        return 1;
    }
}