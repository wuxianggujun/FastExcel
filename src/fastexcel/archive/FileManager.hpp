#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/Path.hpp"
#include <string>
#include <memory>
#include <unordered_map>
#include <vector>

namespace fastexcel {
namespace archive {

class FileManager {
private:
    std::unique_ptr<ZipArchive> archive_;
    std::string filename_;  // 保留用于日志
    core::Path filepath_;   // 用于实际文件操作
    
public:
    explicit FileManager(const core::Path& path);
    ~FileManager();
    
    // 文件操作
    bool open(bool create = true);
    bool close();
    
    // 写入文件
    bool writeFile(const std::string& internal_path, const std::string& content);
    bool writeFile(const std::string& internal_path, const std::vector<uint8_t>& data);
    
    // 批量写入文件 - 高性能模式
    bool writeFiles(const std::vector<std::pair<std::string, std::string>>& files);
    bool writeFiles(std::vector<std::pair<std::string, std::string>>&& files); // 移动语义版本
    
    // 流式写入文件 - 极致性能模式，直接写入ZIP
    bool openStreamingFile(const std::string& internal_path);
    bool writeStreamingChunk(const void* data, size_t size);
    bool writeStreamingChunk(const std::string& data);
    bool closeStreamingFile();
    
    // 读取文件
    bool readFile(const std::string& internal_path, std::string& content);
    bool readFile(const std::string& internal_path, std::vector<uint8_t>& data);
    
    // 检查文件是否存在
    bool fileExists(const std::string& internal_path) const;
    
    // 获取文件列表
    std::vector<std::string> listFiles() const;
    
    // 获取状态
    bool isOpen() const { return archive_ && archive_->isOpen(); }
    
    // 压缩设置
    bool setCompressionLevel(int level);
    
    // Excel文件结构管理 - 公开给Workbook使用
    bool addContentTypes();
    bool addWorkbookRels();
    bool addRootRels();

    // 从现有包中复制未修改的条目（编辑模式保真写回）
    // skip_prefixes: 以这些前缀开头的路径将被跳过，不复制（因为将由新的生成逻辑覆盖）
    bool copyFromExistingPackage(const core::Path& source_package,
                                 const std::vector<std::string>& skip_prefixes);
    
private:
    bool createExcelStructure();
    bool addDocProps();
};

}} // namespace fastexcel::archive
