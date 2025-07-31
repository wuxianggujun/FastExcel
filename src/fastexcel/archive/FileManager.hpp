#pragma once

#include "ZipArchive.h"
#include <string>
#include <memory>
#include <unordered_map>

namespace fastexcel {
namespace archive {

class FileManager {
private:
    std::unique_ptr<ZipArchive> archive_;
    std::string filename_;
    
public:
    explicit FileManager(const std::string& filename);
    ~FileManager();
    
    // 文件操作
    bool open(bool create = true);
    bool close();
    
    // 写入文件
    bool writeFile(const std::string& internal_path, const std::string& content);
    bool writeFile(const std::string& internal_path, const std::vector<uint8_t>& data);
    
    // 读取文件
    bool readFile(const std::string& internal_path, std::string& content);
    bool readFile(const std::string& internal_path, std::vector<uint8_t>& data);
    
    // 检查文件是否存在
    bool fileExists(const std::string& internal_path) const;
    
    // 获取文件列表
    std::vector<std::string> listFiles() const;
    
    // 获取状态
    bool isOpen() const { return archive_ && archive_->isOpen(); }
    
private:
    // Excel文件结构管理
    bool createExcelStructure();
    bool addContentTypes();
    bool addWorkbookRels();
    bool addRootRels();
    bool addDocProps();
};

}} // namespace fastexcel::archive