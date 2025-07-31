#pragma once

#include <string>
#include <vector>
#include <cstdint>

namespace fastexcel {
namespace archive {

class ZipArchive {
private:
    void* zip_handle_ = nullptr;
    void* unzip_handle_ = nullptr;
    std::string filename_;
    bool is_writable_ = false;
    bool is_readable_ = false;
    
public:
    explicit ZipArchive(const std::string& filename);
    ~ZipArchive();
    
    // 文件操作
    bool open(bool create = true);
    bool close();
    
    // 写入操作
    bool addFile(const std::string& internal_path, const std::string& content);
    bool addFile(const std::string& internal_path, const std::vector<uint8_t>& data);
    bool addFile(const std::string& internal_path, const void* data, size_t size);
    
    // 读取操作
    bool extractFile(const std::string& internal_path, std::string& content);
    bool extractFile(const std::string& internal_path, std::vector<uint8_t>& data);
    bool fileExists(const std::string& internal_path) const;
    
    // 文件列表
    std::vector<std::string> listFiles() const;
    
    // 获取状态
    bool isOpen() const { return is_writable_ || is_readable_; }
    bool isWritable() const { return is_writable_; }
    bool isReadable() const { return is_readable_; }
    
private:
    // 内部辅助函数
    void cleanup();
    bool initForWriting();
    bool initForReading();
};

}} // namespace fastexcel::archive