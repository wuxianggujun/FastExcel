#pragma once

#include <string>
#include <vector>
#include <cstdint>
#include <string_view>
#include <mutex>
#include <unordered_set>

namespace fastexcel {
namespace archive {

// 错误码枚举
enum class ZipError {
    Ok,                    // 操作成功
    NotOpen,               // ZIP 文件未打开
    IoFail,                // I/O 操作失败
    BadFormat,             // ZIP 格式错误
    TooLarge,              // 文件太大
    FileNotFound,          // 文件未找到
    InvalidParameter,      // 无效参数
    InternalError          // 内部错误
};

// 为 ZipError 提供 bool 转换操作符，使其可以在布尔上下文中使用
// 只有 ZipError::Ok 被视为 true，其他所有错误都被视为 false
constexpr bool operator!(ZipError error) noexcept {
    return error != ZipError::Ok;
}

// 显式 bool 转换函数，用于更清晰的语义
constexpr bool isSuccess(ZipError error) noexcept {
    return error == ZipError::Ok;
}

constexpr bool isError(ZipError error) noexcept {
    return error != ZipError::Ok;
}

class ZipArchive {
public:
    // 批量写入的文件条目结构
    struct FileEntry {
        std::string internal_path;
        std::string content;
        
        FileEntry() = default;
        
        // 移动构造函数
        FileEntry(std::string&& path, std::string&& data)
            : internal_path(std::move(path)), content(std::move(data)) {}
            
        // 拷贝构造函数
        FileEntry(const std::string& path, const std::string& data)
            : internal_path(path), content(data) {}
            
        // 混合构造函数
        FileEntry(std::string&& path, const std::string& data)
            : internal_path(std::move(path)), content(data) {}
            
        FileEntry(const std::string& path, std::string&& data)
            : internal_path(path), content(std::move(data)) {}
    };

private:
    void* zip_handle_ = nullptr;
    void* unzip_handle_ = nullptr;
    std::string filename_;
    bool is_writable_ = false;
    bool is_readable_ = false;
    bool stream_entry_open_ = false;  // 流式写入条目是否已打开
    int compression_level_ = 6;  // 压缩级别，默认为6
    mutable std::mutex mutex_;  // 线程安全互斥锁
    std::unordered_set<std::string> written_paths_;  // 跟踪已写入的路径，防止重复
    
public:
    explicit ZipArchive(const std::string& filename);
    ~ZipArchive();
    
    // 文件操作
    bool open(bool create = true);
    bool close();
    
    // 写入操作
    ZipError addFile(std::string_view internal_path, std::string_view content);
    ZipError addFile(std::string_view internal_path, const uint8_t* data, size_t size);
    ZipError addFile(std::string_view internal_path, const void* data, size_t size);
    
    // 批量写入操作 - 高性能模式
    ZipError addFiles(const std::vector<FileEntry>& files);
    ZipError addFiles(std::vector<FileEntry>&& files); // 移动语义版本
    
    // 流式写入操作 - 用于大文件
    ZipError openEntry(std::string_view internal_path);
    ZipError writeChunk(const void* data, size_t size);
    ZipError closeEntry();
    
    // 读取操作
    ZipError extractFile(std::string_view internal_path, std::string& content);
    ZipError extractFile(std::string_view internal_path, std::vector<uint8_t>& data);
    ZipError extractFileToStream(std::string_view internal_path, std::ostream& output);
    ZipError fileExists(std::string_view internal_path) const;
    
    // 文件列表
    std::vector<std::string> listFiles() const;
    
    // 获取状态
    bool isOpen() const { return is_writable_ || is_readable_; }
    bool isWritable() const { return is_writable_; }
    bool isReadable() const { return is_readable_; }
    
    // 压缩配置
    ZipError setCompressionLevel(int level);
    
private:
    // 内部辅助函数
    void cleanup();
    bool initForWriting();
    bool initForReading();
    
    // 文件操作辅助方法 - 使用 void* 避免头文件依赖
    void initializeFileInfo(void* file_info, const std::string& path, size_t size);
    ZipError writeFileEntry(const std::string& internal_path, const void* data, size_t size);
};

}} // namespace fastexcel::archive