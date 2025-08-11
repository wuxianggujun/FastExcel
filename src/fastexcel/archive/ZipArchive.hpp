#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipWriter.hpp"
#include <string>
#include <vector>
#include <memory>
#include <cstdint>
#include <string_view>

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

/**
 * @brief ZIP归档类 - 组合ZipReader和ZipWriter提供完整功能
 * 
 * 这个类组合了ZipReader和ZipWriter，提供了同时读写ZIP文件的能力。
 * 如果只需要读或写，建议直接使用ZipReader或ZipWriter以获得更好的性能。
 */
class ZipArchive {
public:
    // 使用ZipWriter的FileEntry
    using FileEntry = ZipWriter::FileEntry;
    
    // 构造/析构
    explicit ZipArchive(const core::Path& path);
    ~ZipArchive();
    
    // ========== 文件操作 ==========
    
    /**
     * 打开ZIP文件
     * @param create true=创建新文件（写模式），false=打开现有文件（读模式）
     * @return 是否成功
     */
    bool open(bool create = true);
    
    /**
     * 关闭ZIP文件
     * @return 是否成功
     */
    bool close();
    
    // ========== 写入操作（委托给ZipWriter） ==========
    
    ZipError addFile(std::string_view internal_path, std::string_view content);
    ZipError addFile(std::string_view internal_path, const uint8_t* data, size_t size);
    ZipError addFile(std::string_view internal_path, const void* data, size_t size);
    
    // 批量写入
    ZipError addFiles(const std::vector<FileEntry>& files);
    ZipError addFiles(std::vector<FileEntry>&& files);
    
    // 流式写入
    ZipError openEntry(std::string_view internal_path);
    ZipError writeChunk(const void* data, size_t size);
    ZipError closeEntry();
    
    // ========== 读取操作（委托给ZipReader） ==========
    
    ZipError extractFile(std::string_view internal_path, std::string& content);
    ZipError extractFile(std::string_view internal_path, std::vector<uint8_t>& data);
    ZipError extractFileToStream(std::string_view internal_path, std::ostream& output);
    ZipError fileExists(std::string_view internal_path) const;
    
    // 文件列表
    std::vector<std::string> listFiles() const;
    
    // ========== 状态查询 ==========
    
    bool isOpen() const { return is_open_; }
    bool isWritable() const { return mode_ == Mode::Write || mode_ == Mode::ReadWrite; }
    bool isReadable() const { return mode_ == Mode::Read || mode_ == Mode::ReadWrite; }
    
    // ========== 配置 ==========
    
    ZipError setCompressionLevel(int level);
    
    // ========== 直接访问底层对象 ==========
    
    /**
     * 获取ZipReader对象（如果可用）
     * @return ZipReader指针，如果不在读模式则返回nullptr
     */
    ZipReader* getReader() { return reader_.get(); }
    const ZipReader* getReader() const { return reader_.get(); }
    
    /**
     * 获取ZipWriter对象（如果可用）
     * @return ZipWriter指针，如果不在写模式则返回nullptr
     */
    ZipWriter* getWriter() { return writer_.get(); }
    const ZipWriter* getWriter() const { return writer_.get(); }
    
private:
    core::Path filepath_;
    std::unique_ptr<ZipReader> reader_;
    std::unique_ptr<ZipWriter> writer_;
    bool is_open_ = false;
    enum class Mode { None, Read, Write, ReadWrite } mode_ = Mode::None;
};

}} // namespace fastexcel::archive