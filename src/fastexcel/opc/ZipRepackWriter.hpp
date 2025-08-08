#pragma once

#include "fastexcel/core/Path.hpp"
#include <memory>
#include <unordered_set>
#include <vector>

namespace fastexcel {
namespace opc {

// 前向声明
class ZipReader;

/**
 * @brief ZIP Repack写入器 - 专门处理repack操作
 * 
 * 核心功能：
 * 1. 支持从源ZIP原样复制（流式，不解压）
 * 2. 支持写入新内容
 * 3. 保证写入顺序和完整性
 */
class ZipRepackWriter {
public:
    explicit ZipRepackWriter(const core::Path& target_path);
    ~ZipRepackWriter();
    
    /**
     * 添加新内容到ZIP
     */
    bool add(const std::string& path, const std::string& content);
    bool add(const std::string& path, const void* data, size_t size);
    
    /**
     * 从源ZIP复制条目（高效流复制，保持原压缩）
     * @param source 源ZIP读取器
     * @param entry_path 要复制的条目路径
     * @return 是否成功
     */
    bool copyFrom(ZipReader* source, const std::string& entry_path);
    
    /**
     * 批量复制多个条目
     */
    bool copyBatch(ZipReader* source, const std::vector<std::string>& paths);
    
    /**
     * 检查条目是否已写入
     */
    bool hasEntry(const std::string& path) const {
        return written_entries_.count(path) > 0;
    }
    
    /**
     * 完成写入并关闭ZIP
     */
    bool finish();
    
    /**
     * 获取写入统计
     */
    struct Stats {
        size_t entries_added = 0;
        size_t entries_copied = 0;
        size_t total_size = 0;
    };
    Stats getStats() const { return stats_; }
    
private:
    class Impl;  // PIMPL模式，隐藏ZIP库细节
    std::unique_ptr<Impl> impl_;
    
    std::unordered_set<std::string> written_entries_;  // 已写入的条目
    Stats stats_;
    
    // 禁止拷贝
    ZipRepackWriter(const ZipRepackWriter&) = delete;
    ZipRepackWriter& operator=(const ZipRepackWriter&) = delete;
};

/**
 * @brief ZIP读取器 - 只读访问ZIP包
 */
class ZipReader {
public:
    explicit ZipReader(const core::Path& path);
    ~ZipReader();
    
    /**
     * 打开ZIP文件
     */
    bool open();
    
    /**
     * 关闭ZIP文件
     */
    void close();
    
    /**
     * 列出所有条目
     */
    std::vector<std::string> listEntries() const;
    
    /**
     * 检查条目是否存在
     */
    bool hasEntry(const std::string& path) const;
    
    /**
     * 读取条目内容
     */
    bool readEntry(const std::string& path, std::string& content) const;
    bool readEntry(const std::string& path, std::vector<uint8_t>& data) const;
    
    /**
     * 获取条目信息
     */
    struct EntryInfo {
        std::string path;
        size_t compressed_size;
        size_t uncompressed_size;
        uint32_t crc32;
        int compression_method;
    };
    bool getEntryInfo(const std::string& path, EntryInfo& info) const;
    
    /**
     * 获取原始压缩数据（用于流复制）
     */
    bool getRawEntry(const std::string& path, std::vector<uint8_t>& raw_data) const;
    
private:
    class Impl;
    std::unique_ptr<Impl> impl_;
    
    // 禁止拷贝
    ZipReader(const ZipReader&) = delete;
    ZipReader& operator=(const ZipReader&) = delete;
};

}} // namespace fastexcel::opc
