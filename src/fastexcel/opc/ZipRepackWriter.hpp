#include "fastexcel/utils/Logger.hpp"
#pragma once

#include "fastexcel/core/Path.hpp"
#include <memory>
#include <unordered_set>
#include <vector>

namespace fastexcel {

// 前向声明
namespace archive {
    class ZipReader;
    class ZipWriter;
}

namespace opc {

/**
 * @brief ZIP Repack写入器 - 专门处理repack操作
 * 
 * 核心功能：
 * 1. 支持从源ZIP原样复制（流式，不解压）
 * 2. 支持写入新内容
 * 3. 保证写入顺序和完整性
 * 
 * @deprecated 请使用 archive::ZipWriter 类，它提供了更完整的功能
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
    bool copyFrom(archive::ZipReader* source, const std::string& entry_path);
    
    /**
     * 批量复制多个条目
     */
    bool copyBatch(archive::ZipReader* source, const std::vector<std::string>& paths);
    
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

}} // namespace fastexcel::opc
