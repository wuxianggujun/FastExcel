#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/opc/ZipRepackWriter.hpp"
#include "fastexcel/archive/ZipArchive.hpp" 
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace opc {

// ========== ZipRepackWriter 实现 ==========

class ZipRepackWriter::Impl {
public:
    archive::ZipArchive zip_;
    core::Path path_;
    bool is_open_ = false;
    
    explicit Impl(const core::Path& path) : zip_(path), path_(path) {}
    
    bool open() {
        is_open_ = zip_.open(true);  // true = 创建模式
        if (is_open_) {
            FASTEXCEL_LOG_DEBUG("Created ZIP for repack: {}", path_.string());
        }
        return is_open_;
    }
};

ZipRepackWriter::ZipRepackWriter(const core::Path& target_path)
    : impl_(std::make_unique<Impl>(target_path)) {
    impl_->open();
}

ZipRepackWriter::~ZipRepackWriter() {
    finish();
}

bool ZipRepackWriter::add(const std::string& path, const std::string& content) {
    if (written_entries_.count(path)) {
        FASTEXCEL_LOG_DEBUG("Entry already written: {}", path);
        return true;
    }
    
    if (impl_->zip_.addFile(path, content) == archive::ZipError::Ok) {
        written_entries_.insert(path);
        stats_.entries_added++;
        stats_.total_size += content.size();
        FASTEXCEL_LOG_DEBUG("Added entry: {} ({} bytes)", path, content.size());
        return true;
    }
    
    FASTEXCEL_LOG_ERROR("Failed to add entry: {}", path);
    return false;
}

bool ZipRepackWriter::add(const std::string& path, const void* data, size_t size) {
    if (written_entries_.count(path)) {
        FASTEXCEL_LOG_DEBUG("Entry already written: {}", path);
        return true;
    }
    
    if (impl_->zip_.addFile(path, data, size) == archive::ZipError::Ok) {
        written_entries_.insert(path);
        stats_.entries_added++;
        stats_.total_size += size;
        FASTEXCEL_LOG_DEBUG("Added entry: {} ({} bytes)", path, size);
        return true;
    }
    
    FASTEXCEL_LOG_ERROR("Failed to add entry: {}", path);
    return false;
}

bool ZipRepackWriter::copyFrom(archive::ZipReader* source, const std::string& entry_path) {
    // Check if file exists using fileExists
    if (!source || source->fileExists(entry_path) != archive::ZipError::Ok) {
        FASTEXCEL_LOG_ERROR("Source entry not found: {}", entry_path);
        return false;
    }
    
    if (written_entries_.count(entry_path)) {
        FASTEXCEL_LOG_DEBUG("Entry already written: {}", entry_path);
        return true;
    }
    
    // 读取源数据 using extractFile
    std::vector<uint8_t> data;
    if (source->extractFile(entry_path, data) != archive::ZipError::Ok) {
        FASTEXCEL_LOG_ERROR("Failed to read source entry: {}", entry_path);
        return false;
    }
    
    // 写入目标
    if (add(entry_path, data.data(), data.size())) {
        stats_.entries_copied++;
        stats_.entries_added--;  // 修正统计
        FASTEXCEL_LOG_DEBUG("Copied entry: {} ({} bytes)", entry_path, data.size());
        return true;
    }
    
    return false;
}

bool ZipRepackWriter::copyBatch(archive::ZipReader* source, const std::vector<std::string>& paths) {
    FASTEXCEL_LOG_DEBUG("Batch copying {} entries", paths.size());
    
    // 可以使用批量写入优化
    std::vector<std::pair<std::string, std::vector<uint8_t>>> batch_data;
    batch_data.reserve(paths.size());
    
    for (const auto& path : paths) {
        if (written_entries_.count(path)) {
            continue;
        }
        
        std::vector<uint8_t> data;
        // Use extractFile instead of readEntry
        if (source->extractFile(path, data) == archive::ZipError::Ok) {
            batch_data.emplace_back(path, std::move(data));
        } else {
            FASTEXCEL_LOG_WARN("Failed to read entry for batch copy: {}", path);
        }
    }
    
    // 批量写入
    for (auto& [path, data] : batch_data) {
        if (impl_->zip_.addFile(path, data.data(), data.size()) == archive::ZipError::Ok) {
            written_entries_.insert(path);
            stats_.entries_copied++;
            stats_.total_size += data.size();
        } else {
            FASTEXCEL_LOG_ERROR("Failed to write entry in batch: {}", path);
            return false;
        }
    }
    
    FASTEXCEL_LOG_DEBUG("Batch copied {} entries", batch_data.size());
    return true;
}

bool ZipRepackWriter::finish() {
    if (!impl_->is_open_) return true;
    
    bool result = impl_->zip_.close();
    impl_->is_open_ = false;
    
    if (result) {
        FASTEXCEL_LOG_INFO("Repack finished: {} entries added, {} entries copied, {} bytes total",
                 stats_.entries_added, stats_.entries_copied, stats_.total_size);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to finalize repacked ZIP");
    }
    
    return result;
}

}} // namespace fastexcel::opc
