#include "fastexcel/opc/ZipRepackWriter.hpp"
#include "fastexcel/archive/ZipArchive.hpp"  // 使用你现有的ZipArchive
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>

namespace fastexcel {
namespace opc {

// ========== ZipReader 实现 ==========

class ZipReader::Impl {
public:
    archive::ZipArchive zip_;
    core::Path path_;
    bool is_open_ = false;
    
    explicit Impl(const core::Path& path) : zip_(path), path_(path) {}
};

ZipReader::ZipReader(const core::Path& path)
    : impl_(std::make_unique<Impl>(path)) {
}

ZipReader::~ZipReader() {
    close();
}

bool ZipReader::open() {
    if (impl_->is_open_) return true;
    
    impl_->is_open_ = impl_->zip_.open(false);  // false = 只读模式
    if (impl_->is_open_) {
        LOG_DEBUG("Opened ZIP for reading: {}", impl_->path_.string());
    }
    return impl_->is_open_;
}

void ZipReader::close() {
    if (impl_->is_open_) {
        impl_->zip_.close();
        impl_->is_open_ = false;
    }
}

std::vector<std::string> ZipReader::listEntries() const {
    if (!impl_->is_open_) return {};
    return impl_->zip_.listFiles();
}

bool ZipReader::hasEntry(const std::string& path) const {
    if (!impl_->is_open_) return false;
    return impl_->zip_.fileExists(path) == archive::ZipError::Ok;
}

bool ZipReader::readEntry(const std::string& path, std::string& content) const {
    if (!impl_->is_open_) return false;
    return impl_->zip_.extractFile(path, content) == archive::ZipError::Ok;
}

bool ZipReader::readEntry(const std::string& path, std::vector<uint8_t>& data) const {
    if (!impl_->is_open_) return false;
    return impl_->zip_.extractFile(path, data) == archive::ZipError::Ok;
}

bool ZipReader::getEntryInfo(const std::string& path, EntryInfo& info) const {
    if (!impl_->is_open_) return false;
    
    // 简化实现，主要用于获取大小信息
    std::vector<uint8_t> data;
    if (impl_->zip_.extractFile(path, data) == archive::ZipError::Ok) {
        info.path = path;
        info.uncompressed_size = data.size();
        info.compressed_size = data.size();  // 简化
        info.crc32 = 0;  // TODO: 计算CRC
        info.compression_method = 8;  // DEFLATE
        return true;
    }
    return false;
}

bool ZipReader::getRawEntry(const std::string& path, std::vector<uint8_t>& raw_data) const {
    // 对于流复制，暂时使用解压后重压的方式
    // TODO: 实现真正的原始数据复制
    return readEntry(path, raw_data);
}

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
            LOG_DEBUG("Created ZIP for repack: {}", path_.string());
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
        LOG_DEBUG("Entry already written: {}", path);
        return true;
    }
    
    if (impl_->zip_.addFile(path, content) == archive::ZipError::Ok) {
        written_entries_.insert(path);
        stats_.entries_added++;
        stats_.total_size += content.size();
        LOG_DEBUG("Added entry: {} ({} bytes)", path, content.size());
        return true;
    }
    
    LOG_ERROR("Failed to add entry: {}", path);
    return false;
}

bool ZipRepackWriter::add(const std::string& path, const void* data, size_t size) {
    if (written_entries_.count(path)) {
        LOG_DEBUG("Entry already written: {}", path);
        return true;
    }
    
    if (impl_->zip_.addFile(path, data, size) == archive::ZipError::Ok) {
        written_entries_.insert(path);
        stats_.entries_added++;
        stats_.total_size += size;
        LOG_DEBUG("Added entry: {} ({} bytes)", path, size);
        return true;
    }
    
    LOG_ERROR("Failed to add entry: {}", path);
    return false;
}

bool ZipRepackWriter::copyFrom(ZipReader* source, const std::string& entry_path) {
    if (!source || !source->hasEntry(entry_path)) {
        LOG_ERROR("Source entry not found: {}", entry_path);
        return false;
    }
    
    if (written_entries_.count(entry_path)) {
        LOG_DEBUG("Entry already written: {}", entry_path);
        return true;
    }
    
    // 读取源数据
    std::vector<uint8_t> data;
    if (!source->readEntry(entry_path, data)) {
        LOG_ERROR("Failed to read source entry: {}", entry_path);
        return false;
    }
    
    // 写入目标
    if (add(entry_path, data.data(), data.size())) {
        stats_.entries_copied++;
        stats_.entries_added--;  // 修正统计
        LOG_DEBUG("Copied entry: {} ({} bytes)", entry_path, data.size());
        return true;
    }
    
    return false;
}

bool ZipRepackWriter::copyBatch(ZipReader* source, const std::vector<std::string>& paths) {
    LOG_DEBUG("Batch copying {} entries", paths.size());
    
    // 可以使用批量写入优化
    std::vector<std::pair<std::string, std::vector<uint8_t>>> batch_data;
    batch_data.reserve(paths.size());
    
    for (const auto& path : paths) {
        if (written_entries_.count(path)) {
            continue;
        }
        
        std::vector<uint8_t> data;
        if (source->readEntry(path, data)) {
            batch_data.emplace_back(path, std::move(data));
        } else {
            LOG_WARN("Failed to read entry for batch copy: {}", path);
        }
    }
    
    // 批量写入
    for (auto& [path, data] : batch_data) {
        if (impl_->zip_.addFile(path, data.data(), data.size()) == archive::ZipError::Ok) {
            written_entries_.insert(path);
            stats_.entries_copied++;
            stats_.total_size += data.size();
        } else {
            LOG_ERROR("Failed to write entry in batch: {}", path);
            return false;
        }
    }
    
    LOG_DEBUG("Batch copied {} entries", batch_data.size());
    return true;
}

bool ZipRepackWriter::finish() {
    if (!impl_->is_open_) return true;
    
    bool result = impl_->zip_.close();
    impl_->is_open_ = false;
    
    if (result) {
        LOG_INFO("Repack finished: {} entries added, {} entries copied, {} bytes total",
                 stats_.entries_added, stats_.entries_copied, stats_.total_size);
    } else {
        LOG_ERROR("Failed to finalize repacked ZIP");
    }
    
    return result;
}

}} // namespace fastexcel::opc
