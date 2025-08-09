#include "fastexcel/read/ReadWorkbook.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <chrono>

namespace fastexcel::read {

// 构造函数
ReadWorkbook::ReadWorkbook(const core::Path& path)
    : path_(path)
    , reader_(std::make_unique<reader::XLSXReader>(path))
    , l1_cache_(1024)  // 1024个槽位的无锁环形缓冲区
    , l2_cache_(256)   // 256个条目的LRU缓存
    , cache_hits_(0)
    , cache_misses_(0) {
    
    // 打开文件
    auto result = reader_->open();
    if (result != core::ErrorCode::Ok) {
        throw std::runtime_error("Failed to open Excel file: " + path.string());
    }
    
    // 获取工作表名称列表
    reader_->getWorksheetNames(worksheet_names_);
    
    // 获取元数据
    reader_->getMetadata(metadata_);
    
    LOG_INFO("ReadWorkbook opened: {}, worksheets: {}", 
             path.string(), worksheet_names_.size());
}

// 析构函数
ReadWorkbook::~ReadWorkbook() {
    // 输出缓存统计信息
    if (cache_hits_ + cache_misses_ > 0) {
        double hit_rate = static_cast<double>(cache_hits_) / 
                         (cache_hits_ + cache_misses_) * 100.0;
        LOG_DEBUG("Cache statistics - Hits: {}, Misses: {}, Hit rate: {:.2f}%",
                 cache_hits_, cache_misses_, hit_rate);
    }
    
    // 关闭reader
    if (reader_) {
        reader_->close();
    }
}

// 获取工作表数量
size_t ReadWorkbook::getWorksheetCount() const {
    return worksheet_names_.size();
}

// 获取工作表名称列表
std::vector<std::string> ReadWorkbook::getWorksheetNames() const {
    return worksheet_names_;
}

// 按索引获取工作表
std::unique_ptr<IReadOnlyWorksheet> ReadWorkbook::getWorksheet(size_t index) const {
    if (index >= worksheet_names_.size()) {
        LOG_ERROR("Worksheet index out of range: {} >= {}", 
                 index, worksheet_names_.size());
        return nullptr;
    }
    
    return getWorksheet(worksheet_names_[index]);
}

// 按名称获取工作表
std::unique_ptr<IReadOnlyWorksheet> ReadWorkbook::getWorksheet(const std::string& name) const {
    // 检查工作表是否存在
    auto it = std::find(worksheet_names_.begin(), worksheet_names_.end(), name);
    if (it == worksheet_names_.end()) {
        LOG_ERROR("Worksheet not found: {}", name);
        return nullptr;
    }
    
    // 尝试从缓存获取
    auto cached = getFromCache(name);
    if (cached) {
        return cached;
    }
    
    // 创建新的工作表对象
    try {
        auto worksheet = std::make_unique<ReadWorksheet>(reader_.get(), name);
        
        // 添加到缓存
        addToCache(name, worksheet.get());
        
        return worksheet;
    } catch (const std::exception& e) {
        LOG_ERROR("Failed to create worksheet {}: {}", name, e.what());
        return nullptr;
    }
}

// 检查工作表是否存在
bool ReadWorkbook::hasWorksheet(const std::string& name) const {
    return std::find(worksheet_names_.begin(), worksheet_names_.end(), name) 
           != worksheet_names_.end();
}

// 获取元数据
reader::WorkbookMetadata ReadWorkbook::getMetadata() const {
    return metadata_;
}

// 获取文件路径
core::Path ReadWorkbook::getPath() const {
    return path_;
}

// 获取访问模式
WorkbookAccessMode ReadWorkbook::getAccessMode() const {
    return WorkbookAccessMode::READ_ONLY;
}

// 从缓存获取工作表
std::unique_ptr<IReadOnlyWorksheet> ReadWorkbook::getFromCache(const std::string& name) const {
    // 先尝试L1缓存（无锁环形缓冲区）
    size_t hash = std::hash<std::string>{}(name);
    size_t slot = hash % l1_cache_.size();
    
    auto& entry = l1_cache_[slot];
    if (entry.name == name && entry.worksheet) {
        cache_hits_++;
        LOG_TRACE("L1 cache hit for worksheet: {}", name);
        
        // 克隆工作表对象
        return entry.worksheet->clone();
    }
    
    // 再尝试L2缓存（LRU缓存）
    auto it = l2_cache_.find(name);
    if (it != l2_cache_.end()) {
        cache_hits_++;
        LOG_TRACE("L2 cache hit for worksheet: {}", name);
        
        // 提升到L1缓存
        l1_cache_[slot] = {name, it->second};
        
        // 更新LRU顺序
        l2_lru_order_.remove(name);
        l2_lru_order_.push_front(name);
        
        return it->second->clone();
    }
    
    cache_misses_++;
    return nullptr;
}

// 添加到缓存
void ReadWorkbook::addToCache(const std::string& name, ReadWorksheet* worksheet) const {
    // 添加到L1缓存
    size_t hash = std::hash<std::string>{}(name);
    size_t slot = hash % l1_cache_.size();
    l1_cache_[slot] = {name, worksheet};
    
    // 添加到L2缓存
    if (l2_cache_.size() >= l2_cache_.max_size()) {
        // 移除最久未使用的条目
        if (!l2_lru_order_.empty()) {
            std::string oldest = l2_lru_order_.back();
            l2_lru_order_.pop_back();
            l2_cache_.erase(oldest);
            LOG_TRACE("Evicted from L2 cache: {}", oldest);
        }
    }
    
    l2_cache_[name] = worksheet;
    l2_lru_order_.push_front(name);
    
    LOG_TRACE("Added to cache: {}", name);
}

// 预加载工作表到缓存
void ReadWorkbook::preloadWorksheets(const std::vector<std::string>& names) {
    for (const auto& name : names) {
        if (hasWorksheet(name)) {
            // 触发加载并缓存
            auto ws = getWorksheet(name);
            if (ws) {
                LOG_DEBUG("Preloaded worksheet: {}", name);
            }
        }
    }
}

// 清空缓存
void ReadWorkbook::clearCache() {
    // 清空L1缓存
    for (auto& entry : l1_cache_) {
        entry = CacheEntry();
    }
    
    // 清空L2缓存
    l2_cache_.clear();
    l2_lru_order_.clear();
    
    // 重置统计
    cache_hits_ = 0;
    cache_misses_ = 0;
    
    LOG_DEBUG("Cache cleared");
}

// 获取缓存统计信息
ReadWorkbook::CacheStats ReadWorkbook::getCacheStats() const {
    CacheStats stats;
    stats.l1_size = 0;
    stats.l2_size = l2_cache_.size();
    stats.hits = cache_hits_;
    stats.misses = cache_misses_;
    
    // 计算L1缓存实际使用量
    for (const auto& entry : l1_cache_) {
        if (entry.worksheet != nullptr) {
            stats.l1_size++;
        }
    }
    
    if (stats.hits + stats.misses > 0) {
        stats.hit_rate = static_cast<double>(stats.hits) / 
                        (stats.hits + stats.misses);
    } else {
        stats.hit_rate = 0.0;
    }
    
    return stats;
}

// 获取内存使用估算
size_t ReadWorkbook::getEstimatedMemoryUsage() const {
    size_t total = sizeof(*this);
    
    // Reader内存
    if (reader_) {
        // 估算reader占用的内存
        total += sizeof(*reader_) + 1024 * 1024; // 假设1MB基础占用
    }
    
    // 缓存内存
    total += l1_cache_.size() * sizeof(CacheEntry);
    total += l2_cache_.size() * (sizeof(std::string) + sizeof(ReadWorksheet*));
    
    // 工作表名称列表
    for (const auto& name : worksheet_names_) {
        total += name.capacity();
    }
    
    return total;
}

// 验证文件完整性
bool ReadWorkbook::validateIntegrity() const {
    try {
        // 检查所有工作表是否可访问
        for (const auto& name : worksheet_names_) {
            auto ws = reader_->getWorksheet(name);
            if (!ws) {
                LOG_ERROR("Worksheet {} is not accessible", name);
                return false;
            }
        }
        
        // 检查共享字符串表
        auto sst = reader_->getSharedStrings();
        if (!sst && metadata_.has_shared_strings) {
            LOG_ERROR("Shared strings table is missing");
            return false;
        }
        
        // 检查样式表
        auto styles = reader_->getStylesLazy();
        if (!styles && metadata_.has_styles) {
            LOG_ERROR("Styles table is missing");
            return false;
        }
        
        LOG_INFO("File integrity check passed");
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Integrity check failed: {}", e.what());
        return false;
    }
}

// 导出为JSON（用于调试）
std::string ReadWorkbook::toJSON() const {
    std::stringstream json;
    json << "{\n";
    json << "  \"path\": \"" << path_.string() << "\",\n";
    json << "  \"worksheets\": [\n";
    
    for (size_t i = 0; i < worksheet_names_.size(); ++i) {
        json << "    \"" << worksheet_names_[i] << "\"";
        if (i < worksheet_names_.size() - 1) {
            json << ",";
        }
        json << "\n";
    }
    
    json << "  ],\n";
    json << "  \"metadata\": {\n";
    json << "    \"has_shared_strings\": " 
         << (metadata_.has_shared_strings ? "true" : "false") << ",\n";
    json << "    \"has_styles\": " 
         << (metadata_.has_styles ? "true" : "false") << ",\n";
    json << "    \"worksheet_count\": " << metadata_.worksheet_count << "\n";
    json << "  },\n";
    
    auto stats = getCacheStats();
    json << "  \"cache_stats\": {\n";
    json << "    \"l1_size\": " << stats.l1_size << ",\n";
    json << "    \"l2_size\": " << stats.l2_size << ",\n";
    json << "    \"hits\": " << stats.hits << ",\n";
    json << "    \"misses\": " << stats.misses << ",\n";
    json << "    \"hit_rate\": " << stats.hit_rate << "\n";
    json << "  }\n";
    json << "}";
    
    return json.str();
}

} // namespace fastexcel::read