#include "FormatRepository.hpp"
#include <algorithm>
#include <mutex>

namespace fastexcel {
namespace core {

FormatRepository::FormatRepository() {
    // 预留空间以减少重新分配
    formats_.reserve(128);
    hash_to_id_.reserve(128);
    
    // 添加默认格式作为ID 0
    auto default_format = std::make_shared<const FormatDescriptor>(
        FormatDescriptor::getDefault()
    );
    formats_.push_back(default_format);
    hash_to_id_[default_format->hash()] = DEFAULT_FORMAT_ID;
}

int FormatRepository::addFormat(const FormatDescriptor& format) {
    total_requests_.fetch_add(1, std::memory_order_relaxed);
    
    size_t format_hash = format.hash();
    
    // 首先尝试共享锁进行读取
    {
        std::shared_lock lock(mutex_);
        auto it = hash_to_id_.find(format_hash);
        if (it != hash_to_id_.end()) {
            // 找到了，但需要验证不是哈希冲突
            int existing_id = it->second;
            if (*formats_[existing_id] == format) {
                cache_hits_.fetch_add(1, std::memory_order_relaxed);
                return existing_id;
            }
        }
    }
    
    // 没有找到或存在哈希冲突，需要写锁
    std::unique_lock lock(mutex_);
    
    // 双重检查，防止在获取写锁期间被其他线程添加
    auto it = hash_to_id_.find(format_hash);
    if (it != hash_to_id_.end()) {
        int existing_id = it->second;
        if (*formats_[existing_id] == format) {
            cache_hits_.fetch_add(1, std::memory_order_relaxed);
            return existing_id;
        }
    }
    
    // 处理哈希冲突：线性搜索所有格式
    for (size_t i = 0; i < formats_.size(); ++i) {
        if (*formats_[i] == format) {
            cache_hits_.fetch_add(1, std::memory_order_relaxed);
            return static_cast<int>(i);
        }
    }
    
    // 真正的新格式，添加到仓储
    int new_id = static_cast<int>(formats_.size());
    auto format_ptr = std::make_shared<const FormatDescriptor>(format);
    formats_.push_back(format_ptr);
    hash_to_id_[format_hash] = new_id;
    
    return new_id;
}

std::shared_ptr<const FormatDescriptor> FormatRepository::getFormat(int id) const {
    std::shared_lock lock(mutex_);
    
    if (id < 0 || static_cast<size_t>(id) >= formats_.size()) {
        // 无效ID，返回默认格式
        return formats_[DEFAULT_FORMAT_ID];
    }
    
    return formats_[id];
}

std::shared_ptr<const FormatDescriptor> FormatRepository::getDefaultFormat() const {
    std::shared_lock lock(mutex_);
    return formats_[DEFAULT_FORMAT_ID];
}

void FormatRepository::clear() {
    std::unique_lock lock(mutex_);
    
    // 保留默认格式
    auto default_format = formats_[DEFAULT_FORMAT_ID];
    
    formats_.clear();
    hash_to_id_.clear();
    
    // 重新添加默认格式
    formats_.push_back(default_format);
    hash_to_id_[default_format->hash()] = DEFAULT_FORMAT_ID;
    
    // 重置统计
    total_requests_.store(0, std::memory_order_relaxed);
    cache_hits_.store(0, std::memory_order_relaxed);
}

double FormatRepository::getCacheHitRate() const {
    size_t total = total_requests_.load(std::memory_order_relaxed);
    size_t hits = cache_hits_.load(std::memory_order_relaxed);
    
    if (total == 0) return 0.0;
    return static_cast<double>(hits) / static_cast<double>(total);
}

FormatRepository::DeduplicationStats FormatRepository::getDeduplicationStats() const {
    size_t total = total_requests_.load(std::memory_order_relaxed);
    size_t unique = getFormatCount();
    
    double ratio = 0.0;
    if (total > 0) {
        ratio = 1.0 - (static_cast<double>(unique) / static_cast<double>(total));
    }
    
    return {total, unique, ratio};
}

size_t FormatRepository::getMemoryUsage() const {
    std::shared_lock lock(mutex_);
    
    size_t total_size = 0;
    
    // 估算formats_向量的内存使用
    total_size += formats_.capacity() * sizeof(std::shared_ptr<const FormatDescriptor>);
    
    // 估算hash_to_id_映射的内存使用
    total_size += hash_to_id_.size() * (sizeof(size_t) + sizeof(int));
    
    // 估算FormatDescriptor对象的内存使用
    total_size += formats_.size() * sizeof(FormatDescriptor);
    
    // 估算字符串的内存使用（粗略估算）
    for (const auto& format : formats_) {
        total_size += format->getFontName().capacity();
        total_size += format->getNumberFormat().capacity();
    }
    
    return total_size;
}

void FormatRepository::importFormats(const FormatRepository& source_repo, 
                                   std::unordered_map<int, int>& id_mapping) {
    id_mapping.clear();
    
    // 获取源仓储的所有格式
    std::shared_lock source_lock(source_repo.mutex_);
    auto source_formats = source_repo.formats_;
    source_lock.unlock();
    
    // 逐个导入格式
    for (size_t source_id = 0; source_id < source_formats.size(); ++source_id) {
        const auto& format = *source_formats[source_id];
        int target_id = addFormat(format);
        id_mapping[static_cast<int>(source_id)] = target_id;
    }
}

}} // namespace fastexcel::core
