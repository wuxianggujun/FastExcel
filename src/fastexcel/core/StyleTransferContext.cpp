#include "StyleTransferContext.hpp"
#include <algorithm>
#include <set>

namespace fastexcel {
namespace core {

StyleTransferContext::StyleTransferContext(const FormatRepository& source_repo, 
                                          FormatRepository& target_repo)
    : source_repository_(source_repo), target_repository_(target_repo) {
    
    // 预留映射缓存空间
    id_mapping_.reserve(source_repository_.getFormatCount());
}

int StyleTransferContext::mapStyleId(int source_id) const {
    // 检查缓存
    auto it = id_mapping_.find(source_id);
    if (it != id_mapping_.end()) {
        return it->second;
    }
    
    // 缓存未命中，执行映射
    return mapStyleIdInternal(source_id);
}

std::vector<int> StyleTransferContext::mapStyleIds(const std::vector<int>& source_ids) const {
    std::vector<int> target_ids;
    target_ids.reserve(source_ids.size());
    
    for (int source_id : source_ids) {
        target_ids.push_back(mapStyleId(source_id));
    }
    
    return target_ids;
}

void StyleTransferContext::preloadAllMappings() const {
    if (bulk_imported_) {
        return; // 已经完成批量导入
    }
    
    // 使用FormatRepository的批量导入功能
    std::unordered_map<int, int> bulk_mapping;
    target_repository_.importFormats(source_repository_, bulk_mapping);
    
    // 更新缓存
    id_mapping_ = std::move(bulk_mapping);
    bulk_imported_ = true;
    
    // 更新统计信息
    transferred_count_ = id_mapping_.size();
    
    // 计算去重数量（目标仓储的格式数量与映射数量的差异）
    size_t target_count_after = target_repository_.getFormatCount();
    size_t source_count = source_repository_.getFormatCount();
    
    if (target_count_after < source_count) {
        deduplicated_count_ = source_count - target_count_after;
    }
}

bool StyleTransferContext::isValidSourceId(int source_id) const {
    return source_repository_.isValidFormatId(source_id);
}

StyleTransferContext::TransferStats StyleTransferContext::getTransferStats() const {
    size_t source_count = source_repository_.getFormatCount();
    size_t target_count = target_repository_.getFormatCount();
    
    double dedup_ratio = 0.0;
    if (transferred_count_ > 0) {
        dedup_ratio = static_cast<double>(deduplicated_count_) / 
                     static_cast<double>(transferred_count_);
    }
    
    return {
        source_count,
        target_count,
        transferred_count_,
        deduplicated_count_,
        dedup_ratio
    };
}

void StyleTransferContext::clearCache() const {
    id_mapping_.clear();
    bulk_imported_ = false;
    transferred_count_ = 0;
    deduplicated_count_ = 0;
}

const std::unordered_map<int, int>& StyleTransferContext::getIdMapping() const {
    if (!bulk_imported_ && id_mapping_.empty()) {
        preloadAllMappings();
    }
    return id_mapping_;
}

size_t StyleTransferContext::transferStyleRange(int start_id, int end_id) const {
    size_t transferred = 0;
    
    for (int id = start_id; id < end_id; ++id) {
        if (isValidSourceId(id)) {
            mapStyleId(id); // 触发传输
            ++transferred;
        }
    }
    
    return transferred;
}

size_t StyleTransferContext::transferAllStyles() const {
    preloadAllMappings(); // 使用批量传输优化
    return transferred_count_;
}

size_t StyleTransferContext::transferStyles(const std::vector<int>& source_ids) const {
    size_t transferred = 0;
    
    for (int source_id : source_ids) {
        if (isValidSourceId(source_id)) {
            auto it = id_mapping_.find(source_id);
            if (it == id_mapping_.end()) {
                mapStyleId(source_id); // 触发传输
                ++transferred;
            }
        }
    }
    
    return transferred;
}

int StyleTransferContext::mapStyleIdInternal(int source_id) const {
    // 检查源ID有效性
    if (!isValidSourceId(source_id)) {
        return target_repository_.getDefaultFormatId();
    }
    
    // 获取源格式
    auto source_format = source_repository_.getFormat(source_id);
    if (!source_format) {
        return target_repository_.getDefaultFormatId();
    }
    
    // 添加到目标仓储（自动去重）
    int target_id = target_repository_.addFormat(*source_format);
    
    // 缓存映射
    id_mapping_[source_id] = target_id;
    
    // 更新统计
    updateStats(true);
    
    return target_id;
}

void StyleTransferContext::updateStats(bool was_new_transfer) const {
    if (was_new_transfer) {
        ++transferred_count_;
    }
    
    // 简单的去重检测：如果目标ID小于源ID，说明发生了去重
    // （这是一个简化的启发式方法）
    // 实际的去重统计应该更复杂，但这里提供基本信息
}

// ========== StyleTransfer命名空间的辅助函数 ==========

namespace StyleTransfer {

std::vector<int> quickCopyStyles(
    const FormatRepository& source_repo,
    FormatRepository& target_repo,
    const std::vector<int>& source_ids) {
    
    std::vector<int> target_ids;
    target_ids.reserve(source_ids.size());
    
    for (int source_id : source_ids) {
        if (!source_repo.isValidFormatId(source_id)) {
            target_ids.push_back(target_repo.getDefaultFormatId());
            continue;
        }
        
        auto source_format = source_repo.getFormat(source_id);
        if (source_format) {
            int target_id = target_repo.addFormat(*source_format);
            target_ids.push_back(target_id);
        } else {
            target_ids.push_back(target_repo.getDefaultFormatId());
        }
    }
    
    return target_ids;
}

std::unique_ptr<FormatRepository> mergeRepositories(
    const FormatRepository& repo1,
    const FormatRepository& repo2) {
    
    auto merged_repo = std::make_unique<FormatRepository>();
    
    // 导入第一个仓储的所有格式
    std::unordered_map<int, int> mapping1;
    merged_repo->importFormats(repo1, mapping1);
    
    // 导入第二个仓储的所有格式（自动去重）
    std::unordered_map<int, int> mapping2;
    merged_repo->importFormats(repo2, mapping2);
    
    return merged_repo;
}

StyleDifference compareRepositories(
    const FormatRepository& repo1,
    const FormatRepository& repo2) {
    
    StyleDifference diff;
    
    // 创建格式哈希值到ID的映射
    std::unordered_map<size_t, std::vector<int>> hash_to_ids1;
    std::unordered_map<size_t, std::vector<int>> hash_to_ids2;
    
    // 收集仓储1的格式
    auto snapshot1 = repo1.createSnapshot();
    for (const auto& format_pair : snapshot1) {
        int id = format_pair.first;
        const auto& format = format_pair.second;
        size_t hash = format->hash();
        hash_to_ids1[hash].push_back(id);
    }
    
    // 收集仓储2的格式
    auto snapshot2 = repo2.createSnapshot();
    for (const auto& format_pair : snapshot2) {
        int id = format_pair.first;
        const auto& format = format_pair.second;
        size_t hash = format->hash();
        hash_to_ids2[hash].push_back(id);
    }
    
    // 查找公共格式
    std::set<size_t> processed_hashes;
    
    for (const auto& pair1 : hash_to_ids1) {
        size_t hash = pair1.first;
        const auto& ids1 = pair1.second;
        
        auto it2 = hash_to_ids2.find(hash);
        if (it2 != hash_to_ids2.end()) {
            // 找到公共格式
            const auto& ids2 = it2->second;
            
            // 创建配对（简化：一对一配对）
            size_t min_size = std::min(ids1.size(), ids2.size());
            for (size_t i = 0; i < min_size; ++i) {
                diff.common_styles.emplace_back(ids1[i], ids2[i]);
            }
            
            // 处理剩余的ID
            for (size_t i = min_size; i < ids1.size(); ++i) {
                diff.only_in_repo1.push_back(ids1[i]);
            }
            for (size_t i = min_size; i < ids2.size(); ++i) {
                diff.only_in_repo2.push_back(ids2[i]);
            }
            
            processed_hashes.insert(hash);
        } else {
            // 只在仓储1中存在
            diff.only_in_repo1.insert(diff.only_in_repo1.end(), ids1.begin(), ids1.end());
        }
    }
    
    // 查找只在仓储2中存在的格式
    for (const auto& pair2 : hash_to_ids2) {
        size_t hash = pair2.first;
        const auto& ids2 = pair2.second;
        
        if (processed_hashes.find(hash) == processed_hashes.end()) {
            // 只在仓储2中存在
            diff.only_in_repo2.insert(diff.only_in_repo2.end(), ids2.begin(), ids2.end());
        }
    }
    
    return diff;
}

} // namespace StyleTransfer

}} // namespace fastexcel::core