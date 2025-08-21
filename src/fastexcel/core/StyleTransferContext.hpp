#pragma once

#include "FormatRepository.hpp"
#include <unordered_map>
#include <memory>

namespace fastexcel {
namespace core {

/**
 * @brief 样式传输上下文 - 跨工作簿样式复制的映射管理
 * 
 * 实现Unit of Work模式，封装样式从源工作簿到目标工作簿的传输过程。
 * 自动处理样式ID的映射，确保样式的一致性和去重。
 */
class StyleTransferContext {
private:
    // 源仓储到目标仓储的引用
    const FormatRepository& source_repository_;
    FormatRepository& target_repository_;
    
    // ID映射缓存：源ID -> 目标ID
    mutable std::unordered_map<int, int> id_mapping_;
    
    // 是否已完成批量导入
    mutable bool bulk_imported_ = false;
    
    // 统计信息
    mutable size_t transferred_count_ = 0;
    mutable size_t deduplicated_count_ = 0;

public:
    /**
     * @brief 构造函数
     * @param source_repo 源格式仓储
     * @param target_repo 目标格式仓储
     */
    StyleTransferContext(const FormatRepository& source_repo, 
                        FormatRepository& target_repo);
    
    ~StyleTransferContext() = default;
    
    // 禁用拷贝和移动（RAII资源管理）
    StyleTransferContext(const StyleTransferContext&) = delete;
    StyleTransferContext& operator=(const StyleTransferContext&) = delete;
    StyleTransferContext(StyleTransferContext&&) = delete;
    StyleTransferContext& operator=(StyleTransferContext&&) = delete;
    
    /**
     * @brief 映射单个样式ID
     * @param source_id 源样式ID
     * @return 目标样式ID，如果源ID无效则返回默认样式ID
     */
    int mapStyleId(int source_id) const;
    
    /**
     * @brief 批量映射样式ID
     * @param source_ids 源样式ID列表
     * @return 目标样式ID列表
     */
    std::vector<int> mapStyleIds(const std::vector<int>& source_ids) const;
    
    /**
     * @brief 预加载所有映射（性能优化）
     * 
     * 一次性建立所有样式的映射关系，适用于大量样式传输的场景。
     * 调用此方法后，mapStyleId将直接从缓存返回结果。
     */
    void preloadAllMappings() const;
    
    /**
     * @brief 检查源样式ID是否有效
     * @param source_id 源样式ID
     * @return 是否有效
     */
    bool isValidSourceId(int source_id) const;
    
    /**
     * @brief 获取映射统计信息
     */
    struct TransferStats {
        size_t source_format_count;    // 源仓储格式数量
        size_t target_format_count;    // 目标仓储格式数量（传输后）
        size_t transferred_count;      // 已传输的格式数量
        size_t deduplicated_count;     // 去重的格式数量
        double deduplication_ratio;    // 去重率
    };
    
    TransferStats getTransferStats() const;
    
    /**
     * @brief 清空映射缓存
     * 
     * 强制重新计算映射关系，适用于目标仓储被修改的情况。
     */
    void clearCache() const;
    
    /**
     * @brief 获取映射缓存大小
     * @return 缓存的映射数量
     */
    size_t getCacheSize() const { return id_mapping_.size(); }
    
    /**
     * @brief 获取完整的ID映射表（只读）
     * @return 源ID到目标ID的映射表
     */
    const std::unordered_map<int, int>& getIdMapping() const;
    
    // 批量操作
    
    /**
     * @brief 传输指定范围的样式
     * @param start_id 起始样式ID（包含）
     * @param end_id 结束样式ID（不包含）
     * @return 传输的样式数量
     */
    size_t transferStyleRange(int start_id, int end_id) const;
    
    /**
     * @brief 传输所有样式
     * @return 传输的样式数量
     */
    size_t transferAllStyles() const;
    
    /**
     * @brief 传输指定的样式列表
     * @param source_ids 要传输的源样式ID列表
     * @return 传输的样式数量
     */
    size_t transferStyles(const std::vector<int>& source_ids) const;

private:
    /**
     * @brief 内部映射方法（不带缓存检查）
     * @param source_id 源样式ID
     * @return 目标样式ID
     */
    int mapStyleIdInternal(int source_id) const;
    
    /**
     * @brief 更新统计信息
     * @param was_new_transfer 是否是新传输的格式
     */
    void updateStats(bool was_new_transfer) const;
};

/**
 * @brief 跨工作簿样式传输的辅助函数集
 */
namespace StyleTransfer {
    
    /**
     * @brief 快速样式复制（无映射管理）
     * @param source_repo 源仓储
     * @param target_repo 目标仓储
     * @param source_ids 要复制的样式ID列表
     * @return 目标样式ID列表
     */
    std::vector<int> quickCopyStyles(
        const FormatRepository& source_repo,
        FormatRepository& target_repo,
        const std::vector<int>& source_ids
    );
    
    /**
     * @brief 合并两个格式仓储
     * @param repo1 第一个仓储
     * @param repo2 第二个仓储
     * @return 包含两个仓储所有格式的新仓储
     */
    std::unique_ptr<FormatRepository> mergeRepositories(
        const FormatRepository& repo1,
        const FormatRepository& repo2
    );
    
    /**
     * @brief 创建样式差异报告
     * @param repo1 第一个仓储
     * @param repo2 第二个仓储
     * @return 差异信息
     */
    struct StyleDifference {
        std::vector<int> only_in_repo1;    // 只在仓储1中的样式
        std::vector<int> only_in_repo2;    // 只在仓储2中的样式
        std::vector<std::pair<int, int>> common_styles; // 相同样式的ID对应关系
    };
    
    StyleDifference compareRepositories(
        const FormatRepository& repo1,
        const FormatRepository& repo2
    );
    
} // namespace StyleTransfer

}} // namespace fastexcel::core
