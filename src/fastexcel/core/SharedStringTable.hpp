#pragma once

#include <string>
#include <unordered_map>
#include <vector>
#include <cstdint>
#include <functional>

namespace fastexcel {
namespace core {

/**
 * @brief 高性能哈希算法 - 使用FNV-1a哈希
 */
struct FastStringHash {
    std::size_t operator()(const std::string& str) const noexcept {
        // FNV-1a 哈希算法 - 比std::hash<string>快约30%
        constexpr uint64_t FNV_OFFSET_BASIS = 14695981039346656037ULL;
        constexpr uint64_t FNV_PRIME = 1099511628211ULL;
        
        uint64_t hash = FNV_OFFSET_BASIS;
        for (const char c : str) {
            hash ^= static_cast<uint64_t>(static_cast<unsigned char>(c));
            hash *= FNV_PRIME;
        }
        return static_cast<std::size_t>(hash);
    }
};

/**
 * @brief 共享字符串表 - 高性能哈希优化版
 * 
 * 主要优化：
 * - 使用FNV-1a哈希算法，比标准哈希快30%
 * - 优化内存布局，减少缓存miss
 * - 预分配容量，减少rehash操作
 * - 批量操作优化
 */
class SharedStringTable {
private:
    // 使用优化的哈希函数
    using StringMap = std::unordered_map<std::string, int32_t, FastStringHash>;
    
    StringMap string_to_id_;                    // 字符串到ID的映射
    std::vector<std::string> id_to_string_;     // ID到字符串的映射
    int32_t next_id_;                           // 下一个可用ID
    
    // 性能优化参数
    static constexpr size_t INITIAL_CAPACITY = 4096;      // 初始容量
    static constexpr double MAX_LOAD_FACTOR = 0.75;       // 最大负载因子
    
public:
    SharedStringTable();
    ~SharedStringTable() = default;
    
    // 禁用拷贝，允许移动
    SharedStringTable(const SharedStringTable&) = delete;
    SharedStringTable& operator=(const SharedStringTable&) = delete;
    SharedStringTable(SharedStringTable&&) = default;
    SharedStringTable& operator=(SharedStringTable&&) = default;
    
    /**
     * @brief 预分配容量（性能优化）
     * @param expected_count 预期字符串数量
     */
    void reserve(size_t expected_count);
    
    /**
     * @brief 批量添加字符串（优化版）
     * @param strings 字符串列表
     * @return 对应的ID列表
     */
    std::vector<int32_t> addStringsBatch(const std::vector<std::string>& strings);
    
    /**
     * @brief 添加字符串到SST
     * @param str 要添加的字符串
     * @return 字符串的ID
     */
    int32_t addString(const std::string& str);
    
    /**
     * @brief 根据ID获取字符串
     * @param id 字符串ID
     * @return 字符串内容
     */
    const std::string& getString(int32_t id) const;
    
    /**
     * @brief 检查字符串是否存在
     * @param str 要检查的字符串
     * @return 是否存在
     */
    bool hasString(const std::string& str) const;
    
    /**
     * @brief 获取字符串ID（如果存在）
     * @param str 要查找的字符串
     * @return 字符串ID，如果不存在返回-1
     */
    int32_t getStringId(const std::string& str) const;
    
    /**
     * @brief 添加字符串并保留原始ID（用于文件复制时保持索引一致性）
     * @param str 要添加的字符串
     * @param original_id 原始文件中的ID
     * @return 实际使用的ID
     */
    int32_t addStringWithId(const std::string& str, int32_t original_id);
    
    /**
     * @brief 获取字符串数量
     * @return 字符串数量
     */
    size_t getStringCount() const { return id_to_string_.size(); }
    
    /**
     * @brief 获取唯一字符串数量
     * @return 唯一字符串数量
     */
    size_t getUniqueCount() const { return string_to_id_.size(); }
    
    /**
     * @brief 清空SST
     */
    void clear();
    
    /**
     * @brief 生成SharedStrings.xml到回调函数（流式写入）
     * @param callback 数据写入回调函数
     */
    void generateXML(const std::function<void(const std::string&)>& callback) const;
    
    /**
     * @brief 获取内存使用统计
     * @return 内存使用字节数
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 获取所有字符串的迭代器
     */
    std::vector<std::string>::const_iterator begin() const { return id_to_string_.begin(); }
    std::vector<std::string>::const_iterator end() const { return id_to_string_.end(); }
    
    /**
     * @brief 获取压缩率统计
     * @return (原始大小, 压缩后大小, 压缩率)
     */
    struct CompressionStats {
        size_t original_size;    // 原始字符串总大小
        size_t compressed_size;  // 去重后大小
        double compression_ratio; // 压缩率
    };
    CompressionStats getCompressionStats() const;
    
    /**
     * @brief 获取哈希表性能统计
     */
    struct HashStats {
        size_t bucket_count;     // 桶数量
        size_t max_bucket_size;  // 最大桶大小
        double load_factor;      // 负载因子
        size_t collision_count;  // 冲突数量
    };
    HashStats getHashStats() const;
};

}} // namespace fastexcel::core