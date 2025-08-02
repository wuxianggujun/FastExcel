#pragma once

#include <string>
#include <unordered_map>
#include <vector>
#include <cstdint>
#include <functional>

namespace fastexcel {
namespace core {

/**
 * @brief 共享字符串表 - 借鉴libxlsxwriter的SST优化
 * 
 * 实现字符串去重存储，减少内存使用和文件大小
 */
class SharedStringTable {
private:
    std::unordered_map<std::string, int32_t> string_to_id_;  // 字符串到ID的映射
    std::vector<std::string> id_to_string_;                  // ID到字符串的映射
    int32_t next_id_;                                        // 下一个可用ID
    
public:
    SharedStringTable();
    ~SharedStringTable() = default;
    
    // 禁用拷贝，允许移动
    SharedStringTable(const SharedStringTable&) = delete;
    SharedStringTable& operator=(const SharedStringTable&) = delete;
    SharedStringTable(SharedStringTable&&) = default;
    SharedStringTable& operator=(SharedStringTable&&) = default;
    
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
    void generateXML(const std::function<void(const char*, size_t)>& callback) const;
    
    /**
     * @brief 生成SharedStrings.xml到文件（流式写入）
     * @param filename 输出文件名
     */
    void generateXMLToFile(const std::string& filename) const;
    
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
};

}} // namespace fastexcel::core