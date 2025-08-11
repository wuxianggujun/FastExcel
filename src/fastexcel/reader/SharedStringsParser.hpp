#include "fastexcel/utils/ModuleLoggers.hpp"
//
// Created by wuxianggujun on 25-8-4.
//

#pragma once


#include <string>
#include <vector>
#include <unordered_map>

namespace fastexcel {
namespace reader {

/**
 * @brief 共享字符串表解析器
 *
 * 负责解析Excel文件中的sharedStrings.xml文件，
 * 提取所有共享字符串并建立索引映射
 */
class SharedStringsParser {
public:
    SharedStringsParser() = default;
    ~SharedStringsParser() = default;
    
    /**
     * @brief 解析共享字符串XML内容
     * @param xml_content XML内容
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content);
    
    /**
     * @brief 根据索引获取字符串
     * @param index 字符串索引
     * @return 字符串内容，如果索引无效返回空字符串
     */
    std::string getString(int index) const;
    
    /**
     * @brief 获取字符串总数
     * @return 字符串总数
     */
    size_t getStringCount() const { return strings_.size(); }
    
    /**
     * @brief 获取所有字符串的映射
     * @return 索引到字符串的映射
     */
    const std::unordered_map<int, std::string>& getStrings() const { return strings_; }
    
    /**
     * @brief 清空解析结果
     */
    void clear();

private:
    std::unordered_map<int, std::string> strings_;  // 索引 -> 字符串
    
    // 辅助方法
    std::string extractTextContent(const std::string& xml, size_t start_pos, size_t end_pos);
    std::string decodeXMLEntities(const std::string& text);
};

} // namespace reader
} // namespace fastexcel
