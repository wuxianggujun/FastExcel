//
// Created by wuxianggujun on 25-8-4.
//

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/span.hpp"
#include <string>
#include <unordered_map>

using fastexcel::core::span;  // Import span into this namespace

namespace fastexcel {
namespace reader {

/**
 * @brief 高性能共享字符串表解析器 - 基于SAX流式解析
 *
 * 负责解析Excel文件中的sharedStrings.xml文件，
 * 使用BaseSAXParser提供的高性能SAX解析能力，
 * 消除所有字符串查找操作，大幅提升解析性能。
 * 
 * 性能优化：
 * - 零字符串查找：基于SAX事件驱动
 * - 文本内容流式收集
 * - 使用项目统一的XML实体解码
 * - 内存预分配优化
 */
class SharedStringsParser : public BaseSAXParser {
public:
    SharedStringsParser() = default;
    ~SharedStringsParser() = default;
    
    /**
     * @brief 解析共享字符串XML内容
     * @param xml_content XML内容 
     * @return 是否解析成功
     */
    bool parse(const std::string& xml_content) {
        clear();
        return parseXML(xml_content);
    }
    
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
    // 解析状态
    struct ParseState {
        int current_string_index = 0;
        bool in_si_element = false;      // 是否在<si>元素内
        bool in_text_element = false;    // 是否在<t>或<r><t>元素内  
        bool in_rich_text = false;       // 是否在富文本<r>元素内
        std::string current_text;        // 当前收集的文本
        
        void reset() {
            current_string_index = 0;
            in_si_element = false;
            in_text_element = false;
            in_rich_text = false;
            current_text.clear();
        }
        
        void startNewString() {
            in_si_element = true;
            in_text_element = false;
            in_rich_text = false;
            current_text.clear();
        }
        
        void endString() {
            in_si_element = false;
            in_text_element = false;
            in_rich_text = false;
        }
    } parse_state_;
    
    std::unordered_map<int, std::string> strings_;  // 索引 -> 字符串
    
    // 重写基类虚函数
    void onStartElement(std::string_view name, span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
};

} // namespace reader
} // namespace fastexcel
