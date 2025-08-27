/**
 * @file DocPropsParser.hpp
 * @brief 文档属性XML解析器
 */

#pragma once

#include "BaseSAXParser.hpp"
#include "fastexcel/core/span.hpp"
#include <string>

using fastexcel::core::span;

namespace fastexcel {
namespace reader {

/**
 * @brief 文档属性信息结构
 */
struct DocPropsInfo {
    // 核心属性 (docProps/core.xml)
    std::string title;          // 标题
    std::string subject;        // 主题
    std::string author;         // 作者 (创建者)
    std::string keywords;       // 关键词
    std::string description;    // 描述
    std::string last_modified_by; // 最后修改者
    std::string created;        // 创建时间
    std::string modified;       // 修改时间
    std::string category;       // 分类
    std::string revision;       // 修订版本

    // 应用属性 (docProps/app.xml)
    std::string application;    // 应用程序
    std::string app_version;    // 应用版本
    std::string company;        // 公司
    std::string manager;        // 管理者
    std::string doc_security;   // 文档安全性
    std::string hyperlinks_changed; // 超链接是否已更改
    std::string shared_doc;     // 是否共享文档
    
    DocPropsInfo() = default;
};

/**
 * @brief 文档属性XML流式解析器
 * 支持解析 docProps/core.xml 和 docProps/app.xml
 */
class DocPropsParser : public BaseSAXParser {
private:
    DocPropsInfo doc_props_;
    
    // 解析状态
    bool in_text_element_ = false;
    std::string current_element_;
    
protected:
    // 重写SAX事件处理方法
    void onStartElement(std::string_view name, 
                       span<const xml::XMLAttribute> attributes, int depth) override;
    void onEndElement(std::string_view name, int depth) override;
    void onText(std::string_view data, int depth) override;
    
public:
    DocPropsParser() = default;
    ~DocPropsParser() = default;
    
    /**
     * @brief 解析核心属性XML (docProps/core.xml)
     */
    bool parseCoreProps(const std::string& xml_content) {
        reset();
        return parseXML(xml_content);
    }
    
    /**
     * @brief 解析应用属性XML (docProps/app.xml)
     */
    bool parseAppProps(const std::string& xml_content) {
        // 不重置，允许合并core和app属性
        return parseXML(xml_content);
    }
    
    /**
     * @brief 获取解析后的文档属性
     */
    const DocPropsInfo& getDocProps() const {
        return doc_props_;
    }
    
    /**
     * @brief 移动文档属性所有权
     */
    DocPropsInfo takeDocProps() {
        return std::move(doc_props_);
    }
    
    /**
     * @brief 清空解析状态
     */
    void reset() {
        state_.reset();
        doc_props_ = DocPropsInfo{};
        in_text_element_ = false;
        current_element_.clear();
    }
};

}} // namespace fastexcel::reader