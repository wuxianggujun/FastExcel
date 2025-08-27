/**
 * @file DocPropsParser.cpp
 * @brief 文档属性XML解析器实现
 */

#include "DocPropsParser.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace reader {

void DocPropsParser::onStartElement(std::string_view name, 
                                   span<const xml::XMLAttribute> /* attributes */, int /* depth */) {
    std::string element_name(name);
    
    // 处理核心属性元素 (docProps/core.xml)
    if (element_name == "dc:title" || element_name == "title") {
        in_text_element_ = true;
        current_element_ = "title";
    } else if (element_name == "dc:subject" || element_name == "subject") {
        in_text_element_ = true;
        current_element_ = "subject";
    } else if (element_name == "dc:creator" || element_name == "creator") {
        in_text_element_ = true;
        current_element_ = "author";
    } else if (element_name == "cp:keywords" || element_name == "keywords") {
        in_text_element_ = true;
        current_element_ = "keywords";
    } else if (element_name == "dc:description" || element_name == "description") {
        in_text_element_ = true;
        current_element_ = "description";
    } else if (element_name == "cp:lastModifiedBy" || element_name == "lastModifiedBy") {
        in_text_element_ = true;
        current_element_ = "last_modified_by";
    } else if (element_name == "dcterms:created" || element_name == "created") {
        in_text_element_ = true;
        current_element_ = "created";
    } else if (element_name == "dcterms:modified" || element_name == "modified") {
        in_text_element_ = true;
        current_element_ = "modified";
    } else if (element_name == "cp:category" || element_name == "category") {
        in_text_element_ = true;
        current_element_ = "category";
    } else if (element_name == "cp:revision" || element_name == "revision") {
        in_text_element_ = true;
        current_element_ = "revision";
    }
    // 处理应用属性元素 (docProps/app.xml)
    else if (element_name == "Application" || element_name == "application") {
        in_text_element_ = true;
        current_element_ = "application";
    } else if (element_name == "AppVersion" || element_name == "appVersion") {
        in_text_element_ = true;
        current_element_ = "app_version";
    } else if (element_name == "Company" || element_name == "company") {
        in_text_element_ = true;
        current_element_ = "company";
    } else if (element_name == "Manager" || element_name == "manager") {
        in_text_element_ = true;
        current_element_ = "manager";
    } else if (element_name == "DocSecurity" || element_name == "docSecurity") {
        in_text_element_ = true;
        current_element_ = "doc_security";
    } else if (element_name == "HyperlinksChanged" || element_name == "hyperlinksChanged") {
        in_text_element_ = true;
        current_element_ = "hyperlinks_changed";
    } else if (element_name == "SharedDoc" || element_name == "sharedDoc") {
        in_text_element_ = true;
        current_element_ = "shared_doc";
    }
    
    if (in_text_element_) {
        FASTEXCEL_LOG_DEBUG("开始解析文档属性元素: {}", current_element_);
    }
}

void DocPropsParser::onEndElement(std::string_view /* name */, int /* depth */) {
    if (in_text_element_) {
        FASTEXCEL_LOG_DEBUG("完成文档属性元素解析: {}", current_element_);
        in_text_element_ = false;
        current_element_.clear();
    }
}

void DocPropsParser::onText(std::string_view data, int /* depth */) {
    if (in_text_element_ && !current_element_.empty()) {
        std::string text_data(data);
        
        // 根据当前元素类型设置对应的属性值
        if (current_element_ == "title") {
            doc_props_.title += text_data;
        } else if (current_element_ == "subject") {
            doc_props_.subject += text_data;
        } else if (current_element_ == "author") {
            doc_props_.author += text_data;
        } else if (current_element_ == "keywords") {
            doc_props_.keywords += text_data;
        } else if (current_element_ == "description") {
            doc_props_.description += text_data;
        } else if (current_element_ == "last_modified_by") {
            doc_props_.last_modified_by += text_data;
        } else if (current_element_ == "created") {
            doc_props_.created += text_data;
        } else if (current_element_ == "modified") {
            doc_props_.modified += text_data;
        } else if (current_element_ == "category") {
            doc_props_.category += text_data;
        } else if (current_element_ == "revision") {
            doc_props_.revision += text_data;
        } else if (current_element_ == "application") {
            doc_props_.application += text_data;
        } else if (current_element_ == "app_version") {
            doc_props_.app_version += text_data;
        } else if (current_element_ == "company") {
            doc_props_.company += text_data;
        } else if (current_element_ == "manager") {
            doc_props_.manager += text_data;
        } else if (current_element_ == "doc_security") {
            doc_props_.doc_security += text_data;
        } else if (current_element_ == "hyperlinks_changed") {
            doc_props_.hyperlinks_changed += text_data;
        } else if (current_element_ == "shared_doc") {
            doc_props_.shared_doc += text_data;
        }
    }
}

}} // namespace fastexcel::reader