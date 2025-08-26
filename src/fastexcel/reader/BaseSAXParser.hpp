#pragma once

#include "fastexcel/xml/XMLStreamReader.hpp" 
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/XMLUtils.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <string>
#include <vector>
#include <stack>
#include <optional>
#include <functional>
#include <unordered_map>

namespace fastexcel {
namespace reader {

/**
 * @brief 通用SAX解析器基类 - 为所有XML解析器提供统一的高性能SAX解析能力
 * 
 * 设计理念：
 * - 零字符串查找：基于XMLStreamReader的SAX事件驱动
 * - 属性快速提取：优化的属性查找方法
 * - 状态管理：元素栈和解析状态跟踪
 * - 工具集成：集成项目现有的CommonUtils和XMLUtils
 * - 模板化设计：支持不同解析器的特化需求
 */
class BaseSAXParser {
protected:
    // 通用解析状态
    struct ParseState {
        // 元素栈用于跟踪嵌套结构
        std::stack<std::string> element_stack;
        
        // 当前深度
        int current_depth = 0;
        
        // 当前文本内容
        std::string current_text;
        
        // 是否正在收集文本
        bool collecting_text = false;
        
        // 错误状态
        bool has_error = false;
        std::string error_message;
        
        void reset() {
            while (!element_stack.empty()) element_stack.pop();
            current_depth = 0;
            current_text.clear();
            collecting_text = false;
            has_error = false;
            error_message.clear();
        }
        
        // 获取当前元素名称
        std::string getCurrentElement() const {
            return element_stack.empty() ? "" : element_stack.top();
        }
        
        // 检查是否在指定元素内
        bool isInElement(const std::string& element_name) const {
            if (element_stack.empty()) return false;
            std::stack<std::string> temp_stack = element_stack;
            while (!temp_stack.empty()) {
                if (temp_stack.top() == element_name) return true;
                temp_stack.pop();
            }
            return false;
        }
    };
    
    ParseState state_;
    
public:
    BaseSAXParser() = default;
    virtual ~BaseSAXParser() = default;
    
    // 禁用拷贝，支持移动
    BaseSAXParser(const BaseSAXParser&) = delete;
    BaseSAXParser& operator=(const BaseSAXParser&) = delete;
    BaseSAXParser(BaseSAXParser&&) = default;
    BaseSAXParser& operator=(BaseSAXParser&&) = default;
    
    /**
     * @brief 解析XML内容的统一入口
     * @param xml_content XML字符串内容
     * @return 是否解析成功
     */
    bool parseXML(const std::string& xml_content) {
        if (xml_content.empty()) {
            state_.has_error = true;
            state_.error_message = "Empty XML content";
            return false;
        }
        
        // 重置状态
        state_.reset();
        
        try {
            // 创建XMLStreamReader实例
            xml::XMLStreamReader reader;
            
            // 设置SAX回调
            reader.setStartElementCallback([this](const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
                handleStartElement(name, attributes, depth);
            });
            
            reader.setEndElementCallback([this](const std::string& name, int depth) {
                handleEndElement(name, depth);
            });
            
            reader.setTextCallback([this](const std::string& text, int depth) {
                handleText(text, depth);
            });
            
            reader.setErrorCallback([this](xml::XMLParseError error, const std::string& message, int line, int column) {
                state_.has_error = true;
                state_.error_message = "XML Parse Error at line " + std::to_string(line) + 
                                     ", column " + std::to_string(column) + ": " + message;
                FASTEXCEL_LOG_ERROR("SAX Parser Error: {}", state_.error_message);
            });
            
            // 执行解析
            auto result = reader.parseFromString(xml_content);
            
            if (result != xml::XMLParseError::Ok) {
                if (!state_.has_error) {
                    state_.has_error = true;
                    state_.error_message = "XML parsing failed";
                }
                return false;
            }
            
            return !state_.has_error;
            
        } catch (const std::exception& e) {
            state_.has_error = true;
            state_.error_message = std::string("Exception during parsing: ") + e.what();
            FASTEXCEL_LOG_ERROR("SAX Parser Exception: {}", state_.error_message);
            return false;
        }
    }
    
    // 错误信息获取
    bool hasError() const { return state_.has_error; }
    std::string getErrorMessage() const { return state_.error_message; }
    
protected:
    /**
     * @brief 子类需要实现的SAX事件处理器
     */
    virtual void handleStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) {
        state_.element_stack.push(name);
        state_.current_depth = depth;
        state_.current_text.clear();
        onStartElement(name, attributes, depth);
    }
    
    virtual void handleEndElement(const std::string& name, int depth) {
        if (!state_.element_stack.empty()) {
            state_.element_stack.pop();
        }
        state_.current_depth = depth;
        onEndElement(name, depth);
        state_.current_text.clear();
    }
    
    virtual void handleText(const std::string& text, int depth) {
        if (state_.collecting_text) {
            state_.current_text += text;
        }
        onText(text, depth);
    }
    
    // 子类重写的虚函数
    virtual void onStartElement(const std::string& name, const std::vector<xml::XMLAttribute>& attributes, int depth) = 0;
    virtual void onEndElement(const std::string& name, int depth) = 0;
    virtual void onText(const std::string& text, int depth) {}
    
    // ==================== 通用工具方法 ====================
    
    /**
     * @brief 快速属性查找 - 零分配优化版本
     */
    std::optional<std::string> findAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) const {
        for (const auto& attr : attributes) {
            if (attr.name == name) {
                return attr.value;
            }
        }
        return std::nullopt;
    }
    
    /**
     * @brief 整数属性提取
     */
    std::optional<int> findIntAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) const {
        auto val = findAttribute(attributes, name);
        if (val) {
            try {
                return std::stoi(*val);
            } catch (const std::exception&) {
                return std::nullopt;
            }
        }
        return std::nullopt;
    }
    
    /**
     * @brief 浮点数属性提取
     */
    std::optional<double> findDoubleAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) const {
        auto val = findAttribute(attributes, name);
        if (val) {
            try {
                return std::stod(*val);
            } catch (const std::exception&) {
                return std::nullopt;
            }
        }
        return std::nullopt;
    }
    
    /**
     * @brief 布尔属性提取
     */
    std::optional<bool> findBoolAttribute(const std::vector<xml::XMLAttribute>& attributes, const std::string& name) const {
        auto val = findAttribute(attributes, name);
        if (val) {
            const std::string& value = *val;
            return (value == "1" || value == "true" || value == "True" || value == "TRUE");
        }
        return std::nullopt;
    }
    
    /**
     * @brief 获取属性值，带默认值
     */
    std::string getAttributeOr(const std::vector<xml::XMLAttribute>& attributes, const std::string& name, const std::string& default_value) const {
        auto val = findAttribute(attributes, name);
        return val ? *val : default_value;
    }
    
    /**
     * @brief 整数属性获取，带默认值
     */
    int getIntAttributeOr(const std::vector<xml::XMLAttribute>& attributes, const std::string& name, int default_value) const {
        auto val = findIntAttribute(attributes, name);
        return val ? *val : default_value;
    }
    
    /**
     * @brief 浮点数属性获取，带默认值
     */
    double getDoubleAttributeOr(const std::vector<xml::XMLAttribute>& attributes, const std::string& name, double default_value) const {
        auto val = findDoubleAttribute(attributes, name);
        return val ? *val : default_value;
    }
    
    /**
     * @brief 布尔属性获取，带默认值
     */
    bool getBoolAttributeOr(const std::vector<xml::XMLAttribute>& attributes, const std::string& name, bool default_value) const {
        auto val = findBoolAttribute(attributes, name);
        return val ? *val : default_value;
    }
    
    // ==================== Excel特化工具方法 ====================
    
    /**
     * @brief 解析单元格引用 (使用CommonUtils)
     */
    std::pair<int, int> parseCellReference(const std::string& ref) const {
        try {
            return utils::CommonUtils::parseReference(ref);
        } catch (const std::exception&) {
            return {-1, -1};
        }
    }
    
    /**
     * @brief 解析范围引用
     */
    bool parseRangeReference(const std::string& ref, int& first_row, int& first_col, int& last_row, int& last_col) const {
        size_t colon = ref.find(':');
        if (colon == std::string::npos) {
            // 单个单元格
            try {
                auto [row, col] = utils::CommonUtils::parseReference(ref);
                first_row = last_row = row;
                first_col = last_col = col;
                return true;
            } catch (const std::exception&) {
                return false;
            }
        }
        
        std::string start_ref = ref.substr(0, colon);
        std::string end_ref = ref.substr(colon + 1);
        
        try {
            auto [r1, c1] = utils::CommonUtils::parseReference(start_ref);
            auto [r2, c2] = utils::CommonUtils::parseReference(end_ref);
            
            first_row = r1;
            first_col = c1;
            last_row = r2;
            last_col = c2;
            return true;
        } catch (const std::exception&) {
            return false;
        }
    }
    
    /**
     * @brief XML实体解码 (使用XMLUtils)
     */
    std::string decodeXMLEntities(const std::string& text) const {
        return utils::XMLUtils::unescapeXML(text);
    }
    
    /**
     * @brief 获取当前收集的文本内容（自动解码XML实体）
     */
    std::string getCurrentTextDecoded() const {
        return decodeXMLEntities(state_.current_text);
    }
    
    /**
     * @brief 开始收集文本内容
     */
    void startCollectingText() {
        state_.collecting_text = true;
        state_.current_text.clear();
    }
    
    /**
     * @brief 停止收集文本内容
     */
    void stopCollectingText() {
        state_.collecting_text = false;
    }
    
    /**
     * @brief 设置错误状态
     */
    void setError(const std::string& message) {
        state_.has_error = true;
        state_.error_message = message;
        FASTEXCEL_LOG_ERROR("Parser Error: {}", message);
    }
    
    // ==================== 状态查询方法 ====================
    
    /**
     * @brief 获取当前元素名称
     */
    std::string getCurrentElement() const {
        return state_.getCurrentElement();
    }
    
    /**
     * @brief 检查是否在指定元素内
     */
    bool isInElement(const std::string& element_name) const {
        return state_.isInElement(element_name);
    }
    
    /**
     * @brief 获取当前深度
     */
    int getCurrentDepth() const {
        return state_.current_depth;
    }
    
    /**
     * @brief 获取当前文本内容（原始）
     */
    std::string getCurrentText() const {
        return state_.current_text;
    }
};

}} // namespace fastexcel::reader