#pragma once

#include <string>
#include <vector>
#include <fmt/format.h>
#include <unordered_map>
#include <algorithm>

namespace fastexcel {
namespace core {

/**
 * @brief 自定义属性条目
 */
struct CustomProperty {
    enum Type {
        String,
        Number,
        Boolean
    };
    
    std::string name;
    std::string value;
    Type type;
    
    CustomProperty() = default;
    CustomProperty(const std::string& n, const std::string& v) : name(n), value(v), type(String) {}
    CustomProperty(const std::string& n, double v) : name(n), value(fmt::format("{}", v)), type(Number) {}
    CustomProperty(const std::string& n, bool v) : name(n), value(v ? "true" : "false"), type(Boolean) {}
};

/**
 * @brief 自定义属性管理器
 * 
 * 封装自定义属性的增删改查逻辑，消除Workbook中的重复代码
 */
class CustomPropertyManager {
private:
    std::vector<CustomProperty> properties_;
    
    /**
     * @brief 查找指定名称的属性
     * @param name 属性名称
     * @return 指向属性的迭代器
     */
    auto findProperty(const std::string& name) {
        return std::find_if(properties_.begin(), properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    }
    
    /**
     * @brief 查找指定名称的属性（const版本）
     * @param name 属性名称
     * @return 指向属性的常量迭代器
     */
    auto findProperty(const std::string& name) const {
        return std::find_if(properties_.begin(), properties_.end(),
                          [&name](const CustomProperty& prop) {
                              return prop.name == name;
                          });
    }

public:
    CustomPropertyManager() = default;
    ~CustomPropertyManager() = default;
    
    // 禁用拷贝，允许移动
    CustomPropertyManager(const CustomPropertyManager&) = delete;
    CustomPropertyManager& operator=(const CustomPropertyManager&) = delete;
    CustomPropertyManager(CustomPropertyManager&&) = default;
    CustomPropertyManager& operator=(CustomPropertyManager&&) = default;
    
    /**
     * @brief 设置字符串类型属性
     * @param name 属性名称
     * @param value 属性值
     */
    void setProperty(const std::string& name, const std::string& value) {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            it->value = value;
            it->type = CustomProperty::String;
        } else {
            properties_.emplace_back(name, value);
        }
    }
    
    /**
     * @brief 设置数值类型属性
     * @param name 属性名称
     * @param value 属性值
     */
    void setProperty(const std::string& name, double value) {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            it->value = std::to_string(value);
            it->type = CustomProperty::Number;
        } else {
            properties_.emplace_back(name, value);
        }
    }
    
    /**
     * @brief 设置布尔类型属性
     * @param name 属性名称
     * @param value 属性值
     */
    void setProperty(const std::string& name, bool value) {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            it->value = value ? "true" : "false";
            it->type = CustomProperty::Boolean;
        } else {
            properties_.emplace_back(name, value);
        }
    }
    
    /**
     * @brief 获取属性值
     * @param name 属性名称
     * @return 属性值字符串，如果不存在则返回空字符串
     */
    std::string getProperty(const std::string& name) const {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            return it->value;
        }
        return "";
    }
    
    /**
     * @brief 获取属性类型
     * @param name 属性名称
     * @param type 输出参数，属性类型
     * @return 是否找到该属性
     */
    bool getPropertyType(const std::string& name, CustomProperty::Type& type) const {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            type = it->type;
            return true;
        }
        return false;
    }
    
    /**
     * @brief 删除属性
     * @param name 属性名称
     * @return 是否成功删除
     */
    bool removeProperty(const std::string& name) {
        auto it = findProperty(name);
        if (it != properties_.end()) {
            properties_.erase(it);
            return true;
        }
        return false;
    }
    
    /**
     * @brief 获取所有属性的映射
     * @return 属性名到属性值的映射
     */
    std::unordered_map<std::string, std::string> getAllProperties() const {
        std::unordered_map<std::string, std::string> result;
        for (const auto& prop : properties_) {
            result[prop.name] = prop.value;
        }
        return result;
    }
    
    /**
     * @brief 获取所有属性的详细信息
     * @return 属性列表的只读引用
     */
    const std::vector<CustomProperty>& getAllDetailedProperties() const {
        return properties_;
    }
    
    /**
     * @brief 清空所有属性
     */
    void clear() {
        properties_.clear();
    }
    
    /**
     * @brief 获取属性数量
     * @return 属性数量
     */
    size_t size() const {
        return properties_.size();
    }
    
    /**
     * @brief 检查是否为空
     * @return 是否没有任何属性
     */
    bool empty() const {
        return properties_.empty();
    }
    
    /**
     * @brief 检查属性是否存在
     * @param name 属性名称
     * @return 属性是否存在
     */
    bool hasProperty(const std::string& name) const {
        return findProperty(name) != properties_.end();
    }
};

}} // namespace fastexcel::core
