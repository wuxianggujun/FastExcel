#pragma once

#include <string>
#include <vector>
#include <unordered_map>
#include <fmt/format.h>
#include <algorithm>

namespace fastexcel {
namespace core {

/**
 * @brief 定义名称条目
 */
struct DefinedName {
    std::string name;     ///< 名称
    std::string formula;  ///< 公式或引用
    std::string scope;    ///< 作用域（工作表名或空字符串表示工作簿级）
    
    DefinedName() = default;
    DefinedName(const std::string& n, const std::string& f, const std::string& s = "")
        : name(n), formula(f), scope(s) {}
        
    bool operator==(const DefinedName& other) const {
        return name == other.name && formula == other.formula && scope == other.scope;
    }
};

/**
 * @brief 定义名称管理器
 * 
 * 封装定义名称的增删改查逻辑，消除Workbook中的重复代码
 */
class DefinedNameManager {
private:
    std::vector<DefinedName> defined_names_;
    
    /**
     * @brief 查找指定名称和作用域的定义名称
     * @param name 名称
     * @param scope 作用域
     * @return 指向定义名称的迭代器
     */
    auto findDefinition(const std::string& name, const std::string& scope) {
        return std::find_if(defined_names_.begin(), defined_names_.end(),
                          [&name, &scope](const DefinedName& dn) {
                              return dn.name == name && dn.scope == scope;
                          });
    }
    
    /**
     * @brief 查找指定名称和作用域的定义名称（const版本）
     * @param name 名称
     * @param scope 作用域
     * @return 指向定义名称的常量迭代器
     */
    auto findDefinition(const std::string& name, const std::string& scope) const {
        return std::find_if(defined_names_.begin(), defined_names_.end(),
                          [&name, &scope](const DefinedName& dn) {
                              return dn.name == name && dn.scope == scope;
                          });
    }

public:
    DefinedNameManager() = default;
    ~DefinedNameManager() = default;
    
    // 禁用拷贝，允许移动
    DefinedNameManager(const DefinedNameManager&) = delete;
    DefinedNameManager& operator=(const DefinedNameManager&) = delete;
    DefinedNameManager(DefinedNameManager&&) = default;
    DefinedNameManager& operator=(DefinedNameManager&&) = default;
    
    /**
     * @brief 定义名称
     * @param name 名称
     * @param formula 公式或引用
     * @param scope 作用域（工作表名或空字符串表示工作簿级）
     */
    void define(const std::string& name, const std::string& formula, const std::string& scope = "") {
        // 验证名称是否有效
        if (!isValidName(name)) {
            throw std::invalid_argument(fmt::format("Invalid defined name: {}", name));
        }
        
        auto it = findDefinition(name, scope);
        if (it != defined_names_.end()) {
            // 更新已存在的定义
            it->formula = formula;
        } else {
            // 添加新定义
            defined_names_.emplace_back(name, formula, scope);
        }
    }
    
    /**
     * @brief 获取定义名称的公式
     * @param name 名称
     * @param scope 作用域
     * @return 公式字符串，如果不存在则返回空字符串
     */
    std::string get(const std::string& name, const std::string& scope = "") const {
        auto it = findDefinition(name, scope);
        if (it != defined_names_.end()) {
            return it->formula;
        }
        return "";
    }
    
    /**
     * @brief 获取完整的定义名称信息
     * @param name 名称
     * @param scope 作用域
     * @return 指向定义名称的指针，如果不存在则返回nullptr
     */
    const DefinedName* getDefinition(const std::string& name, const std::string& scope = "") const {
        auto it = findDefinition(name, scope);
        if (it != defined_names_.end()) {
            return &(*it);
        }
        return nullptr;
    }
    
    /**
     * @brief 删除定义名称
     * @param name 名称
     * @param scope 作用域
     * @return 是否成功删除
     */
    bool remove(const std::string& name, const std::string& scope = "") {
        auto it = findDefinition(name, scope);
        if (it != defined_names_.end()) {
            defined_names_.erase(it);
            return true;
        }
        return false;
    }
    
    /**
     * @brief 获取所有定义名称
     * @return 定义名称列表的只读引用
     */
    const std::vector<DefinedName>& getAllDefinitions() const {
        return defined_names_;
    }
    
    /**
     * @brief 获取指定作用域的定义名称
     * @param scope 作用域
     * @return 该作用域下的定义名称列表
     */
    std::vector<DefinedName> getDefinitionsByScope(const std::string& scope = "") const {
        std::vector<DefinedName> result;
        for (const auto& def : defined_names_) {
            if (def.scope == scope) {
                result.push_back(def);
            }
        }
        return result;
    }
    
    /**
     * @brief 获取所有定义名称的简单映射
     * @return 名称到公式的映射（不包含作用域信息）
     */
    std::unordered_map<std::string, std::string> getSimpleMapping() const {
        std::unordered_map<std::string, std::string> result;
        for (const auto& def : defined_names_) {
            // 对于有作用域的名称，使用 "scope!name" 作为键
            std::string key = def.scope.empty() ? def.name : fmt::format("{}!{}", def.scope, def.name);
            result[key] = def.formula;
        }
        return result;
    }
    
    /**
     * @brief 清空所有定义名称
     */
    void clear() {
        defined_names_.clear();
    }
    
    /**
     * @brief 获取定义名称数量
     * @return 定义名称数量
     */
    size_t size() const {
        return defined_names_.size();
    }
    
    /**
     * @brief 检查是否为空
     * @return 是否没有任何定义名称
     */
    bool empty() const {
        return defined_names_.empty();
    }
    
    /**
     * @brief 检查定义名称是否存在
     * @param name 名称
     * @param scope 作用域
     * @return 定义名称是否存在
     */
    bool hasDefinition(const std::string& name, const std::string& scope = "") const {
        return findDefinition(name, scope) != defined_names_.end();
    }
    
    /**
     * @brief 重命名定义名称
     * @param old_name 旧名称
     * @param new_name 新名称
     * @param scope 作用域
     * @return 是否成功重命名
     */
    bool rename(const std::string& old_name, const std::string& new_name, const std::string& scope = "") {
        if (!isValidName(new_name)) {
            return false;
        }
        
        // 检查新名称是否已存在
        if (hasDefinition(new_name, scope)) {
            return false;
        }
        
        auto it = findDefinition(old_name, scope);
        if (it != defined_names_.end()) {
            it->name = new_name;
            return true;
        }
        
        return false;
    }
    
    /**
     * @brief 更新定义名称的作用域
     * @param name 名称
     * @param old_scope 旧作用域
     * @param new_scope 新作用域
     * @return 是否成功更新
     */
    bool updateScope(const std::string& name, const std::string& old_scope, const std::string& new_scope) {
        // 检查新作用域下是否已存在同名定义
        if (hasDefinition(name, new_scope)) {
            return false;
        }
        
        auto it = findDefinition(name, old_scope);
        if (it != defined_names_.end()) {
            it->scope = new_scope;
            return true;
        }
        
        return false;
    }

private:
    /**
     * @brief 验证定义名称是否有效
     * @param name 名称
     * @return 是否有效
     */
    bool isValidName(const std::string& name) const {
        if (name.empty() || name.length() > 255) {
            return false;
        }
        
        // 名称不能以数字开头
        if (std::isdigit(name[0])) {
            return false;
        }
        
        // 名称不能包含特殊字符（简化版检查）
        for (char c : name) {
            if (!std::isalnum(c) && c != '_' && c != '.' && c != '\\') {
                return false;
            }
        }
        
        // 不能与单元格引用冲突（如A1, B2等）
        if (name.length() <= 3) {
            size_t i = 0;
            
            // 检查是否以字母开头
            while (i < name.length() && std::isalpha(name[i])) {
                ++i;
            }
            
            // 检查后面是否都是数字
            while (i < name.length() && std::isdigit(name[i])) {
                ++i;
            }
            
            if (i == name.length() && i > 1) {
                return false; // 可能是单元格引用
            }
        }
        
        return true;
    }
};

}} // namespace fastexcel::core
