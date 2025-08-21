#pragma once

#include "fastexcel/utils/TimeUtils.hpp"
#include <string>
#include <unordered_map>
#include <vector>
#include <ctime>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;
class DirtyManager;
class DefinedNameManager;

/**
 * @brief 文档属性结构
 */
struct DocumentProperties {
    std::string title;
    std::string subject;
    std::string author;
    std::string manager;
    std::string company;
    std::string category;
    std::string keywords;
    std::string comments;
    std::string status;
    std::string hyperlink_base;
    std::string application = "FastExcel";
    std::tm created_time;
    std::tm modified_time;
    
    DocumentProperties();
};

/**
 * @brief 工作簿文档管理器 - 负责管理文档属性和元数据
 * 
 * 设计原则：
 * - 单一职责：专注于文档元数据管理
 * - 高性能：使用哈希表提供快速查找
 * - 类型安全：支持多种属性值类型
 */
class WorkbookDocumentManager {
public:
    // 自定义属性值类型
    enum class PropertyType {
        STRING,
        DOUBLE, 
        BOOLEAN,
        DATE
    };
    
    struct CustomProperty {
        PropertyType type;
        std::string value;  // 统一存储为字符串，使用时转换
        
        CustomProperty() : type(PropertyType::STRING) {}
        explicit CustomProperty(const std::string& val) : type(PropertyType::STRING), value(val) {}
        explicit CustomProperty(double val) : type(PropertyType::DOUBLE), value(std::to_string(val)) {}
        explicit CustomProperty(bool val) : type(PropertyType::BOOLEAN), value(val ? "true" : "false") {}
        explicit CustomProperty(const std::tm& val) : type(PropertyType::DATE), value(utils::TimeUtils::formatTimeISO8601(val)) {}
        
        // 类型转换方法
        std::string asString() const { return value; }
        double asDouble() const;
        bool asBoolean() const;
        std::tm asDate() const;
    };

private:
    Workbook* workbook_;  // 父工作簿引用（不拥有）
    
    // 核心文档属性
    DocumentProperties doc_properties_;
    
    // 自定义属性
    std::unordered_map<std::string, CustomProperty> custom_properties_;
    
    // 脏数据管理器
    std::unique_ptr<DirtyManager> dirty_manager_;
    
    // 定义名称管理器  
    std::unique_ptr<DefinedNameManager> defined_name_manager_;
    
    // 配置
    struct Configuration {
        bool auto_update_modified_time = true;    // 自动更新修改时间
        bool validate_properties = true;          // 验证属性值
        size_t max_custom_properties = 100;       // 最大自定义属性数
        size_t max_property_length = 1000;        // 最大属性值长度
    } config_;

public:
    explicit WorkbookDocumentManager(Workbook* workbook);
    ~WorkbookDocumentManager() = default;
    
    // 禁用拷贝，允许移动
    WorkbookDocumentManager(const WorkbookDocumentManager&) = delete;
    WorkbookDocumentManager& operator=(const WorkbookDocumentManager&) = delete;
    WorkbookDocumentManager(WorkbookDocumentManager&&) = default;
    WorkbookDocumentManager& operator=(WorkbookDocumentManager&&) = default;
    
    // === 核心文档属性管理 ===
    
    /**
     * @brief 设置文档标题
     */
    void setTitle(const std::string& title);
    const std::string& getTitle() const { return doc_properties_.title; }
    
    /**
     * @brief 设置文档主题
     */
    void setSubject(const std::string& subject);
    const std::string& getSubject() const { return doc_properties_.subject; }
    
    /**
     * @brief 设置文档作者
     */
    void setAuthor(const std::string& author);
    const std::string& getAuthor() const { return doc_properties_.author; }
    
    /**
     * @brief 设置管理者
     */
    void setManager(const std::string& manager);
    const std::string& getManager() const { return doc_properties_.manager; }
    
    /**
     * @brief 设置公司
     */
    void setCompany(const std::string& company);
    const std::string& getCompany() const { return doc_properties_.company; }
    
    /**
     * @brief 设置类别
     */
    void setCategory(const std::string& category);
    const std::string& getCategory() const { return doc_properties_.category; }
    
    /**
     * @brief 设置关键词
     */
    void setKeywords(const std::string& keywords);
    const std::string& getKeywords() const { return doc_properties_.keywords; }
    
    /**
     * @brief 设置注释
     */
    void setComments(const std::string& comments);
    const std::string& getComments() const { return doc_properties_.comments; }
    
    /**
     * @brief 设置状态
     */
    void setStatus(const std::string& status);
    const std::string& getStatus() const { return doc_properties_.status; }
    
    /**
     * @brief 设置超链接基础
     */
    void setHyperlinkBase(const std::string& hyperlink_base);
    const std::string& getHyperlinkBase() const { return doc_properties_.hyperlink_base; }
    
    /**
     * @brief 设置应用程序名称
     */
    void setApplication(const std::string& application);
    const std::string& getApplication() const { return doc_properties_.application; }
    
    /**
     * @brief 批量设置文档属性
     */
    void setDocumentProperties(const std::string& title = "",
                              const std::string& subject = "",
                              const std::string& author = "",
                              const std::string& company = "",
                              const std::string& comments = "");
    
    // === 时间属性管理 ===
    
    /**
     * @brief 设置创建时间
     */
    void setCreatedTime(const std::tm& created_time);
    const std::tm& getCreatedTime() const { return doc_properties_.created_time; }
    
    /**
     * @brief 设置修改时间
     */
    void setModifiedTime(const std::tm& modified_time);
    const std::tm& getModifiedTime() const { return doc_properties_.modified_time; }
    
    /**
     * @brief 更新修改时间到当前时间
     */
    void updateModifiedTime();
    
    // === 自定义属性管理 ===
    
    /**
     * @brief 设置自定义属性
     */
    void setCustomProperty(const std::string& name, const std::string& value);
    void setCustomProperty(const std::string& name, double value);
    void setCustomProperty(const std::string& name, bool value);
    void setCustomProperty(const std::string& name, const std::tm& value);
    
    /**
     * @brief 获取自定义属性
     */
    std::string getCustomProperty(const std::string& name, const std::string& default_value = "") const;
    double getCustomPropertyAsDouble(const std::string& name, double default_value = 0.0) const;
    bool getCustomPropertyAsBoolean(const std::string& name, bool default_value = false) const;
    std::tm getCustomPropertyAsDate(const std::string& name) const;
    
    /**
     * @brief 检查自定义属性是否存在
     */
    bool hasCustomProperty(const std::string& name) const;
    
    /**
     * @brief 获取自定义属性类型
     */
    PropertyType getCustomPropertyType(const std::string& name) const;
    
    /**
     * @brief 删除自定义属性
     */
    bool removeCustomProperty(const std::string& name);
    
    /**
     * @brief 获取所有自定义属性名称
     */
    std::vector<std::string> getCustomPropertyNames() const;
    
    /**
     * @brief 获取自定义属性数量
     */
    size_t getCustomPropertyCount() const { return custom_properties_.size(); }
    
    /**
     * @brief 清除所有自定义属性
     */
    void clearCustomProperties();
    
    /**
     * @brief 批量设置自定义属性
     */
    void setCustomProperties(const std::unordered_map<std::string, std::string>& properties);
    
    /**
     * @brief 获取所有自定义属性（作为字符串映射）
     */
    std::unordered_map<std::string, std::string> getAllCustomProperties() const;
    
    // === 访问器 ===
    
    /**
     * @brief 获取完整的文档属性结构
     */
    const DocumentProperties& getDocumentProperties() const { return doc_properties_; }
    
    /**
     * @brief 设置完整的文档属性结构
     */
    void setDocumentProperties(const DocumentProperties& properties);
    
    // === 配置管理 ===
    
    Configuration& getConfiguration() { return config_; }
    const Configuration& getConfiguration() const { return config_; }
    
    // === 验证和工具 ===
    
    /**
     * @brief 验证属性名称是否有效
     */
    bool isValidPropertyName(const std::string& name) const;
    
    /**
     * @brief 验证属性值是否有效
     */
    bool isValidPropertyValue(const std::string& value) const;
    
    // === DirtyManager 访问 ===
    
    /**
     * @brief 获取脏数据管理器
     * @return 脏数据管理器指针
     */
    DirtyManager* getDirtyManager() { return dirty_manager_.get(); }
    const DirtyManager* getDirtyManager() const { return dirty_manager_.get(); }
    
    // === 定义名称管理 ===
    
    /**
     * @brief 定义名称
     * @param name 名称
     * @param formula 公式
     * @param scope 作用域（工作表名或空表示全局）
     */
    void defineName(const std::string& name, const std::string& formula, const std::string& scope = "");
    
    /**
     * @brief 获取定义名称的公式
     * @param name 名称
     * @param scope 作用域
     * @return 公式（如果不存在返回空字符串）
     */
    std::string getDefinedName(const std::string& name, const std::string& scope = "") const;
    
    /**
     * @brief 删除定义名称
     * @param name 名称
     * @param scope 作用域
     * @return 是否成功
     */
    bool removeDefinedName(const std::string& name, const std::string& scope = "");

private:
    void markAsModified();  // 标记工作簿为已修改
    bool validateProperty(const std::string& name, const std::string& value) const;
};

}} // namespace fastexcel::core
