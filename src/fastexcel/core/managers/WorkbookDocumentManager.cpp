#include "WorkbookDocumentManager.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <cctype>
#include <sstream>

namespace fastexcel {
namespace core {

// DocumentProperties 实现
DocumentProperties::DocumentProperties() {
    // 使用 TimeUtils 获取当前时间
    created_time = utils::TimeUtils::getCurrentTime();
    modified_time = created_time;
}

// CustomProperty 类型转换实现
double WorkbookDocumentManager::CustomProperty::asDouble() const {
    try {
        return std::stod(value);
    } catch (const std::exception&) {
        return 0.0;
    }
}

bool WorkbookDocumentManager::CustomProperty::asBoolean() const {
    if (value == "true" || value == "1" || value == "yes") return true;
    if (value == "false" || value == "0" || value == "no") return false;
    return false;
}

std::tm WorkbookDocumentManager::CustomProperty::asDate() const {
    return utils::TimeUtils::parseTimeISO8601(value);
}

// WorkbookDocumentManager 实现
WorkbookDocumentManager::WorkbookDocumentManager(Workbook* workbook) 
    : workbook_(workbook),
      dirty_manager_(std::make_unique<DirtyManager>()),
      defined_name_manager_(std::make_unique<DefinedNameManager>()) {
    // 设置默认值
    doc_properties_.author = "FastExcel";
    doc_properties_.company = "FastExcel Library";
}

void WorkbookDocumentManager::setTitle(const std::string& title) {
    if (!validateProperty("title", title)) return;
    doc_properties_.title = title;
    markAsModified();
}

void WorkbookDocumentManager::setSubject(const std::string& subject) {
    if (!validateProperty("subject", subject)) return;
    doc_properties_.subject = subject;
    markAsModified();
}

void WorkbookDocumentManager::setAuthor(const std::string& author) {
    if (!validateProperty("author", author)) return;
    doc_properties_.author = author;
    markAsModified();
}

void WorkbookDocumentManager::setManager(const std::string& manager) {
    if (!validateProperty("manager", manager)) return;
    doc_properties_.manager = manager;
    markAsModified();
}

void WorkbookDocumentManager::setCompany(const std::string& company) {
    if (!validateProperty("company", company)) return;
    doc_properties_.company = company;
    markAsModified();
}

void WorkbookDocumentManager::setCategory(const std::string& category) {
    if (!validateProperty("category", category)) return;
    doc_properties_.category = category;
    markAsModified();
}

void WorkbookDocumentManager::setKeywords(const std::string& keywords) {
    if (!validateProperty("keywords", keywords)) return;
    doc_properties_.keywords = keywords;
    markAsModified();
}

void WorkbookDocumentManager::setComments(const std::string& comments) {
    if (!validateProperty("comments", comments)) return;
    doc_properties_.comments = comments;
    markAsModified();
}

void WorkbookDocumentManager::setStatus(const std::string& status) {
    if (!validateProperty("status", status)) return;
    doc_properties_.status = status;
    markAsModified();
}

void WorkbookDocumentManager::setHyperlinkBase(const std::string& hyperlink_base) {
    if (!validateProperty("hyperlink_base", hyperlink_base)) return;
    doc_properties_.hyperlink_base = hyperlink_base;
    markAsModified();
}

void WorkbookDocumentManager::setApplication(const std::string& application) {
    if (!validateProperty("application", application)) return;
    doc_properties_.application = application;
    markAsModified();
}

void WorkbookDocumentManager::setDocumentProperties(const std::string& title,
                                                   const std::string& subject,
                                                   const std::string& author,
                                                   const std::string& company,
                                                   const std::string& comments) {
    bool changed = false;
    
    if (!title.empty() && validateProperty("title", title)) {
        doc_properties_.title = title;
        changed = true;
    }
    
    if (!subject.empty() && validateProperty("subject", subject)) {
        doc_properties_.subject = subject;
        changed = true;
    }
    
    if (!author.empty() && validateProperty("author", author)) {
        doc_properties_.author = author;
        changed = true;
    }
    
    if (!company.empty() && validateProperty("company", company)) {
        doc_properties_.company = company;
        changed = true;
    }
    
    if (!comments.empty() && validateProperty("comments", comments)) {
        doc_properties_.comments = comments;
        changed = true;
    }
    
    if (changed) {
        markAsModified();
    }
}

void WorkbookDocumentManager::setCreatedTime(const std::tm& created_time) {
    doc_properties_.created_time = created_time;
    markAsModified();
}

void WorkbookDocumentManager::setModifiedTime(const std::tm& modified_time) {
    doc_properties_.modified_time = modified_time;
    markAsModified();
}

void WorkbookDocumentManager::updateModifiedTime() {
    doc_properties_.modified_time = utils::TimeUtils::getCurrentTime();
    markAsModified();
}

void WorkbookDocumentManager::setCustomProperty(const std::string& name, const std::string& value) {
    if (!isValidPropertyName(name) || !isValidPropertyValue(value)) return;
    
    if (custom_properties_.size() >= config_.max_custom_properties && 
        custom_properties_.find(name) == custom_properties_.end()) {
        CORE_WARN("Custom property limit reached ({}), ignoring property: {}", 
                  config_.max_custom_properties, name);
        return;
    }
    
    custom_properties_[name] = CustomProperty(value);
    markAsModified();
}

void WorkbookDocumentManager::setCustomProperty(const std::string& name, double value) {
    if (!isValidPropertyName(name)) return;
    
    if (custom_properties_.size() >= config_.max_custom_properties && 
        custom_properties_.find(name) == custom_properties_.end()) {
        CORE_WARN("Custom property limit reached ({}), ignoring property: {}", 
                  config_.max_custom_properties, name);
        return;
    }
    
    custom_properties_[name] = CustomProperty(value);
    markAsModified();
}

void WorkbookDocumentManager::setCustomProperty(const std::string& name, bool value) {
    if (!isValidPropertyName(name)) return;
    
    if (custom_properties_.size() >= config_.max_custom_properties && 
        custom_properties_.find(name) == custom_properties_.end()) {
        CORE_WARN("Custom property limit reached ({}), ignoring property: {}", 
                  config_.max_custom_properties, name);
        return;
    }
    
    custom_properties_[name] = CustomProperty(value);
    markAsModified();
}

void WorkbookDocumentManager::setCustomProperty(const std::string& name, const std::tm& value) {
    if (!isValidPropertyName(name)) return;
    
    if (custom_properties_.size() >= config_.max_custom_properties && 
        custom_properties_.find(name) == custom_properties_.end()) {
        CORE_WARN("Custom property limit reached ({}), ignoring property: {}", 
                  config_.max_custom_properties, name);
        return;
    }
    
    custom_properties_[name] = CustomProperty(value);
    markAsModified();
}

std::string WorkbookDocumentManager::getCustomProperty(const std::string& name, const std::string& default_value) const {
    auto it = custom_properties_.find(name);
    return (it != custom_properties_.end()) ? it->second.asString() : default_value;
}

double WorkbookDocumentManager::getCustomPropertyAsDouble(const std::string& name, double default_value) const {
    auto it = custom_properties_.find(name);
    return (it != custom_properties_.end()) ? it->second.asDouble() : default_value;
}

bool WorkbookDocumentManager::getCustomPropertyAsBoolean(const std::string& name, bool default_value) const {
    auto it = custom_properties_.find(name);
    return (it != custom_properties_.end()) ? it->second.asBoolean() : default_value;
}

std::tm WorkbookDocumentManager::getCustomPropertyAsDate(const std::string& name) const {
    auto it = custom_properties_.find(name);
    if (it != custom_properties_.end()) {
        return it->second.asDate();
    }
    return utils::TimeUtils::getCurrentTime();  // 返回当前时间作为默认值
}

bool WorkbookDocumentManager::hasCustomProperty(const std::string& name) const {
    return custom_properties_.find(name) != custom_properties_.end();
}

WorkbookDocumentManager::PropertyType WorkbookDocumentManager::getCustomPropertyType(const std::string& name) const {
    auto it = custom_properties_.find(name);
    return (it != custom_properties_.end()) ? it->second.type : PropertyType::STRING;
}

bool WorkbookDocumentManager::removeCustomProperty(const std::string& name) {
    auto it = custom_properties_.find(name);
    if (it != custom_properties_.end()) {
        custom_properties_.erase(it);
        markAsModified();
        return true;
    }
    return false;
}

std::vector<std::string> WorkbookDocumentManager::getCustomPropertyNames() const {
    std::vector<std::string> names;
    names.reserve(custom_properties_.size());
    
    for (const auto& [name, property] : custom_properties_) {
        names.push_back(name);
    }
    
    std::sort(names.begin(), names.end());
    return names;
}

void WorkbookDocumentManager::clearCustomProperties() {
    if (!custom_properties_.empty()) {
        custom_properties_.clear();
        markAsModified();
    }
}

void WorkbookDocumentManager::setCustomProperties(const std::unordered_map<std::string, std::string>& properties) {
    bool changed = false;
    
    for (const auto& [name, value] : properties) {
        if (isValidPropertyName(name) && isValidPropertyValue(value)) {
            if (custom_properties_.size() < config_.max_custom_properties || 
                custom_properties_.find(name) != custom_properties_.end()) {
                custom_properties_[name] = CustomProperty(value);
                changed = true;
            }
        }
    }
    
    if (changed) {
        markAsModified();
    }
}

std::unordered_map<std::string, std::string> WorkbookDocumentManager::getAllCustomProperties() const {
    std::unordered_map<std::string, std::string> result;
    result.reserve(custom_properties_.size());
    
    for (const auto& [name, property] : custom_properties_) {
        result[name] = property.asString();
    }
    
    return result;
}

void WorkbookDocumentManager::setDocumentProperties(const DocumentProperties& properties) {
    doc_properties_ = properties;
    markAsModified();
}

bool WorkbookDocumentManager::isValidPropertyName(const std::string& name) const {
    if (name.empty() || name.length() > config_.max_property_length) {
        return false;
    }
    
    // 检查是否包含无效字符
    return std::all_of(name.begin(), name.end(), [](char c) {
        return std::isalnum(c) || c == '_' || c == '-' || c == '.';
    });
}

bool WorkbookDocumentManager::isValidPropertyValue(const std::string& value) const {
    return value.length() <= config_.max_property_length;
}

void WorkbookDocumentManager::markAsModified() {
    if (config_.auto_update_modified_time) {
        doc_properties_.modified_time = utils::TimeUtils::getCurrentTime();
    }
    
    // 通知脏数据管理器已修改
    if (dirty_manager_) {
        dirty_manager_->markDirty("docProps/core.xml", DirtyManager::DirtyLevel::CONTENT);
        dirty_manager_->markDirty("docProps/app.xml", DirtyManager::DirtyLevel::CONTENT);
        dirty_manager_->markDirty("docProps/custom.xml", DirtyManager::DirtyLevel::CONTENT);
    }
}

bool WorkbookDocumentManager::validateProperty(const std::string& name, const std::string& value) const {
    if (!config_.validate_properties) return true;
    return isValidPropertyName(name) && isValidPropertyValue(value);
}

// === 定义名称管理实现 ===

void WorkbookDocumentManager::defineName(const std::string& name, const std::string& formula, const std::string& scope) {
    if (!defined_name_manager_) return;
    defined_name_manager_->define(name, formula, scope);
    markAsModified();
}

std::string WorkbookDocumentManager::getDefinedName(const std::string& name, const std::string& scope) const {
    if (!defined_name_manager_) return std::string();
    return defined_name_manager_->get(name, scope);
}

bool WorkbookDocumentManager::removeDefinedName(const std::string& name, const std::string& scope) {
    if (!defined_name_manager_) return false;
    bool result = defined_name_manager_->remove(name, scope);
    if (result) {
        markAsModified();
    }
    return result;
}

}} // namespace fastexcel::core