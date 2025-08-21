#include "WorkbookSecurityManager.hpp"
#include <vector>
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Path.hpp"
#include <algorithm>
#include <random>
#include <fstream>
#include <iomanip>
#include <sstream>
#include <cctype>
#include <set>

namespace fastexcel {
namespace core {

WorkbookSecurityManager::WorkbookSecurityManager(Workbook* workbook)
    : workbook_(workbook) {
}

bool WorkbookSecurityManager::protect(const ProtectionOptions& options) {
    if (is_protected_) {
        CORE_WARN("Workbook is already protected");
        return false;
    }
    
    // 验证密码（如果提供了密码）
    if (!options.password.empty() && !isPasswordValid(options.password)) {
        CORE_ERROR("Password does not meet security policy requirements");
        return false;
    }
    
    // 应用保护设置
    is_protected_ = true;
    structure_locked_ = options.lock_structure;
    windows_locked_ = options.lock_windows;
    read_only_recommended_ = options.read_only_recommended;
    
    if (!options.password.empty()) {
        protection_password_hash_ = hashPassword(options.password);
    }
    
    markAsModified();
    CORE_INFO("Workbook protection enabled");
    return true;
}

bool WorkbookSecurityManager::protect(const std::string& password, bool lock_structure, bool lock_windows) {
    ProtectionOptions options;
    options.password = password;
    options.lock_structure = lock_structure;
    options.lock_windows = lock_windows;
    return protect(options);
}

bool WorkbookSecurityManager::unprotect(const std::string& password) {
    if (!is_protected_) {
        return true;  // 已经未保护状态
    }
    
    // 如果设置了密码，验证密码
    if (!protection_password_hash_.empty()) {
        if (password.empty() || !verifyPasswordHash(password, protection_password_hash_)) {
            CORE_ERROR("Invalid password for workbook unprotection");
            return false;
        }
    }
    
    // 移除保护
    is_protected_ = false;
    structure_locked_ = false;
    windows_locked_ = false;
    protection_password_hash_.clear();
    
    markAsModified();
    CORE_INFO("Workbook protection removed");
    return true;
}

bool WorkbookSecurityManager::verifyPassword(const std::string& password) const {
    if (protection_password_hash_.empty()) {
        return password.empty();  // 没有设置密码时，空密码有效
    }
    return verifyPasswordHash(password, protection_password_hash_);
}

bool WorkbookSecurityManager::changePassword(const std::string& old_password, const std::string& new_password) {
    if (!is_protected_) {
        CORE_ERROR("Cannot change password: workbook is not protected");
        return false;
    }
    
    // 验证旧密码
    if (!verifyPassword(old_password)) {
        CORE_ERROR("Invalid old password");
        return false;
    }
    
    // 验证新密码
    if (!new_password.empty() && !isPasswordValid(new_password)) {
        CORE_ERROR("New password does not meet security policy requirements");
        return false;
    }
    
    // 更新密码
    if (new_password.empty()) {
        protection_password_hash_.clear();
    } else {
        protection_password_hash_ = hashPassword(new_password);
    }
    
    markAsModified();
    CORE_INFO("Workbook password changed successfully");
    return true;
}

void WorkbookSecurityManager::setReadOnlyRecommended(bool recommend) {
    read_only_recommended_ = recommend;
    markAsModified();
}

bool WorkbookSecurityManager::addVbaProject(const std::string& vba_project_path) {
    if (vba_project_path.empty()) {
        CORE_ERROR("VBA project path cannot be empty");
        return false;
    }
    
    Path project_path(vba_project_path);
    if (!project_path.exists()) {
        CORE_ERROR("VBA project file does not exist: {}", vba_project_path);
        return false;
    }
    
    if (!loadVbaProjectInfo(vba_project_path)) {
        return false;
    }
    
    markAsModified();
    CORE_INFO("VBA project added: {}", vba_project_path);
    return true;
}

bool WorkbookSecurityManager::removeVbaProject() {
    if (!hasVbaProject()) {
        return true;  // 已经没有VBA项目
    }
    
    vba_project_.reset();
    markAsModified();
    CORE_INFO("VBA project removed");
    return true;
}

bool WorkbookSecurityManager::protectVbaProject(const std::string& password) {
    if (!hasVbaProject()) {
        CORE_ERROR("No VBA project to protect");
        return false;
    }
    
    if (!isPasswordValid(password)) {
        CORE_ERROR("VBA protection password does not meet security requirements");
        return false;
    }
    
    vba_project_->is_protected = true;
    vba_project_->protection_password = hashPassword(password);
    
    markAsModified();
    CORE_INFO("VBA project protection enabled");
    return true;
}

bool WorkbookSecurityManager::unprotectVbaProject(const std::string& password) {
    if (!hasVbaProject() || !vba_project_->is_protected) {
        return true;  // 没有VBA项目或未保护
    }
    
    if (!verifyPasswordHash(password, vba_project_->protection_password)) {
        CORE_ERROR("Invalid VBA project password");
        return false;
    }
    
    vba_project_->is_protected = false;
    vba_project_->protection_password.clear();
    
    markAsModified();
    CORE_INFO("VBA project protection removed");
    return true;
}

bool WorkbookSecurityManager::verifyVbaPassword(const std::string& password) const {
    if (!hasVbaProject() || !vba_project_->is_protected) {
        return true;  // 没有保护
    }
    return verifyPasswordHash(password, vba_project_->protection_password);
}

int WorkbookSecurityManager::checkPasswordStrength(const std::string& password) const {
    if (password.empty()) return 0;
    
    int score = 0;
    
    // 长度评分（最多25分）
    score += std::min(25, static_cast<int>(password.length()) * 2);
    
    // 字符多样性评分
    if (containsMixedCase(password)) score += 15;
    if (containsNumbers(password)) score += 15;
    if (containsSpecialChars(password)) score += 15;
    
    // 长度额外奖励
    if (password.length() >= 12) score += 10;
    if (password.length() >= 16) score += 10;
    
    // 复杂性奖励
    std::set<char> unique_chars(password.begin(), password.end());
    if (unique_chars.size() >= password.length() * 0.7) {
        score += 10;  // 字符重复率低
    }
    
    return std::min(100, score);
}

bool WorkbookSecurityManager::isPasswordValid(const std::string& password) const {
    if (config_.allow_weak_passwords) {
        return !password.empty();
    }
    
    if (password.length() < config_.min_password_length) {
        return false;
    }
    
    if (config_.require_mixed_case && !containsMixedCase(password)) {
        return false;
    }
    
    if (config_.require_numbers && !containsNumbers(password)) {
        return false;
    }
    
    if (config_.require_special_chars && !containsSpecialChars(password)) {
        return false;
    }
    
    return true;
}

std::string WorkbookSecurityManager::generateSecurePassword(size_t length, bool include_special_chars) const {
    if (length < 4) length = 4;  // 最小长度
    
    const std::string lowercase = "abcdefghijklmnopqrstuvwxyz";
    const std::string uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    const std::string numbers = "0123456789";
    const std::string special = "!@#$%^&*()_+-=[]{}|;:,.<>?";
    
    std::string charset = lowercase + uppercase + numbers;
    if (include_special_chars) {
        charset += special;
    }
    
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> dis(0, charset.length() - 1);
    
    std::string password;
    password.reserve(length);
    
    // 确保至少包含各种字符类型
    password += lowercase[dis(gen) % lowercase.length()];
    password += uppercase[dis(gen) % uppercase.length()];
    password += numbers[dis(gen) % numbers.length()];
    if (include_special_chars && length > 3) {
        password += special[dis(gen) % special.length()];
    }
    
    // 填充剩余字符
    while (password.length() < length) {
        password += charset[dis(gen)];
    }
    
    // 随机打乱
    std::shuffle(password.begin(), password.end(), gen);
    
    return password;
}

bool WorkbookSecurityManager::isOperationAllowed(Operation operation) const {
    if (!is_protected_) {
        return true;  // 未保护时允许所有操作
    }
    
    switch (operation) {
        case Operation::ADD_WORKSHEET:
        case Operation::DELETE_WORKSHEET:
        case Operation::RENAME_WORKSHEET:
        case Operation::MOVE_WORKSHEET:
        case Operation::MODIFY_STRUCTURE:
            return !structure_locked_;
            
        case Operation::CHANGE_WINDOW_STATE:
            return !windows_locked_;
    }
    
    return false;
}

bool WorkbookSecurityManager::requestPermission(Operation operation, const std::string& password) const {
    if (isOperationAllowed(operation)) {
        return true;
    }
    
    // 需要密码验证
    return verifyPassword(password);
}

void WorkbookSecurityManager::setPasswordPolicy(size_t min_length, bool require_mixed_case,
                                                bool require_numbers, bool require_special_chars) {
    config_.min_password_length = min_length;
    config_.require_mixed_case = require_mixed_case;
    config_.require_numbers = require_numbers;
    config_.require_special_chars = require_special_chars;
}

WorkbookSecurityManager::SecuritySummary WorkbookSecurityManager::getSecuritySummary() const {
    SecuritySummary summary;
    summary.workbook_protected = is_protected_;
    summary.vba_project_exists = hasVbaProject();
    summary.vba_project_protected = hasVbaProject() && vba_project_->is_protected;
    summary.read_only_recommended = read_only_recommended_;
    
    // 计算密码强度（仅作为示例，实际不应暴露密码）
    summary.password_strength_score = protection_password_hash_.empty() ? 0 : 75;  // 估算值
    
    return summary;
}

std::vector<std::string> WorkbookSecurityManager::performSecurityAudit() const {
    std::vector<std::string> issues;
    
    if (!is_protected_) {
        issues.push_back("Workbook is not protected");
    } else {
        if (protection_password_hash_.empty()) {
            issues.push_back("Workbook protection has no password");
        }
        
        if (!structure_locked_ && !windows_locked_) {
            issues.push_back("No structural protection enabled");
        }
    }
    
    if (hasVbaProject() && !vba_project_->is_protected) {
        issues.push_back("VBA project is not password protected");
    }
    
    if (!read_only_recommended_ && hasVbaProject()) {
        issues.push_back("Consider enabling read-only recommendation for VBA projects");
    }
    
    return issues;
}

// 私有方法实现

std::string WorkbookSecurityManager::hashPassword(const std::string& password) const {
    // 这是一个简化的哈希实现，实际应用中应该使用更安全的哈希算法
    // 如 bcrypt, scrypt 或 Argon2
    std::hash<std::string> hasher;
    size_t hash = hasher(password + "fastexcel_salt");
    
    std::ostringstream oss;
    oss << std::hex << hash;
    return oss.str();
}

bool WorkbookSecurityManager::verifyPasswordHash(const std::string& password, const std::string& hash) const {
    return hashPassword(password) == hash;
}

void WorkbookSecurityManager::markAsModified() {
    // 通知工作簿已修改
    if (workbook_ && workbook_->getDirtyManager()) {
        workbook_->getDirtyManager()->markDirty("xl/workbook.xml", DirtyManager::DirtyLevel::CONTENT);
    }
}

bool WorkbookSecurityManager::loadVbaProjectInfo(const std::string& path) {
    try {
        vba_project_ = std::make_unique<VbaProjectInfo>();
        vba_project_->path = path;
        
        Path project_path(path);
        vba_project_->file_size = project_path.fileSize();
        vba_project_->checksum = calculateFileChecksum(path);
        
        return true;
    } catch (const std::exception& e) {
        CORE_ERROR("Failed to load VBA project info: {}", e.what());
        vba_project_.reset();
        return false;
    }
}

std::string WorkbookSecurityManager::calculateFileChecksum(const std::string& file_path) const {
    // 简化的校验和计算（实际应用中应该使用SHA-256等）
    std::ifstream file(file_path, std::ios::binary);
    if (!file.is_open()) {
        return "";
    }
    
    std::hash<std::string> hasher;
    std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    size_t hash = hasher(content);
    
    std::ostringstream oss;
    oss << std::hex << hash;
    return oss.str();
}

bool WorkbookSecurityManager::containsMixedCase(const std::string& str) const {
    bool has_lower = std::any_of(str.begin(), str.end(), [](char c) { return std::islower(c); });
    bool has_upper = std::any_of(str.begin(), str.end(), [](char c) { return std::isupper(c); });
    return has_lower && has_upper;
}

bool WorkbookSecurityManager::containsNumbers(const std::string& str) const {
    return std::any_of(str.begin(), str.end(), [](char c) { return std::isdigit(c); });
}

bool WorkbookSecurityManager::containsSpecialChars(const std::string& str) const {
    return std::any_of(str.begin(), str.end(), [](char c) {
        return !std::isalnum(c) && !std::isspace(c);
    });
}

}} // namespace fastexcel::core