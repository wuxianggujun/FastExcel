#pragma once

#include <string>
#include <memory>
#include <vector>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;

/**
 * @brief 工作簿安全管理器 - 负责工作簿的安全功能
 * 
 * 设计原则：
 * - 单一职责：专注于安全相关功能
 * - 数据保护：敏感信息（如密码）的安全处理
 * - 功能完整：支持多种保护级别
 */
class WorkbookSecurityManager {
public:
    // 保护选项
    struct ProtectionOptions {
        bool lock_structure = true;      // 锁定结构（不能添加/删除工作表）
        bool lock_windows = false;       // 锁定窗口
        bool read_only_recommended = false;  // 建议只读
        std::string password;            // 保护密码
        
        ProtectionOptions() = default;
        explicit ProtectionOptions(const std::string& pwd) : password(pwd) {}
    };
    
    // VBA项目信息
    struct VbaProjectInfo {
        std::string path;                // VBA项目文件路径
        bool is_protected = false;       // 是否受保护
        std::string protection_password; // 保护密码
        size_t file_size = 0;           // 文件大小
        std::string checksum;            // 文件校验和
    };

private:
    Workbook* workbook_;  // 父工作簿引用（不拥有）
    
    // 保护状态
    bool is_protected_ = false;
    bool structure_locked_ = false;
    bool windows_locked_ = false;
    bool read_only_recommended_ = false;
    
    // 密码管理（注意：实际应用中应该使用更安全的方式存储）
    std::string protection_password_hash_;  // 密码哈希
    
    // VBA项目
    std::unique_ptr<VbaProjectInfo> vba_project_;
    
    // 配置
    struct Configuration {
        bool allow_weak_passwords = false;    // 允许弱密码
        size_t min_password_length = 8;       // 最小密码长度
        bool require_mixed_case = false;      // 要求大小写混合
        bool require_numbers = false;         // 要求包含数字
        bool require_special_chars = false;   // 要求包含特殊字符
    } config_;

public:
    explicit WorkbookSecurityManager(Workbook* workbook);
    ~WorkbookSecurityManager() = default;
    
    // 禁用拷贝，允许移动
    WorkbookSecurityManager(const WorkbookSecurityManager&) = delete;
    WorkbookSecurityManager& operator=(const WorkbookSecurityManager&) = delete;
    WorkbookSecurityManager(WorkbookSecurityManager&&) = default;
    WorkbookSecurityManager& operator=(WorkbookSecurityManager&&) = default;
    
    // === 工作簿保护功能 ===
    
    /**
     * @brief 保护工作簿
     * @param options 保护选项
     * @return 是否成功
     */
    bool protect(const ProtectionOptions& options);
    // 便捷重载：使用默认选项
    bool protect();
    
    /**
     * @brief 保护工作簿（简化版本）
     * @param password 密码
     * @param lock_structure 锁定结构
     * @param lock_windows 锁定窗口
     * @return 是否成功
     */
    bool protect(const std::string& password, bool lock_structure = true, bool lock_windows = false);
    
    /**
     * @brief 取消保护
     * @param password 密码（如果之前设置了密码）
     * @return 是否成功
     */
    bool unprotect(const std::string& password = "");
    
    /**
     * @brief 检查是否受保护
     */
    bool isProtected() const { return is_protected_; }
    
    /**
     * @brief 检查结构是否被锁定
     */
    bool isStructureLocked() const { return structure_locked_; }
    
    /**
     * @brief 检查窗口是否被锁定
     */
    bool isWindowsLocked() const { return windows_locked_; }
    
    /**
     * @brief 检查是否建议只读
     */
    bool isReadOnlyRecommended() const { return read_only_recommended_; }
    
    /**
     * @brief 验证保护密码
     * @param password 要验证的密码
     * @return 是否正确
     */
    bool verifyPassword(const std::string& password) const;
    
    /**
     * @brief 更改保护密码
     * @param old_password 旧密码
     * @param new_password 新密码
     * @return 是否成功
     */
    bool changePassword(const std::string& old_password, const std::string& new_password);
    
    /**
     * @brief 设置建议只读
     * @param recommend 是否建议只读
     */
    void setReadOnlyRecommended(bool recommend);
    
    // === VBA项目管理 ===
    
    /**
     * @brief 添加VBA项目
     * @param vba_project_path VBA项目文件路径
     * @return 是否成功
     */
    bool addVbaProject(const std::string& vba_project_path);
    
    /**
     * @brief 删除VBA项目
     * @return 是否成功
     */
    bool removeVbaProject();
    
    /**
     * @brief 检查是否有VBA项目
     */
    bool hasVbaProject() const { return vba_project_ != nullptr; }
    
    /**
     * @brief 获取VBA项目信息
     */
    const VbaProjectInfo* getVbaProjectInfo() const { return vba_project_.get(); }
    
    /**
     * @brief 保护VBA项目
     * @param password 保护密码
     * @return 是否成功
     */
    bool protectVbaProject(const std::string& password);
    
    /**
     * @brief 取消VBA项目保护
     * @param password 保护密码
     * @return 是否成功
     */
    bool unprotectVbaProject(const std::string& password);
    
    /**
     * @brief 验证VBA项目密码
     * @param password 密码
     * @return 是否正确
     */
    bool verifyVbaPassword(const std::string& password) const;
    
    // === 密码验证和强度检查 ===
    
    /**
     * @brief 检查密码强度
     * @param password 密码
     * @return 强度分数（0-100）
     */
    int checkPasswordStrength(const std::string& password) const;
    
    /**
     * @brief 验证密码是否符合策略
     * @param password 密码
     * @return 是否符合策略
     */
    bool isPasswordValid(const std::string& password) const;
    
    /**
     * @brief 生成安全的随机密码
     * @param length 密码长度
     * @param include_special_chars 是否包含特殊字符
     * @return 生成的密码
     */
    std::string generateSecurePassword(size_t length = 12, bool include_special_chars = true) const;
    
    // === 访问控制 ===
    
    /**
     * @brief 检查是否允许操作（基于保护状态）
     * @param operation 操作类型
     * @return 是否允许
     */
    enum class Operation {
        ADD_WORKSHEET,
        DELETE_WORKSHEET,
        RENAME_WORKSHEET,
        MOVE_WORKSHEET,
        MODIFY_STRUCTURE,
        CHANGE_WINDOW_STATE
    };
    bool isOperationAllowed(Operation operation) const;
    
    /**
     * @brief 请求操作权限
     * @param operation 操作类型
     * @param password 密码（如果需要）
     * @return 是否允许
     */
    bool requestPermission(Operation operation, const std::string& password = "") const;
    
    // === 配置管理 ===
    
    Configuration& getConfiguration() { return config_; }
    const Configuration& getConfiguration() const { return config_; }
    
    /**
     * @brief 设置密码策略
     */
    void setPasswordPolicy(size_t min_length, bool require_mixed_case, 
                          bool require_numbers, bool require_special_chars);
    
    // === 安全审计 ===
    
    /**
     * @brief 获取安全状态摘要
     */
    struct SecuritySummary {
        bool workbook_protected;
        bool vba_project_exists;
        bool vba_project_protected;
        bool read_only_recommended;
        int password_strength_score;
    };
    SecuritySummary getSecuritySummary() const;
    
    /**
     * @brief 执行安全检查
     * @return 安全问题列表
     */
    std::vector<std::string> performSecurityAudit() const;

private:
    // 内部辅助方法
    std::string hashPassword(const std::string& password) const;
    bool verifyPasswordHash(const std::string& password, const std::string& hash) const;
    void markAsModified();  // 标记工作簿为已修改
    bool loadVbaProjectInfo(const std::string& path);
    std::string calculateFileChecksum(const std::string& file_path) const;
    bool containsMixedCase(const std::string& str) const;
    bool containsNumbers(const std::string& str) const;
    bool containsSpecialChars(const std::string& str) const;
};

}} // namespace fastexcel::core
