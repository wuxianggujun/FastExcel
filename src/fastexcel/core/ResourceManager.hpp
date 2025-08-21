#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include <memory>
#include <string>
#include <vector>
#include <functional>

namespace fastexcel {
namespace core {

// 前向声明
class Workbook;
class IFileWriter;

/**
 * @brief 资源管理器 - 负责文件资源的管理和I/O操作
 * 
 * 遵循单一职责原则(SRP)：专注于文件资源管理
 * 实现延迟写入、原子保存、透传复制等核心功能
 */
class ResourceManager {
public:
    // 资源管理模式
    enum class Mode {
        READ_ONLY,      // 只读模式
        WRITE_NEW,      // 写入新文件
        EDIT_EXISTING   // 编辑现有文件
    };

    // 保存策略
    struct SaveStrategy {
        bool use_temp_file = true;          // 使用临时文件
        bool atomic_replace = true;         // 原子替换
        bool preserve_backup = false;       // 保留备份文件
        std::vector<std::string> skip_components;  // 跳过的组件列表
    };

private:
    std::unique_ptr<archive::FileManager> file_manager_;
    std::string filename_;
    std::string original_package_path_;
    Mode mode_;
    bool is_open_ = false;
    bool delayed_write_mode_ = true;  // 延迟写入模式
    
    // 默认的核心组件跳过列表
    static const std::vector<std::string> CORE_COMPONENTS_TO_REBUILD;
    static const std::vector<std::string> SAFE_TO_PASSTHROUGH;

public:
    ResourceManager();
    explicit ResourceManager(const Path& path, Mode mode = Mode::WRITE_NEW);
    ~ResourceManager();

    // 禁用拷贝，允许移动
    ResourceManager(const ResourceManager&) = delete;
    ResourceManager& operator=(const ResourceManager&) = delete;
    ResourceManager(ResourceManager&&) = default;
    ResourceManager& operator=(ResourceManager&&) = default;

    // 核心文件操作
    
    /**
     * @brief 打开资源管理器
     * @param create_if_not_exists 如果文件不存在是否创建
     * @return 是否成功
     */
    bool open(bool create_if_not_exists = false);
    
    /**
     * @brief 关闭资源管理器
     * @return 是否成功
     */
    bool close();
    
    /**
     * @brief 检查是否已打开
     */
    bool isOpen() const { return is_open_; }
    
    /**
     * @brief 为编辑模式准备（延迟打开）
     * @param path 文件路径
     * @param original_path 原始文件路径（用于透传复制）
     * @return 是否成功
     */
    bool prepareForEditing(const Path& path, const std::string& original_path);
    
    /**
     * @brief 原子保存操作
     * @param workbook 工作簿对象
     * @param strategy 保存策略
     * @return 是否成功
     */
    bool atomicSave(const Workbook* workbook, const SaveStrategy& strategy = SaveStrategy());
    
    /**
     * @brief 另存为
     * @param new_path 新路径
     * @param workbook 工作簿对象
     * @return 是否成功
     */
    bool saveAs(const Path& new_path, const Workbook* workbook);
    
    // 透传复制功能
    
    /**
     * @brief 从原始包复制非核心组件
     * @param source_path 源文件路径
     * @param skip_prefixes 跳过的前缀列表
     * @return 是否成功
     */
    bool copyFromOriginalPackage(const Path& source_path, 
                                 const std::vector<std::string>& skip_prefixes);
    
    /**
     * @brief 智能透传复制（自动判断跳过列表）
     * @param source_path 源文件路径
     * @param preserve_media 是否保留媒体文件
     * @param preserve_vba 是否保留VBA项目
     * @return 是否成功
     */
    bool smartPassthrough(const Path& source_path, 
                         bool preserve_media = true, 
                         bool preserve_vba = true);
    
    // 文件写入接口
    
    /**
     * @brief 写入文件内容
     * @param internal_path 内部路径
     * @param content 内容
     * @return 是否成功
     */
    bool writeFile(const std::string& internal_path, const std::string& content);
    bool writeFile(const std::string& internal_path, const std::vector<uint8_t>& data);
    
    /**
     * @brief 批量写入文件
     * @param files 文件列表
     * @return 是否成功
     */
    bool writeFiles(const std::vector<std::pair<std::string, std::string>>& files);
    
    /**
     * @brief 获取文件写入器
     * @param use_streaming 是否使用流式写入
     * @return 文件写入器
     */
    std::unique_ptr<IFileWriter> createFileWriter(bool use_streaming = false);
    
    // 压缩设置
    
    /**
     * @brief 设置压缩级别
     * @param level 压缩级别 (0-9)
     * @return 是否成功
     */
    bool setCompressionLevel(int level);
    
    // 临时文件管理
    
    /**
     * @brief 创建临时文件路径
     * @param base_path 基础路径
     * @return 临时文件路径
     */
    static std::string createTempPath(const std::string& base_path);
    
    /**
     * @brief 原子替换文件
     * @param temp_path 临时文件路径
     * @param target_path 目标文件路径
     * @return 是否成功
     */
    static bool atomicReplace(const Path& temp_path, const Path& target_path);
    
    // 状态查询
    
    Mode getMode() const { return mode_; }
    const std::string& getFilename() const { return filename_; }
    const std::string& getOriginalPath() const { return original_package_path_; }
    archive::FileManager* getFileManager() { return file_manager_.get(); }
    const archive::FileManager* getFileManager() const { return file_manager_.get(); }
    
    /**
     * @brief 检查是否需要透传复制
     * @return 是否需要
     */
    bool needsPassthrough() const {
        return mode_ == Mode::EDIT_EXISTING && !original_package_path_.empty();
    }

private:
    bool openInternal(bool for_writing);
    bool saveInternal(const Workbook* workbook, const SaveStrategy& strategy);
    std::vector<std::string> getSkipPrefixes(const SaveStrategy& strategy) const;
    bool handleSameFileSave(const SaveStrategy& strategy);
    void cleanupTempFiles(const std::vector<std::string>& temp_files);
};

}} // namespace fastexcel::core
