#pragma once

#include "fastexcel/opc/IPackageManager.hpp"
#include "fastexcel/opc/PackageManagerService.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/opc/ZipRepackWriter.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <memory>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <algorithm>

namespace fastexcel {
namespace opc {

/**
 * @brief PackageEditor专用的包管理器实现
 * 
 * 职责：
 * - ZIP文件的读取和写入
 * - 部件的增删改查
 * - 状态管理和错误处理
 * 
 * 设计原则：
 * - 单一职责：专注于包文件管理
 * - 依赖倒置：实现IPackageManager接口
 * - 错误安全：提供详细的状态检查和错误报告
 */
class PackageEditorManager : public IPackageManager {
public:
    explicit PackageEditorManager(std::unique_ptr<archive::ZipReader> reader = nullptr);
    ~PackageEditorManager() = default;

    // 禁用拷贝，支持移动
    PackageEditorManager(const PackageEditorManager&) = delete;
    PackageEditorManager& operator=(const PackageEditorManager&) = delete;
    PackageEditorManager(PackageEditorManager&&) = default;
    PackageEditorManager& operator=(PackageEditorManager&&) = default;

    // ========== IPackageManager 接口实现 ==========
    
    // 读取操作
    bool openForReading(const core::Path& path) override;
    std::string readPart(const std::string& part_name) override;
    bool partExists(const std::string& part_name) const override;
    std::vector<std::string> listParts() const override;
    
    // 写入操作
    bool openForWriting(const core::Path& path) override;
    bool writePart(const std::string& part_name, const std::string& content) override;
    bool removePart(const std::string& part_name) override;
    bool commit() override;
    
    // 状态查询
    bool isReadable() const override;
    bool isWritable() const override;
    size_t getPartCount() const override;

private:
    std::unique_ptr<archive::ZipReader> zip_reader_;
    std::unordered_map<std::string, std::string> pending_writes_;
    std::unordered_set<std::string> removed_parts_;
    core::Path target_path_;
    bool readable_ = false;
    bool writable_ = false;
};

}} // namespace fastexcel::opc