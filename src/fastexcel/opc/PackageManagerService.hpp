#include "fastexcel/utils/ModuleLoggers.hpp"
#pragma once

#include "fastexcel/opc/IPackageManager.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipWriter.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <memory>
#include <string>
#include <vector>
#include <unordered_set>
#include <unordered_map>

namespace fastexcel {
namespace opc {

/**
 * @brief 标准OPC包管理器实现
 * 遵循单一职责原则 (SRP)：专注于ZIP包的读写操作
 * 遵循依赖倒置原则 (DIP)：依赖抽象的ZipReader/ZipWriter接口
 */
class StandardPackageManager : public IPackageManager {
private:
    std::unique_ptr<archive::ZipReader> reader_;
    std::unique_ptr<archive::ZipWriter> writer_;
    core::Path package_path_;
    
    // 缓存部件列表，提高性能
    mutable std::vector<std::string> cached_parts_;
    mutable bool parts_cached_ = false;
    
    // 跟踪修改的部件
    std::unordered_set<std::string> modified_parts_;
    std::unordered_map<std::string, std::string> new_content_;
    std::unordered_set<std::string> removed_parts_;
    
    void invalidateCache() { parts_cached_ = false; cached_parts_.clear(); }
    
public:
    StandardPackageManager() = default;
    ~StandardPackageManager() = default;
    
    // 禁用拷贝构造和赋值，因为管理着资源
    StandardPackageManager(const StandardPackageManager&) = delete;
    StandardPackageManager& operator=(const StandardPackageManager&) = delete;
    
    // 支持移动语义
    StandardPackageManager(StandardPackageManager&&) = default;
    StandardPackageManager& operator=(StandardPackageManager&&) = default;
    
    // 读取操作
    
    bool openForReading(const core::Path& path) override {
        package_path_ = path;
        
        if (!path.exists()) {
            OPC_ERROR("Package file does not exist: {}", path.string());
            return false;
        }
        
        try {
            reader_ = std::make_unique<archive::ZipReader>(path);
            if (!reader_->isOpen()) {
                OPC_ERROR("Failed to open package for reading: {}", path.string());
                reader_.reset();
                return false;
            }
            
            OPC_INFO("Opened package for reading: {}", path.string());
            return true;
        }
        catch (const std::exception& e) {
            OPC_ERROR("Exception opening package for reading: {} - {}", path.string(), e.what());
            reader_.reset();
            return false;
        }
    }
    
    std::string readPart(const std::string& part_name) override {
        if (!reader_) {
            OPC_ERROR("Package not open for reading");
            return "";
        }
        
        std::string content;
        auto result = reader_->extractFile(part_name, content);
        
        if (result != archive::ZipError::Ok) {
            OPC_WARN("Failed to read part '{}': error code {}", part_name, static_cast<int>(result));
            return "";
        }
        
        OPC_DEBUG("Read part '{}': {} bytes", part_name, content.size());
        return content;
    }
    
    bool partExists(const std::string& part_name) const override {
        if (!reader_) return false;
        
        // 先检查缓存
        if (!parts_cached_) {
            cached_parts_ = reader_->listFiles();
            parts_cached_ = true;
        }
        
        return std::find(cached_parts_.begin(), cached_parts_.end(), part_name) != cached_parts_.end();
    }
    
    std::vector<std::string> listParts() const override {
        if (!reader_) return {};
        
        if (!parts_cached_) {
            cached_parts_ = reader_->listFiles();
            parts_cached_ = true;
        }
        
        return cached_parts_;
    }
    
    // 写入操作
    
    bool openForWriting(const core::Path& path) override {
        package_path_ = path;
        
        try {
            writer_ = std::make_unique<archive::ZipWriter>(path);
            if (!writer_->isOpen()) {
                OPC_ERROR("Failed to open package for writing: {}", path.string());
                writer_.reset();
                return false;
            }
            
            OPC_INFO("Opened package for writing: {}", path.string());
            return true;
        }
        catch (const std::exception& e) {
            OPC_ERROR("Exception opening package for writing: {} - {}", path.string(), e.what());
            writer_.reset();
            return false;
        }
    }
    
    bool writePart(const std::string& part_name, const std::string& content) override {
        if (!writer_) {
            OPC_ERROR("Package not open for writing");
            return false;
        }
        
        // 记录修改，但还不写入（延迟写入策略）
        new_content_[part_name] = content;
        modified_parts_.insert(part_name);
        removed_parts_.erase(part_name);  // 如果之前标记删除，现在取消
        
        OPC_DEBUG("Staged part '{}' for writing: {} bytes", part_name, content.size());
        return true;
    }
    
    bool removePart(const std::string& part_name) override {
        if (!writer_) {
            OPC_ERROR("Package not open for writing");
            return false;
        }
        
        removed_parts_.insert(part_name);
        new_content_.erase(part_name);
        modified_parts_.erase(part_name);
        
        OPC_DEBUG("Staged part '{}' for removal", part_name);
        return true;
    }
    
    bool commit() override {
        if (!writer_) {
            OPC_ERROR("Package not open for writing");
            return false;
        }
        
        try {
            // 阶段1：写入新的和修改的部件
            for (const auto& [part_name, content] : new_content_) {
                auto result = writer_->addFile(part_name, content);
                if (result != archive::ZipError::Ok) {
                    OPC_ERROR("Failed to write part '{}': error code {}", part_name, static_cast<int>(result));
                    return false;
                }
                OPC_DEBUG("Written part '{}': {} bytes", part_name, content.size());
            }
            
            // 阶段2：如果需要从原包复制未修改的部件
            if (reader_ && reader_->isOpen()) {
                auto all_parts = listParts();
                for (const auto& part : all_parts) {
                    // 跳过已删除和已修改的部件
                    if (removed_parts_.count(part) || modified_parts_.count(part)) {
                        continue;
                    }
                    
                    std::string content = readPart(part);
                    if (!content.empty() || partExists(part)) {  // 空内容也是有效的
                        auto result = writer_->addFile(part, content);
                        if (result != archive::ZipError::Ok) {
                            OPC_WARN("Failed to copy unchanged part '{}'", part);
                            // 继续处理其他部件，不要因为一个部件失败就全部失败
                        } else {
                            OPC_DEBUG("Copied unchanged part '{}': {} bytes", part, content.size());
                        }
                    }
                }
            }
            
            // 关闭写入器，实际提交到磁盘
            writer_.reset();
            
            // 清空修改跟踪
            new_content_.clear();
            modified_parts_.clear();
            removed_parts_.clear();
            invalidateCache();
            
            OPC_INFO("Successfully committed changes to package: {}", package_path_.string());
            return true;
        }
        catch (const std::exception& e) {
            OPC_ERROR("Exception committing package changes: {}", e.what());
            return false;
        }
    }
    
    // 状态查询
    
    bool isReadable() const override {
        return reader_ && reader_->isOpen();
    }
    
    bool isWritable() const override {
        return writer_ && writer_->isOpen();
    }
    
    size_t getPartCount() const override {
        return listParts().size();
    }
    
    // 扩展功能
    
    /**
     * @brief 获取修改统计信息
     */
    struct ModificationStats {
        size_t modified_parts_count = 0;
        size_t removed_parts_count = 0;
        size_t total_content_size = 0;
    };
    
    ModificationStats getModificationStats() const {
        ModificationStats stats;
        stats.modified_parts_count = modified_parts_.size();
        stats.removed_parts_count = removed_parts_.size();
        
        for (const auto& [part, content] : new_content_) {
            stats.total_content_size += content.size();
        }
        
        return stats;
    }
    
    /**
     * @brief 检查是否有待提交的修改
     */
    bool hasPendingChanges() const {
        return !new_content_.empty() || !removed_parts_.empty();
    }
};

/**
 * @brief 内存包管理器 - 用于测试和简单场景
 * 将所有内容保存在内存中，不涉及实际文件操作
 */
class MemoryPackageManager : public IPackageManager {
private:
    std::unordered_map<std::string, std::string> parts_;
    bool readable_ = false;
    bool writable_ = false;
    
public:
    bool openForReading(const core::Path& path) override {
        (void)path;
        readable_ = true;
        return true;
    }
    
    std::string readPart(const std::string& part_name) override {
        auto it = parts_.find(part_name);
        return it != parts_.end() ? it->second : "";
    }
    
    bool partExists(const std::string& part_name) const override {
        return parts_.count(part_name) > 0;
    }
    
    std::vector<std::string> listParts() const override {
        std::vector<std::string> result;
        for (const auto& [part, content] : parts_) {
            result.push_back(part);
        }
        return result;
    }
    
    bool openForWriting(const core::Path& path) override {
        (void)path;
        writable_ = true;
        return true;
    }
    
    bool writePart(const std::string& part_name, const std::string& content) override {
        parts_[part_name] = content;
        return true;
    }
    
    bool removePart(const std::string& part_name) override {
        parts_.erase(part_name);
        return true;
    }
    
    bool commit() override {
        return true;  // 内存版本立即生效
    }
    
    bool isReadable() const override { return readable_; }
    bool isWritable() const override { return writable_; }
    size_t getPartCount() const override { return parts_.size(); }
};

} // namespace opc
} // namespace fastexcel
