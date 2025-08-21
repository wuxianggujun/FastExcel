#include "fastexcel/utils/ModuleLoggers.hpp"
#include "fastexcel/opc/PackageEditorManager.hpp"

namespace fastexcel {
namespace opc {

PackageEditorManager::PackageEditorManager(std::unique_ptr<archive::ZipReader> reader) 
    : zip_reader_(std::move(reader)) {
}

// 读取操作实现

bool PackageEditorManager::openForReading(const core::Path& path) {
    zip_reader_ = std::make_unique<archive::ZipReader>(path);
    readable_ = zip_reader_->open();
    
    if (!readable_) {
        OPC_ERROR("Failed to open ZIP file for reading: {}", path.string());
    } else {
        OPC_DEBUG("Opened ZIP file for reading: {} ({} parts)", 
                  path.string(), zip_reader_->listFiles().size());
    }
    
    return readable_;
}

std::string PackageEditorManager::readPart(const std::string& part_name) {
    if (!zip_reader_) {
        OPC_ERROR("No ZIP reader available for reading part: {}", part_name);
        return "";
    }
    
    std::string content;
    if (zip_reader_->extractFile(part_name, content) == archive::ZipError::Ok) {
        OPC_DEBUG("Successfully read part: {} ({} bytes)", part_name, content.size());
        return content;
    }
    
    OPC_WARN("Failed to read part: {}", part_name);
    return "";
}

bool PackageEditorManager::partExists(const std::string& part_name) const {
    if (!zip_reader_) return false;
    
    auto files = zip_reader_->listFiles();
    return std::find(files.begin(), files.end(), part_name) != files.end();
}

std::vector<std::string> PackageEditorManager::listParts() const {
    if (!zip_reader_) {
        OPC_WARN("No ZIP reader available for listing parts");
        return {};
    }
    
    return zip_reader_->listFiles();
}

// 写入操作实现

bool PackageEditorManager::openForWriting(const core::Path& path) {
    target_path_ = path;
    writable_ = true;
    
    OPC_DEBUG("Opened for writing to: {}", path.string());
    return true;
}

bool PackageEditorManager::writePart(const std::string& part_name, const std::string& content) {
    if (!writable_) {
        OPC_ERROR("Package not opened for writing, cannot write part: {}", part_name);
        return false;
    }
    
    pending_writes_[part_name] = content;
    OPC_DEBUG("Queued part for writing: {} ({} bytes)", part_name, content.size());
    return true;
}

bool PackageEditorManager::removePart(const std::string& part_name) {
    removed_parts_.insert(part_name);
    pending_writes_.erase(part_name); // 确保不会写入被删除的部件
    
    OPC_DEBUG("Marked part for removal: {}", part_name);
    return true;
}

bool PackageEditorManager::commit() {
    if (!writable_ || target_path_.empty()) {
        OPC_ERROR("Cannot commit: package not properly opened for writing");
        return false;
    }

    OPC_INFO("Committing {} pending writes, {} removals to: {}", 
             pending_writes_.size(), removed_parts_.size(), target_path_.string());

    // 使用ZipRepackWriter实现实际的写入
    ZipRepackWriter writer(target_path_);
    
    // 写入所有待写入的部件
    for (const auto& [path, content] : pending_writes_) {
        if (removed_parts_.count(path)) {
            OPC_DEBUG("Skipping removed part: {}", path);
            continue;
        }
        
        if (!writer.add(path, content)) {
            OPC_ERROR("Failed to add part to writer: {}", path);
            return false;
        }
        
        OPC_DEBUG("Added part to writer: {} ({} bytes)", path, content.size());
    }
    
    // 复制未修改的部件
    if (zip_reader_) {
        auto all_files = zip_reader_->listFiles();
        std::vector<std::string> to_copy;
        
        for (const auto& file : all_files) {
            if (pending_writes_.count(file) || removed_parts_.count(file)) {
                continue;
            }
            to_copy.push_back(file);
        }
        
        if (!to_copy.empty()) {
            OPC_DEBUG("Copying {} unchanged parts", to_copy.size());
            if (!writer.copyBatch(zip_reader_.get(), to_copy)) {
                OPC_ERROR("Failed to copy unchanged parts");
                return false;
            }
        }
    }
    
    // 完成写入
    bool success = writer.finish();
    
    if (success) {
        OPC_INFO("Successfully committed package to: {}", target_path_.string());
        
        // 清理状态
        pending_writes_.clear();
        removed_parts_.clear();
    } else {
        OPC_ERROR("Failed to commit package to: {}", target_path_.string());
    }
    
    return success;
}

// 状态查询实现

bool PackageEditorManager::isReadable() const {
    return readable_ && zip_reader_ != nullptr;
}

bool PackageEditorManager::isWritable() const {
    return writable_;
}

size_t PackageEditorManager::getPartCount() const {
    if (zip_reader_) {
        return zip_reader_->listFiles().size();
    }
    return 0;
}

}} // namespace fastexcel::opc
