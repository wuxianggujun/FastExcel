#include "FileWrapper.hpp"
#include <random>
#include <chrono>
#include <filesystem>
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace utils {

TempFileWrapper::TempFileWrapper(const std::string& prefix, const std::string& suffix)
    : temp_path_(generateTempPath(prefix, suffix))
    , file_(temp_path_, "wb")
    , should_delete_(true) {
    FASTEXCEL_LOG_DEBUG("Created temporary file: {}", temp_path_);
}

TempFileWrapper::~TempFileWrapper() {
    if (should_delete_ && !temp_path_.empty()) {
        try {
            std::filesystem::remove(temp_path_);
            FASTEXCEL_LOG_DEBUG("Deleted temporary file: {}", temp_path_);
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_WARN("Failed to delete temporary file {}: {}", temp_path_, e.what());
        }
    }
}

TempFileWrapper::TempFileWrapper(TempFileWrapper&& other) noexcept
    : temp_path_(std::move(other.temp_path_))
    , file_(std::move(other.file_))
    , should_delete_(other.should_delete_) {
    other.should_delete_ = false; // 转移所有权
}

TempFileWrapper& TempFileWrapper::operator=(TempFileWrapper&& other) noexcept {
    if (this != &other) {
        // 清理当前资源
        if (should_delete_ && !temp_path_.empty()) {
            try {
                std::filesystem::remove(temp_path_);
            } catch (const std::filesystem::filesystem_error& e) {
                // 记录文件系统错误但不抛出（移动赋值操作符中）
                FASTEXCEL_LOG_WARN("Failed to remove temp file '{}' during move assignment: {}", temp_path_.string(), e.what());
            } catch (const std::exception& e) {
                // 记录其他错误
                FASTEXCEL_LOG_WARN("Exception while removing temp file '{}': {}", temp_path_.string(), e.what());
            }
        }
        
        // 转移资源
        temp_path_ = std::move(other.temp_path_);
        file_ = std::move(other.file_);
        should_delete_ = other.should_delete_;
        other.should_delete_ = false;
    }
    return *this;
}

std::string TempFileWrapper::generateTempPath(const std::string& prefix, const std::string& suffix) {
    // 获取系统临时目录
    std::filesystem::path temp_dir = std::filesystem::temp_directory_path();
    
    // 生成随机文件名
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> dis(1000, 9999);
    
    auto now = std::chrono::system_clock::now();
    auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
    
    std::string filename = prefix + std::to_string(timestamp) + "_" + std::to_string(dis(gen)) + suffix;
    
    return (temp_dir / filename).string();
}

} // namespace utils
} // namespace fastexcel