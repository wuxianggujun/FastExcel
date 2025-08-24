/**
 * @file FileWrapper.hpp
 * @brief RAII文件句柄包装器，提供异常安全的文件管理
 */

#pragma once

#include <memory>
#include <cstdio>
#include <string>
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Path.hpp"

namespace fastexcel {
namespace utils {

/**
 * @brief RAII文件句柄包装器
 * 
 * 提供异常安全的文件资源管理，自动关闭文件句柄
 */
class FileWrapper {
public:
    /**
     * @brief 构造函数，打开文件
     * @param filename 文件名
     * @param mode 文件打开模式
     * @throws FileException 文件打开失败时
     */
    FileWrapper(const std::string& filename, const char* mode) {
        core::Path path(filename);
        FILE* raw_file = nullptr;
        if (std::strcmp(mode, "w") == 0 || std::strcmp(mode, "wb") == 0) {
            raw_file = path.openForWrite(true);
        } else {
            raw_file = path.openForRead();
        }
        
        if (raw_file) {
            file_.reset(raw_file);
            file_.get_deleter() = &fclose;
        }
        
        if (!file_) {
            throw core::FileException(
                "Failed to open file: " + filename, 
                filename,
                core::ErrorCode::FileNotFound,
                __FILE__, __LINE__
            );
        }
    }
    
    /**
     * @brief 从已有FILE*构造，接管所有权
     * @param file 文件指针
     * @param take_ownership 是否接管所有权
     */
    explicit FileWrapper(FILE* file, bool take_ownership = true) {
        if (take_ownership) {
            file_.reset(file);
            file_.get_deleter() = &fclose;
        } else {
            file_.reset(file);
            file_.get_deleter() = +[](FILE*) -> int { return 0; }; // No-op deleter returning int
        }
    }
    
    /**
     * @brief 移动构造函数
     */
    FileWrapper(FileWrapper&& other) noexcept = default;
    
    /**
     * @brief 移动赋值操作符
     */
    FileWrapper& operator=(FileWrapper&& other) noexcept = default;
    
    // 禁用拷贝构造和拷贝赋值
    FileWrapper(const FileWrapper&) = delete;
    FileWrapper& operator=(const FileWrapper&) = delete;
    
    /**
     * @brief 获取文件指针
     */
    FILE* get() const noexcept { 
        return file_.get(); 
    }
    
    /**
     * @brief 检查文件是否有效
     */
    operator bool() const noexcept { 
        return file_ != nullptr; 
    }
    
    /**
     * @brief 刷新文件缓冲区
     */
    void flush() {
        if (file_) {
            std::fflush(file_.get());
        }
    }
    
    /**
     * @brief 释放文件句柄的所有权
     * @return 文件指针
     */
    FILE* release() noexcept {
        return file_.release();
    }
    
    /**
     * @brief 重置文件句柄
     */
    void reset(FILE* file = nullptr, bool take_ownership = true) {
        if (take_ownership) {
            file_.reset(file);
            if (file) {
                file_.get_deleter() = &fclose;
            }
        } else {
            file_.reset(file);
            file_.get_deleter() = +[](FILE*) -> int { return 0; }; // No-op deleter returning int
        }
    }

private:
    std::unique_ptr<FILE, int(*)(FILE*)> file_{nullptr, nullptr};
};

/**
 * @brief 临时文件包装器
 * 
 * 在析构时自动删除临时文件
 */
class TempFileWrapper {
public:
    /**
     * @brief 构造函数，创建临时文件
     * @param prefix 文件名前缀
     * @param suffix 文件名后缀
     */
    TempFileWrapper(const std::string& prefix = "fastexcel_temp_", 
                   const std::string& suffix = ".tmp");
    
    /**
     * @brief 析构函数，自动删除临时文件
     */
    ~TempFileWrapper();
    
    // 禁用拷贝
    TempFileWrapper(const TempFileWrapper&) = delete;
    TempFileWrapper& operator=(const TempFileWrapper&) = delete;
    
    // 允许移动
    TempFileWrapper(TempFileWrapper&& other) noexcept;
    TempFileWrapper& operator=(TempFileWrapper&& other) noexcept;
    
    /**
     * @brief 获取文件包装器
     */
    FileWrapper& getFile() { return file_; }
    const FileWrapper& getFile() const { return file_; }
    
    /**
     * @brief 获取文件路径
     */
    const std::string& getPath() const { return temp_path_; }

private:
    std::string temp_path_;
    FileWrapper file_;
    bool should_delete_ = true;
    
    std::string generateTempPath(const std::string& prefix, const std::string& suffix);
};

} // namespace utils
} // namespace fastexcel