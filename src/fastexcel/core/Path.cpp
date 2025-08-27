#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <fstream>
#include <cstdio>

#ifdef _WIN32
#include <windows.h>
#include <utf8.h>
#else
#include <filesystem>
#endif

namespace fastexcel {
namespace core {

// 构造函数实现
Path::Path(const std::string& path) : utf8_path_(path) {
#ifdef _WIN32
    // 在Windows下验证UTF-8字符串的有效性
    if (!utf8::is_valid(path.begin(), path.end())) {
        // 如果不是有效的UTF-8，假设它是本地编码，尝试转换
        // 这里可以添加更复杂的编码检测和转换逻辑
        utf8_path_ = path; // 暂时直接使用
    }
#endif
}

Path::Path(const char* path) : Path(std::string(path)) {}

#ifdef _WIN32
std::wstring Path::getWidePath() const {
    if (utf8_path_.empty()) return std::wstring();
    
    try {
        // 使用utf8cpp库进行转换
        std::wstring result;
        utf8::utf8to16(utf8_path_.begin(), utf8_path_.end(), std::back_inserter(result));
        return result;
    } catch (const utf8::exception& e) {
        // UTF-8转换失败，可能是编码问题
        // 作为后备方案，使用Windows API
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, utf8_path_.c_str(), -1, NULL, 0);
        if (size_needed == 0) return std::wstring();
        
        std::wstring result(size_needed - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, utf8_path_.c_str(), -1, &result[0], size_needed);
        return result;
    }
}
#endif

// 文件操作实现
bool Path::exists() const {
    if (utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    DWORD attributes = GetFileAttributesW(wide_path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES;
#else
    return std::filesystem::exists(utf8_path_);
#endif
}

bool Path::isFile() const {
    if (utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    DWORD attributes = GetFileAttributesW(wide_path.c_str());
    return (attributes != INVALID_FILE_ATTRIBUTES && !(attributes & FILE_ATTRIBUTE_DIRECTORY));
#else
    return std::filesystem::is_regular_file(utf8_path_);
#endif
}

bool Path::isDirectory() const {
    if (utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    DWORD attributes = GetFileAttributesW(wide_path.c_str());
    return (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY));
#else
    return std::filesystem::is_directory(utf8_path_);
#endif
}

uintmax_t Path::fileSize() const {
    if (utf8_path_.empty()) return 0;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    WIN32_FILE_ATTRIBUTE_DATA file_data;
    if (GetFileAttributesExW(wide_path.c_str(), GetFileExInfoStandard, &file_data)) {
        ULARGE_INTEGER size;
        size.HighPart = file_data.nFileSizeHigh;
        size.LowPart = file_data.nFileSizeLow;
        return size.QuadPart;
    }
    return 0;
#else
    try {
        return std::filesystem::file_size(utf8_path_);
    } catch (const std::filesystem::filesystem_error& e) {
        FASTEXCEL_LOG_DEBUG("Filesystem error getting file size '{}': {}", utf8_path_, e.what());
        return 0;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("Exception getting file size '{}': {}", utf8_path_, e.what());
        return 0;
    }
#endif
}

bool Path::remove() const {
    if (utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    return DeleteFileW(wide_path.c_str()) != 0;
#else
    try {
        return std::filesystem::remove(utf8_path_);
    } catch (const std::filesystem::filesystem_error& e) {
        FASTEXCEL_LOG_DEBUG("Filesystem error removing file '{}': {}", utf8_path_, e.what());
        return false;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("Exception removing file '{}': {}", utf8_path_, e.what());
        return false;
    }
#endif
}

bool Path::copyTo(const Path& target, bool overwrite) const {
    if (utf8_path_.empty() || target.utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring src_wide = getWidePath();
    std::wstring dst_wide = target.getWidePath();
    return CopyFileW(src_wide.c_str(), dst_wide.c_str(), overwrite ? FALSE : TRUE) != 0;
#else
    try {
        auto copy_options = overwrite ? 
            std::filesystem::copy_options::overwrite_existing : 
            std::filesystem::copy_options::none;
        std::filesystem::copy_file(utf8_path_, target.utf8_path_, copy_options);
        return true;
    } catch (const std::filesystem::filesystem_error& e) {
        FASTEXCEL_LOG_DEBUG("Filesystem error copying file '{}' to '{}': {}", 
                           utf8_path_, target.utf8_path_, e.what());
        return false;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("Exception copying file '{}' to '{}': {}", 
                           utf8_path_, target.utf8_path_, e.what());
        return false;
    }
#endif
}

bool Path::moveTo(const Path& target) const {
    if (utf8_path_.empty() || target.utf8_path_.empty()) return false;
    
#ifdef _WIN32
    std::wstring src_wide = getWidePath();
    std::wstring dst_wide = target.getWidePath();
    return MoveFileW(src_wide.c_str(), dst_wide.c_str()) != 0;
#else
    try {
        std::filesystem::rename(utf8_path_, target.utf8_path_);
        return true;
    } catch (const std::filesystem::filesystem_error& e) {
        FASTEXCEL_LOG_DEBUG("Filesystem error moving file '{}' to '{}': {}", 
                           utf8_path_, target.utf8_path_, e.what());
        return false;
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_DEBUG("Exception moving file '{}' to '{}': {}", 
                           utf8_path_, target.utf8_path_, e.what());
        return false;
    }
#endif
}

FILE* Path::openForRead(bool binary) const {
    if (utf8_path_.empty()) return nullptr;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    std::wstring mode = binary ? L"rb" : L"r";
    FILE* file = nullptr;
    errno_t err = _wfopen_s(&file, wide_path.c_str(), mode.c_str());
    return (err == 0) ? file : nullptr;
#else
    const char* mode = binary ? "rb" : "r";
    return fopen(utf8_path_.c_str(), mode);
#endif
}

FILE* Path::openForWrite(bool binary) const {
    if (utf8_path_.empty()) return nullptr;
    
#ifdef _WIN32
    std::wstring wide_path = getWidePath();
    std::wstring mode = binary ? L"wb" : L"w";
    FILE* file = nullptr;
    errno_t err = _wfopen_s(&file, wide_path.c_str(), mode.c_str());
    return (err == 0) ? file : nullptr;
#else
    const char* mode = binary ? "wb" : "w";
    return fopen(utf8_path_.c_str(), mode);
#endif
}

} // namespace core
} // namespace fastexcel
