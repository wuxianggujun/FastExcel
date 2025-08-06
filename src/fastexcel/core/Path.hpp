#pragma once

#include <string>
#include <filesystem>

#ifdef _WIN32
#include <utf8.h>
#endif

namespace fastexcel {
namespace core {

/**
 * @brief UTF-8路径处理类，封装跨平台文件路径操作
 * 
 * 这个类使用utf8cpp库来处理Unicode文件路径，提供统一的接口
 * 在Windows下自动处理UTF-8到UTF-16的转换，在其他平台直接使用UTF-8
 */
class Path {
private:
    std::string utf8_path_;
    

public:
    /**
     * @brief 构造函数
     * @param path UTF-8编码的路径字符串
     */
    explicit Path(const std::string& path);
    
    /**
     * @brief 构造函数
     * @param path C风格字符串路径
     */
    explicit Path(const char* path);
    
    /**
     * @brief 默认构造函数
     */
    Path() = default;
    
    /**
     * @brief 复制构造函数
     */
    Path(const Path& other) = default;
    
    /**
     * @brief 移动构造函数
     */
    Path(Path&& other) noexcept = default;
    
    /**
     * @brief 赋值操作符
     */
    Path& operator=(const Path& other) = default;
    
    /**
     * @brief 移动赋值操作符
     */
    Path& operator=(Path&& other) noexcept = default;
    
    /**
     * @brief 析构函数
     */
    ~Path() = default;
    
    // 路径操作
    /**
     * @brief 获取UTF-8编码的路径字符串
     * @return UTF-8路径字符串
     */
    const std::string& string() const { return utf8_path_; }
    
    /**
     * @brief 获取C风格字符串
     * @return C风格路径字符串
     */
    const char* c_str() const { return utf8_path_.c_str(); }
    
    /**
     * @brief 检查路径是否为空
     * @return 是否为空
     */
    bool empty() const { return utf8_path_.empty(); }
    
    /**
     * @brief 清空路径
     */
    void clear() { utf8_path_.clear(); }
    
    // 文件操作
    /**
     * @brief 检查文件是否存在
     * @return 文件是否存在
     */
    bool exists() const;
    
    /**
     * @brief 检查是否为文件
     * @return 是否为文件
     */
    bool isFile() const;
    
    /**
     * @brief 检查是否为目录
     * @return 是否为目录
     */
    bool isDirectory() const;
    
    /**
     * @brief 获取文件大小
     * @return 文件大小（字节），失败返回0
     */
    uintmax_t fileSize() const;
    
    /**
     * @brief 删除文件
     * @return 是否删除成功
     */
    bool remove() const;
    
    /**
     * @brief 复制文件到目标路径
     * @param target 目标路径
     * @param overwrite 是否覆盖现有文件
     * @return 是否复制成功
     */
    bool copyTo(const Path& target, bool overwrite = true) const;
    
    /**
     * @brief 移动文件到目标路径
     * @param target 目标路径
     * @return 是否移动成功
     */
    bool moveTo(const Path& target) const;
    
    // 文件流操作
    /**
     * @brief 打开文件进行读取
     * @param binary 是否以二进制模式打开
     * @return 文件句柄，失败返回nullptr
     */
    FILE* openForRead(bool binary = true) const;
    
    /**
     * @brief 打开文件进行写入
     * @param binary 是否以二进制模式打开
     * @return 文件句柄，失败返回nullptr
     */
    FILE* openForWrite(bool binary = true) const;
    
#ifdef _WIN32
    /**
     * @brief 获取Windows宽字符路径
     * @return 宽字符路径字符串
     */
    std::wstring getWidePath() const;
#endif
    
    // 比较操作符
    bool operator==(const Path& other) const { return utf8_path_ == other.utf8_path_; }
    bool operator!=(const Path& other) const { return utf8_path_ != other.utf8_path_; }
    bool operator<(const Path& other) const { return utf8_path_ < other.utf8_path_; }
    
    // 字符串转换操作符
    operator const std::string&() const { return utf8_path_; }
    
    // 流输出支持
    friend std::ostream& operator<<(std::ostream& os, const Path& path) {
        return os << path.utf8_path_;
    }
};

} // namespace core
} // namespace fastexcel