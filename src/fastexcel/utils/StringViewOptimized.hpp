/**
 * @file StringViewOptimized.hpp
 * @brief string_view优化工具，减少字符串分配开销
 */

#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <memory>
#include <cstring>
#include <cstdarg>
#include <unordered_set>
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/utils/Logger.hpp"

namespace fastexcel {
namespace utils {

/**
 * @brief 优化的字符串视图工具类
 * 
 * 提供高效的字符串操作，减少不必要的内存分配
 */
class StringViewOptimized {
public:
    /**
     * @brief 字符串连接器（避免多次内存分配）
     */
    class StringJoiner {
    public:
        explicit StringJoiner(const std::string& separator = "")
            : separator_(separator) {
            parts_.reserve(16); // 预分配常见大小
        }
        
        /**
         * @brief 添加字符串片段
         */
        StringJoiner& add(std::string_view part) {
            if (!part.empty()) {
                parts_.push_back(part);
            }
            return *this;
        }
        
        StringJoiner& add(const char* part) {
            if (part && *part) {
                parts_.emplace_back(part);
            }
            return *this;
        }
        
        StringJoiner& add(const std::string& part) {
            if (!part.empty()) {
                parts_.push_back(part);
            }
            return *this;
        }
        
        /**
         * @brief 构建最终字符串（一次性分配内存）
         */
        std::string build() const {
            if (parts_.empty()) {
                return {};
            }
            
            if (parts_.size() == 1) {
                return std::string(parts_[0]);
            }
            
            // 计算总大小
            size_t total_size = 0;
            for (const auto& part : parts_) {
                total_size += part.size();
            }
            
            // 添加分隔符的大小
            if (!separator_.empty() && parts_.size() > 1) {
                total_size += separator_.size() * (parts_.size() - 1);
            }
            
            // 一次性分配内存并构建
            std::string result;
            result.reserve(total_size);
            
            bool first = true;
            for (const auto& part : parts_) {
                if (!first && !separator_.empty()) {
                    result.append(separator_);
                }
                result.append(part);
                first = false;
            }
            
            return result;
        }
        
        /**
         * @brief 获取当前部分数量
         */
        size_t size() const noexcept {
            return parts_.size();
        }
        
        /**
         * @brief 清空所有部分
         */
        void clear() {
            parts_.clear();
        }
        
        /**
         * @brief 预分配容量
         */
        void reserve(size_t capacity) {
            parts_.reserve(capacity);
        }
        
    private:
        std::vector<std::string_view> parts_;
        std::string separator_;
    };
    
    /**
     * @brief 字符串构建器（类似StringBuilder）
     */
    class StringBuilder {
    public:
        explicit StringBuilder(size_t initial_capacity = 256) {
            buffer_.reserve(initial_capacity);
        }
        
        StringBuilder& append(std::string_view str) {
            buffer_.append(str);
            return *this;
        }
        
        StringBuilder& append(const char* str) {
            if (str) {
                buffer_.append(str);
            }
            return *this;
        }
        
        StringBuilder& append(char c) {
            buffer_.push_back(c);
            return *this;
        }
        
        StringBuilder& append(int value) {
            buffer_.append(std::to_string(value));
            return *this;
        }
        
        StringBuilder& append(double value) {
            buffer_.append(std::to_string(value));
            return *this;
        }
        
        StringBuilder& appendFormat(const char* format, ...) {
            va_list args;
            va_start(args, format);
            
            // 计算需要的大小
            int size = std::vsnprintf(nullptr, 0, format, args);
            va_end(args);
            
            if (size > 0) {
                size_t old_size = buffer_.size();
                buffer_.resize(old_size + size);
                
                va_start(args, format);
                std::vsnprintf(buffer_.data() + old_size, size + 1, format, args);
                va_end(args);
            }
            
            return *this;
        }
        
        std::string build() {
            return std::move(buffer_);
        }
        
        std::string_view view() const noexcept {
            return buffer_;
        }
        
        void clear() {
            buffer_.clear();
        }
        
        void reserve(size_t capacity) {
            buffer_.reserve(capacity);
        }
        
        size_t size() const noexcept {
            return buffer_.size();
        }
        
        bool empty() const noexcept {
            return buffer_.empty();
        }
        
    private:
        std::string buffer_;
    };

public:
    /**
     * @brief 高效的字符串分割（避免创建临时字符串）
     */
    static std::vector<std::string_view> split(std::string_view str, char delimiter) {
        std::vector<std::string_view> result;
        result.reserve(8); // 预分配常见大小
        
        size_t start = 0;
        size_t end = 0;
        
        while (end < str.length()) {
            if (str[end] == delimiter) {
                if (end > start) {
                    result.push_back(str.substr(start, end - start));
                }
                start = end + 1;
            }
            ++end;
        }
        
        // 添加最后一部分
        if (start < str.length()) {
            result.push_back(str.substr(start));
        }
        
        return result;
    }
    
    /**
     * @brief 分割字符串（多字符分隔符）
     */
    static std::vector<std::string_view> split(std::string_view str, std::string_view delimiter) {
        std::vector<std::string_view> result;
        
        if (delimiter.empty()) {
            result.push_back(str);
            return result;
        }
        
        size_t start = 0;
        size_t pos = str.find(delimiter);
        
        while (pos != std::string_view::npos) {
            result.push_back(str.substr(start, pos - start));
            start = pos + delimiter.length();
            pos = str.find(delimiter, start);
        }
        
        // 添加最后一部分
        result.push_back(str.substr(start));
        
        return result;
    }
    
    /**
     * @brief 去除前后空白字符
     */
    static std::string_view trim(std::string_view str) {
        size_t start = 0;
        size_t end = str.length();
        
        // 去除前导空白
        while (start < end && std::isspace(str[start])) {
            ++start;
        }
        
        // 去除尾随空白
        while (end > start && std::isspace(str[end - 1])) {
            --end;
        }
        
        return str.substr(start, end - start);
    }
    
    /**
     * @brief 去除前导空白
     */
    static std::string_view trimLeft(std::string_view str) {
        size_t start = 0;
        while (start < str.length() && std::isspace(str[start])) {
            ++start;
        }
        return str.substr(start);
    }
    
    /**
     * @brief 去除尾随空白
     */
    static std::string_view trimRight(std::string_view str) {
        size_t end = str.length();
        while (end > 0 && std::isspace(str[end - 1])) {
            --end;
        }
        return str.substr(0, end);
    }
    
    /**
     * @brief 不区分大小写的比较
     */
    static bool equalsIgnoreCase(std::string_view a, std::string_view b) {
        if (a.size() != b.size()) {
            return false;
        }
        
        for (size_t i = 0; i < a.size(); ++i) {
            if (std::tolower(a[i]) != std::tolower(b[i])) {
                return false;
            }
        }
        
        return true;
    }
    
    /**
     * @brief 检查字符串是否以指定前缀开始
     */
    static bool startsWith(std::string_view str, std::string_view prefix) {
        if (prefix.size() > str.size()) {
            return false;
        }
        return str.substr(0, prefix.size()) == prefix;
    }
    
    /**
     * @brief 检查字符串是否以指定后缀结束
     */
    static bool endsWith(std::string_view str, std::string_view suffix) {
        if (suffix.size() > str.size()) {
            return false;
        }
        return str.substr(str.size() - suffix.size()) == suffix;
    }
    
    /**
     * @brief 查找并替换（返回新字符串，但尽可能减少分配）
     */
    static std::string replace(std::string_view str, std::string_view from, std::string_view to) {
        if (from.empty()) {
            return std::string(str);
        }
        
        std::string result;
        result.reserve(str.size() + (to.size() - from.size()) * 4); // 估算大小
        
        size_t start = 0;
        size_t pos = str.find(from);
        
        while (pos != std::string_view::npos) {
            result.append(str.substr(start, pos - start));
            result.append(to);
            start = pos + from.size();
            pos = str.find(from, start);
        }
        
        result.append(str.substr(start));
        return result;
    }
    
    /**
     * @brief 高效的字符串格式化
     */
    template<typename... Args>
    static std::string format(const char* format, Args&&... args) {
        // 首先计算需要的大小
        int size = std::snprintf(nullptr, 0, format, args...);
        if (size <= 0) {
            return {};
        }
        
        // 一次性分配并格式化
        std::string result(size, '\0');
        std::snprintf(result.data(), size + 1, format, args...);
        
        return result;
    }
    
    /**
     * @brief 检查字符串是否包含指定子串
     */
    static bool contains(std::string_view str, std::string_view substr) {
        return str.find(substr) != std::string_view::npos;
    }
    
    /**
     * @brief 计算字符串中某字符的出现次数
     */
    static size_t count(std::string_view str, char c) {
        size_t count = 0;
        for (char ch : str) {
            if (ch == c) {
                ++count;
            }
        }
        return count;
    }
    
    /**
     * @brief 安全的字符串转数字
     */
    static bool tryParseInt(std::string_view str, int& result) {
        str = trim(str);
        if (str.empty()) {
            return false;
        }
        
        try {
            size_t pos;
            result = std::stoi(std::string(str), &pos);
            return pos == str.size();
        } catch (...) {
            return false;
        }
    }
    
    static bool tryParseDouble(std::string_view str, double& result) {
        str = trim(str);
        if (str.empty()) {
            return false;
        }
        
        try {
            size_t pos;
            result = std::stod(std::string(str), &pos);
            return pos == str.size();
        } catch (...) {
            return false;
        }
    }
};

/**
 * @brief 字符串池，用于避免重复字符串的内存分配
 */
class StringPool {
public:
    /**
     * @brief 获取或创建字符串
     * @param str 输入字符串视图
     * @return 指向池中字符串的指针
     */
    const std::string* intern(std::string_view str) {
        // Convert string_view to string for the find operation
        std::string str_copy(str);
        auto it = pool_.find(str_copy);
        if (it != pool_.end()) {
            return &(*it);
        }
        
        // 插入新字符串
        auto [inserted_it, success] = pool_.insert(std::move(str_copy));
        return &(*inserted_it);
    }
    
    /**
     * @brief 清空字符串池
     */
    void clear() {
        pool_.clear();
    }
    
    /**
     * @brief 获取池大小
     */
    size_t size() const noexcept {
        return pool_.size();
    }
    
private:
    std::unordered_set<std::string> pool_;
};

} // namespace utils
} // namespace fastexcel

// 便捷的字符串操作宏
#define FASTEXCEL_STRING_JOIN(...) \
    fastexcel::utils::StringViewOptimized::StringJoiner().add(__VA_ARGS__).build()

#define FASTEXCEL_STRING_BUILD(capacity) \
    fastexcel::utils::StringViewOptimized::StringBuilder(capacity)

#define FASTEXCEL_FORMAT(format, ...) \
    fastexcel::utils::StringViewOptimized::format(format, __VA_ARGS__)