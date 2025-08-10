#pragma once

#include <string>
#include <vector>
#include <functional>
#include <type_traits>
#include <algorithm>
#include <chrono>
#include <limits>
#include <cctype>
#include <stdexcept>
#include "fastexcel/core/Exception.hpp"

namespace fastexcel {
namespace utils {

/**
 * @brief 通用工具类 - 提供常用的辅助函数
 *
 * 注意：XML生成功能请使用 fastexcel::xml::XMLStreamWriter，
 * 它提供了更高效的流式XML写入能力。
 */
class CommonUtils {
public:
    // ========== 字符串工具 ==========
    
    /**
     * @brief 列号转换为字母表示（A, B, ..., Z, AA, AB, ...）
     * @param col 列号（0开始）
     * @return 字母表示
     */
    static std::string columnToLetter(int col) {
        std::string result;
        while (col >= 0) {
            result = static_cast<char>('A' + (col % 26)) + result;
            col = col / 26 - 1;
        }
        return result;
    }
    
    /**
     * @brief 生成单元格引用（如A1, B2等）
     * @param row 行号（0开始）
     * @param col 列号（0开始）
     * @return 单元格引用
     */
    static std::string cellReference(int row, int col) {
        return columnToLetter(col) + std::to_string(row + 1);
    }
    
    /**
     * @brief 生成范围引用（如A1:B2）
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @return 范围引用
     */
    static std::string rangeReference(int first_row, int first_col, int last_row, int last_col) {
        return cellReference(first_row, first_col) + ":" + cellReference(last_row, last_col);
    }
    
    /**
     * @brief 解析单元格引用（如A1 -> (0, 0)）
     * @param reference 单元格引用字符串
     * @return std::pair<int, int> 行列坐标（0开始）
     * @throws std::invalid_argument 如果引用格式不正确
     */
    static std::pair<int, int> parseReference(const std::string& reference) {
        if (reference.empty()) {
            throw std::invalid_argument("Empty cell reference");
        }
        
        size_t i = 0;
        
        // 解析列部分 (A-Z)
        int col = 0;
        while (i < reference.length() && std::isalpha(reference[i])) {
            char c = std::toupper(reference[i]);
            if (c < 'A' || c > 'Z') {
                throw std::invalid_argument("Invalid column letter in reference: " + reference);
            }
            col = col * 26 + (c - 'A' + 1);
            ++i;
        }
        
        if (col == 0) {
            throw std::invalid_argument("No column part in reference: " + reference);
        }
        col -= 1; // 转换为0开始的索引
        
        // 解析行部分 (1-9)
        if (i >= reference.length() || !std::isdigit(reference[i])) {
            throw std::invalid_argument("No row part in reference: " + reference);
        }
        
        int row = 0;
        while (i < reference.length() && std::isdigit(reference[i])) {
            row = row * 10 + (reference[i] - '0');
            ++i;
        }
        
        if (row == 0) {
            throw std::invalid_argument("Invalid row number in reference: " + reference);
        }
        row -= 1; // 转换为0开始的索引
        
        // 检查是否有剩余字符
        if (i < reference.length()) {
            throw std::invalid_argument("Invalid characters at end of reference: " + reference);
        }
        
        return std::make_pair(row, col);
    }
    
    // ========== 验证工具 ==========
    
    /**
     * @brief 验证单元格位置是否有效
     * @param row 行号
     * @param col 列号
     * @return 是否有效
     */
    static bool isValidCellPosition(int row, int col) {
        return row >= 0 && col >= 0 && row <= 1048575 && col <= 16383; // Excel 2007+ limits
    }
    
    /**
     * @brief 验证范围是否有效
     * @param first_row 起始行
     * @param first_col 起始列
     * @param last_row 结束行
     * @param last_col 结束列
     * @return 是否有效
     */
    static bool isValidRange(int first_row, int first_col, int last_row, int last_col) {
        return isValidCellPosition(first_row, first_col) &&
               isValidCellPosition(last_row, last_col) &&
               first_row <= last_row && first_col <= last_col;
    }
    
    /**
     * @brief 验证工作表名称是否有效
     * @param name 工作表名称
     * @return 是否有效
     */
    static bool isValidSheetName(const std::string& name) {
        if (name.empty() || name.length() > 31) {
            return false;
        }
        
        // 检查非法字符
        const std::string invalid_chars = "[]*/\\?:";
        return name.find_first_of(invalid_chars) == std::string::npos;
    }
    
    // ========== 模板工具 ==========
    
    /**
     * @brief 安全的数值转换
     * @tparam To 目标类型
     * @tparam From 源类型
     * @param value 源值
     * @param default_value 默认值（转换失败时使用）
     * @return 转换后的值
     */
    template<typename To, typename From>
    static To safeCast(const From& value, const To& default_value = To{}) {
        if constexpr (std::is_arithmetic_v<From> && std::is_arithmetic_v<To>) {
            if (value >= std::numeric_limits<To>::min() &&
                value <= std::numeric_limits<To>::max()) {
                return static_cast<To>(value);
            }
        }
        return default_value;
    }
    
    /**
     * @brief 条件执行函数
     * @tparam Func 函数类型
     * @param condition 条件
     * @param func 要执行的函数
     */
    template<typename Func>
    static void executeIf(bool condition, Func&& func) {
        if (condition) {
            std::forward<Func>(func)();
        }
    }
    
    /**
     * @brief 批量操作辅助函数
     * @tparam Container 容器类型
     * @tparam Func 函数类型
     * @param container 容器
     * @param func 对每个元素执行的函数
     */
    template<typename Container, typename Func>
    static void forEach(Container&& container, Func&& func) {
        std::for_each(std::begin(container), std::end(container), std::forward<Func>(func));
    }
    
    // ========== 内存和性能工具 ==========
    
    /**
     * @brief 计算容器的内存使用量
     * @tparam Container 容器类型
     * @param container 容器
     * @return 内存使用量（字节）
     */
    template<typename Container>
    static size_t calculateMemoryUsage(const Container& container) {
        size_t usage = sizeof(Container);
        if constexpr (std::is_same_v<Container, std::string>) {
            usage += container.capacity();
        } else if constexpr (std::is_same_v<Container, std::vector<typename Container::value_type>>) {
            usage += container.capacity() * sizeof(typename Container::value_type);
        }
        return usage;
    }
    
    /**
     * @brief RAII计时器
     */
    class ScopedTimer {
    private:
        std::chrono::high_resolution_clock::time_point start_;
        std::function<void(double)> callback_;
        
    public:
        explicit ScopedTimer(std::function<void(double)> callback)
            : start_(std::chrono::high_resolution_clock::now()), callback_(callback) {}
        
        ~ScopedTimer() {
            auto end = std::chrono::high_resolution_clock::now();
            auto duration = std::chrono::duration<double, std::milli>(end - start_);
            if (callback_) {
                callback_(duration.count());
            }
        }
    };
};

/**
 * @brief 错误处理辅助宏
 */
#define FASTEXCEL_VALIDATE_CELL_POSITION(row, col) \
    do { \
        if (!fastexcel::utils::CommonUtils::isValidCellPosition(row, col)) { \
            FASTEXCEL_THROW_CELL("Invalid cell position: (" + std::to_string(row) + ", " + std::to_string(col) + ")", row, col); \
        } \
    } while(0)

#define FASTEXCEL_VALIDATE_RANGE(first_row, first_col, last_row, last_col) \
    do { \
        if (!fastexcel::utils::CommonUtils::isValidRange(first_row, first_col, last_row, last_col)) { \
            FASTEXCEL_THROW_PARAM("Invalid range"); \
        } \
    } while(0)

}} // namespace fastexcel::utils
