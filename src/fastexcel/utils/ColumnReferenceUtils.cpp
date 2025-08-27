/**
 * @file ColumnReferenceUtils.cpp
 * @brief 列引用解析优化工具实现
 */

#include "ColumnReferenceUtils.hpp"
#include <cctype>
#include <algorithm>

namespace fastexcel {
namespace utils {

// 静态成员定义
std::array<uint32_t, ColumnReferenceUtils::MAX_COLUMNS> ColumnReferenceUtils::column_lookup_{};
bool ColumnReferenceUtils::initialized_ = false;

void ColumnReferenceUtils::initializeLookupTable() {
    if (initialized_) return;
    
    // 清零数组
    column_lookup_.fill(0);
    
    size_t index = 0;
    
    // 单字母列 A-Z (1-26)
    for (char c1 = 'A'; c1 <= 'Z'; ++c1) {
        if (index < MAX_COLUMNS) {
            column_lookup_[index++] = static_cast<uint32_t>(c1 - 'A' + 1);
        }
    }
    
    // 双字母列 AA-ZZ (27-702)
    for (char c1 = 'A'; c1 <= 'Z'; ++c1) {
        for (char c2 = 'A'; c2 <= 'Z'; ++c2) {
            if (index < MAX_COLUMNS) {
                uint32_t col = static_cast<uint32_t>((c1 - 'A' + 1) * 26 + (c2 - 'A' + 1));
                column_lookup_[index++] = col;
            }
        }
    }
    
    // 三字母列 AAA-ZZZ (703-18278)
    for (char c1 = 'A'; c1 <= 'Z'; ++c1) {
        for (char c2 = 'A'; c2 <= 'Z'; ++c2) {
            for (char c3 = 'A'; c3 <= 'Z'; ++c3) {
                if (index < MAX_COLUMNS) {
                    uint32_t col = static_cast<uint32_t>((c1 - 'A' + 1) * 26 * 26 + 
                                                        (c2 - 'A' + 1) * 26 + 
                                                        (c3 - 'A' + 1));
                    column_lookup_[index++] = col;
                }
            }
        }
    }
    
    initialized_ = true;
}

bool ColumnReferenceUtils::isInitialized() {
    return initialized_;
}

uint32_t ColumnReferenceUtils::parseColumnFast(std::string_view cell_ref) {
    if (!initialized_) {
        initializeLookupTable();
    }
    
    if (cell_ref.empty()) return 0;
    
    // 提取列部分（字母）
    size_t col_end = 0;
    while (col_end < cell_ref.size() && std::isalpha(cell_ref[col_end])) {
        ++col_end;
    }
    
    if (col_end == 0) return 0; // 没有字母部分
    
    return parseColumnOnly(cell_ref.substr(0, col_end));
}

uint32_t ColumnReferenceUtils::parseColumnOnly(std::string_view col_ref) {
    if (!initialized_) {
        initializeLookupTable();
    }
    
    if (col_ref.empty() || col_ref.size() > 3) return 0;
    
    // 直接计算（对于小的查找表，直接计算比哈希查找更快）
    uint32_t col = 0;
    for (char c : col_ref) {
        if (!std::isalpha(c)) return 0;
        col = col * 26 + static_cast<uint32_t>(std::toupper(c) - 'A' + 1);
    }
    
    return col;
}

}} // namespace fastexcel::utils