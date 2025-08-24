#include "XMLEscapeSIMD.hpp"

// 这个文件会被Highway多次编译，所以我们需要小心处理重复定义
// 只在第一次编译时定义这些

#if !defined(HWY_TARGET_INCLUDE)
// 这是主编译单元，定义静态成员和非SIMD函数

namespace fastexcel {
namespace xml {

// 静态成员变量定义
bool XMLEscapeSIMD::simd_initialized_ = false;
bool XMLEscapeSIMD::simd_supported_ = false;

void XMLEscapeSIMD::initialize() {
    if (simd_initialized_) return;
    
#ifdef FASTEXCEL_HAS_HIGHWAY
    simd_supported_ = true;
#else
    simd_supported_ = false;
#endif
    
    simd_initialized_ = true;
}

bool XMLEscapeSIMD::isSIMDSupported() {
    if (!simd_initialized_) {
        initialize();
    }
    return simd_supported_;
}

// 标量实现
void XMLEscapeSIMD::escapeAttributesScalar(const char* text, size_t length, WriteCallback writer) {
    size_t last_write_pos = 0;
    
    for (size_t i = 0; i < length; i++) {
        const char* replacement = nullptr;
        size_t replacement_len = 0;
        
        switch (text[i]) {
            case '&':
                replacement = "&amp;";
                replacement_len = 5;
                break;
            case '<':
                replacement = "&lt;";
                replacement_len = 4;
                break;
            case '>':
                replacement = "&gt;";
                replacement_len = 4;
                break;
            case '"':
                replacement = "&quot;";
                replacement_len = 6;
                break;
            case '\'':
                replacement = "&apos;";
                replacement_len = 6;
                break;
            case '\n':
                replacement = "&#xA;";
                replacement_len = 5;
                break;
            default:
                continue;
        }
        
        if (i > last_write_pos) {
            writer(std::string(text + last_write_pos, i - last_write_pos));
        }
        writer(std::string(replacement, replacement_len));
        last_write_pos = i + 1;
    }
    
    if (last_write_pos < length) {
        writer(std::string(text + last_write_pos, length - last_write_pos));
    }
}

void XMLEscapeSIMD::escapeDataScalar(const char* text, size_t length, WriteCallback writer) {
    size_t last_write_pos = 0;
    
    for (size_t i = 0; i < length; i++) {
        const char* replacement = nullptr;
        size_t replacement_len = 0;
        
        switch (text[i]) {
            case '&':
                replacement = "&amp;";
                replacement_len = 5;
                break;
            case '<':
                replacement = "&lt;";
                replacement_len = 4;
                break;
            case '>':
                replacement = "&gt;";
                replacement_len = 4;
                break;
            default:
                continue;
        }
        
        if (i > last_write_pos) {
            writer(std::string(text + last_write_pos, i - last_write_pos));
        }
        writer(std::string(replacement, replacement_len));
        last_write_pos = i + 1;
    }
    
    if (last_write_pos < length) {
        writer(std::string(text + last_write_pos, length - last_write_pos));
    }
}

}} // namespace fastexcel::xml
#endif // !defined(HWY_TARGET_INCLUDE)

#ifdef FASTEXCEL_HAS_HIGHWAY
#include "hwy/highway.h"

namespace fastexcel {
namespace xml {

// 使用Highway的简化SIMD实现，避免复杂的HWY_EXPORT机制
// Highway 大多数 API 位于 hwy::HWY_NAMESPACE 命名空间
namespace hn = hwy::HWY_NAMESPACE;

void XMLEscapeSIMD::escapeAttributesSIMD(const char* text, size_t length, WriteCallback writer) {
    if (!isSIMDSupported()) {
        escapeAttributesScalar(text, length, writer);
        return;
    }
    
    // 使用Highway SIMD优化的属性转义
    const hn::ScalableTag<uint8_t> d;
    const size_t N = hn::Lanes(d);
    
    size_t i = 0;
    size_t last_write_pos = 0;
    
    // SIMD处理主循环
    while (i + N <= length) {
        const auto chunk = hn::Load(d, reinterpret_cast<const uint8_t*>(text + i));
        
        // 检查需要转义的字符
        const auto amp_mask = hn::Eq(chunk, hn::Set(d, uint8_t('&')));
        const auto lt_mask = hn::Eq(chunk, hn::Set(d, uint8_t('<')));
        const auto gt_mask = hn::Eq(chunk, hn::Set(d, uint8_t('>')));
        const auto quot_mask = hn::Eq(chunk, hn::Set(d, uint8_t('"')));
        const auto apos_mask = hn::Eq(chunk, hn::Set(d, uint8_t('\'')));
        const auto nl_mask = hn::Eq(chunk, hn::Set(d, uint8_t('\n')));
        
        auto needs_escape = hn::Or(amp_mask, lt_mask);
        needs_escape = hn::Or(needs_escape, gt_mask);
        needs_escape = hn::Or(needs_escape, quot_mask);
        needs_escape = hn::Or(needs_escape, apos_mask);
        needs_escape = hn::Or(needs_escape, nl_mask);
        
        // 如果有字符需要转义，逐个处理
        if (!hn::AllFalse(d, needs_escape)) {
            for (size_t j = 0; j < N && i + j < length; j++) {
                char c = text[i + j];
                const char* replacement = nullptr;
                size_t replacement_len = 0;
                
                switch (c) {
                    case '&': replacement = "&amp;"; replacement_len = 5; break;
                    case '<': replacement = "&lt;"; replacement_len = 4; break;
                    case '>': replacement = "&gt;"; replacement_len = 4; break;
                    case '"': replacement = "&quot;"; replacement_len = 6; break;
                    case '\'': replacement = "&apos;"; replacement_len = 6; break;
                    case '\n': replacement = "&#xA;"; replacement_len = 5; break;
                    default: continue;
                }
                
                if (i + j > last_write_pos) {
                    writer(std::string(text + last_write_pos, (i + j) - last_write_pos));
                }
                writer(std::string(replacement, replacement_len));
                last_write_pos = i + j + 1;
            }
        }
        i += N;
    }
    
    // 处理剩余的字符
    for (; i < length; i++) {
        char c = text[i];
        const char* replacement = nullptr;
        size_t replacement_len = 0;
        
        switch (c) {
            case '&': replacement = "&amp;"; replacement_len = 5; break;
            case '<': replacement = "&lt;"; replacement_len = 4; break;
            case '>': replacement = "&gt;"; replacement_len = 4; break;
            case '"': replacement = "&quot;"; replacement_len = 6; break;
            case '\'': replacement = "&apos;"; replacement_len = 6; break;
            case '\n': replacement = "&#xA;"; replacement_len = 5; break;
            default: continue;
        }
        
        if (i > last_write_pos) {
            writer(std::string(text + last_write_pos, i - last_write_pos));
        }
        writer(std::string(replacement, replacement_len));
        last_write_pos = i + 1;
    }
    
    if (last_write_pos < length) {
        writer(std::string(text + last_write_pos, length - last_write_pos));
    }
}

void XMLEscapeSIMD::escapeDataSIMD(const char* text, size_t length, WriteCallback writer) {
    if (!isSIMDSupported()) {
        escapeDataScalar(text, length, writer);
        return;
    }
    
    // 使用Highway SIMD优化的数据转义
    const hn::ScalableTag<uint8_t> d;
    const size_t N = hn::Lanes(d);
    
    size_t i = 0;
    size_t last_write_pos = 0;
    
    // SIMD处理主循环
    while (i + N <= length) {
        const auto chunk = hn::Load(d, reinterpret_cast<const uint8_t*>(text + i));
        
        // 检查需要转义的字符（数据转义只需要 &, <, >）
        const auto amp_mask = hn::Eq(chunk, hn::Set(d, uint8_t('&')));
        const auto lt_mask = hn::Eq(chunk, hn::Set(d, uint8_t('<')));
        const auto gt_mask = hn::Eq(chunk, hn::Set(d, uint8_t('>')));
        
        auto needs_escape = hn::Or(amp_mask, lt_mask);
        needs_escape = hn::Or(needs_escape, gt_mask);
        
        // 如果有字符需要转义，逐个处理
        if (!hn::AllFalse(d, needs_escape)) {
            for (size_t j = 0; j < N && i + j < length; j++) {
                char c = text[i + j];
                const char* replacement = nullptr;
                size_t replacement_len = 0;
                
                switch (c) {
                    case '&': replacement = "&amp;"; replacement_len = 5; break;
                    case '<': replacement = "&lt;"; replacement_len = 4; break;
                    case '>': replacement = "&gt;"; replacement_len = 4; break;
                    default: continue;
                }
                
                if (i + j > last_write_pos) {
                    writer(std::string(text + last_write_pos, (i + j) - last_write_pos));
                }
                writer(std::string(replacement, replacement_len));
                last_write_pos = i + j + 1;
            }
        }
        i += N;
    }
    
    // 处理剩余的字符
    for (; i < length; i++) {
        char c = text[i];
        const char* replacement = nullptr;
        size_t replacement_len = 0;
        
        switch (c) {
            case '&': replacement = "&amp;"; replacement_len = 5; break;
            case '<': replacement = "&lt;"; replacement_len = 4; break;
            case '>': replacement = "&gt;"; replacement_len = 4; break;
            default: continue;
        }
        
        if (i > last_write_pos) {
            writer(std::string(text + last_write_pos, i - last_write_pos));
        }
        writer(std::string(replacement, replacement_len));
        last_write_pos = i + 1;
    }
    
    if (last_write_pos < length) {
        writer(std::string(text + last_write_pos, length - last_write_pos));
    }
}

}} // namespace fastexcel::xml

#else
// FASTEXCEL_HAS_HIGHWAY 未定义时的回退实现

namespace fastexcel {
namespace xml {

void XMLEscapeSIMD::escapeAttributesSIMD(const char* text, size_t length, WriteCallback writer) {
    escapeAttributesScalar(text, length, writer);
}

void XMLEscapeSIMD::escapeDataSIMD(const char* text, size_t length, WriteCallback writer) {
    escapeDataScalar(text, length, writer);
}

}} // namespace fastexcel::xml

#endif // FASTEXCEL_HAS_HIGHWAY
