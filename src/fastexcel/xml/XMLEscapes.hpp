#pragma once

#include <cstddef>

namespace fastexcel {
namespace xml {

// XML 转义常量（非工具类）：集中提供实体字面量及长度
struct XMLEscapes {
    inline static constexpr char AMP[]  = "&amp;";   // &  → &amp;
    inline static constexpr char LT[]   = "&lt;";    // <  → &lt;
    inline static constexpr char GT[]   = "&gt;";    // >  → &gt;
    inline static constexpr char QUOT[] = "&quot;";  // " → &quot;
    inline static constexpr char APOS[] = "&apos;";  // '  → &apos;
    inline static constexpr char NL[]   = "&#xA;";  // \n（属性上下文）

    inline static constexpr size_t AMP_LEN  = 5;
    inline static constexpr size_t LT_LEN   = 4;
    inline static constexpr size_t GT_LEN   = 4;
    inline static constexpr size_t QUOT_LEN = 6;
    inline static constexpr size_t APOS_LEN = 6;
    inline static constexpr size_t NL_LEN   = 5;
};

} // namespace xml
} // namespace fastexcel

