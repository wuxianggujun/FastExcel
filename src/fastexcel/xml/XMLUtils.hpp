#pragma once

#include <cstddef>

namespace fastexcel {
namespace xml {

// XML 转义常量与轻量工具
struct XMLEscapes {
    static constexpr char AMP[]  = "&amp;";   // &
    static constexpr char LT[]   = "&lt;";    // <
    static constexpr char GT[]   = "&gt;";    // >
    static constexpr char QUOT[] = "&quot;";  // "
    static constexpr char APOS[] = "&apos;";  // '
    static constexpr char NL[]   = "&#xA;";  // \n (attributes)
};

} // namespace xml
} // namespace fastexcel

