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

    // XML标签构建字符（非转义用）
    inline static constexpr char TAG_OPEN    = '<';     // 标签开始
    inline static constexpr char TAG_CLOSE   = '>';     // 标签结束
    inline static constexpr char TAG_SLASH   = '/';     // 闭合标签斜杠
    inline static constexpr char ATTR_QUOTE  = '"';     // 属性引号
    inline static constexpr char SPACE       = ' ';     // 空格

    // XML字符常量（用于比较和赋值）
    inline static constexpr char CHAR_LT     = '<';     // 小于号字符
    inline static constexpr char CHAR_GT     = '>';     // 大于号字符
    inline static constexpr char CHAR_AMP    = '&';     // 和号字符
    inline static constexpr char CHAR_QUOT   = '"';     // 双引号字符
    inline static constexpr char CHAR_APOS   = '\'';    // 单引号字符
    inline static constexpr char CHAR_EQUAL  = '=';     // 等号字符
    inline static constexpr char CHAR_UNDER  = '_';     // 下划线字符
    inline static constexpr char CHAR_COLON  = ':';     // 冒号字符
    inline static constexpr char CHAR_HYPHEN = '-';     // 连字符
    inline static constexpr char CHAR_DOT    = '.';     // 句点字符
};

} // namespace xml
} // namespace fastexcel

