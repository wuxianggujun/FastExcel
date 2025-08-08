#pragma once

#include <memory>
#include <string>

#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/xml/XMLStreamReader.hpp"

namespace fastexcel {
namespace theme {

// 高性能主题解析器：解析 xl/theme/theme1.xml 到 Theme 对象
class ThemeParser {
public:
    // 解析完整XML字符串
    static std::unique_ptr<Theme> parseFromXML(const std::string& xml);

private:
    static void parseColorScheme(xml::XMLStreamReader::SimpleElement* clrScheme, Theme& out);
    static void parseFontScheme(xml::XMLStreamReader::SimpleElement* fontScheme, Theme& out);
};

} // namespace theme
} // namespace fastexcel


