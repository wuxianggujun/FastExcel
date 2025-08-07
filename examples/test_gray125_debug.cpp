#include "../src/fastexcel/reader/StylesParser.hpp"
#include <iostream>
#include <string>

int main() {
    fastexcel::reader::StylesParser parser;
    
    // 测试gray125模式解析
    std::string test_xml = R"(
    <styleSheet>
        <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
        </fills>
        <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            <xf numFmtId="0" fontId="0" fillId="1" borderId="0"/>
        </cellXfs>
    </styleSheet>
    )";
    
    std::cout << "Testing gray125 parsing..." << std::endl;
    
    bool success = parser.parse(test_xml);
    if (!success) {
        std::cout << "ERROR: Failed to parse test XML" << std::endl;
        return 1;
    }
    
    std::cout << "Parsed " << parser.getFormatCount() << " formats" << std::endl;
    
    // 获取第二个格式（应该是gray125）
    auto format = parser.getFormat(1);
    if (!format) {
        std::cout << "ERROR: Failed to get format 1" << std::endl;
        return 1;
    }
    
    std::cout << "Format 1 pattern type: " << static_cast<int>(format->getPattern()) << std::endl;
    std::cout << "Expected Gray125 enum value: " << static_cast<int>(fastexcel::core::PatternType::Gray125) << std::endl;
    
    if (format->getPattern() == fastexcel::core::PatternType::Gray125) {
        std::cout << "SUCCESS: Gray125 pattern correctly parsed!" << std::endl;
    } else {
        std::cout << "ERROR: Gray125 pattern not correctly parsed" << std::endl;
    }
    
    return 0;
}