// 验证 rfind 修复
#include <iostream>
#include <string>

int main() {
    std::string part = "xl/worksheets/sheet1.xml";
    
    std::cout << "String: " << part << std::endl;
    std::cout << "Length: " << part.length() << std::endl;
    std::cout << std::endl;
    
    // 找到所有 "sheet" 的位置
    size_t pos = part.find("sheet");
    while (pos != std::string::npos) {
        std::cout << "find('sheet') at position: " << pos 
                  << " -> '" << part.substr(pos, 5) << "'" << std::endl;
        pos = part.find("sheet", pos + 1);
    }
    
    std::cout << std::endl;
    
    // 使用 rfind 找最后一个
    auto pos1 = part.rfind("sheet");
    auto pos2 = part.find(".xml");
    
    std::cout << "rfind('sheet') position: " << pos1 << std::endl;
    std::cout << "find('.xml') position: " << pos2 << std::endl;
    
    if (pos1 != std::string::npos && pos2 != std::string::npos) {
        size_t number_start = pos1 + 5;
        std::cout << "Number start position: " << number_start << std::endl;
        
        if (number_start < pos2) {
            std::string number_str = part.substr(number_start, pos2 - number_start);
            std::cout << "Extracted number string: '" << number_str << "'" << std::endl;
            
            try {
                int idx = std::stoi(number_str) - 1;
                std::cout << "SUCCESS! Parsed index: " << idx << std::endl;
            } catch (const std::exception& e) {
                std::cout << "ERROR: " << e.what() << std::endl;
            }
        } else {
            std::cout << "ERROR: Invalid positions!" << std::endl;
        }
    }
    
    return 0;
}