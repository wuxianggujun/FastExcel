#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include "fastexcel/reader/StylesParser.hpp"
#include <fstream>
#include <filesystem>

namespace fastexcel {
namespace utils {

int ColumnWidthCalculator::parseRealMDW(const std::string& styles_xml_path) {
    try {
        // 读取styles.xml文件
        std::ifstream file(styles_xml_path);
        if (!file.is_open()) {
            return 7; // 默认Calibri 11pt的MDW
        }
        
        std::string xml_content((std::istreambuf_iterator<char>(file)),
                               std::istreambuf_iterator<char>());
        
        // 使用高性能StylesParser解析
        reader::StylesParser parser;
        if (!parser.parse(xml_content)) {
            return 7; // 解析失败，返回默认值
        }
        
        // 获取默认字体信息
        auto [font_name, font_size] = parser.getDefaultFontInfo();
        
        // 使用estimateMDW方法计算MDW
        return estimateMDW(font_name, font_size);
        
    } catch (const std::exception&) {
        // 发生异常，返回默认值
        return 7;
    }
}

int ColumnWidthCalculator::parseRealMDWFromWorkbook(const std::string& workbook_dir) {
    std::filesystem::path styles_path = std::filesystem::path(workbook_dir) / "xl" / "styles.xml";
    
    if (std::filesystem::exists(styles_path)) {
        return parseRealMDW(styles_path.string());
    }
    
    return 7; // 默认Calibri 11pt的MDW
}

}} // namespace fastexcel::utils