#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include "fastexcel/utils/StylesParser.hpp"

namespace fastexcel {
namespace utils {

int ColumnWidthCalculator::parseRealMDW(const std::string& styles_xml_path) {
    auto font_info = StylesParser::parseDefaultFont(styles_xml_path);
    
    if (font_info.is_parsed) {
        return font_info.mdw;
    }
    
    // 解析失败，返回默认值
    return 7; // Calibri 11pt的默认MDW
}

int ColumnWidthCalculator::parseRealMDWFromWorkbook(const std::string& workbook_dir) {
    auto font_info = StylesParser::parseFromWorkbook(workbook_dir);
    
    if (font_info.is_parsed) {
        return font_info.mdw;
    }
    
    // 解析失败，返回默认值
    return 7; // Calibri 11pt的默认MDW
}

}} // namespace fastexcel::utils