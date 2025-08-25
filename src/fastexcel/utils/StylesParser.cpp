#include "fastexcel/utils/StylesParser.hpp"
#include "fastexcel/utils/ColumnWidthCalculator.hpp"
#include <fstream>
#include <sstream>
#include <regex>
#include <filesystem>

namespace fastexcel {
namespace utils {

// DefaultFontInfo::calculateMDW实现
int StylesParser::DefaultFontInfo::calculateMDW(const std::string& font_name, double font_size) {
    return ColumnWidthCalculator::estimateMDW(font_name, font_size);
}

StylesParser::DefaultFontInfo StylesParser::parseDefaultFont(const std::string& styles_xml_path) {
    std::ifstream file(styles_xml_path);
    if (!file.is_open()) {
        return DefaultFontInfo(); // 返回默认值
    }
    
    std::stringstream buffer;
    buffer << file.rdbuf();
    return parseDefaultFontFromContent(buffer.str());
}

StylesParser::DefaultFontInfo StylesParser::parseDefaultFontFromContent(const std::string& xml_content) {
    try {
        // 1. 查找Normal样式（cellXfs中index=0的样式）的fontId
        int font_id = findNormalStyleFontId(xml_content);
        if (font_id == -1) {
            font_id = 0; // 默认使用第一个字体
        }
        
        // 2. 根据fontId解析字体信息
        auto [font_name, font_size] = parseFontById(xml_content, font_id);
        
        if (!font_name.empty()) {
            return DefaultFontInfo(font_name, font_size);
        }
        
    } catch (const std::exception& e) {
        // 解析失败，返回默认值
    }
    
    return DefaultFontInfo(); // 返回默认值
}

StylesParser::DefaultFontInfo StylesParser::parseFromWorkbook(const std::string& workbook_dir) {
    std::filesystem::path styles_path = std::filesystem::path(workbook_dir) / "xl" / "styles.xml";
    
    if (std::filesystem::exists(styles_path)) {
        return parseDefaultFont(styles_path.string());
    }
    
    return DefaultFontInfo(); // 返回默认值
}

std::pair<std::string, double> StylesParser::parseFontById(const std::string& xml_content, int font_id) {
    try {
        // 查找fonts节点下的第font_id个font
        std::regex fonts_regex(R"(<fonts[^>]*>(.*?)</fonts>)", std::regex_constants::icase | std::regex_constants::dotall);
        std::smatch fonts_match;
        
        if (!std::regex_search(xml_content, fonts_match, fonts_regex)) {
            return {"", 0.0};
        }
        
        std::string fonts_content = fonts_match[1].str();
        
        // 查找所有font节点
        std::regex font_regex(R"(<font[^>]*>(.*?)</font>)", std::regex_constants::icase | std::regex_constants::dotall);
        std::sregex_iterator font_iter(fonts_content.begin(), fonts_content.end(), font_regex);
        std::sregex_iterator font_end;
        
        int current_id = 0;
        for (; font_iter != font_end; ++font_iter, ++current_id) {
            if (current_id == font_id) {
                std::string font_content = font_iter->str();
                
                // 解析字体名称（Latin字体）
                std::string font_name = extractTagValue(font_content, "name");
                if (font_name.empty()) {
                    font_name = extractTagAttribute(font_content, "name", "val");
                }
                
                // 解析字体大小
                std::string size_str = extractTagValue(font_content, "sz");
                if (size_str.empty()) {
                    size_str = extractTagAttribute(font_content, "sz", "val");
                }
                
                double font_size = 11.0; // 默认值
                if (!size_str.empty()) {
                    try {
                        font_size = std::stod(size_str);
                    } catch (const std::invalid_argument& e) {
                        FASTEXCEL_LOG_DEBUG("Invalid font size string '{}': {}", size_str, e.what());
                        font_size = 11.0;
                    } catch (const std::out_of_range& e) {
                        FASTEXCEL_LOG_DEBUG("Font size '{}' out of range: {}", size_str, e.what());
                        font_size = 11.0;
                    } catch (const std::exception& e) {
                        FASTEXCEL_LOG_DEBUG("Exception parsing font size '{}': {}", size_str, e.what());
                        font_size = 11.0;
                    }
                }
                
                return {font_name, font_size};
            }
        }
        
    } catch (const std::exception& e) {
        // 解析失败
    }
    
    return {"", 0.0};
}

int StylesParser::findNormalStyleFontId(const std::string& xml_content) {
    try {
        // 查找cellXfs节点
        std::regex cellxfs_regex(R"(<cellXfs[^>]*>(.*?)</cellXfs>)", std::regex_constants::icase | std::regex_constants::dotall);
        std::smatch cellxfs_match;
        
        if (!std::regex_search(xml_content, cellxfs_match, cellxfs_regex)) {
            return -1;
        }
        
        std::string cellxfs_content = cellxfs_match[1].str();
        
        // 查找第一个xf节点（通常是Normal样式）
        std::regex xf_regex(R"(<xf[^>]*>)", std::regex_constants::icase);
        std::smatch xf_match;
        
        if (std::regex_search(cellxfs_content, xf_match, xf_regex)) {
            std::string xf_tag = xf_match[0].str();
            
            // 提取fontId属性
            std::regex fontid_regex(R"(fontId\s*=\s*[\"'](\d+)[\"'])", std::regex_constants::icase);
            std::smatch fontid_match;
            
            if (std::regex_search(xf_tag, fontid_match, fontid_regex)) {
                return std::stoi(fontid_match[1].str());
            }
        }
        
    } catch (const std::exception& e) {
        // 解析失败
    }
    
    return -1;
}

std::string StylesParser::extractTagValue(const std::string& xml, const std::string& tag_name) {
    try {
        std::regex tag_regex("<" + tag_name + R"([^>]*>(.*?)<\/)" + tag_name + ">", 
                           std::regex_constants::icase | std::regex_constants::dotall);
        std::smatch match;
        
        if (std::regex_search(xml, match, tag_regex)) {
            return match[1].str();
        }
        
    } catch (const std::exception& e) {
        // 解析失败
    }
    
    return "";
}

std::string StylesParser::extractTagAttribute(const std::string& xml, const std::string& tag_name, const std::string& attr_name) {
    try {
        std::regex tag_regex("<" + tag_name + R"([^>]*)", std::regex_constants::icase);
        std::smatch tag_match;
        
        if (std::regex_search(xml, tag_match, tag_regex)) {
            std::string tag_content = tag_match[0].str();
            
            std::regex attr_regex(attr_name + R"(\s*=\s*[\"']([^\"']*)[\"'])", std::regex_constants::icase);
            std::smatch attr_match;
            
            if (std::regex_search(tag_content, attr_match, attr_regex)) {
                return attr_match[1].str();
            }
        }
        
    } catch (const std::exception& e) {
        // 解析失败
    }
    
    return "";
}

}} // namespace fastexcel::utils