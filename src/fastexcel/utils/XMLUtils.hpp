#pragma once

#include "fastexcel/xml/XMLEscapes.hpp"
#include "fmt/format.h"
#include <string>

namespace fastexcel {
namespace utils {

/**
 * @brief XML工具类 - 提供XML相关的辅助函数
 */
class XMLUtils {
public:
  /**
   * @brief XML转义 - 将特殊字符转换为XML实体
   * @param text 需要转义的文本
   * @return 转义后的文本
   *
   * 转义规则：
   * - < -> &lt;
   * - > -> &gt;
   * - & -> &amp;
   * - " -> &quot;
   * - ' -> &apos;
   * - 跳过无效控制字符（保留制表符、换行符、回车符）
   */
  static std::string escapeXML(const std::string &text) {
    std::string result;
    result.reserve(static_cast<size_t>(text.size() * 1.2));

    for (char c : text) {
      switch (c) {
      case xml::XMLEscapes::CHAR_LT:
        result.append(xml::XMLEscapes::LT, xml::XMLEscapes::LT_LEN);
        break;
      case xml::XMLEscapes::CHAR_GT:
        result.append(xml::XMLEscapes::GT, xml::XMLEscapes::GT_LEN);
        break;
      case xml::XMLEscapes::CHAR_AMP:
        result.append(xml::XMLEscapes::AMP, xml::XMLEscapes::AMP_LEN);
        break;
      case xml::XMLEscapes::CHAR_QUOT:
        result.append(xml::XMLEscapes::QUOT, xml::XMLEscapes::QUOT_LEN);
        break;
      case xml::XMLEscapes::CHAR_APOS:
        result.append(xml::XMLEscapes::APOS, xml::XMLEscapes::APOS_LEN);
        break;
      default:
        // 只过滤真正的控制字符，不要过滤UTF-8多字节字符
        // 使用unsigned char比较避免负数问题
        unsigned char uc = static_cast<unsigned char>(c);
        if (uc < 0x20 && uc != 0x09 && uc != 0x0A && uc != 0x0D) {
          continue; // 跳过无效控制字符
        }
        result.push_back(c);
        break;
      }
    }

    return result;
  }

  /**
   * @brief XML反转义 - 将XML实体转换为特殊字符
   * @param text 需要反转义的文本
   * @return 反转义后的文本
   */
  static std::string unescapeXML(const std::string &text) {
    std::string result;
    result.reserve(text.size());

    for (size_t i = 0; i < text.length(); ++i) {
      if (text[i] == xml::XMLEscapes::CHAR_AMP && i + 1 < text.length()) {
        // 查找实体结束位置
        size_t end = text.find(';', i + 1);
        if (end != std::string::npos) {
          std::string entity = text.substr(i, end - i + 1);
          if (entity == xml::XMLEscapes::LT) {
            result.push_back(xml::XMLEscapes::CHAR_LT);
          } else if (entity == xml::XMLEscapes::GT) {
            result.push_back(xml::XMLEscapes::CHAR_GT);
          } else if (entity == xml::XMLEscapes::AMP) {
            result.push_back(xml::XMLEscapes::CHAR_AMP);
          } else if (entity == xml::XMLEscapes::QUOT) {
            result.push_back(xml::XMLEscapes::CHAR_QUOT);
          } else if (entity == xml::XMLEscapes::APOS) {
            result.push_back(xml::XMLEscapes::CHAR_APOS);
          } else {
            // 不认识的实体，保持原样
            result.append(entity);
          }
          i = end; // 跳过整个实体
        } else {
          result.push_back(text[i]);
        }
      } else {
        result.push_back(text[i]);
      }
    }

    return result;
  }

  /**
   * @brief 验证XML名称是否有效
   * @param name XML名称（标签名、属性名等）
   * @return 是否有效
   */
  static bool isValidXMLName(const std::string &name) {
    if (name.empty()) {
      return false;
    }

    // XML名称必须以字母、下划线或冒号开头
    char first = name[0];
    if (!((first >= 'A' && first <= 'Z') || (first >= 'a' && first <= 'z') ||
          first == xml::XMLEscapes::CHAR_UNDER ||
          first == xml::XMLEscapes::CHAR_COLON)) {
      return false;
    }

    // 其余字符可以是字母、数字、连字符、句点、下划线或冒号
    for (size_t i = 1; i < name.length(); ++i) {
      char c = name[i];
      if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
            (c >= '0' && c <= '9') || c == xml::XMLEscapes::CHAR_HYPHEN ||
            c == xml::XMLEscapes::CHAR_DOT ||
            c == xml::XMLEscapes::CHAR_UNDER ||
            c == xml::XMLEscapes::CHAR_COLON)) {
        return false;
      }
    }

    return true;
  }

  /**
   * @brief 生成XML属性字符串
   * @param name 属性名
   * @param value 属性值
   * @return 格式化的属性字符串，如 name="value"
   */
  static std::string formatAttribute(const std::string &name,
                                     const std::string &value) {
    return fmt::format("{}=\"{}\"", name, escapeXML(value));
  }

  /**
   * @brief 生成XML开始标签
   * @param tagName 标签名
   * @param attributes 属性字符串（可选）
   * @param selfClosing 是否为自闭合标签
   * @return 格式化的开始标签
   */
  static std::string formatStartTag(const std::string &tagName,
                                    const std::string &attributes = "",
                                    bool selfClosing = false) {
    if (attributes.empty()) {
      return selfClosing ? fmt::format("<{}/>", tagName)
                         : fmt::format("<{}>", tagName);
    }
    return selfClosing ? fmt::format("<{} {}/>", tagName, attributes)
                       : fmt::format("<{} {}>", tagName, attributes);
  }

  /**
   * @brief 生成XML结束标签
   * @param tagName 标签名
   * @return 格式化的结束标签
   */
  static std::string formatEndTag(const std::string &tagName) {
    return fmt::format("</{}>", tagName);
  }

  /**
   * @brief 生成完整的XML元素（开始标签+内容+结束标签）
   * @param tagName 标签名
   * @param content 元素内容
   * @param attributes 属性字符串（可选）
   * @param escapeContent 是否转义内容
   * @return 完整的XML元素
   */
  static std::string formatElement(const std::string &tagName,
                                   const std::string &content,
                                   const std::string &attributes = "",
                                   bool escapeContent = true) {
    return fmt::format("{}{}{}", formatStartTag(tagName, attributes),
                       escapeContent ? escapeXML(content) : content,
                       formatEndTag(tagName));
  }
};

} // namespace utils
} // namespace fastexcel
