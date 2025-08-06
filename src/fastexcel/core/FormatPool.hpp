#pragma once

#include "fastexcel/core/Format.hpp"
#include <functional>
#include <memory>
#include <unordered_map>
#include <vector>

namespace fastexcel::xml {
class XMLStreamWriter;
}

namespace fastexcel {
namespace core {

/**
 * @brief 格式键 - 用于格式去重的唯一标识
 */
struct FormatKey {
  // 字体属性
  std::string font_name;
  double font_size;
  bool bold;
  bool italic;
  bool underline;
  bool strikethrough;
  uint32_t font_color;

  // 对齐属性
  int horizontal_align;
  int vertical_align;
  bool text_wrap;
  int text_rotation;

  // 边框属性
  int border_style;
  uint32_t border_color;

  // 填充属性
  int pattern;
  uint32_t bg_color;
  uint32_t fg_color;

  // 数字格式
  std::string number_format;

  // 保护属性
  bool locked;
  bool hidden;

  FormatKey();
  explicit FormatKey(const Format &format);

  bool operator==(const FormatKey &other) const;
  bool operator!=(const FormatKey &other) const { return !(*this == other); }
};

} // namespace core
} // namespace fastexcel

// 为FormatKey提供哈希函数
namespace std {
template <> struct hash<fastexcel::core::FormatKey> {
  size_t operator()(const fastexcel::core::FormatKey &key) const;
};
} // namespace std

namespace fastexcel {
namespace core {

/**
 * @brief 格式池 - 借鉴libxlsxwriter的格式去重优化
 *
 * 实现格式对象的去重存储和管理，避免重复的格式定义
 */
class FormatPool {
private:
  std::vector<std::unique_ptr<Format>> formats_;         // 格式对象存储
  std::unordered_map<FormatKey, Format *> format_cache_; // 格式缓存
  std::unordered_map<Format *, size_t> format_to_index_; // 格式到索引的映射
  size_t next_index_;                                    // 下一个可用索引

  // 默认格式
  std::unique_ptr<Format> default_format_;

  // 格式复制专用：原始样式数据（绕过去重机制）
  std::unordered_map<int, std::shared_ptr<core::Format>> raw_styles_for_copy_;

public:
  FormatPool();
  ~FormatPool() = default;

  // 禁用拷贝，允许移动
  FormatPool(const FormatPool &) = delete;
  FormatPool &operator=(const FormatPool &) = delete;
  FormatPool(FormatPool &&) = default;
  FormatPool &operator=(FormatPool &&) = default;

  /**
   * @brief 获取或创建格式
   * @param key 格式键
   * @return 格式指针
   */
  Format *getOrCreateFormat(const FormatKey &key);

  /**
   * @brief 获取或创建格式（从现有格式）
   * @param format 现有格式
   * @return 格式指针
   */
  Format *getOrCreateFormat(const Format &format);

  /**
   * @brief 添加格式到池中
   * @param format 要添加的格式
   * @return 格式指针
   */
  Format *addFormat(std::unique_ptr<Format> format);

  /**
   * @brief 批量导入样式（用于格式复制）
   * @param styles 样式映射（索引->格式）
   */
  void importStyles(
      const std::unordered_map<int, std::shared_ptr<core::Format>> &styles);

  /**
   * @brief 设置原始样式用于XML生成（格式复制专用）
   * @param styles 原始样式映射
   */
  void setRawStylesForCopy(
      const std::unordered_map<int, std::shared_ptr<core::Format>> &styles);

  /**
   * @brief 检查是否有原始样式用于复制
   * @return 是否有原始样式
   */
  bool hasRawStylesForCopy() const { return !raw_styles_for_copy_.empty(); }

  /**
   * @brief 获取原始样式数据用于复制（只读访问）
   * @return 原始样式数据的引用
   */
  const std::unordered_map<int, std::shared_ptr<core::Format>>& getRawStylesForCopy() const {
    return raw_styles_for_copy_;
  }

  /**
   * @brief 获取格式索引
   * @param format 格式指针
   * @return 格式索引
   */
  size_t getFormatIndex(Format *format) const;

  /**
   * @brief 根据索引获取格式
   * @param index 格式索引
   * @return 格式指针
   */
  Format *getFormatByIndex(size_t index) const;

  /**
   * @brief 获取默认格式
   * @return 默认格式指针
   */
  Format *getDefaultFormat() const { return default_format_.get(); }

  /**
   * @brief 获取格式数量
   * @return 格式数量
   */
  size_t getFormatCount() const { return formats_.size(); }

  /**
   * @brief 获取缓存命中率
   * @return 缓存命中率
   */
  double getCacheHitRate() const;

  /**
   * @brief 清空格式池
   */
  void clear();

  /**
   * @brief 生成样式XML到回调函数（流式写入）
   * @param callback 数据写入回调函数
   */
  void generateStylesXML(
      const std::function<void(const char *, size_t)> &callback) const;

  /**
   * @brief 生成样式XML到文件（流式写入）
   * @param filename 输出文件名
   */
  void generateStylesXMLToFile(const std::string &filename) const;

  /**
   * @brief 获取内存使用统计
   * @return 内存使用字节数
   */
  size_t getMemoryUsage() const;

  /**
   * @brief 获取去重统计
   * @return (原始请求数, 实际格式数, 去重率)
   */
  struct DeduplicationStats {
    size_t total_requests;      // 总请求数
    size_t unique_formats;      // 唯一格式数
    double deduplication_ratio; // 去重率
  };
  DeduplicationStats getDeduplicationStats() const;

  /**
   * @brief 格式迭代器
   */
  std::vector<std::unique_ptr<Format>>::const_iterator begin() const {
    return formats_.begin();
  }
  std::vector<std::unique_ptr<Format>>::const_iterator end() const {
    return formats_.end();
  }

private:
  // 内部辅助方法
  Format *createFormatFromKey(const FormatKey &key);
  void updateFormatIndex(Format *format, size_t index);

  /**
   * @brief 通用的样式XML生成方法
   * @param writer XML写入器引用
   */
  void generateStylesXMLInternal(xml::XMLStreamWriter &writer) const;

  // 统计数据
  mutable size_t total_requests_;
  mutable size_t cache_hits_;
};

} // namespace core
} // namespace fastexcel
