#include "fastexcel/utils/Logger.hpp"
#pragma once

#include <functional>
#include <memory>
#include <string>
#include <vector>

namespace fastexcel {
namespace theme {
class Theme;
}
namespace core {
class SharedStringTable;
class FormatRepository;
class Workbook;
class IFileWriter;
} // namespace core
namespace xml {
class UnifiedXMLGenerator;
}
} // namespace fastexcel

namespace fastexcel {
namespace xml {

// 轻量上下文由 UnifiedXMLGenerator 承担，生成器只负责自身部件
struct XMLContextView {
  const class ::fastexcel::core::Workbook *workbook = nullptr;
  const class ::fastexcel::core::FormatRepository *format_repo = nullptr;
  const class ::fastexcel::core::SharedStringTable *sst = nullptr;
  const class ::fastexcel::theme::Theme *theme = nullptr;
};

class IXMLPartGenerator {
public:
  virtual ~IXMLPartGenerator() = default;
  // 返回该生成器负责的包内路径（一个或多个）
  virtual std::vector<std::string>
  partNames(const XMLContextView &ctx) const = 0;
  // 生成指定部件，写入到 IFileWriter（内部可选择批量或流式）
  virtual bool generatePart(const std::string &part, const XMLContextView &ctx,
                            ::fastexcel::core::IFileWriter &writer) = 0;
};

} // namespace xml
} // namespace fastexcel
