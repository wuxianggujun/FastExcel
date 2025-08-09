# 主题实现指南（Theme Implementation Guide）

本文档说明 FastExcel 在主题（Theme）与配色/字体方案方面的实现与用法，并给出与源码一致的 API 示例，便于在代码中正确应用与修改主题。

## 模块概览

- 核心类型：`fastexcel::theme::Theme`
- 颜色方案：`fastexcel::theme::ThemeColorScheme`
  - 枚举：`ThemeColorScheme::ColorType`（Background1、Text1、Background2、Text2、Accent1..6、Hyperlink、FollowedHyperlink）
- 字体方案：`fastexcel::theme::ThemeFontScheme`（Major/Minor 字体族，Latin/EastAsia/Complex）
- 解析器：`fastexcel::theme::ThemeParser`（从 `xl/theme/theme1.xml` 解析到 `Theme` 对象）
- 写回：`Theme::toXML()` 生成最小可用的 OOXML 主题 XML
- 工作簿接口（节选，定义于 `core/Workbook.hpp`）：
  - `void setThemeXML(const std::string&)`
  - `void setOriginalThemeXML(const std::string&)`
  - `void setTheme(const theme::Theme&)`
  - `const theme::Theme* getTheme() const`
  - `void setThemeName(const std::string&)`
  - `void setThemeColor(theme::ThemeColorScheme::ColorType, const core::Color&)`
  - `bool setThemeColorByName(const std::string& name, const core::Color& color)`
  - `void setThemeMajorFontLatin/EastAsia/Complex(...)`
  - `void setThemeMinorFontLatin/EastAsia/Complex(...)`

## 常见用法

### 创建并设置主题

```cpp
#include "fastexcel/FastExcel.hpp"

using namespace fastexcel;
using namespace fastexcel::core;
using namespace fastexcel::theme;

int main() {
    auto wb = Workbook::create(Path("theme_demo.xlsx"));

    // 构建主题对象
    Theme thm("My Company Theme");

    // 配色：设置 accent1 与超链接色
    thm.colors().setColor(ThemeColorScheme::ColorType::Accent1, core::Color(0x33, 0x99, 0xFF));
    thm.colors().setColor(ThemeColorScheme::ColorType::Hyperlink, core::Color(0x00, 0x66, 0xCC));

    // 字体：设置 major/minor 字体族
    thm.fonts().setMajorFontLatin("Segoe UI");
    thm.fonts().setMinorFontLatin("Calibri");

    // 应用到工作簿
    wb->setTheme(thm);

    // 也可直接改名与单色项
    wb->setThemeName("My Company Theme");
    wb->setThemeColor(ThemeColorScheme::ColorType::Accent2, core::Color(0x22, 0x88, 0x44));

    // 正常写表/保存
    auto ws = wb->addWorksheet("Sheet1");
    ws->writeString(0, 0, "Hello Theme");

    wb->save();
    return 0;
}
```

### 按名称设置颜色（映射自 OOXML 名称）

```cpp
// 支持如 "accent1", "hlink", "lt1"/"lt2"(Background1/2), "dk1"/"dk2"(Text1/2) 等名称
bool ok = wb->setThemeColorByName("accent1", core::Color(0x33,0x99,0xFF));
```

### 从 XML 解析并保真写回

```cpp
// 解析现有主题 XML（例如通过读取器获取）
std::string theme_xml = /* xl/theme/theme1.xml 内容 */;
auto parsed = ThemeParser::parseFromXML(theme_xml);
if (parsed) {
    wb->setTheme(*parsed);
}

// 若需要：保留原始 XML，以便未编辑时保真写回
wb->setOriginalThemeXML(theme_xml);
```

## 与样式系统的关系

- 主题仅提供“主题色/主题字体”的来源，具体样式（`FormatDescriptor`）在序列化时可引用主题色槽位。
- 当修改主题后，使用主题色的样式在 Excel 中会体现为新颜色；无需逐一重写所有单元格样式。

## 最佳实践

- 优先使用 `Theme` 对象进行结构化编辑，便于验证与维护。
- 大批量变更颜色时，优先通过主题色槽位调整，避免逐格写死颜色。
- 仅在确需完全保真时持有原始 `theme.xml`，否则以 `Theme::toXML()` 生成更轻量的主题文件。
- 与样式仓储（`FormatRepository`）配合，在主题调整后可调用工作簿的优化/去重能力以减少重复样式。

---

最后更新：2025-08-10
