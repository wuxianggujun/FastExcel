# 主题（theme）API 与类关系

描述 Excel 主题（颜色/字体方案）并能序列化最小可用主题 XML；被 `Workbook` 使用并可由读取器返回解析结果。

## 关系图（简述）
- `theme::Theme` 组合 `ThemeColorScheme` 与 `ThemeFontScheme`
- `core::Workbook` 可持有 `unique_ptr<theme::Theme>` 并暴露主题相关设置 API
- `reader::XLSXReader` 可提供原始主题 XML 与解析后的 `Theme` 对象

---

## class fastexcel::theme::ThemeColorScheme
- 职责：维护 12 项标准主题颜色（背景/文本/6 个强调色/超链接/已访问链接）。
- 主要 API：
  - `getColor(ColorType)`、`setColor(ColorType, Color)`
  - `getColorByName("accent1"|"lt1"|"hlink"...)`、`setColorByName(name, color)`

## class fastexcel::theme::ThemeFontScheme
- 职责：维护 Major/Minor 两套字体集（Latin/EastAsia/Complex）。
- 主要 API：`getMajorFonts()/getMinorFonts()`；`setMajorFontLatin/EastAsia/Complex()`；`setMinorFont*()`

## class fastexcel::theme::Theme
- 职责：组合主题颜色与字体方案，提供名字与序列化。
- 主要 API：`getName()/setName()`、`colors()/fonts()`、`toXML()`

