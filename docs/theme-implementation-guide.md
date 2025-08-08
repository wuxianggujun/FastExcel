# FastExcel 主题系统实施指南

## 概述

本文档详细说明如何实现和完善FastExcel的主题解析和编辑功能，这是当前架构中的一个关键缺失部分。

## 1. 主题系统架构设计

### 1.1 核心组件

```cpp
namespace fastexcel::theme {

// 主题颜色方案
class ThemeColorScheme {
public:
    enum ColorType {
        Background1,    // lt1
        Text1,         // dk1
        Background2,   // lt2
        Text2,         // dk2
        Accent1,
        Accent2,
        Accent3,
        Accent4,
        Accent5,
        Accent6,
        Hyperlink,
        FollowedHyperlink
    };
    
private:
    std::array<Color, 12> colors_;
    std::unordered_map<std::string, std::vector<TintShade>> tint_shades_;
    
public:
    // 获取基础颜色
    Color getColor(ColorType type) const;
    Color getColor(const std::string& name) const;
    
    // 获取带有色调/阴影的颜色
    Color getTintedColor(ColorType type, double tint) const;
    
    // 设置颜色
    void setColor(ColorType type, const Color& color);
    
    // 从XML解析
    bool parseFromXML(const std::string& xml);
    std::string toXML() const;
};

// 主题字体方案
class ThemeFontScheme {
public:
    struct FontSet {
        std::string latin;
        std::string eastAsia;
        std::string complexScript;
        std::string symbol;
    };
    
private:
    FontSet major_fonts_;  // 标题字体
    FontSet minor_fonts_;  // 正文字体
    
public:
    // 获取字体
    std::string getMajorFont(const std::string& script = "latin") const;
    std::string getMinorFont(const std::string& script = "latin") const;
    
    // 设置字体
    void setMajorFont(const std::string& script, const std::string& font);
    void setMinorFont(const std::string& script, const std::string& font);
    
    // XML序列化
    bool parseFromXML(const std::string& xml);
    std::string toXML() const;
};

// 主题格式方案
class ThemeFormatScheme {
public:
    struct FillStyle {
        PatternType pattern;
        Color foreground;
        Color background;
        double transparency;
    };
    
    struct LineStyle {
        BorderStyle style;
        Color color;
        double width;
    };
    
    struct EffectStyle {
        bool shadow;
        bool reflection;
        bool glow;
        Color glowColor;
        double glowRadius;
    };
    
private:
    std::vector<FillStyle> fill_styles_;
    std::vector<LineStyle> line_styles_;
    std::vector<EffectStyle> effect_styles_;
    
public:
    // 样式访问
    const FillStyle& getFillStyle(size_t index) const;
    const LineStyle& getLineStyle(size_t index) const;
    const EffectStyle& getEffectStyle(size_t index) const;
    
    // 样式修改
    void addFillStyle(const FillStyle& style);
    void addLineStyle(const LineStyle& style);
    void addEffectStyle(const EffectStyle& style);
    
    // XML序列化
    bool parseFromXML(const std::string& xml);
    std::string toXML() const;
};

// 完整的主题对象
class Theme {
private:
    std::string name_;
    ThemeColorScheme color_scheme_;
    ThemeFontScheme font_scheme_;
    ThemeFormatScheme format_scheme_;
    
    // 缓存解析后的样式
    mutable std::unordered_map<std::string, FormatDescriptor> style_cache_;
    
public:
    Theme() = default;
    explicit Theme(const std::string& name);
    
    // 基本属性
    const std::string& getName() const { return name_; }
    void setName(const std::string& name) { name_ = name; }
    
    // 组件访问
    ThemeColorScheme& getColorScheme() { return color_scheme_; }
    ThemeFontScheme& getFontScheme() { return font_scheme_; }
    ThemeFormatScheme& getFormatScheme() { return format_scheme_; }
    
    const ThemeColorScheme& getColorScheme() const { return color_scheme_; }
    const ThemeFontScheme& getFontScheme() const { return font_scheme_; }
    const ThemeFormatScheme& getFormatScheme() const { return format_scheme_; }
    
    // 应用主题到样式
    FormatDescriptor applyToStyle(const FormatDescriptor& base,
                                  const ThemeReference& ref) const;
    
    // 预定义主题样式
    FormatDescriptor getTableHeaderStyle() const;
    FormatDescriptor getTableDataStyle() const;
    FormatDescriptor getTitleStyle() const;
    FormatDescriptor getSubtitleStyle() const;
    
    // XML序列化
    bool loadFromFile(const std::string& path);
    bool saveToFile(const std::string& path) const;
    bool parseFromXML(const std::string& xml);
    std::string toXML() const;
    
    // 预定义主题
    static Theme createOfficeTheme();
    static Theme createModernTheme();
    static Theme createClassicTheme();
    static Theme createDarkTheme();
};

}
```

### 1.2 主题管理器

```cpp
namespace fastexcel::theme {

// 主题管理器 - 单例模式
class ThemeManager {
private:
    std::unordered_map<std::string, std::unique_ptr<Theme>> themes_;
    Theme* active_theme_ = nullptr;
    std::string default_theme_name_ = "Office";
    
    // 主题变更监听器
    std::vector<std::function<void(const Theme&)>> listeners_;
    
    ThemeManager();
    
public:
    static ThemeManager& getInstance();
    
    // 主题管理
    void registerTheme(std::unique_ptr<Theme> theme);
    void unregisterTheme(const std::string& name);
    Theme* getTheme(const std::string& name);
    std::vector<std::string> getAvailableThemes() const;
    
    // 活动主题
    void setActiveTheme(const std::string& name);
    Theme* getActiveTheme() { return active_theme_; }
    const Theme* getActiveTheme() const { return active_theme_; }
    
    // 主题应用
    FormatDescriptor applyTheme(const FormatDescriptor& format,
                                const std::string& theme_element) const;
    
    // 监听器
    void addThemeChangeListener(std::function<void(const Theme&)> listener);
    void removeAllListeners();
    
    // 初始化
    void loadBuiltinThemes();
    void loadCustomThemes(const std::string& directory);
};

}
```

## 2. XML解析实现

### 2.1 主题XML解析器

```cpp
namespace fastexcel::theme::parser {

class ThemeXMLParser {
private:
    // XML命名空间
    static constexpr const char* THEME_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static constexpr const char* RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
public:
    // 解析主题文件
    static std::unique_ptr<Theme> parseTheme(const std::string& xml_content);
    
    // 解析各个部分
    static ThemeColorScheme parseColorScheme(const std::string& xml_node);
    static ThemeFontScheme parseFontScheme(const std::string& xml_node);
    static ThemeFormatScheme parseFormatScheme(const std::string& xml_node);
    
private:
    // 辅助解析方法
    static Color parseColor(const std::string& color_node);
    static std::vector<TintShade> parseTintShades(const std::string& node);
    static PatternType parsePattern(const std::string& pattern_str);
    static BorderStyle parseBorderStyle(const std::string& style_str);
    
    // RGB颜色解析
    static Color parseRGBColor(const std::string& rgb_str);
    static Color parseSystemColor(const std::string& sys_color);
    static Color parseSchemeColor(const std::string& scheme_ref);
};

// 使用libexpat的具体实现
class ExpatThemeParser : public ThemeXMLParser {
private:
    struct ParseContext {
        Theme* theme;
        std::string current_element;
        std::stack<std::string> element_stack;
        std::map<std::string, std::string> attributes;
    };
    
    static void XMLCALL startElement(void* userData, const char* name, const char** atts);
    static void XMLCALL endElement(void* userData, const char* name);
    static void XMLCALL characterData(void* userData, const char* s, int len);
    
public:
    static std::unique_ptr<Theme> parse(const std::string& xml_content);
};

}
```

### 2.2 主题XML生成器

```cpp
namespace fastexcel::theme::generator {

class ThemeXMLGenerator {
public:
    // 生成完整的主题XML
    static std::string generateThemeXML(const Theme& theme);
    
    // 生成各个部分
    static std::string generateColorScheme(const ThemeColorScheme& colors);
    static std::string generateFontScheme(const ThemeFontScheme& fonts);
    static std::string generateFormatScheme(const ThemeFormatScheme& formats);
    
private:
    // 辅助生成方法
    static std::string colorToXML(const Color& color);
    static std::string patternToXML(PatternType pattern);
    static std::string borderStyleToXML(BorderStyle style);
    
    // XML格式化
    static std::string formatXML(const std::string& raw_xml);
};

}
```

## 3. 主题与样式集成

### 3.1 主题感知的样式构建器

```cpp
namespace fastexcel::core {

// 扩展的样式构建器，支持主题
class ThemeAwareStyleBuilder : public StyleBuilder {
private:
    theme::Theme* theme_ = nullptr;
    std::vector<theme::ThemeReference> theme_refs_;
    
public:
    ThemeAwareStyleBuilder() = default;
    explicit ThemeAwareStyleBuilder(theme::Theme* theme);
    
    // 主题颜色引用
    ThemeAwareStyleBuilder& themeColor(theme::ThemeColorScheme::ColorType type,
                                       double tint = 0.0);
    
    // 主题字体引用
    ThemeAwareStyleBuilder& themeMajorFont(double size = 11.0);
    ThemeAwareStyleBuilder& themeMinorFont(double size = 11.0);
    
    // 主题样式引用
    ThemeAwareStyleBuilder& themeTableStyle(const std::string& style_name);
    ThemeAwareStyleBuilder& themeCellStyle(const std::string& style_name);
    
    // 构建时应用主题
    FormatDescriptor build() const override;
    FormatDescriptor buildWithTheme(const theme::Theme& theme) const;
};

}
```

### 3.2 工作簿主题集成

```cpp
namespace fastexcel::core {

// 扩展Workbook类以支持主题
class Workbook {
    // ... 现有成员 ...
    
private:
    std::unique_ptr<theme::Theme> theme_;
    bool use_default_theme_ = true;
    
public:
    // 主题操作
    void setTheme(std::unique_ptr<theme::Theme> theme);
    void setTheme(const std::string& theme_name);
    theme::Theme* getTheme() { return theme_.get(); }
    const theme::Theme* getTheme() const { return theme_.get(); }
    
    // 应用主题样式
    int addThemedStyle(const ThemeAwareStyleBuilder& builder);
    void applyThemeToAllStyles();
    
    // 主题导出/导入
    bool exportTheme(const std::string& path) const;
    bool importTheme(const std::string& path);
};

}
```

## 4. 实现步骤

### 第一步：基础数据结构（2-3天）

1. 实现Color类的扩展，支持主题颜色
2. 实现ThemeColorScheme类
3. 实现ThemeFontScheme类
4. 实现ThemeFormatScheme类
5. 编写单元测试

### 第二步：XML解析（3-4天）

1. 实现基于libexpat的XML解析器
2. 解析颜色方案
3. 解析字体方案
4. 解析格式方案
5. 测试标准主题文件解析

### 第三步：XML生成（2-3天）

1. 实现XML生成器
2. 生成颜色方案XML
3. 生成字体方案XML
4. 生成格式方案XML
5. 验证生成的XML符合规范

### 第四步：主题管理器（2天）

1. 实现ThemeManager单例
2. 加载内置主题
3. 实现主题切换
4. 添加监听器机制

### 第五步：样式集成（3-4天）

1. 扩展StyleBuilder支持主题
2. 修改FormatRepository支持主题引用
3. 更新Workbook类集成主题
4. 实现主题样式应用

### 第六步：测试和优化（2-3天）

1. 完整的单元测试
2. 集成测试
3. 性能测试
4. 内存泄漏检查
5. 文档编写

## 5. 使用示例

### 5.1 基本使用

```cpp
// 创建工作簿并设置主题
auto workbook = Workbook::create("output.xlsx");
workbook->setTheme("Modern");

// 使用主题颜色
auto header_style = workbook->createStyleBuilder()
    .themeColor(theme::ThemeColorScheme::Accent1)
    .themeMajorFont(14.0)
    .bold()
    .build();

// 应用样式
worksheet->getCell(0, 0).setStyle(header_style);
```

### 5.2 自定义主题

```cpp
// 创建自定义主题
auto custom_theme = std::make_unique<theme::Theme>("MyTheme");

// 设置颜色方案
auto& colors = custom_theme->getColorScheme();
colors.setColor(theme::ThemeColorScheme::Accent1, Color::fromRGB(0x4472C4));
colors.setColor(theme::ThemeColorScheme::Accent2, Color::fromRGB(0xED7D31));

// 设置字体方案
auto& fonts = custom_theme->getFontScheme();
fonts.setMajorFont("latin", "Arial");
fonts.setMinorFont("latin", "Calibri");

// 应用到工作簿
workbook->setTheme(std::move(custom_theme));
```

### 5.3 主题编辑

```cpp
// 获取当前主题
auto* theme = workbook->getTheme();

// 修改主题颜色
theme->getColorScheme().setColor(
    theme::ThemeColorScheme::Hyperlink,
    Color::fromRGB(0x0563C1)
);

// 修改主题字体
theme->getFontScheme().setMajorFont("eastAsia", "微软雅黑");

// 应用修改到所有样式
workbook->applyThemeToAllStyles();
```

## 6. 性能考虑

### 6.1 缓存策略

```cpp
class ThemeStyleCache {
private:
    struct CacheKey {
        size_t base_style_hash;
        size_t theme_hash;
        std::string theme_element;
        
        bool operator==(const CacheKey& other) const;
    };
    
    struct CacheKeyHash {
        size_t operator()(const CacheKey& key) const;
    };
    
    std::unordered_map<CacheKey, FormatDescriptor, CacheKeyHash> cache_;
    size_t max_size_ = 1000;
    
public:
    std::optional<FormatDescriptor> get(const CacheKey& key) const;
    void put(const CacheKey& key, const FormatDescriptor& format);
    void clear();
    void setMaxSize(size_t size);
};
```

### 6.2 延迟解析

```cpp
class LazyTheme {
private:
    mutable std::optional<ThemeColorScheme> color_scheme_;
    mutable std::optional<ThemeFontScheme> font_scheme_;
    mutable std::optional<ThemeFormatScheme> format_scheme_;
    std::string xml_content_;
    
public:
    // 延迟解析，只在需要时解析
    const ThemeColorScheme& getColorScheme() const {
        if (!color_scheme_) {
            color_scheme_ = parseColorScheme(xml_content_);
        }
        return *color_scheme_;
    }
};
```

## 7. 错误处理

```cpp
namespace fastexcel::theme {

// 主题相关异常
class ThemeException : public std::exception {
private:
    std::string message_;
    
public:
    explicit ThemeException(const std::string& msg) : message_(msg) {}
    const char* what() const noexcept override { return message_.c_str(); }
};

class ThemeParseException : public ThemeException {
public:
    ThemeParseException(const std::string& msg, int line = -1, int column = -1)
        : ThemeException(formatMessage(msg, line, column)) {}
    
private:
    static std::string formatMessage(const std::string& msg, int line, int column);
};

class ThemeNotFound : public ThemeException {
public:
    explicit ThemeNotFound(const std::string& theme_name)
        : ThemeException("Theme not found: " + theme_name) {}
};

}
```

## 8. 测试计划

### 8.1 单元测试

```cpp
// 颜色方案测试
TEST(ThemeColorScheme, ParseXML) {
    const char* xml = R"(
        <a:clrScheme name="Office">
            <a:dk1><a:sysClr val="windowText"/></a:dk1>
            <a:lt1><a:sysClr val="window"/></a:lt1>
            <!-- ... -->
        </a:clrScheme>
    )";
    
    ThemeColorScheme scheme;
    ASSERT_TRUE(scheme.parseFromXML(xml));
    EXPECT_EQ(scheme.getColor(ThemeColorScheme::Text1), Color::BLACK);
}

// 主题应用测试
TEST(Theme, ApplyToStyle) {
    Theme theme = Theme::createOfficeTheme();
    StyleBuilder builder;
    builder.themeColor(ThemeColorScheme::Accent1);
    
    FormatDescriptor format = builder.buildWithTheme(theme);
    EXPECT_EQ(format.getFontColor(), theme.getColorScheme().getColor(ThemeColorScheme::Accent1));
}
```

### 8.2 集成测试

```cpp
TEST(WorkbookTheme, SaveAndLoad) {
    // 创建带主题的工作簿
    auto workbook = Workbook::create("test_theme.xlsx");
    workbook->setTheme("Modern");
    
    auto worksheet = workbook->addWorksheet("Sheet1");
    worksheet->writeString(0, 0, "Test");
    
    // 应用主题样式
    auto style = workbook->createStyleBuilder()
        .themeColor(ThemeColorScheme::Accent1)
        .build();
    worksheet->getCell(0, 0).setStyle(style);
    
    ASSERT_TRUE(workbook->save());
    
    // 重新加载并验证
    auto loaded = Workbook::open("test_theme.xlsx");
    ASSERT_NE(loaded->getTheme(), nullptr);
    EXPECT_EQ(loaded->getTheme()->getName(), "Modern");
}
```

## 9. 兼容性注意事项

### 9.1 Excel版本兼容性

- Excel 2007+: 完全支持主题
- Excel 2003: 降级到最接近的颜色
- LibreOffice Calc: 部分支持

### 9.2 主题文件格式

遵循ECMA-376标准：
- 主题文件位置: xl/theme/theme1.xml
- 命名空间: drawingml/2006/main
- 关系类型: officeDocument/2006/relationships/theme

## 10. 总结

通过实现完整的主题系统，FastExcel将能够：

1. **完整支持Excel主题功能**
   - 解析和编辑主题文件
   - 应用主题到样式
   - 创建自定义主题

2. **提供更好的用户体验**
   - 简化样式管理
   - 支持主题切换
   - 保持与Excel的一致性

3. **提高代码质量**
   - 模块化设计
   - 清晰的接口
   - 完善的测试覆盖

实施此方案后，FastExcel将具备与商业Excel库相当的主题处理能力，同时保持高性能和易用性。