#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/WorkbookModeSelector.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/CustomPropertyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/WorkbookCoordinator.hpp"
#include "fastexcel/core/WorksheetManager.hpp"
#include "fastexcel/core/ResourceManager.hpp"
#include "fastexcel/core/SharedStringCollector.hpp"
#include "fastexcel/core/managers/WorkbookDocumentManager.hpp"
#include "fastexcel/core/managers/WorkbookSecurityManager.hpp"
#include "fastexcel/core/managers/WorkbookDataManager.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/utils/AddressParser.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "FormatDescriptor.hpp"
#include "FormatRepository.hpp"
#include "StyleTransferContext.hpp"
#include "StyleBuilder.hpp"
#include "SharedStringTable.hpp"
// 内存管理组件
#include "fastexcel/memory/WorkbookMemoryManager.hpp"
#include "fastexcel/utils/SafeConstruction.hpp"
#include "fastexcel/utils/StringViewOptimized.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <unordered_map>
#include <set>
#include <ctime>
#include <functional>
#include <stdexcept>

namespace fastexcel {
namespace xml {
    class DocPropsXMLGenerator; // 前向声明
}
namespace opc {
    class PackageEditor;  // 前向声明PackageEditor
}

namespace reader {
    class XLSXReader;  // 前向声明XLSXReader
}

namespace core {

// 前向声明
class NamedStyle;
class ExcelStructureGenerator;

/**
 * @brief Workbook类 - Excel工作簿
 * 
 * 采用全新的样式管理系统，提供线程安全、高性能的Excel操作接口。
 * 
 * 核心特性：
 * - 不可变样式系统：使用FormatDescriptor值对象，线程安全
 * - 样式去重优化：FormatRepository自动去重，节省内存
 * - 流式样式构建：StyleBuilder提供链式调用API
 * - 跨工作簿操作：支持样式和工作表的复制传输
 * - 性能监控：内存使用统计、样式优化工具
 * - 工作表管理：完整的工作表生命周期管理
 * - 文档属性：丰富的元数据管理
 * - 多种保存选项：支持不同的性能优化模式
 */
class Workbook {
    friend class ExcelStructureGenerator;
    friend class ::fastexcel::opc::PackageEditor;  // 让PackageEditor能访问私有方法
    friend class ::fastexcel::reader::XLSXReader;  // 让XLSXReader能访问私有open方法
    friend class ::fastexcel::xml::DocPropsXMLGenerator;  // 让XML生成器能访问私有方法
private:
    // 核心架构组件
    std::string filename_;
    std::unique_ptr<WorkbookCoordinator> coordinator_;
    std::unique_ptr<WorksheetManager> worksheet_manager_;
    std::unique_ptr<ResourceManager> resource_manager_;
    
    // 样式和格式管理
    std::unique_ptr<FormatRepository> format_repo_;
    std::unique_ptr<SharedStringTable> shared_string_table_;
    std::unique_ptr<SharedStringCollector> string_collector_;
    
    // 主题管理
    std::unique_ptr<theme::Theme> theme_;
    std::string theme_xml_;
    std::string theme_xml_original_;
    bool theme_dirty_ = false;
    
    // 状态管理
    WorkbookState state_ = WorkbookState::CLOSED;
    FileSource file_source_ = FileSource::NEW_FILE;
    std::string original_package_path_;
    
    // 专门管理器（职责分离）
    std::unique_ptr<WorkbookDocumentManager> document_manager_;
    std::unique_ptr<WorkbookSecurityManager> security_manager_;
    
    // 统一内存管理器
    std::unique_ptr<memory::WorkbookMemoryManager> memory_manager_;
    
    // 性能统计（已移除）：改用各内存池自身统计
    std::unique_ptr<WorkbookDataManager> data_manager_;
    
    // 配置选项
    WorkbookOptions options_;
    bool preserve_unknown_parts_ = true;
    
    // 兼容性保留
    std::vector<std::shared_ptr<Worksheet>> worksheets_;    // 暂时保留以兼容
    std::unique_ptr<archive::FileManager> file_manager_;    // 暂时保留以兼容
    int next_sheet_id_ = 1;                                // 迁移到WorksheetManager
    size_t active_worksheet_index_ = 0;                    // 迁移到WorksheetManager

public:
    void setOptions(const WorkbookOptions& opts) {
        options_ = opts;
    }

    static std::unique_ptr<Workbook> create(const std::string& filepath);
    static std::unique_ptr<Workbook> openReadOnly(const std::string& filepath);
    static std::unique_ptr<Workbook> openEditable(const std::string& filepath);
    

    
    explicit Workbook(const Path& path);
    ~Workbook();
    
    // 禁用拷贝构造和赋值
    Workbook(const Workbook&) = delete;
    Workbook& operator=(const Workbook&) = delete;
    
    // 允许移动构造和赋值
    Workbook(Workbook&&) = default;
    Workbook& operator=(Workbook&&) = default;
    
    // 文件操作
    bool save();
    bool saveAs(const std::string& filename);
    
    // CSV功能
    
    /**
     * @brief 从CSV文件创建新的工作表
     * @param filepath CSV文件路径
     * @param sheet_name 工作表名称（可选）
     * @param options CSV解析选项
     * @return 新建的工作表指针，失败时返回nullptr
     */
    std::shared_ptr<Worksheet> loadCSV(const std::string& filepath, 
                                      const std::string& sheet_name = "",
                                      const CSVOptions& options = CSVOptions::standard()) {
        if (data_manager_) {
            auto result = data_manager_->importCSV(filepath, "", options);
            if (result.success && result.worksheet) {
                if (!sheet_name.empty()) {
                    // TODO: Rename worksheet if name provided  
                }
                return result.worksheet;
            }
        }
        return nullptr;
    }
    
    /**
     * @brief 从CSV字符串创建新的工作表
     * @param csv_content CSV内容字符串
     * @param sheet_name 工作表名称（可选）
     * @param options CSV解析选项
     * @return 新建的工作表指针，失败时返回nullptr
     */
    std::shared_ptr<Worksheet> loadCSVString(const std::string& csv_content,
                                            const std::string& sheet_name = "Sheet1",
                                            const CSVOptions& options = CSVOptions::standard()) {
        if (data_manager_) {
            auto result = data_manager_->importCSVString(csv_content, "", options);
            if (result.success && result.worksheet) {
                if (!sheet_name.empty()) {
                    // TODO: Rename worksheet if name provided
                }
                return result.worksheet;
            }
        }
        return nullptr;
    }
    
    /**
     * @brief 将工作簿的指定工作表导出为CSV
     * @param sheet_index 工作表索引（0-based）
     * @param filepath 目标文件路径
     * @param options CSV导出选项
     * @return 是否成功
     */
    bool exportSheetAsCSV(size_t sheet_index, const std::string& filepath,
                         const CSVOptions& options = CSVOptions::standard()) const {
        return data_manager_ ? data_manager_->exportCSV(sheet_index, filepath, options).success : false;
    }
    
    /**
     * @brief 将工作簿的指定工作表导出为CSV（按名称）
     * @param sheet_name 工作表名称
     * @param filepath 目标文件路径
     * @param options CSV导出选项
     * @return 是否成功
     */
    bool exportSheetAsCSV(const std::string& sheet_name, const std::string& filepath,
                         const CSVOptions& options = CSVOptions::standard()) const {
        return data_manager_ ? data_manager_->exportCSVByName(sheet_name, filepath, options).success : false;
    }
    
    /**
     * @brief 检查文件是否为CSV格式
     * @param filepath 文件路径
     * @return 是否为CSV文件
     */
    static bool isCSVFile(const std::string& filepath) {
        return WorkbookDataManager::isCSVFile(filepath);
    }
    
    bool isOpen() const;
    bool close();
    
    // 编辑模式/保真写回配置
    void setPreserveUnknownParts(bool enable) { preserve_unknown_parts_ = enable; }
    bool getPreserveUnknownParts() const { return preserve_unknown_parts_; }

    // 工作表管理
    std::shared_ptr<Worksheet> addSheet(const std::string& name = "");
    std::shared_ptr<Worksheet> insertSheet(size_t index, const std::string& name = "");
    bool removeSheet(const std::string& name);
    bool removeSheet(size_t index);
    
    std::shared_ptr<Worksheet> getSheet(const std::string& name);
    std::shared_ptr<Worksheet> getSheet(size_t index);
    std::shared_ptr<const Worksheet> getSheet(const std::string& name) const;
    std::shared_ptr<const Worksheet> getSheet(size_t index) const;
    
    // 便捷的工作表访问操作符
    std::shared_ptr<Worksheet> operator[](size_t index) {
        return getSheet(index);
    }
    
    std::shared_ptr<const Worksheet> operator[](size_t index) const {
        return getSheet(index);
    }
    
    std::shared_ptr<Worksheet> operator[](const std::string& name) {
        return getSheet(name);
    }
    
    std::shared_ptr<const Worksheet> operator[](const std::string& name) const {
        return getSheet(name);
    }
    
    size_t getSheetCount() const { return worksheets_.size(); }
    bool isEmpty() const { return worksheets_.empty(); }
    
    /**
     * @brief 获取第一个工作表
     * @return 第一个工作表指针，如果无工作表返回nullptr
     * 
     * @example
     * if (auto firstSheet = workbook.getFirstSheet()) {
     *     firstSheet->setValue("A1", std::string("Hello"));
     * }
     */
    std::shared_ptr<Worksheet> getFirstSheet();
    std::shared_ptr<const Worksheet> getFirstSheet() const;
    std::shared_ptr<Worksheet> getLastSheet();
    std::shared_ptr<const Worksheet> getLastSheet() const;
    std::vector<std::string> getSheetNames() const;
    
    // 便捷的工作表查找方法
    bool hasSheet(const std::string& name) const;
    std::shared_ptr<Worksheet> findSheet(const std::string& name);
    std::shared_ptr<const Worksheet> findSheet(const std::string& name) const;
    std::vector<std::shared_ptr<Worksheet>> getAllSheets();
    std::vector<std::shared_ptr<const Worksheet>> getAllSheets() const;
    int clearAllSheets();
    bool renameSheet(const std::string& old_name, const std::string& new_name);
    bool moveSheet(size_t from_index, size_t to_index);
    std::shared_ptr<Worksheet> copyWorksheet(const std::string& source_name, const std::string& new_name);
    std::shared_ptr<Worksheet> copyWorksheetFrom(const std::shared_ptr<const Worksheet>& source_worksheet,
                                const std::string& new_name = "");
    void setActiveWorksheet(size_t index);
    std::shared_ptr<Worksheet> getActiveWorksheet();
    std::shared_ptr<const Worksheet> getActiveWorksheet() const;
    
    // 便捷的单元格访问方法（跨工作表）
    /**
     * @brief 通过工作表索引和单元格坐标获取值
     * @tparam T 返回值类型
     * @param sheet_index 工作表索引
     * @param row 行号
     * @param col 列号
     * @return 单元格值
     * 
     * @example
     * auto value = workbook.getValue<std::string>(0, 0, 0);  // 获取第一个工作表A1的字符串值
     * auto number = workbook.getValue<double>(1, 1, 1);      // 获取第二个工作表B2的数字值
     */
    template<typename T>
    T getValue(size_t sheet_index, const Address& address) const {
        auto worksheet = getSheet(sheet_index);
        if (!worksheet) {
            throw std::runtime_error("Invalid worksheet index: " + std::to_string(sheet_index));
        }
        return worksheet->getValue<T>(address);
    }
    
    /**
     * @brief 通过工作表名称和单元格坐标设置值
     * @tparam T 值类型
     * @param sheet_name 工作表名称
     * @param row 行号
     * @param col 列号
     * @param value 要设置的值
     * 
     * @example
     * workbook.setValue("Sheet1", 0, 0, std::string("Hello"));  // 在Sheet1的A1设置字符串
     * workbook.setValue("Data", 1, 1, 123.45);                  // 在Data工作表的B2设置数字
     */
    template<typename T>
    void setValue(const std::string& sheet_name, const Address& address, const T& value) {
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        worksheet->setValue<T>(address, value);
    }
    
    /**
     * @brief 通过工作表索引和单元格坐标设置值
     * @tparam T 值类型
     * @param sheet_index 工作表索引
     * @param row 行号
     * @param col 列号
     * @param value 要设置的值
     */
    template<typename T>
    void setValue(size_t sheet_index, const Address& address, const T& value) {
        auto worksheet = getSheet(sheet_index);
        if (!worksheet) {
            throw std::runtime_error("Invalid worksheet index: " + std::to_string(sheet_index));
        }
        worksheet->setValue<T>(address, value);
    }
    
    // 安全版本的跨工作表访问方法（不抛异常）
    /**
     * @brief 安全获取工作表指针
     * @param name 工作表名称
     * @return 工作表指针的可选值，失败时返回std::nullopt
     */
    std::optional<std::shared_ptr<Worksheet>> tryGetSheet(const std::string& name) noexcept {
        try {
            auto sheet = findSheet(name);
            return sheet ? std::make_optional(sheet) : std::nullopt;
        } catch (const std::out_of_range& e) {
            FASTEXCEL_LOG_DEBUG("Sheet not found '{}': {}", name, e.what());
            return std::nullopt;
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_DEBUG("Exception in tryGetSheet '{}': {}", name, e.what());
            return std::nullopt;
        }
    }
    
    /**
     * @brief 安全获取工作表指针（索引版本）
     * @param index 工作表索引
     * @return 工作表指针的可选值，失败时返回std::nullopt
     */
    std::optional<std::shared_ptr<Worksheet>> tryGetSheet(size_t index) noexcept {
        try {
            auto sheet = getSheet(index);
            return sheet ? std::make_optional(sheet) : std::nullopt;
        } catch (const std::out_of_range& e) {
            FASTEXCEL_LOG_DEBUG("Sheet index out of range {}: {}", index, e.what());
            return std::nullopt;
        } catch (const std::exception& e) {
            FASTEXCEL_LOG_DEBUG("Exception in tryGetSheet index {}: {}", index, e.what());
            return std::nullopt;
        }
    }
    
    /**
     * @brief 安全获取单元格值（通过工作表名称）
     * @tparam T 返回值类型
     * @param sheet_name 工作表名称
     * @param row 行号
     * @param col 列号
     * @return 可选值，失败时返回std::nullopt
     */
    template<typename T>
    std::optional<T> tryGetValue(const std::string& sheet_name, const Address& address) const noexcept {
        try {
            auto worksheet = findSheet(sheet_name);
            if (!worksheet) {
                return std::nullopt;
            }
            return worksheet->tryGetValue<T>(address);
        } catch (...) {
            return std::nullopt;
        }
    }
    
    /**
     * @brief 安全获取单元格值（通过工作表索引）
     * @tparam T 返回值类型
     * @param sheet_index 工作表索引
     * @param row 行号
     * @param col 列号
     * @return 可选值，失败时返回std::nullopt
     */
    template<typename T>
    std::optional<T> tryGetValue(size_t sheet_index, const Address& address) const noexcept {
        try {
            if (sheet_index >= worksheets_.size()) {
                return std::nullopt;
            }
            auto worksheet = worksheets_[sheet_index];
            if (!worksheet) {
                return std::nullopt;
            }
            return worksheet->tryGetValue<T>(address);
        } catch (...) {
            return std::nullopt;
        }
    }
    
    /**
     * @brief 安全设置单元格值（通过工作表名称）
     * @tparam T 值类型
     * @param sheet_name 工作表名称
     * @param row 行号
     * @param col 列号
     * @param value 要设置的值
     * @return 是否成功设置
     */
    template<typename T>
    bool trySetValue(const std::string& sheet_name, const Address& address, const T& value) noexcept {
        try {
            auto worksheet = findSheet(sheet_name);
            if (!worksheet) {
                return false;
            }
            worksheet->setValue<T>(address, value);
            return true;
        } catch (...) {
            return false;
        }
    }
    
    /**
     * @brief 通过完整地址字符串访问单元格（支持跨工作表）
     * @tparam T 返回值类型
     * @param full_address 完整地址（如"Sheet1!A1", "Data!B2"）
     * @return 单元格值
     * 
     * @example
     * auto value = workbook.getValue<std::string>("Sheet1!A1");  // 获取Sheet1工作表A1的值
     * auto value = workbook.getValue<double>("Data!C3");         // 获取Data工作表C3的值
     */
    template<typename T>
    T getValue(const std::string& full_address) const {
        auto [sheet_name, row, col] = utils::AddressParser::parseAddress(full_address);
        if (sheet_name.empty()) {
            // 如果没有工作表名，使用第一个工作表
            return getValue<T>(0, Address(row, col));
        }
        
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        return worksheet->getValue<T>(Address(row, col));
    }
    
    /**
     * @brief 通过完整地址字符串设置单元格值（支持跨工作表）
     * @tparam T 值类型
     * @param full_address 完整地址（如"Sheet1!A1", "Data!B2"）
     * @param value 要设置的值
     * 
     * @example
     * workbook.setValue("Sheet1!A1", std::string("Hello"));  // 在Sheet1的A1设置字符串
     * workbook.setValue("Data!C3", 123.45);                  // 在Data工作表的C3设置数字
     */
    template<typename T>
    void setValue(const std::string& full_address, const T& value) {
        auto [sheet_name, row, col] = utils::AddressParser::parseAddress(full_address);
        if (sheet_name.empty()) {
            // 如果没有工作表名，使用第一个工作表
            setValue<T>(0, Address(row, col), value);
            return;
        }
        
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        worksheet->setValue<T>(Address(row, col), value);
    }
    
    // 样式管理
    
    /**
     * @brief 添加样式到工作簿
     * @param style 样式描述符
     * @return 样式ID
     */
    int addStyle(const FormatDescriptor& style);
    
    /**
     * @brief 添加样式到工作簿（使用Builder）
     * @param builder 样式构建器
     * @return 样式ID
     */
    int addStyle(const StyleBuilder& builder);
    
    /**
     * @brief 添加命名样式
     * @param named_style 命名样式
     * @return 样式ID
     */
    int addNamedStyle(const NamedStyle& named_style);
    
    /**
     * @brief 创建样式构建器
     * @return 样式构建器
     */
    StyleBuilder createStyleBuilder() const;
    
    /**
     * @brief 根据ID获取样式
     * @param style_id 样式ID
     * @return 样式描述符，如果ID无效则返回默认样式
     */
    std::shared_ptr<const FormatDescriptor> getStyle(int style_id) const;
    
    /**
     * @brief 获取默认样式ID
     * @return 默认样式ID
     */
    int getDefaultStyleId() const;
    
    /**
     * @brief 检查样式ID是否有效
     * @param style_id 样式ID
     * @return 是否有效
     */
    bool isValidStyleId(int style_id) const;
    
    /**
     * @brief 获取样式数量
     * @return 样式数量
     */
    size_t getStyleCount() const;
    
    /**
     * @brief 获取样式仓储（只读访问）
     * @return 样式仓储的常量引用
     */
    const FormatRepository& getStyles() const;
    
    /**
     * @brief 设置自定义主题XML
     * @param theme_xml 主题XML内容
     */
    void setThemeXML(const std::string& theme_xml);
    
    /**
     * @brief 设置原始主题XML（仅供读取器使用，保持保真）
     */
    void setOriginalThemeXML(const std::string& theme_xml);
    
    /**
     * @brief 设置主题（结构化对象）
     * @param theme 主题对象
     */
    void setTheme(const theme::Theme& theme);
    
    /**
     * @brief 获取当前主题对象（只读，可能为nullptr）
     */
    const theme::Theme* getTheme() const { return theme_.get(); }
    
    /**
     * @brief 设置主题名称
     */
    void setThemeName(const std::string& name);
    
    /**
     * @brief 通过类型设置主题颜色
     */
    void setThemeColor(theme::ThemeColorScheme::ColorType type, const core::Color& color);
    
    /**
     * @brief 通过名称设置主题颜色（如 "accent1"/"lt1"/"hlink" 等）
     * @return 是否设置成功
     */
    bool setThemeColorByName(const std::string& name, const core::Color& color);
    
    /**
     * @brief 设置主题的major字体族
     */
    void setThemeMajorFontLatin(const std::string& name);
    void setThemeMajorFontEastAsia(const std::string& name);
    void setThemeMajorFontComplex(const std::string& name);
    
    /**
     * @brief 设置主题的minor字体族
     */
    void setThemeMinorFontLatin(const std::string& name);
    void setThemeMinorFontEastAsia(const std::string& name);
    void setThemeMinorFontComplex(const std::string& name);
    
    /**
     * @brief 获取当前主题XML
     * @return 主题XML内容
     */
    const std::string& getThemeXML() const;
    
    /**
     * @brief 从另一个工作簿复制样式
     * @param source_workbook 源工作簿
     * @return 样式传输上下文（用于ID映射）
     */
    std::unique_ptr<StyleTransferContext> copyStylesFrom(
        const Workbook& source_workbook);
    
    /**
     * @brief 获取样式去重统计
     * @return 去重统计信息
     */
    FormatRepository::DeduplicationStats getStyleStats() const;
    
    // 文档属性
    
    /**
     * @brief 设置文档标题
     * @param title 标题
     */
    void setTitle(const std::string& title) { 
        if (document_manager_) document_manager_->setTitle(title); 
    }
    
    std::string getTitle() const {
        return document_manager_ ? document_manager_->getTitle() : std::string();
    }
    
    void setSubject(const std::string& subject) {
        if (document_manager_) document_manager_->setSubject(subject);
    }
    
    std::string getSubject() const {
        return document_manager_ ? document_manager_->getSubject() : std::string();
    }
    
    void setAuthor(const std::string& author) {
        if (document_manager_) document_manager_->setAuthor(author);
    }
    
    std::string getAuthor() const {
        return document_manager_ ? document_manager_->getAuthor() : std::string();
    }
    
    void setManager(const std::string& manager) {
        if (document_manager_) document_manager_->setManager(manager);
    }
    
    void setCompany(const std::string& company) {
        if (document_manager_) document_manager_->setCompany(company);
    }
    
    void setCategory(const std::string& category) {
        if (document_manager_) document_manager_->setCategory(category);
    }
    
    void setKeywords(const std::string& keywords) {
        if (document_manager_) document_manager_->setKeywords(keywords);
    }
    
    void setComments(const std::string& comments) {
        if (document_manager_) document_manager_->setComments(comments);
    }
    
    void setStatus(const std::string& status) {
        if (document_manager_) document_manager_->setStatus(status);
    }
    
    void setHyperlinkBase(const std::string& hyperlink_base) {
        if (document_manager_) document_manager_->setHyperlinkBase(hyperlink_base);
    }
    
    void setCreatedTime(const std::tm& created_time) {
        if (document_manager_) document_manager_->setCreatedTime(created_time);
    }
    
    void setModifiedTime(const std::tm& modified_time) {
        if (document_manager_) document_manager_->setModifiedTime(modified_time);
    }
    
    void setDocumentProperties(const std::string& title = "",
                              const std::string& subject = "",
                              const std::string& author = "",
                              const std::string& company = "",
                              const std::string& comments = "");
    
    void setApplication(const std::string& application);
    
    // 自定义属性
    
    /**
     * @brief 添加自定义属性（字符串）
     * @param name 属性名
     * @param value 属性值
     */
    void setProperty(const std::string& name, const std::string& value) {
        if (document_manager_) document_manager_->setCustomProperty(name, value);
    }
    
    /**
     * @brief 添加自定义属性（字符串字面量）- 优先匹配字符串字面量
     * @param name 属性名
     * @param value 属性值
     */
    void setProperty(const std::string& name, const char* value) {
        if (document_manager_) document_manager_->setCustomProperty(name, std::string(value));
    }
    
    /**
     * @brief 添加自定义属性（数字）
     * @param name 属性名
     * @param value 属性值
     */
    void setProperty(const std::string& name, double value) {
        if (document_manager_) document_manager_->setCustomProperty(name, value);
    }
    
    void setProperty(const std::string& name, bool value) {
        if (document_manager_) document_manager_->setCustomProperty(name, value);
    }
    
    std::string getProperty(const std::string& name) const {
        return document_manager_ ? document_manager_->getCustomProperty(name) : std::string();
    }
    
    bool removeProperty(const std::string& name) {
        return document_manager_ ? document_manager_->removeCustomProperty(name) : false;
    }
    
    std::unordered_map<std::string, std::string> getAllProperties() const {
        return document_manager_ ? document_manager_->getAllCustomProperties() : std::unordered_map<std::string, std::string>();
    }
    
    // 定义名称
    void defineName(const std::string& name, const std::string& formula, const std::string& scope = "") {
        if (document_manager_) document_manager_->defineName(name, formula, scope);
    }
    
    std::string getDefinedName(const std::string& name, const std::string& scope = "") const {
        return document_manager_ ? document_manager_->getDefinedName(name, scope) : std::string();
    }
    
    bool removeDefinedName(const std::string& name, const std::string& scope = "") {
        return document_manager_ ? document_manager_->removeDefinedName(name, scope) : false;
    }
    
    // VBA项目
    bool addVbaProject(const std::string& vba_project_path) {
        return security_manager_ ? security_manager_->addVbaProject(vba_project_path) : false;
    }
    
    bool hasVbaProject() const {
        return security_manager_ ? security_manager_->hasVbaProject() : false;
    }
    
    // 工作簿保护
    void protect(const std::string& password = "", bool lock_structure = true, bool lock_windows = false) {
        if (security_manager_) {
            WorkbookSecurityManager::ProtectionOptions options;
            options.password = password;
            options.lock_structure = lock_structure;
            options.lock_windows = lock_windows;
            security_manager_->protect(options);
        }
    }
    
    void unprotect() {
        if (security_manager_) security_manager_->unprotect();
    }
    
    bool isProtected() const {
        return security_manager_ ? security_manager_->isProtected() : false;
    }
    
    // 工作簿选项
    
    /**
     * @brief 设置常量内存模式
     * @param constant_memory 是否启用
     */
    void setConstantMemory(bool constant_memory) { options_.constant_memory = constant_memory; }
    
    /**
     * @brief 设置临时目录
     * @param tmpdir 临时目录路径
     */
    void setTmpDir(const std::string& tmpdir) { options_.tmpdir = tmpdir; }
    
    /**
     * @brief 设置建议只读
     * @param read_only_recommended 是否建议只读
     */
    void setReadOnlyRecommended(bool read_only_recommended) { options_.read_only_recommended = read_only_recommended; }
    
    /**
     * @brief 设置计算选项
     * @param calc_on_load 加载时计算
     * @param full_calc_on_load 加载时完全计算
     */
    void setCalcOptions(bool calc_on_load, bool full_calc_on_load = false);
    
    void setUseSharedStrings(bool enable) { options_.use_shared_strings = enable; }
    void setMode(WorkbookMode mode) {
        options_.mode = mode;
    }
    WorkbookMode getMode() const { return options_.mode; }
    void setAutoModeThresholds(size_t cell_threshold, size_t memory_threshold) {
        options_.auto_mode_cell_threshold = cell_threshold;
        options_.auto_mode_memory_threshold = memory_threshold;
    }
    void setRowBufferSize(size_t size) { options_.row_buffer_size = size; }
    void setCompressionLevel(int level) { options_.compression_level = level; }
    void setXMLBufferSize(size_t size) { options_.xml_buffer_size = size; }
    void setHighPerformanceMode(bool enable);
    
    // 获取状态

    // 管理器访问器（获取管理器实例）
    WorkbookDocumentManager* getDocumentManager() { return document_manager_.get(); }
    const WorkbookDocumentManager* getDocumentManager() const { return document_manager_.get(); }
    
    WorkbookSecurityManager* getSecurityManager() { return security_manager_.get(); }
    const WorkbookSecurityManager* getSecurityManager() const { return security_manager_.get(); }
    
    WorkbookDataManager* getDataManager() { return data_manager_.get(); }
    const WorkbookDataManager* getDataManager() const { return data_manager_.get(); }
    
    // 脏数据管理器访问
    DirtyManager* getDirtyManager() { return document_manager_ ? document_manager_->getDirtyManager() : nullptr; }
    const DirtyManager* getDirtyManager() const { return document_manager_ ? document_manager_->getDirtyManager() : nullptr; }
    
    // 生成控制（基于DirtyManager的新实现）
    bool shouldGenerateContentTypes() const;
    bool shouldGenerateRootRels() const;
    bool shouldGenerateWorkbookCore() const;
    bool shouldGenerateStyles() const;
    bool shouldGenerateTheme() const;
    bool shouldGenerateSharedStrings() const;
    bool shouldGenerateDocPropsCore() const;
    bool shouldGenerateDocPropsApp() const;
    bool shouldGenerateDocPropsCustom() const;
    bool shouldGenerateSheet(size_t index) const;
    bool shouldGenerateSheetRels(size_t index) const;
    
    bool isReadOnly() const { return state_ == WorkbookState::READING; }
    bool isEditable() const { return state_ == WorkbookState::EDITING || state_ == WorkbookState::CREATING; }
    const std::string& getFilename() const { return filename_; }
    const DocumentProperties& getDocumentProperties() const {
        static DocumentProperties empty_properties;
        return document_manager_ ? document_manager_->getDocumentProperties() : empty_properties;
    }

    bool hasCustomProperties() const {
        return document_manager_ && (document_manager_->getCustomPropertyCount() > 0);
    }

    std::unordered_map<std::string, std::string> getAllCustomProperties() const {
        if (!document_manager_) return {};
        return document_manager_->getAllCustomProperties();
    }
    
    WorkbookOptions& getOptions() { return options_; }
    const WorkbookOptions& getOptions() const { return options_; }
    
    // 共享字符串管理
    int addSharedString(const std::string& str);
    int addSharedStringWithIndex(const std::string& str, int original_index);
    int getSharedStringIndex(const std::string& str) const;
    const SharedStringTable* getSharedStrings() const;
    SharedStringTable* getSharedStringTable() { return shared_string_table_.get(); }
    const SharedStringTable* getSharedStringTable() const { return shared_string_table_.get(); }
    
    bool isEditMode() const { return state_ == WorkbookState::EDITING; }
    const std::string& getOriginalPackagePath() const { return original_package_path_; }
    size_t getEstimatedSize() const;

    
    // 工作簿编辑功能
    
    /**
     * @brief 打开现有文件进行编辑（直接可用，无需再调用open）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> open(const Path& path);
    
    /**
     * @brief 打开现有文件进行编辑（字符串重载版本）
     * @param filepath 文件路径字符串
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> open(const std::string& filepath);
    
    /**
     * @brief 刷新工作簿（重新读取文件内容）
     * @return 是否成功
     */
    bool refresh();
    
    /**
     * @brief 合并另一个工作簿的内容
     * @param other_workbook 其他工作簿
     * @param merge_options 合并选项
     * @return 是否成功
     */
    struct MergeOptions {
        bool merge_worksheets = true;      // 合并工作表
        bool merge_formats = true;         // 合并格式
        bool merge_properties = false;     // 合并文档属性
        bool overwrite_existing = false;   // 覆盖现有内容
        std::string name_prefix;           // 工作表名称前缀
    };
    bool mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, const MergeOptions& options);
    // 便捷重载：默认选项
    bool mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook);
    
    /**
     * @brief 导出工作表到新工作簿
     * @param worksheet_names 要导出的工作表名称列表
     * @param output_filename 输出文件名
     * @return 是否成功
     */
    bool exportWorksheets(const std::vector<std::string>& worksheet_names, const std::string& output_filename);
    
    /**
     * @brief 批量重命名工作表
     * @param rename_map 重命名映射 (旧名称 -> 新名称)
     * @return 成功重命名的数量
     */
    int batchRenameWorksheets(const std::unordered_map<std::string, std::string>& rename_map);
    
    /**
     * @brief 批量删除工作表
     * @param worksheet_names 要删除的工作表名称列表
     * @return 成功删除的数量
     */
    int batchRemoveWorksheets(const std::vector<std::string>& worksheet_names);
    
    /**
     * @brief 重新排序工作表
     * @param new_order 新的工作表顺序（工作表名称列表）
     * @return 是否成功
     */
    bool reorderWorksheets(const std::vector<std::string>& new_order);
    
    /**
     * @brief 查找并替换（全工作簿）
     * @param find_text 查找的文本
     * @param replace_text 替换的文本
     * @param options 查找替换选项
     * @return 替换的总数量
     */
    struct FindReplaceOptions {
        bool match_case = false;
        bool match_entire_cell = false;
        std::vector<std::string> worksheet_filter; // 限制在特定工作表中查找
    };
    int findAndReplaceAll(const std::string& find_text, const std::string& replace_text,
                         const FindReplaceOptions& options);
    // 便捷重载：默认选项
    int findAndReplaceAll(const std::string& find_text, const std::string& replace_text);
    
    /**
     * @brief 全局查找
     * @param search_text 搜索文本
     * @param options 查找选项
     * @return 匹配结果列表 (工作表名, 行, 列)
     */
    std::vector<std::tuple<std::string, int, int>> findAll(const std::string& search_text,
                                                           const FindReplaceOptions& options);
    // 便捷重载：默认选项
    std::vector<std::tuple<std::string, int, int>> findAll(const std::string& search_text);
    
    /**
     * @brief 获取工作簿统计信息
     */
    struct WorkbookStats {
        size_t total_worksheets = 0;
        size_t total_cells = 0;
        size_t total_formats = 0;
        size_t memory_usage = 0;
        std::unordered_map<std::string, size_t> worksheet_cell_counts;
    };
    WorkbookStats getStatistics() const;
    
    /**
     * @brief 检查工作簿是否已修改（新API）
     * @return 是否已修改
     */
    bool isModified() const;
    
    /**
     * @brief 获取内存使用总量（新API）
     * @return 内存使用字节数
     */
    size_t getTotalMemoryUsage() const;
    
    /**
     * @brief 优化工作簿（压缩样式、清理未使用资源，新API）
     * @return 优化的项目数
     */
    size_t optimize();
    
    /**
     * @brief 内存池优化的Cell创建
     */
    template<typename... Args>
    ::fastexcel::pool_ptr<Cell> createOptimizedCell(Args&&... args) {
        if (!memory_manager_) {
            memory_manager_ = std::make_unique<memory::WorkbookMemoryManager>();
        }
        
        return memory_manager_->createOptimizedCell(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 内存池优化的FormatDescriptor创建
     */
    template<typename... Args>
    ::fastexcel::pool_ptr<FormatDescriptor> createOptimizedFormat(Args&&... args) {
        if (!memory_manager_) {
            memory_manager_ = std::make_unique<memory::WorkbookMemoryManager>();
        }
        
        return memory_manager_->createOptimizedFormat(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 创建默认格式的FormatDescriptor对象
     */
    ::fastexcel::pool_ptr<FormatDescriptor> createDefaultFormat() {
        if (!memory_manager_) {
            memory_manager_ = std::make_unique<memory::WorkbookMemoryManager>();
        }
        
        return memory_manager_->createDefaultFormat();
    }
    
    /**
     * @brief 基础单元格值设置方法
     */
    void setCellValue(int row, int col, const std::string& value) {
        // 确保至少存在一个工作表
        size_t sheet_count = worksheet_manager_ ? worksheet_manager_->count() : worksheets_.size();
        if (sheet_count == 0) {
            addSheet("Sheet1");
        }
        // 获取第一个工作表
        auto sheet = worksheet_manager_ ? worksheet_manager_->getByIndex(0)
                                        : (worksheets_.empty() ? nullptr : worksheets_[0]);
        if (sheet) {
            sheet->setCellValue(row, col, value);
        }
    }
    
    void setCellValue(int row, int col, double value) {
        size_t sheet_count = worksheet_manager_ ? worksheet_manager_->count() : worksheets_.size();
        if (sheet_count == 0) {
            addSheet("Sheet1");
        }
        auto sheet = worksheet_manager_ ? worksheet_manager_->getByIndex(0)
                                        : (worksheets_.empty() ? nullptr : worksheets_[0]);
        if (sheet) {
            sheet->setCellValue(row, col, value);
        }
    }
    
    void setCellValue(int row, int col, int value) {
        setCellValue(row, col, static_cast<double>(value));
    }
    
    void setCellValue(int row, int col, bool value) {
        size_t sheet_count = worksheet_manager_ ? worksheet_manager_->count() : worksheets_.size();
        if (sheet_count == 0) {
            addSheet("Sheet1");
        }
        auto sheet = worksheet_manager_ ? worksheet_manager_->getByIndex(0)
                                        : (worksheets_.empty() ? nullptr : worksheets_[0]);
        if (sheet) {
            sheet->setCellValue(row, col, value);
        }
    }
    
    // 已移除：显式“优化版”接口；优化已深度集成到 setCellValue
    
    /**
     * @brief 使用StringJoiner优化的复合值设置
     */
    void setCellComplexValue(int row, int col, 
                           const std::vector<std::string_view>& parts,
                           std::string_view separator = " ") {
        std::string result;
        bool first = true;
        for (const auto& part : parts) {
            if (!first) {
                result += separator;
            }
            result += part;
            first = false;
        }
        setCellValue(row, col, result);
    }
    
    /**
     * @brief 使用StringBuilder优化的格式化值设置
     */
    template<typename... Args>
    void setCellFormattedValue(int row, int col, const char* format, Args&&... args) {
        std::string formatted = utils::StringViewOptimized::format(format, std::forward<Args>(args)...);
        setCellValue(row, col, formatted);
    }
    
    // 已移除：Workbook 级别内存使用统计接口（统一使用内存池统计）
    
    /**
     * @brief 内存收缩（释放未使用的内存）
     */
    void shrinkMemory() {
        if (memory_manager_) {
            memory_manager_->shrinkAll();
            FASTEXCEL_LOG_DEBUG("Memory shrinking completed for Workbook");
        }
    }

private:
    // 文档属性通过 WorkbookDocumentManager 管理
    // 内部方法
    
    /**
     * @brief 内部方法：打开工作簿文件管理器
     * @return 是否成功
     */
    bool open();
    
    // 生成Excel文件结构
    bool generateExcelStructure();
    bool generateWithGenerator(bool use_streaming_writer);

    // DirtyManager 访问（由 WorkbookDocumentManager 持有）
public:
    // 注意：避免重复声明 getDirtyManager（MSVC C2535）
    
    
    // 辅助函数
    std::string generateUniqueSheetName(const std::string& base_name) const;
    bool validateSheetName(const std::string& name) const;
    void collectSharedStrings();
    
    // 访问模式检查辅助方法
    void ensureEditable(const std::string& operation = "") const;
    void ensureReadable(const std::string& operation = "") const;
    
    // 文件路径生成
    std::string getWorksheetPath(int sheet_id) const;
    std::string getWorksheetRelPath(int sheet_id) const;
    
    
    
    // 智能模式选择辅助方法
    size_t estimateMemoryUsage() const;
    size_t getTotalCellCount() const;

    // 状态验证和转换辅助方法
    /**
     * @brief 检查当前状态是否允许指定操作
     * @param required_state 要求的最低状态
     * @return 是否允许操作
     */
    bool isStateValid(WorkbookState required_state) const;
    
    /**
     * @brief 获取当前工作簿状态
     * @return 当前状态
     */
    WorkbookState getCurrentState() const { return state_; }
    
    /**
     * @brief 获取文件来源类型
     * @return 文件来源
     */
    FileSource getFileSource() const { return file_source_; }
    
    /**
     * @brief 状态转换方法
     * @param new_state 新状态
     * @param reason 转换原因（用于日志）
     */
    void transitionToState(WorkbookState new_state, const std::string& reason = "");
    
    // 内部：根据编辑/透传状态返回"是否在编辑模式下且启用透传"
    bool isPassThroughEditMode() const { return file_source_ == FileSource::EXISTING_FILE && preserve_unknown_parts_; }

};

}} // namespace fastexcel::core
