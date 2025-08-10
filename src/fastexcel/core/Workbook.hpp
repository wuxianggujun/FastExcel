#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/WorkbookModeSelector.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/CustomPropertyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "FormatDescriptor.hpp"
#include "FormatRepository.hpp"
#include "StyleTransferContext.hpp"
#include "StyleBuilder.hpp"
#include "SharedStringTable.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <unordered_map>
#include <set>
#include <ctime>
#include <functional>

namespace fastexcel {
namespace opc {
    class PackageEditor;  // 前向声明PackageEditor
}

namespace core {

// 前向声明
class NamedStyle;
class ExcelStructureGenerator;

// 工作簿状态枚举 - 统一的状态管理
enum class WorkbookState {
    CLOSED,      // 未打开状态
    CREATING,    // 正在创建新文件
    READING,     // 只读模式打开
    EDITING      // 编辑模式打开
};

// 文件来源类型
enum class FileSource {
    NEW_FILE,        // 全新创建的文件
    EXISTING_FILE    // 从现有文件加载
};

// 文档属性结构
struct DocumentProperties {
    std::string title;
    std::string subject;
    std::string author;
    std::string manager;
    std::string company;
    std::string category;
    std::string keywords;
    std::string comments;
    std::string status;
    std::string hyperlink_base;
    std::tm created_time;
    std::tm modified_time;
    
    DocumentProperties();
};

// 工作簿选项
struct WorkbookOptions {
    bool constant_memory = false;      // 常量内存模式
    bool use_zip64 = false;           // 使用ZIP64格式
    std::string tmpdir;               // 临时目录
    bool optimize_for_speed = false;  // 速度优化
    bool read_only_recommended = false; // 建议只读
    
    // 计算选项
    bool calc_on_load = true;         // 加载时计算
    bool full_calc_on_load = false;   // 加载时完全计算
    
    // 安全选项
    std::string password;             // 工作簿密码
    bool encrypt_metadata = false;    // 加密元数据
    
    // 性能优化选项
    bool use_shared_strings = true;   // 使用共享字符串（默认启用以匹配Excel格式）
    WorkbookMode mode = WorkbookMode::AUTO;  // 工作簿模式（默认自动选择）
    size_t row_buffer_size = 5000;    // 行缓冲大小（默认较大缓冲）
    int compression_level = 6;        // ZIP压缩级别（默认平衡压缩）
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小（4MB）
    
    // 自动模式阈值
    size_t auto_mode_cell_threshold = 1000000;     // 100万单元格
    size_t auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
};

/**
 * @brief Workbook类 - Excel工作簿（新架构）
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
private:
    std::string filename_;
    std::vector<std::shared_ptr<Worksheet>> worksheets_;
    std::unique_ptr<archive::FileManager> file_manager_;
    
    // 格式管理 - 新样式架构
    std::unique_ptr<FormatRepository> format_repo_;
    
    // 主题管理
    std::string theme_xml_; // 自定义主题XML内容（编辑或外部设置）
    std::string theme_xml_original_; // 从文件读取的原始主题XML（用于未编辑时的保真写回）
    bool theme_dirty_ = false; // 主题是否被编辑
    std::unique_ptr<theme::Theme> theme_; // 结构化主题对象（优先用于生成）
    
    // ID管理
    int next_sheet_id_ = 1;
    
    // 🔧 统一的状态管理系统（重构后）
    WorkbookState state_ = WorkbookState::CLOSED;        // 当前工作簿状态
    FileSource file_source_ = FileSource::NEW_FILE;     // 文件来源类型
    std::string original_package_path_;                  // 原始文件路径（用于保留未修改部件）
    
    // 文档属性
    DocumentProperties doc_properties_;
    std::unique_ptr<CustomPropertyManager> custom_property_manager_;
    
    // 定义名称管理
    std::unique_ptr<DefinedNameManager> defined_name_manager_;
    
    // 工作簿选项
    WorkbookOptions options_;
    
    // 共享字符串表
    std::unique_ptr<SharedStringTable> shared_string_table_;
    
    // VBA项目
    std::string vba_project_path_;
    bool has_vba_ = false;
    
    // 工作簿保护
    bool protected_ = false;
    std::string protection_password_;
    bool lock_structure_ = false;
    bool lock_windows_ = false;

    // 新的智能脏数据管理器
    std::unique_ptr<DirtyManager> dirty_manager_;
    
    // 工作簿选项（包含保留未修改部件的设置）
    bool preserve_unknown_parts_ = true; // 保留未修改的Excel部件（如绘图、打印设置等）

public:
    /**
     * @brief 创建工作簿（直接可用，无需再调用open）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> create(const Path& path);
    
    /**
     * @brief 只读方式打开Excel文件（新API - 推荐）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     * 
     * 特点：
     * - 轻量级：内存占用小，加载速度快
     * - 安全：编译期和运行期防止意外修改
     * - 高性能：针对只读场景优化，支持懒加载
     * 
     * 使用场景：
     * - 数据分析和统计
     * - 大文件处理
     * - 模板数据提取
     * - 数据转换和导入
     */
    static std::unique_ptr<Workbook> openForReading(const Path& path);
    
    /**
     * @brief 编辑方式打开Excel文件（新API - 推荐）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     * 
     * 特点：
     * - 完整功能：支持所有编辑操作
     * - 变更追踪：精确跟踪修改状态
     * - 格式支持：完整的样式和格式处理
     * 
     * 使用场景：
     * - 修改现有Excel文件
     * - 复杂的格式设置
     * - 需要保存更改的场景
     */
    static std::unique_ptr<Workbook> openForEditing(const Path& path);
    

    
    /**
     * @brief 构造函数
     * @param path 文件路径
     */
    explicit Workbook(const Path& path);
    
    /**
     * @brief 析构函数
     */
    ~Workbook();
    
    // 禁用拷贝构造和赋值
    Workbook(const Workbook&) = delete;
    Workbook& operator=(const Workbook&) = delete;
    
    // 允许移动构造和赋值
    Workbook(Workbook&&) = default;
    Workbook& operator=(Workbook&&) = default;
    
    // ========== 文件操作 ==========
    
    /**
     * @brief 保存工作簿
     * @return 是否成功
     */
    bool save();
    
    /**
     * @brief 另存为
     * @param filename 新文件名
     * @return 是否成功
     */
    bool saveAs(const std::string& filename);
    
    /**
     * @brief 检查工作簿是否已打开
     * @return 是否已打开（处于可用状态）
     */
    bool isOpen() const;
    
    /**
     * @brief 关闭工作簿
     * @return 是否成功
     */
    bool close();
    
    // ========== 编辑模式/保真写回配置 ==========
    void setPreserveUnknownParts(bool enable) { preserve_unknown_parts_ = enable; }
    bool getPreserveUnknownParts() const { return preserve_unknown_parts_; }

    // ========== 工作表管理 ==========
    
    /**
     * @brief 添加工作表
     * @param name 工作表名称（空则自动生成）
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> addWorksheet(const std::string& name = "");
    
    /**
     * @brief 插入工作表
     * @param index 插入位置
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> insertWorksheet(size_t index, const std::string& name = "");
    
    /**
     * @brief 删除工作表
     * @param name 工作表名称
     * @return 是否成功
     */
    bool removeWorksheet(const std::string& name);
    
    /**
     * @brief 删除工作表
     * @param index 工作表索引
     * @return 是否成功
     */
    bool removeWorksheet(size_t index);
    
    /**
     * @brief 获取工作表（按名称）
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> getWorksheet(const std::string& name);
    
    /**
     * @brief 获取工作表（按索引）
     * @param index 工作表索引
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> getWorksheet(size_t index);
    
    /**
     * @brief 获取工作表（按名称，只读）
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> getWorksheet(const std::string& name) const;
    
    /**
     * @brief 获取工作表（按索引，只读）
     * @param index 工作表索引
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> getWorksheet(size_t index) const;
    
    /**
     * @brief 获取工作表数量
     * @return 工作表数量
     */
    size_t getWorksheetCount() const { return worksheets_.size(); }
    
    /**
     * @brief 获取所有工作表名称
     * @return 工作表名称列表
     */
    std::vector<std::string> getWorksheetNames() const;
    
    /**
     * @brief 重命名工作表
     * @param old_name 旧名称
     * @param new_name 新名称
     * @return 是否成功
     */
    bool renameWorksheet(const std::string& old_name, const std::string& new_name);
    
    /**
     * @brief 移动工作表
     * @param from_index 源位置
     * @param to_index 目标位置
     * @return 是否成功
     */
    bool moveWorksheet(size_t from_index, size_t to_index);
    
    /**
     * @brief 复制工作表
     * @param source_name 源工作表名称
     * @param new_name 新工作表名称
     * @return 新工作表指针
     */
    std::shared_ptr<Worksheet> copyWorksheet(const std::string& source_name, const std::string& new_name);
    
    /**
     * @brief 复制工作表从另一个工作簿
     * @param source_worksheet 源工作表
     * @param new_name 新工作表名称（空则使用源名称）
     * @return 新创建的工作表指针
     */
    std::shared_ptr<Worksheet> copyWorksheetFrom(const std::shared_ptr<const Worksheet>& source_worksheet, 
                                const std::string& new_name = "");
    
    /**
     * @brief 设置活动工作表
     * @param index 工作表索引
     */
    void setActiveWorksheet(size_t index);
    
    // ========== 样式管理 ==========
    
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
    const FormatRepository& getStyleRepository() const;
    
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
    
    // ========== 文档属性 ==========
    
    /**
     * @brief 设置文档标题
     * @param title 标题
     */
    void setTitle(const std::string& title) { doc_properties_.title = title; }
    
    /**
     * @brief 获取文档标题
     * @return 标题
     */
    const std::string& getTitle() const { return doc_properties_.title; }
    
    /**
     * @brief 设置文档主题
     * @param subject 主题
     */
    void setSubject(const std::string& subject) { doc_properties_.subject = subject; }
    
    /**
     * @brief 获取文档主题
     * @return 主题
     */
    const std::string& getSubject() const { return doc_properties_.subject; }
    
    /**
     * @brief 设置文档作者
     * @param author 作者
     */
    void setAuthor(const std::string& author) { doc_properties_.author = author; }
    
    /**
     * @brief 获取文档作者
     * @return 作者
     */
    const std::string& getAuthor() const { return doc_properties_.author; }
    
    /**
     * @brief 设置文档管理者
     * @param manager 管理者
     */
    void setManager(const std::string& manager) { doc_properties_.manager = manager; }
    
    /**
     * @brief 设置公司
     * @param company 公司
     */
    void setCompany(const std::string& company) { doc_properties_.company = company; }
    
    /**
     * @brief 设置类别
     * @param category 类别
     */
    void setCategory(const std::string& category) { doc_properties_.category = category; }
    
    /**
     * @brief 设置关键词
     * @param keywords 关键词
     */
    void setKeywords(const std::string& keywords) { doc_properties_.keywords = keywords; }
    
    /**
     * @brief 设置注释
     * @param comments 注释
     */
    void setComments(const std::string& comments) { doc_properties_.comments = comments; }
    
    /**
     * @brief 设置状态
     * @param status 状态
     */
    void setStatus(const std::string& status) { doc_properties_.status = status; }
    
    /**
     * @brief 设置超链接基础
     * @param hyperlink_base 超链接基础
     */
    void setHyperlinkBase(const std::string& hyperlink_base) { doc_properties_.hyperlink_base = hyperlink_base; }
    
    /**
     * @brief 设置创建时间
     * @param created_time 创建时间
     */
    void setCreatedTime(const std::tm& created_time) { doc_properties_.created_time = created_time; }
    
    /**
     * @brief 设置修改时间
     * @param modified_time 修改时间
     */
    void setModifiedTime(const std::tm& modified_time) { doc_properties_.modified_time = modified_time; }
    
    /**
     * @brief 批量设置文档属性（新API）
     * @param title 标题
     * @param subject 主题
     * @param author 作者
     * @param company 公司
     * @param comments 注释
     */
    void setDocumentProperties(const std::string& title = "",
                              const std::string& subject = "",
                              const std::string& author = "",
                              const std::string& company = "",
                              const std::string& comments = "");
    
    /**
     * @brief 设置应用程序名称（新API）
     * @param application 应用程序名称
     */
    void setApplication(const std::string& application);
    
    // ========== 自定义属性 ==========
    
    /**
     * @brief 添加自定义属性（字符串）
     * @param name 属性名
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, const std::string& value);
    
    /**
     * @brief 添加自定义属性（数字）
     * @param name 属性名
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, double value);
    
    /**
     * @brief 添加自定义属性（布尔）
     * @param name 属性名
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, bool value);
    
    /**
     * @brief 获取自定义属性
     * @param name 属性名
     * @return 属性值（如果不存在返回空字符串）
     */
    std::string getCustomProperty(const std::string& name) const;
    
    /**
     * @brief 删除自定义属性
     * @param name 属性名
     * @return 是否成功
     */
    bool removeCustomProperty(const std::string& name);
    
    /**
     * @brief 获取所有自定义属性
     * @return 自定义属性映射 (名称 -> 值)
     */
    std::unordered_map<std::string, std::string> getCustomProperties() const;
    
    // ========== 定义名称 ==========
    
    /**
     * @brief 定义名称
     * @param name 名称
     * @param formula 公式
     * @param scope 作用域（工作表名或空表示全局）
     */
    void defineName(const std::string& name, const std::string& formula, const std::string& scope = "");
    
    /**
     * @brief 获取定义名称的公式
     * @param name 名称
     * @param scope 作用域
     * @return 公式（如果不存在返回空字符串）
     */
    std::string getDefinedName(const std::string& name, const std::string& scope = "") const;
    
    /**
     * @brief 删除定义名称
     * @param name 名称
     * @param scope 作用域
     * @return 是否成功
     */
    bool removeDefinedName(const std::string& name, const std::string& scope = "");
    
    // ========== VBA项目 ==========
    
    /**
     * @brief 添加VBA项目
     * @param vba_project_path VBA项目文件路径
     * @return 是否成功
     */
    bool addVbaProject(const std::string& vba_project_path);
    
    /**
     * @brief 检查是否有VBA项目
     * @return 是否有VBA项目
     */
    bool hasVbaProject() const { return has_vba_; }
    
    // ========== 工作簿保护 ==========
    
    /**
     * @brief 保护工作簿
     * @param password 密码（可选）
     * @param lock_structure 锁定结构
     * @param lock_windows 锁定窗口
     */
    void protect(const std::string& password = "", bool lock_structure = true, bool lock_windows = false);
    
    /**
     * @brief 取消保护
     */
    void unprotect();
    
    /**
     * @brief 检查是否受保护
     * @return 是否受保护
     */
    bool isProtected() const { return protected_; }
    
    // ========== 工作簿选项 ==========
    
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
    
    /**
     * @brief 启用/禁用共享字符串
     * @param enable 是否启用共享字符串
     */
    void setUseSharedStrings(bool enable) { options_.use_shared_strings = enable; }
    
    /**
     * @brief 设置工作簿模式
     * @param mode 工作簿模式（AUTO/BATCH/STREAMING）
     */
    void setMode(WorkbookMode mode) {
        options_.mode = mode;
    }
    
    /**
     * @brief 获取当前工作簿模式
     * @return 当前模式
     */
    WorkbookMode getMode() const { return options_.mode; }
    
    /**
     * @brief 设置自动模式阈值
     * @param cell_threshold 单元格数量阈值
     * @param memory_threshold 内存使用阈值（字节）
     */
    void setAutoModeThresholds(size_t cell_threshold, size_t memory_threshold) {
        options_.auto_mode_cell_threshold = cell_threshold;
        options_.auto_mode_memory_threshold = memory_threshold;
    }
    
    /**
     * @brief 设置行缓冲大小
     * @param size 缓冲大小
     */
    void setRowBufferSize(size_t size) { options_.row_buffer_size = size; }
    
    /**
     * @brief 设置ZIP压缩级别
     * @param level 压缩级别（0-9）
     */
    void setCompressionLevel(int level) { options_.compression_level = level; }
    
    /**
     * @brief 设置XML缓冲区大小
     * @param size 缓冲区大小（字节）
     */
    void setXMLBufferSize(size_t size) { options_.xml_buffer_size = size; }
    
    /**
     * @brief 启用高性能模式（自动配置最佳性能参数）
     * @param enable 是否启用
     */
    void setHighPerformanceMode(bool enable);
    
    // ========== 获取状态 ==========

    // 获取脏数据管理器
    DirtyManager* getDirtyManager() { return dirty_manager_.get(); }
    const DirtyManager* getDirtyManager() const { return dirty_manager_.get(); }
    
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
    
    /**
     * @brief 检查是否只读模式
     * @return 是否只读
     */
    bool isReadOnly() const { return state_ == WorkbookState::READING; }
    
    /**
     * @brief 检查是否编辑模式
     * @return 是否可编辑
     */
    bool isEditable() const { return state_ == WorkbookState::EDITING || state_ == WorkbookState::CREATING; }
    
    /**
     * @brief 获取文件名
     * @return 文件名
     */
    const std::string& getFilename() const { return filename_; }
    
    /**
     * @brief 获取文档属性
     * @return 文档属性
     */
    const DocumentProperties& getDocumentProperties() const { return doc_properties_; }
    
    /**
     * @brief 获取工作簿选项
     * @return 工作簿选项引用
     */
    WorkbookOptions& getOptions() { return options_; }
    
    /**
     * @brief 获取工作簿选项（只读）
     * @return 工作簿选项引用
     */
    const WorkbookOptions& getOptions() const { return options_; }
    
    // ========== 共享字符串管理 ==========
    
    /**
     * @brief 添加共享字符串
     * @param str 字符串
     * @return 字符串索引
     */
    int addSharedString(const std::string& str);

    /**
     * @brief 添加共享字符串并保持原始索引（用于文件复制）
     * @param str 字符串
     * @param original_index 原始文件中的索引
     * @return 实际使用的索引
     */
    int addSharedStringWithIndex(const std::string& str, int original_index);
    
    /**
     * @brief 获取共享字符串索引
     * @param str 字符串
     * @return 索引（如果不存在返回-1）
     */
    int getSharedStringIndex(const std::string& str) const;
    
    /**
     * @brief 获取共享字符串表
     * @return 共享字符串表指针（可能为nullptr）
     */
    const SharedStringTable* getSharedStringTable() const;

    
    // ========== 工作簿编辑功能 ==========
    
    /**
     * @brief 打开现有文件进行编辑（直接可用，无需再调用open）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> open(const Path& path);
    
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
    bool mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, const MergeOptions& options = {});
    
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
                         const FindReplaceOptions& options = {});
    
    /**
     * @brief 全局查找
     * @param search_text 搜索文本
     * @param options 查找选项
     * @return 匹配结果列表 (工作表名, 行, 列)
     */
    std::vector<std::tuple<std::string, int, int>> findAll(const std::string& search_text,
                                                           const FindReplaceOptions& options = {});
    
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

private:
    // ========== 内部方法 ==========
    
    /**
     * @brief 内部方法：打开工作簿文件管理器
     * @return 是否成功
     */
    bool open();
    
    // 生成Excel文件结构
    bool generateExcelStructure();
    bool generateWithGenerator(bool use_streaming_writer);
    
    
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

    // 🔧 状态验证和转换辅助方法
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
