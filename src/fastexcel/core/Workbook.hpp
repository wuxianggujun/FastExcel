#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Format.hpp"
#include "fastexcel/core/FormatPool.hpp"
#include "fastexcel/core/WorkbookModeSelector.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/utils/CommonUtils.hpp"
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <unordered_map>
#include <set>
#include <ctime>
#include <functional>

namespace fastexcel {
namespace core {

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

// 自定义属性
struct CustomProperty {
    std::string name;
    std::string value;
    enum Type { String, Number, Date, Boolean } type;
    
    CustomProperty(const std::string& n, const std::string& v) 
        : name(n), value(v), type(String) {}
    CustomProperty(const std::string& n, double v) 
        : name(n), value(std::to_string(v)), type(Number) {}
    CustomProperty(const std::string& n, bool v) 
        : name(n), value(v ? "true" : "false"), type(Boolean) {}
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
    bool streaming_xml = true;        // 流式XML写入（已废弃，使用mode代替）
    size_t row_buffer_size = 5000;    // 行缓冲大小（默认较大缓冲）
    int compression_level = 1;        // ZIP压缩级别（默认轻度压缩，避免STORE模式）
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小（4MB）
    
    // 自动模式阈值
    size_t auto_mode_cell_threshold = 1000000;     // 100万单元格
    size_t auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
};

// 定义名称
struct DefinedName {
    std::string name;
    std::string formula;
    std::string scope;  // 作用域（工作表名或空表示全局）
    bool hidden = false;
    
    DefinedName(const std::string& n, const std::string& f, const std::string& s = "")
        : name(n), formula(f), scope(s) {}
};

/**
 * @brief Workbook类 - Excel工作簿
 * 
 * 提供完整的Excel工作簿功能，包括：
 * - 工作表管理
 * - 格式管理
 * - 文档属性设置
 * - 自定义属性
 * - 定义名称
 * - VBA项目支持
 * - 工作簿保护
 * - 多种保存选项
 */
class Workbook {
private:
    std::string filename_;
    std::vector<std::shared_ptr<Worksheet>> worksheets_;
    std::unique_ptr<archive::FileManager> file_manager_;
    
    // 格式管理 - 使用FormatPool
    std::unique_ptr<FormatPool> format_pool_;
    
    // ID管理
    int next_sheet_id_ = 1;
    
    // 状态
    bool is_open_ = false;
    bool read_only_ = false;
    
    // 文档属性
    DocumentProperties doc_properties_;
    std::vector<CustomProperty> custom_properties_;
    
    // 定义名称
    std::vector<DefinedName> defined_names_;
    
    // 工作簿选项
    WorkbookOptions options_;
    
    // 共享字符串
    std::unordered_map<std::string, int> shared_strings_;
    std::vector<std::string> shared_strings_list_;
    
    // VBA项目
    std::string vba_project_path_;
    bool has_vba_ = false;
    
    // 工作簿保护
    bool protected_ = false;
    std::string protection_password_;
    bool lock_structure_ = false;
    bool lock_windows_ = false;

public:
    /**
     * @brief 创建工作簿
     * @param path 文件路径
     * @return 工作簿智能指针
     */
    static std::unique_ptr<Workbook> create(const Path& path);
    
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
     * @brief 打开工作簿进行编辑
     * @return 是否成功
     */
    bool open();
    
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
     * @brief 关闭工作簿
     * @return 是否成功
     */
    bool close();
    
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
     * @brief 设置活动工作表
     * @param index 工作表索引
     */
    void setActiveWorksheet(size_t index);
    
    // ========== 格式管理 ==========
    
    /**
     * @brief 创建格式
     * @return 格式指针
     */
    std::shared_ptr<Format> createFormat();
    
    /**
     * @brief 根据ID获取格式
     * @param format_id 格式ID
     * @return 格式指针，如果不存在返回nullptr
     */
    std::shared_ptr<Format> getFormat(size_t format_id) const;
    
    /**
     * @brief 获取格式池
     * @return 格式池指针
     */
    FormatPool* getFormatPool() const { return format_pool_.get(); }
    
    /**
     * @brief 从另一个工作簿复制样式数据（用于格式复制）
     * @param source_workbook 源工作簿
     */
    void copyStylesFrom(const Workbook* source_workbook);
    
    /**
     * @brief 获取格式数量
     * @return 格式数量
     */
    size_t getFormatCount() const { return format_pool_->getFormatCount(); }
    
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
        // 为了向后兼容，同步更新streaming_xml
        options_.streaming_xml = (mode == WorkbookMode::STREAMING);
    }
    
    /**
     * @brief 获取当前工作簿模式
     * @return 当前模式
     */
    WorkbookMode getMode() const { return options_.mode; }
    
    /**
     * @brief 启用/禁用流式XML写入（已废弃，请使用setMode）
     * @param enable 是否启用流式XML写入
     * @deprecated 使用 setMode() 代替
     */
    void setStreamingXML(bool enable) {
        options_.streaming_xml = enable;
        // 同步更新mode
        if (enable) {
            options_.mode = WorkbookMode::STREAMING;
        } else {
            options_.mode = WorkbookMode::BATCH;
        }
    }
    
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
    
    /**
     * @brief 检查是否已打开
     * @return 是否已打开
     */
    bool isOpen() const { return is_open_; }
    
    /**
     * @brief 获取文件名
     * @return 文件名
     */
    const std::string& getFilename() const { return filename_; }
    
    /**
     * @brief 检查是否只读
     * @return 是否只读
     */
    bool isReadOnly() const { return read_only_; }
    
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
     * @brief 获取共享字符串索引
     * @param str 字符串
     * @return 索引（如果不存在返回-1）
     */
    int getSharedStringIndex(const std::string& str) const;
    
    // ========== 工作簿编辑功能 ==========
    
    /**
     * @brief 从现有文件加载工作簿进行编辑
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> loadForEdit(const Path& path);
    
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

private:
    // ========== 内部方法 ==========
    
    // 生成Excel文件结构
    bool generateExcelStructure();
    bool generateExcelStructureBatch();
    bool generateExcelStructureStreaming();
    bool generateWorksheetXMLStreaming(const std::shared_ptr<Worksheet>& worksheet, const std::string& path);
    
    // 生成各种XML文件（流式写入）
    void generateWorkbookXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateStylesXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateSharedStringsXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateWorksheetXML(const std::shared_ptr<Worksheet>& worksheet, const std::function<void(const char*, size_t)>& callback) const;
    void generateDocPropsAppXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateDocPropsCoreXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateDocPropsCustomXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateContentTypesXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateRelsXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateWorkbookRelsXML(const std::function<void(const char*, size_t)>& callback) const;
    void generateThemeXML(const std::function<void(const char*, size_t)>& callback) const;
    
    // 辅助函数
    std::string generateUniqueSheetName(const std::string& base_name) const;
    bool validateSheetName(const std::string& name) const;
    void collectSharedStrings();
    
    // 文件路径生成
    std::string getWorksheetPath(int sheet_id) const;
    std::string getWorksheetRelPath(int sheet_id) const;
    
    // 时间格式化
    std::string formatTime(const std::tm& time) const;
    
    // 密码哈希
    std::string hashPassword(const std::string& password) const;
    
    // 流式XML辅助方法
    std::string escapeXML(const std::string& text) const;
    
    // 智能模式选择辅助方法
    size_t estimateMemoryUsage() const;
    size_t getTotalCellCount() const;
};

}} // namespace fastexcel::core
