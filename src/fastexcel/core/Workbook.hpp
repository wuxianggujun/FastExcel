#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/Format.hpp"
#include "fastexcel/archive/FileManager.hpp"
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
    bool use_shared_strings = false;  // 使用共享字符串（默认禁用以提高性能）
    bool streaming_xml = true;        // 流式XML写入（默认启用以优化内存）
    size_t row_buffer_size = 5000;    // 行缓冲大小（默认较大缓冲）
    int compression_level = 1;        // ZIP压缩级别（默认快速压缩）
    size_t xml_buffer_size = 4 * 1024 * 1024; // XML缓冲区大小（4MB）
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
    std::unordered_map<int, std::shared_ptr<Format>> formats_;
    std::unique_ptr<archive::FileManager> file_manager_;
    
    // ID管理
    int next_format_id_ = 164;        // Excel内置格式之后从164开始
    int next_sheet_id_ = 1;
    int next_font_id_ = 1;
    int next_fill_id_ = 2;            // 0和1是默认填充
    int next_border_id_ = 1;
    int next_xf_id_ = 1;
    
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
    
    // 格式索引管理
    std::unordered_map<size_t, int> font_hash_to_id_;
    std::unordered_map<size_t, int> fill_hash_to_id_;
    std::unordered_map<size_t, int> border_hash_to_id_;
    std::unordered_map<size_t, int> xf_hash_to_id_;
    
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
     * @param filename 文件名
     * @return 工作簿智能指针
     */
    static std::unique_ptr<Workbook> create(const std::string& filename);
    
    /**
     * @brief 构造函数
     * @param filename 文件名
     */
    explicit Workbook(const std::string& filename);
    
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
     * @brief 获取格式
     * @param format_id 格式ID
     * @return 格式指针
     */
    std::shared_ptr<Format> getFormat(int format_id) const;
    
    /**
     * @brief 获取格式数量
     * @return 格式数量
     */
    size_t getFormatCount() const { return formats_.size(); }
    
    // ========== 文档属性 ==========
    
    /**
     * @brief 设置文档标题
     * @param title 标题
     */
    void setTitle(const std::string& title) { doc_properties_.title = title; }
    
    /**
     * @brief 设置文档主题
     * @param subject 主题
     */
    void setSubject(const std::string& subject) { doc_properties_.subject = subject; }
    
    /**
     * @brief 设置文档作者
     * @param author 作者
     */
    void setAuthor(const std::string& author) { doc_properties_.author = author; }
    
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
     * @brief 启用/禁用流式XML写入
     * @param enable 是否启用流式XML写入
     */
    void setStreamingXML(bool enable) { options_.streaming_xml = enable; }
    
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
    
    // 格式管理内部方法
    int getFontId(const Format& format);
    int getFillId(const Format& format);
    int getBorderId(const Format& format);
    int getXfId(const Format& format);
    
    // 辅助函数
    std::string generateUniqueSheetName(const std::string& base_name) const;
    bool validateSheetName(const std::string& name) const;
    void collectSharedStrings();
    void updateFormatIndices();
    
    // 文件路径生成
    std::string getWorksheetPath(int sheet_id) const;
    std::string getWorksheetRelPath(int sheet_id) const;
    
    // 时间格式化
    std::string formatTime(const std::tm& time) const;
    
    // 密码哈希
    std::string hashPassword(const std::string& password) const;
    
    // 流式XML辅助方法
    std::string columnToLetter(int col) const;
    std::string escapeXML(const std::string& text) const;
};

}} // namespace fastexcel::core
