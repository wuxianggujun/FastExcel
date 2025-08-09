#pragma once

#include "fastexcel/core/Path.hpp"
#include <memory>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <string>

namespace fastexcel {

// 前向声明
namespace core {
    class Workbook;
}

namespace archive {
    class ZipReader;
    class ZipWriter;
}

namespace opc {

// 前向声明
class ZipRepackWriter;
class PartGraph;
class ContentTypes;

/**
 * @brief Excel包文件编辑器 - 专注于高效文件保存
 * 
 * 设计原则：
 * 1. 单一职责：只负责文件级别的增量保存和ZIP重新打包
 * 2. 简洁接口：不重新定义业务对象，直接使用现有的Workbook API
 * 3. 高效实现：Copy-On-Write策略，只重新生成修改的部分
 * 
 * 使用方式：
 * ```cpp
 * auto workbook = core::Workbook::open("file.xlsx");
 * auto worksheet = workbook->getWorksheet("Sheet1");
 * worksheet->setCell(1, 1, "Hello");  // 使用现有Cell API
 * 
 * auto editor = PackageEditor::fromWorkbook(workbook.get());
 * editor->save();  // 高效保存修改
 * ```
 */
class PackageEditor {
public:
    // ========== 工厂方法 ==========
    
    /**
     * 打开现有Excel文件进行编辑
     * @param xlsx_path Excel文件路径
     * @return PackageEditor实例，失败返回nullptr
     */
    static std::unique_ptr<PackageEditor> open(const core::Path& xlsx_path);
    
    /**
     * 从现有Workbook创建文件编辑器
     * @param workbook 现有的Workbook对象（不会获取所有权）
     * @return PackageEditor实例，失败返回nullptr
     */
    static std::unique_ptr<PackageEditor> fromWorkbook(core::Workbook* workbook);
    
    /**
     * 创建新的Excel文件编辑器
     * @return PackageEditor实例，包含默认工作表，失败返回nullptr
     */
    static std::unique_ptr<PackageEditor> create();
    
    ~PackageEditor();

    // ========== 核心文件操作 ==========
    
    /**
     * 保存所有修改到原文件
     * @return 成功返回true
     */
    bool save();
    
    /**
     * 将所有修改保存到指定文件
     * @param dst 目标文件路径
     * @return 成功返回true
     */
    bool commit(const core::Path& dst);
    
    // ========== 工作簿访问 ==========
    
    /**
     * 获取关联的Workbook对象
     * @return Workbook指针，用于进行业务操作
     */
    core::Workbook* getWorkbook() const { return workbook_; }
    
    // ========== 状态查询 ==========
    
    /**
     * 检查是否有未保存的修改
     */
    bool isDirty() const;
    
    /**
     * 获取工作表名称列表
     */
    std::vector<std::string> getSheetNames() const;
    
    /**
     * 获取需要重新生成的部件列表（调试用）
     */
    std::vector<std::string> getDirtyParts() const;
    
    // ========== 验证方法 ==========
    
    /**
     * 验证工作表名称是否合法
     */
    static bool isValidSheetName(const std::string& name);
    
    /**
     * 验证单元格引用是否合法
     */
    static bool isValidCellRef(int row, int col);
    
private:
    PackageEditor();
    
    // ========== 初始化 ==========
    
    bool initialize(const core::Path& xlsx_path);
    bool initializeFromWorkbook();
    
    // ========== 编辑计划 ==========
    
    struct EditPlan {
        std::unordered_set<std::string> dirty_parts;      // 修改的部件
        std::unordered_set<std::string> removed_parts;    // 删除的部件
        std::unordered_map<std::string, std::string> new_parts; // 新增的部件及内容
        
        // 标记部件为dirty
        void markDirty(const std::string& part);
        
        // 标记关联部件
        void markRelatedDirty(const std::string& part);
        
        // 添加新部件
        void addNewPart(const std::string& path, const std::string& content);
        
        // 删除部件
        void removePart(const std::string& path);
    };
    
    // ========== 部件生成 ==========
    
    /**
     * 为dirty部件生成内容
     */
    std::string generatePart(const std::string& path) const;
    
    /**
     * 生成workbook.xml
     */
    std::string generateWorkbook() const;
    
    /**
     * 生成worksheet XML
     */
    std::string generateWorksheet(const std::string& sheet_name) const;
    
    /**
     * 生成styles.xml
     */
    std::string generateStyles() const;
    
    /**
     * 生成sharedStrings.xml
     */
    std::string generateSharedStrings() const;
    
    /**
     * 生成Content_Types.xml
     */
    std::string generateContentTypes() const;
    
    /**
     * 生成关系文件
     */
    std::string generateRels(const std::string& rels_path) const;
    
    // ========== 数据模型前向声明 ==========
    
    struct WorkbookModel;
    struct WorksheetModel;
    struct StylesModel;
    struct SharedStringsModel;
    
    // ========== 成员变量 ==========
    
    // 源包信息
    core::Path source_path_;
    std::unique_ptr<archive::ZipReader> source_reader_;
    
    // 部件图和类型
    std::unique_ptr<PartGraph> part_graph_;
    std::unique_ptr<ContentTypes> content_types_;
    
    // 编辑计划
    EditPlan edit_plan_;
    
    // Workbook对象（用于XML生成）
    core::Workbook* workbook_ = nullptr;
    bool owns_workbook_ = false;  // 是否拥有workbook_的所有权
    
    // 懒加载的模型（只加载需要编辑的）
    mutable std::unique_ptr<WorkbookModel> workbook_model_;
    mutable std::unordered_map<std::string, std::unique_ptr<WorksheetModel>> sheet_models_;
    mutable std::unique_ptr<StylesModel> styles_model_;
    mutable std::unique_ptr<SharedStringsModel> sst_model_;
    
    // 工作表名称到ID的映射
    mutable std::unordered_map<std::string, int> sheet_name_to_id_;
    mutable std::unordered_map<int, std::string> sheet_id_to_name_;
    
    // 配置选项
    struct Options {
        bool preserve_unknown = true;       // 保留未知部件
        bool remove_calc_chain = true;      // 删除calcChain
        bool full_calc_on_load = true;      // 设置重新计算
        bool use_shared_strings = true;     // 使用共享字符串
        bool maintain_compat = true;        // 保持兼容性
    } options_;
    
    // 确保模型已加载
    void ensureWorkbookModel() const;
    void ensureSheetModel(const std::string& sheet) const;
    void ensureStylesModel() const;
    void ensureSharedStringsModel() const;
    
    // 自动检测Workbook修改
    void detectAndMarkDirtyParts();
    
    // 工作表映射管理
    void rebuildSheetMapping() const;
    int getSheetId(const std::string& sheet_name) const;
    std::string getSheetName(int sheet_id) const;
    std::string getSheetPath(const std::string& sheet_name) const;
    std::string getSheetPath(int sheet_id) const;
    
    // 常量
    static constexpr int MAX_ROWS = 1048576;    // Excel 2007+ 最大行数
    static constexpr int MAX_COLS = 16384;      // Excel 2007+ 最大列数 (XFD)
};

}} // namespace fastexcel::opc
