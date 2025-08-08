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

namespace opc {

// 前向声明
class ZipReader;
class ZipRepackWriter;
class PartGraph;
class ContentTypes;

/**
 * @brief OPC包编辑器 - 主流成熟架构实现
 * 
 * 实现策略：
 * 1. 读入阶段：索引所有条目，解析关系图，只对编辑部件建模
 * 2. 编辑阶段：Copy-On-Write，标记dirty
 * 3. 提交阶段：Repack到新ZIP，智能复制
 */
class PackageEditor {
public:
    // ========== 类型定义 ==========
    
    // 单元格引用
    struct CellRef {
        int row;
        int col;
        
        CellRef(int r, int c) : row(r), col(c) {}
        std::string toString() const; // 返回 "A1" 格式
    };
    
    // 单元格值
    struct CellValue {
        enum Type { Empty, String, Number, Formula, Boolean };
        Type type = Empty;
        std::string str_value;
        double num_value = 0.0;
        bool bool_value = false;
        
        static CellValue string(const std::string& s) {
            CellValue v;
            v.type = String;
            v.str_value = s;
            return v;
        }
        
        static CellValue number(double n) {
            CellValue v;
            v.type = Number;
            v.num_value = n;
            return v;
        }
        
        static CellValue formula(const std::string& f) {
            CellValue v;
            v.type = Formula;
            v.str_value = f;
            return v;
        }
    };
    
    using SheetId = std::string;  // 工作表名称作为ID
    
    // ========== 工厂方法 ==========
    
    /**
     * 打开现有Excel文件进行编辑
     */
    static std::unique_ptr<PackageEditor> open(const core::Path& xlsx_path);
    
    /**
     * 从现有Workbook创建编辑器
     */
    static std::unique_ptr<PackageEditor> fromWorkbook(core::Workbook* workbook);
    
    /**
     * 创建新的Excel文件
     */
    static std::unique_ptr<PackageEditor> create();
    
    ~PackageEditor();
    
    // ========== 核心编辑API ==========
    
    /**
     * 设置单元格值
     */
    void setCell(const SheetId& sheet, const CellRef& ref, const CellValue& value);
    
    /**
     * 获取单元格值
     */
    CellValue getCell(const SheetId& sheet, const CellRef& ref) const;
    
    /**
     * 添加工作表
     */
    void addSheet(const std::string& name);
    
    /**
     * 重命名工作表
     */
    void renameSheet(const SheetId& old_name, const std::string& new_name);
    
    /**
     * 删除工作表（第一期可以不支持）
     */
    void removeSheet(const SheetId& sheet);
    
    /**
     * 设置样式（只追加）
     */
    int appendStyle(const class StyleDescriptor& style);
    
    /**
     * 添加图片
     */
    void addImage(const SheetId& sheet, const CellRef& anchor, 
                  const std::vector<uint8_t>& image_data, const std::string& mime_type);
    
    // ========== 提交操作 ==========
    
    /**
     * 提交到文件（repack）
     * @param dst 目标文件路径
     * @return 是否成功
     */
    bool commit(const core::Path& dst);
    
    /**
     * 覆盖原文件
     */
    bool save();
    
    // ========== 状态查询 ==========
    
    /**
     * 获取所有工作表名称
     */
    std::vector<std::string> getSheetNames() const;
    
    /**
     * 检查是否有修改
     */
    bool isDirty() const { return !edit_plan_.dirty_parts.empty(); }
    
    /**
     * 获取脏部件列表
     */
    std::vector<std::string> getDirtyParts() const;
    
private:
    PackageEditor();
    
    // ========== 初始化 ==========
    
    bool initialize(const core::Path& xlsx_path);
    bool initializeEmpty();
    
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
    std::string generateWorksheet(const SheetId& sheet) const;
    
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
    std::unique_ptr<ZipReader> source_reader_;
    
    // 部件图和类型
    std::unique_ptr<PartGraph> part_graph_;
    std::unique_ptr<ContentTypes> content_types_;
    
    // 编辑计划
    EditPlan edit_plan_;
    
    // Workbook对象（用于XML生成）
    core::Workbook* workbook_ = nullptr;
    
    // 懒加载的模型（只加载需要编辑的）
    mutable std::unique_ptr<WorkbookModel> workbook_model_;
    mutable std::unordered_map<SheetId, std::unique_ptr<WorksheetModel>> sheet_models_;
    mutable std::unique_ptr<StylesModel> styles_model_;
    mutable std::unique_ptr<SharedStringsModel> sst_model_;
    
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
    void ensureSheetModel(const SheetId& sheet) const;
    void ensureStylesModel() const;
    void ensureSharedStringsModel() const;
};

}} // namespace fastexcel::opc
