#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/archive/FileManager.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/opc/PackageEditorManager.hpp"
#include "fastexcel/read/ReadWorkbook.hpp"
#include <memory>
#include <string>
#include <vector>

namespace fastexcel {
namespace edit {

// 前向声明
class EditWorksheet;
class RowWriter;

/**
 * @brief 可编辑工作簿接口
 * 
 * 继承只读功能，添加编辑能力
 */
class IEditableWorkbook : public read::IReadOnlyWorkbook {
public:
    virtual ~IEditableWorkbook() = default;
    
    // 编辑方法
    virtual std::shared_ptr<EditWorksheet> getWorksheetForEdit(const std::string& name) = 0;
    virtual std::shared_ptr<EditWorksheet> getWorksheetForEdit(size_t index) = 0;
    virtual std::shared_ptr<EditWorksheet> addWorksheet(const std::string& name) = 0;
    virtual bool removeWorksheet(const std::string& name) = 0;
    virtual bool removeWorksheet(size_t index) = 0;
    
    // 文件操作
    virtual bool save() = 0;
    virtual bool saveAs(const std::string& filename) = 0;
    
    // 状态管理
    virtual bool hasUnsavedChanges() const = 0;
    virtual void markAsModified() = 0;
    virtual void discardChanges() = 0;
    
    // 工作表管理
    virtual bool renameWorksheet(const std::string& old_name, const std::string& new_name) = 0;
    virtual bool moveWorksheet(size_t from_index, size_t to_index) = 0;
};

/**
 * @brief 编辑会话
 * 
 * 管理Excel文件的编辑操作
 * - 支持创建新文件
 * - 支持编辑现有文件
 * - 智能保存策略
 * - 流式写入支持
 */
class EditSession : public IEditableWorkbook {
private:
    // 核心组件
    std::unique_ptr<core::Workbook> workbook_;  // 复用现有Workbook作为内部实现
    std::unique_ptr<archive::FileManager> file_manager_;
    std::unique_ptr<xml::UnifiedXMLGenerator> xml_generator_;
    std::unique_ptr<core::ExcelStructureGenerator> structure_generator_;
    std::unique_ptr<core::DirtyManager> dirty_manager_;
    std::unique_ptr<opc::PackageEditorManager> package_manager_;
    
    // 编辑状态
    const read::WorkbookAccessMode access_mode_;
    bool has_unsaved_changes_ = false;
    std::optional<std::string> original_file_path_;
    std::string filename_;
    
    // 工作表管理
    std::vector<std::shared_ptr<EditWorksheet>> worksheets_;
    
    // 模式选择
    core::WorkbookMode mode_ = core::WorkbookMode::AUTO;
    
public:
    /**
     * @brief 创建新工作簿
     * @param path 文件路径
     * @return 编辑会话
     */
    static std::unique_ptr<EditSession> createNew(const core::Path& path);
    
    /**
     * @brief 从现有文件创建编辑会话
     * @param path 文件路径
     * @return 编辑会话
     */
    static std::unique_ptr<EditSession> fromExistingFile(const core::Path& path);
    
    /**
     * @brief 从只读工作簿创建编辑会话
     * @param read_workbook 只读工作簿
     * @param target_path 目标路径
     * @return 编辑会话
     */
    static std::unique_ptr<EditSession> beginEdit(const read::ReadWorkbook& read_workbook, 
                                                  const core::Path& target_path);
    
private:
    /**
     * @brief 私有构造函数
     * @param path 文件路径
     * @param mode 访问模式
     */
    EditSession(const core::Path& path, read::WorkbookAccessMode mode);
    
public:
    /**
     * @brief 析构函数
     */
    ~EditSession();
    
    // ========== IReadOnlyWorkbook 接口实现 ==========
    
    read::WorkbookAccessMode getAccessMode() const override { 
        return access_mode_; 
    }
    
    bool isReadOnly() const override { 
        return false; 
    }
    
    std::string getFilename() const override {
        return filename_;
    }
    
    size_t getWorksheetCount() const override;
    std::vector<std::string> getWorksheetNames() const override;
    std::shared_ptr<const read::ReadWorksheet> getWorksheet(const std::string& name) const override;
    std::shared_ptr<const read::ReadWorksheet> getWorksheet(size_t index) const override;
    bool hasWorksheet(const std::string& name) const override;
    int getWorksheetIndex(const std::string& name) const override;
    
    // ========== IEditableWorkbook 接口实现 ==========
    
    std::shared_ptr<EditWorksheet> getWorksheetForEdit(const std::string& name) override;
    std::shared_ptr<EditWorksheet> getWorksheetForEdit(size_t index) override;
    std::shared_ptr<EditWorksheet> addWorksheet(const std::string& name) override;
    bool removeWorksheet(const std::string& name) override;
    bool removeWorksheet(size_t index) override;
    
    bool save() override;
    bool saveAs(const std::string& filename) override;
    
    bool hasUnsavedChanges() const override;
    void markAsModified() override;
    void discardChanges() override;
    
    bool renameWorksheet(const std::string& old_name, const std::string& new_name) override;
    bool moveWorksheet(size_t from_index, size_t to_index) override;
    
    // ========== 高性能编辑方法 ==========
    
    /**
     * @brief 创建行写入器
     * @param worksheet_name 工作表名称
     * @return 行写入器
     */
    std::unique_ptr<RowWriter> createRowWriter(const std::string& worksheet_name);
    
    /**
     * @brief 设置工作簿模式
     * @param mode 模式（AUTO/BATCH/STREAMING）
     */
    void setMode(core::WorkbookMode mode) { mode_ = mode; }
    
    /**
     * @brief 获取工作簿模式
     * @return 当前模式
     */
    core::WorkbookMode getMode() const { return mode_; }
    
    /**
     * @brief 启用/禁用共享字符串
     * @param enable 是否启用
     */
    void setUseSharedStrings(bool enable);
    
    /**
     * @brief 设置压缩级别
     * @param level 压缩级别（0-9）
     */
    void setCompressionLevel(int level);
    
    /**
     * @brief 获取内部Workbook（用于兼容）
     * @return Workbook指针
     */
    core::Workbook* getInternalWorkbook() { return workbook_.get(); }
    
private:
    /**
     * @brief 初始化新工作簿
     */
    void initializeNewWorkbook();
    
    /**
     * @brief 从文件加载
     * @param path 文件路径
     * @return 是否成功
     */
    bool loadFromFile(const core::Path& path);
    
    // 注意：Excel 结构生成已统一到 ExcelStructureGenerator
    // 原 generateExcelStructure() 方法已删除，使用工作簿的保存接口替代
    
    /**
     * @brief 验证工作表名称
     * @param name 名称
     * @return 是否有效
     */
    bool isValidWorksheetName(const std::string& name) const;
};

/**
 * @brief 可编辑工作表
 * 
 * 提供完整的编辑功能
 */
class EditWorksheet {
private:
    std::shared_ptr<core::Worksheet> worksheet_;  // 复用现有Worksheet
    EditSession* parent_session_;
    
public:
    EditWorksheet(const std::string& name, EditSession* session);
    EditWorksheet(std::shared_ptr<core::Worksheet> worksheet, EditSession* session);
    
    // 基本信息
    std::string getName() const;
    void setName(const std::string& name);
    
    // 数据写入
    void writeString(int row, int col, const std::string& value);
    void writeNumber(int row, int col, double value);
    void writeBoolean(int row, int col, bool value);
    void writeFormula(int row, int col, const std::string& formula);
    
    // 格式设置
    void setCellFormat(int row, int col, int format_id);
    void setRowHeight(int row, double height);
    void setColumnWidth(int col, double width);
    
    // 范围操作
    void mergeCells(int first_row, int first_col, int last_row, int last_col);
    void unmergeCells(int first_row, int first_col, int last_row, int last_col);
    void clearRange(int first_row, int first_col, int last_row, int last_col);
    
    // 获取内部Worksheet
    core::Worksheet* getInternalWorksheet() { return worksheet_.get(); }
};

}} // namespace fastexcel::edit