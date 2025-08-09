#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/opc/PackageManagerService.hpp"
#include "fastexcel/opc/IPackageManager.hpp"
#include "fastexcel/tracking/IChangeTracker.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include <memory>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <functional>

namespace fastexcel {

// 前向声明
namespace core {
    class Workbook;
    class Worksheet;
}

namespace archive {
    class ZipReader;
}

namespace opc {
// 前向声明
class ZipRepackWriter;
}

namespace opc {

/**
 * @brief Excel包文件编辑器 - 重构版
 * 
 * 新架构设计原则：
 * 1. 单一职责：PackageEditor专注于协调各个服务
 * 2. 依赖注入：通过接口依赖抽象实现
 * 3. 服务分离：XML生成、包管理、变更跟踪分别由专门服务处理
 * 
 * 使用方式：
 * ```cpp
 * auto editor = PackageEditor::fromWorkbook(workbook.get());
 * editor->detectChanges();  // 新功能：智能变更检测
 * editor->commit("output.xlsx");  // 保存修改
 * ```
 */
class PackageEditor {
    // 前向声明友元类
    friend class WorkbookXMLGenerator;
    
private:
    // 核心服务组件（依赖注入）
    std::unique_ptr<IPackageManager> package_manager_;
    std::unique_ptr<xml::UnifiedXMLGenerator> xml_generator_;
    std::unique_ptr<tracking::IChangeTracker> change_tracker_;
    
    // 关联的Workbook对象（不拥有所有权）
    core::Workbook* workbook_ = nullptr;
    
    // 内部状态
    core::Path source_path_;
    bool initialized_ = false;
    
    // 配置选项
    struct Options {
        bool preserve_unknown_parts = true;   // 保留未知部件
        bool remove_calc_chain = true;        // 删除calcChain
        bool auto_detect_changes = true;      // 自动检测变更
        bool use_shared_strings = true;       // 使用共享字符串
        bool validate_xml = false;            // XML验证
    } options_;
    
public:
    PackageEditor();
    ~PackageEditor();
    
    // 禁用拷贝，支持移动
    PackageEditor(const PackageEditor&) = delete;
    PackageEditor& operator=(const PackageEditor&) = delete;
    PackageEditor(PackageEditor&&) = default;
    PackageEditor& operator=(PackageEditor&&) = default;
    
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
    
    // ========== 新功能：智能变更管理 ==========
    
    /**
     * 智能检测并标记变更的部件
     * 这是新架构的核心功能，自动分析Workbook变更
     */
    void detectChanges();
    
    /**
     * 手动标记部件为需要重新生成
     * @param part_name 部件名称，如 "xl/workbook.xml"
     */
    void markPartDirty(const std::string& part_name);
    
    /**
     * 获取变更统计信息
     */
    struct ChangeStats {
        size_t modified_parts = 0;
        size_t created_parts = 0;
        size_t deleted_parts = 0;
        size_t total_size_bytes = 0;
    };
    
    ChangeStats getChangeStats() const;
    
    // ========== 配置管理 ==========
    
    /**
     * 设置是否保留未知部件
     */
    void setPreserveUnknownParts(bool preserve) {
        options_.preserve_unknown_parts = preserve;
    }
    
    /**
     * 设置是否删除计算链
     */
    void setRemoveCalcChain(bool remove) {
        options_.remove_calc_chain = remove;
    }
    
    /**
     * 设置是否自动检测变更
     */
    void setAutoDetectChanges(bool auto_detect) {
        options_.auto_detect_changes = auto_detect;
    }
    
    const Options& getOptions() const { return options_; }
    
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
     * 获取需要重新生成的部件列表
     */
    std::vector<std::string> getDirtyParts() const;
    
    /**
     * 获取包中所有部件列表
     */
    std::vector<std::string> getAllParts() const;
    
    // ========== 验证方法 ==========
    
    /**
     * 验证工作表名称是否合法
     */
    static bool isValidSheetName(const std::string& name);
    
    /**
     * 验证单元格引用是否合法
     */
    static bool isValidCellRef(int row, int col);
    
    // ========== 高级功能 ==========
    
    /**
     * 生成指定部件的内容（用于调试和自定义）
     * @param part_path 部件路径
     * @return 生成的XML内容
     */
    std::string generatePart(const std::string& part_path) const;
    
    /**
     * 验证生成的XML内容（可选功能）
     * @param xml_content XML内容
     * @return 验证结果
     */
    bool validateXML(const std::string& xml_content) const;
    
private:
    // ========== 初始化方法 ==========
    
    bool initialize(const core::Path& xlsx_path);
    bool initializeFromWorkbook();
    bool initializeServices(std::unique_ptr<archive::ZipReader> zip_reader, core::Workbook* workbook);
    
    // ========== 私有辅助方法 ==========
    
    std::string extractSheetNameFromPath(const std::string& part_path) const;
    bool isRequiredPart(const std::string& part) const;
    void logOperationStats() const;
    
    // 常量
    static constexpr int MAX_ROWS = 1048576;    // Excel 2007+ 最大行数
    static constexpr int MAX_COLS = 16384;      // Excel 2007+ 最大列数 (XFD)
};

} // namespace opc
} // namespace fastexcel
