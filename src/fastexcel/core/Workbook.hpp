#pragma once

#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/WorkbookModeSelector.hpp"
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

// DocumentProperties is defined in WorkbookDocumentManager.hpp

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
    /**
     * @brief 创建工作簿（直接可用，无需再调用open）
     * @param path 文件路径
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> create(const Path& path);
    
    /**
     * @brief 创建工作簿（字符串重载版本）
     * @param filepath 文件路径字符串
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> create(const std::string& filepath);
    
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
     * @brief 只读方式打开Excel文件（字符串重载版本）
     * @param filepath 文件路径字符串
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> openForReading(const std::string& filepath);
    
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
     * @brief 编辑方式打开Excel文件（字符串重载版本）
     * @param filepath 文件路径字符串
     * @return 工作簿智能指针，失败返回nullptr
     */
    static std::unique_ptr<Workbook> openForEditing(const std::string& filepath);
    

    
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
    
    // 文件操作
    
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
    
    // 编辑模式/保真写回配置
    void setPreserveUnknownParts(bool enable) { preserve_unknown_parts_ = enable; }
    bool getPreserveUnknownParts() const { return preserve_unknown_parts_; }

    // 工作表管理
    
    /**
     * @brief 添加工作表
     * @param name 工作表名称（空则自动生成）
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> addSheet(const std::string& name = "");
    
    /**
     * @brief 插入工作表
     * @param index 插入位置
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> insertSheet(size_t index, const std::string& name = "");
    
    /**
     * @brief 删除工作表
     * @param name 工作表名称
     * @return 是否成功
     */
    bool removeSheet(const std::string& name);
    
    /**
     * @brief 删除工作表
     * @param index 工作表索引
     * @return 是否成功
     */
    bool removeSheet(size_t index);
    
    /**
     * @brief 获取工作表（按名称）
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> getSheet(const std::string& name);
    
    /**
     * @brief 获取工作表（按索引）
     * @param index 工作表索引
     * @return 工作表指针
     */
    std::shared_ptr<Worksheet> getSheet(size_t index);
    
    /**
     * @brief 获取工作表（按名称，只读）
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> getSheet(const std::string& name) const;
    
    /**
     * @brief 获取工作表（按索引，只读）
     * @param index 工作表索引
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> getSheet(size_t index) const;
    
    // 便捷的工作表访问操作符
    /**
     * @brief 通过索引访问工作表（操作符重载）
     * @param index 工作表索引
     * @return 工作表指针
     * 
     * @example
     * auto worksheet = workbook[0];  // 获取第一个工作表
     * workbook[0]->setValue("A1", std::string("Hello"));  // 直接操作
     */
    std::shared_ptr<Worksheet> operator[](size_t index) {
        return getSheet(index);
    }
    
    /**
     * @brief 通过索引访问工作表（只读操作符重载）
     * @param index 工作表索引
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> operator[](size_t index) const {
        return getSheet(index);
    }
    
    /**
     * @brief 通过名称访问工作表（操作符重载）
     * @param name 工作表名称
     * @return 工作表指针
     * 
     * @example
     * auto worksheet = workbook["Sheet1"];  // 获取名为"Sheet1"的工作表
     * workbook["Sheet1"]->setValue("A1", std::string("Hello"));  // 直接操作
     */
    std::shared_ptr<Worksheet> operator[](const std::string& name) {
        return getSheet(name);
    }
    
    /**
     * @brief 通过名称访问工作表（只读操作符重载）
     * @param name 工作表名称
     * @return 工作表指针
     */
    std::shared_ptr<const Worksheet> operator[](const std::string& name) const {
        return getSheet(name);
    }
    
    /**
     * @brief 获取工作表数量
     * @return 工作表数量
     */
    size_t getSheetCount() const { return worksheets_.size(); }
    
    /**
     * @brief 检查工作簿是否为空（无工作表）
     * @return 是否为空
     */
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
    
    /**
     * @brief 获取第一个工作表（只读版本）
     * @return 第一个工作表指针，如果无工作表返回nullptr
     */
    std::shared_ptr<const Worksheet> getFirstSheet() const;
    
    /**
     * @brief 获取最后一个工作表
     * @return 最后一个工作表指针，如果无工作表返回nullptr
     */
    std::shared_ptr<Worksheet> getLastSheet();
    
    /**
     * @brief 获取最后一个工作表（只读版本）
     * @return 最后一个工作表指针，如果无工作表返回nullptr
     */
    std::shared_ptr<const Worksheet> getLastSheet() const;
    
    /**
     * @brief 获取所有工作表名称
     * @return 工作表名称列表
     */
    std::vector<std::string> getSheetNames() const;
    
    // 便捷的工作表查找方法
    /**
     * @brief 检查是否存在指定名称的工作表
     * @param name 工作表名称
     * @return 是否存在
     * 
     * @example
     * if (workbook.hasSheet("Data")) {
     *     auto sheet = workbook.getSheet("Data");
     * }
     */
    bool hasSheet(const std::string& name) const;
    
    /**
     * @brief 查找工作表（可能不存在）
     * @param name 工作表名称
     * @return 工作表指针，如果不存在返回nullptr
     * 
     * @example
     * if (auto sheet = workbook.findSheet("Data")) {
     *     // 工作表存在，可以安全使用
     *     sheet->setValue("A1", std::string("Hello"));
     * }
     */
    std::shared_ptr<Worksheet> findSheet(const std::string& name);
    
    /**
     * @brief 查找工作表（只读版本）
     * @param name 工作表名称
     * @return 工作表指针，如果不存在返回nullptr
     */
    std::shared_ptr<const Worksheet> findSheet(const std::string& name) const;
    
    /**
     * @brief 获取所有工作表
     * @return 工作表指针列表
     * 
     * @example
     * for (auto& sheet : workbook.getAllSheets()) {
     *     std::cout << "工作表: " << sheet->getName() << std::endl;
     * }
     */
    std::vector<std::shared_ptr<Worksheet>> getAllSheets();
    
    /**
     * @brief 获取所有工作表（只读版本）
     * @return 工作表指针列表
     */
    std::vector<std::shared_ptr<const Worksheet>> getAllSheets() const;
    
    /**
     * @brief 清空所有工作表
     * @return 清理的工作表数量
     * 
     * @example
     * int count = workbook.clearAllSheets();
     * std::cout << "已清理 " << count << " 个工作表" << std::endl;
     */
    int clearAllSheets();
    
    /**
     * @brief 重命名工作表
     * @param old_name 旧名称
     * @param new_name 新名称
     * @return 是否成功
     */
    bool renameSheet(const std::string& old_name, const std::string& new_name);
    
    /**
     * @brief 移动工作表
     * @param from_index 源位置
     * @param to_index 目标位置
     * @return 是否成功
     */
    bool moveSheet(size_t from_index, size_t to_index);
    
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
    
    /**
     * @brief 获取活动工作表
     * @return 活动工作表指针，如果没有工作表则返回nullptr
     * 
     * @example
     * auto activeSheet = workbook.getActiveWorksheet();
     * if (activeSheet) {
     *     activeSheet->setValue("A1", std::string("Hello"));
     * }
     */
    std::shared_ptr<Worksheet> getActiveWorksheet();
    
    /**
     * @brief 获取活动工作表（只读）
     * @return 活动工作表指针，如果没有工作表则返回nullptr
     */
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
    T getValue(size_t sheet_index, int row, int col) const {
        auto worksheet = getSheet(sheet_index);
        if (!worksheet) {
            throw std::runtime_error("Invalid worksheet index: " + std::to_string(sheet_index));
        }
        return worksheet->getValue<T>(row, col);
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
    void setValue(const std::string& sheet_name, int row, int col, const T& value) {
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        worksheet->setValue<T>(row, col, value);
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
    void setValue(size_t sheet_index, int row, int col, const T& value) {
        auto worksheet = getSheet(sheet_index);
        if (!worksheet) {
            throw std::runtime_error("Invalid worksheet index: " + std::to_string(sheet_index));
        }
        worksheet->setValue<T>(row, col, value);
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
        } catch (...) {
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
        } catch (...) {
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
    std::optional<T> tryGetValue(const std::string& sheet_name, int row, int col) const noexcept {
        try {
            auto worksheet = findSheet(sheet_name);
            if (!worksheet) {
                return std::nullopt;
            }
            return worksheet->tryGetValue<T>(row, col);
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
    std::optional<T> tryGetValue(size_t sheet_index, int row, int col) const noexcept {
        try {
            if (sheet_index >= worksheets_.size()) {
                return std::nullopt;
            }
            auto worksheet = worksheets_[sheet_index];
            if (!worksheet) {
                return std::nullopt;
            }
            return worksheet->tryGetValue<T>(row, col);
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
    bool trySetValue(const std::string& sheet_name, int row, int col, const T& value) noexcept {
        try {
            auto worksheet = findSheet(sheet_name);
            if (!worksheet) {
                return false;
            }
            worksheet->setValue<T>(row, col, value);
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
            return getValue<T>(0, row, col);
        }
        
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        return worksheet->getValue<T>(row, col);
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
            setValue<T>(0, row, col, value);
            return;
        }
        
        auto worksheet = getSheet(sheet_name);
        if (!worksheet) {
            throw std::runtime_error("Worksheet not found: " + sheet_name);
        }
        worksheet->setValue<T>(row, col, value);
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
    
    /**
     * @brief 获取文档标题
     * @return 标题
     */
    std::string getTitle() const { 
        return document_manager_ ? document_manager_->getTitle() : std::string(); 
    }
    
    /**
     * @brief 设置文档主题
     * @param subject 主题
     */
    void setSubject(const std::string& subject) { 
        if (document_manager_) document_manager_->setSubject(subject); 
    }
    
    /**
     * @brief 获取文档主题
     * @return 主题
     */
    std::string getSubject() const { 
        return document_manager_ ? document_manager_->getSubject() : std::string(); 
    }
    
    /**
     * @brief 设置文档作者
     * @param author 作者
     */
    void setAuthor(const std::string& author) { 
        if (document_manager_) document_manager_->setAuthor(author); 
    }
    
    /**
     * @brief 获取文档作者
     * @return 作者
     */
    std::string getAuthor() const { 
        return document_manager_ ? document_manager_->getAuthor() : std::string(); 
    }
    
    /**
     * @brief 设置文档管理者
     * @param manager 管理者
     */
    void setManager(const std::string& manager) { 
        if (document_manager_) document_manager_->setManager(manager); 
    }
    
    /**
     * @brief 设置公司
     * @param company 公司
     */
    void setCompany(const std::string& company) { 
        if (document_manager_) document_manager_->setCompany(company); 
    }
    
    /**
     * @brief 设置类别
     * @param category 类别
     */
    void setCategory(const std::string& category) { 
        if (document_manager_) document_manager_->setCategory(category); 
    }
    
    /**
     * @brief 设置关键词
     * @param keywords 关键词
     */
    void setKeywords(const std::string& keywords) { 
        if (document_manager_) document_manager_->setKeywords(keywords); 
    }
    
    /**
     * @brief 设置注释
     * @param comments 注释
     */
    void setComments(const std::string& comments) { 
        if (document_manager_) document_manager_->setComments(comments); 
    }
    
    /**
     * @brief 设置状态
     * @param status 状态
     */
    void setStatus(const std::string& status) { 
        if (document_manager_) document_manager_->setStatus(status); 
    }
    
    /**
     * @brief 设置超链接基础
     * @param hyperlink_base 超链接基础
     */
    void setHyperlinkBase(const std::string& hyperlink_base) { 
        if (document_manager_) document_manager_->setHyperlinkBase(hyperlink_base); 
    }
    
    /**
     * @brief 设置创建时间
     * @param created_time 创建时间
     */
    void setCreatedTime(const std::tm& created_time) { 
        if (document_manager_) document_manager_->setCreatedTime(created_time); 
    }
    
    /**
     * @brief 设置修改时间
     * @param modified_time 修改时间
     */
    void setModifiedTime(const std::tm& modified_time) { 
        if (document_manager_) document_manager_->setModifiedTime(modified_time); 
    }
    
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
     * @brief 添加自定义属性（数字）
     * @param name 属性名
     * @param value 属性值
     */
    void setProperty(const std::string& name, double value) {
        if (document_manager_) document_manager_->setCustomProperty(name, value);
    }
    
    /**
     * @brief 添加自定义属性（布尔）
     * @param name 属性名
     * @param value 属性值
     */
    void setProperty(const std::string& name, bool value) {
        if (document_manager_) document_manager_->setCustomProperty(name, value);
    }
    
    /**
     * @brief 获取自定义属性
     * @param name 属性名
     * @return 属性值（如果不存在返回空字符串）
     */
    std::string getProperty(const std::string& name) const {
        return document_manager_ ? document_manager_->getCustomProperty(name) : std::string();
    }
    
    /**
     * @brief 删除自定义属性
     * @param name 属性名
     * @return 是否成功
     */
    bool removeProperty(const std::string& name) {
        return document_manager_ ? document_manager_->removeCustomProperty(name) : false;
    }
    
    /**
     * @brief 获取所有自定义属性
     * @return 自定义属性映射 (名称 -> 值)
     */
    std::unordered_map<std::string, std::string> getAllProperties() const {
        return document_manager_ ? document_manager_->getAllCustomProperties() : std::unordered_map<std::string, std::string>();
    }
    
    // 定义名称
    
    /**
     * @brief 定义名称
     * @param name 名称
     * @param formula 公式
     * @param scope 作用域（工作表名或空表示全局）
     */
    void defineName(const std::string& name, const std::string& formula, const std::string& scope = "") {
        if (document_manager_) document_manager_->defineName(name, formula, scope);
    }
    
    /**
     * @brief 获取定义名称的公式
     * @param name 名称
     * @param scope 作用域
     * @return 公式（如果不存在返回空字符串）
     */
    std::string getDefinedName(const std::string& name, const std::string& scope = "") const {
        return document_manager_ ? document_manager_->getDefinedName(name, scope) : std::string();
    }
    
    /**
     * @brief 删除定义名称
     * @param name 名称
     * @param scope 作用域
     * @return 是否成功
     */
    bool removeDefinedName(const std::string& name, const std::string& scope = "") {
        return document_manager_ ? document_manager_->removeDefinedName(name, scope) : false;
    }
    
    // VBA项目
    
    /**
     * @brief 添加VBA项目
     * @param vba_project_path VBA项目文件路径
     * @return 是否成功
     */
    bool addVbaProject(const std::string& vba_project_path) {
        return security_manager_ ? security_manager_->addVbaProject(vba_project_path) : false;
    }
    
    /**
     * @brief 检查是否有VBA项目
     * @return 是否有VBA项目
     */
    bool hasVbaProject() const { 
        return security_manager_ ? security_manager_->hasVbaProject() : false; 
    }
    
    // 工作簿保护
    
    /**
     * @brief 保护工作簿
     * @param password 密码（可选）
     * @param lock_structure 锁定结构
     * @param lock_windows 锁定窗口
     */
    void protect(const std::string& password = "", bool lock_structure = true, bool lock_windows = false) {
        if (security_manager_) {
            WorkbookSecurityManager::ProtectionOptions options;
            options.password = password;
            options.lock_structure = lock_structure;
            options.lock_windows = lock_windows;
            security_manager_->protect(options);
        }
    }
    
    /**
     * @brief 取消保护
     */
    void unprotect() {
        if (security_manager_) security_manager_->unprotect();
    }
    
    /**
     * @brief 检查是否受保护
     * @return 是否受保护
     */
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
    
    // 获取状态

    // 管理器访问器（获取管理器实例）
    WorkbookDocumentManager* getDocumentManager() { return document_manager_.get(); }
    const WorkbookDocumentManager* getDocumentManager() const { return document_manager_.get(); }
    
    WorkbookSecurityManager* getSecurityManager() { return security_manager_.get(); }
    const WorkbookSecurityManager* getSecurityManager() const { return security_manager_.get(); }
    
    WorkbookDataManager* getDataManager() { return data_manager_.get(); }
    const WorkbookDataManager* getDataManager() const { return data_manager_.get(); }
    
    // 同步方法（保持遗留系统兼容）
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
    const DocumentProperties& getDocumentProperties() const { 
        static DocumentProperties empty_properties;
        return document_manager_ ? document_manager_->getDocumentProperties() : empty_properties; 
    }
    
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
    
    // 共享字符串管理
    
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
    const SharedStringTable* getSharedStrings() const;
    
    /**
     * @brief 获取共享字符串表
     * @return 共享字符串表指针
     */
    SharedStringTable* getSharedStringTable() { return shared_string_table_.get(); }
    const SharedStringTable* getSharedStringTable() const { return shared_string_table_.get(); }
    
    /**
     * @brief 检查是否处于编辑模式
     * @return 是否处于编辑模式
     */
    bool isEditMode() const { return state_ == WorkbookState::EDITING; }
    
    /**
     * @brief 获取原始包路径（用于编辑模式）
     * @return 原始文件路径
     */
    const std::string& getOriginalPackagePath() const { return original_package_path_; }
    
    /**
     * @brief 获取估计大小（用于决定是否使用流式写入）
     * @return 估计的文件大小（字节）
     */
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
    // 内部方法
    
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
