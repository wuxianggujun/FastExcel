#pragma once

#include "FormatDescriptor.hpp"
#include "FormatRepository.hpp"
#include "StyleTransferContext.hpp"
#include "StyleBuilder.hpp"
#include <memory>
#include <string>
#include <vector>

namespace fastexcel {
namespace core {

// 前向声明
class Worksheet;

/**
 * @brief Excel工作簿 - 重构后的高级API
 * 
 * 使用新的样式管理系统，提供线程安全、高性能的Excel操作接口。
 * 采用PIMPL模式隐藏实现细节，确保API稳定性。
 */
class Workbook {
private:
    class WorkbookImpl;
    std::unique_ptr<WorkbookImpl> pImpl_;

public:
    /**
     * @brief 创建新的工作簿
     * @param filename 文件名
     * @return 工作簿智能指针
     */
    static std::unique_ptr<Workbook> create(const std::string& filename);
    
    /**
     * @brief 打开现有工作簿
     * @param filename 文件名
     * @return 工作簿智能指针
     */
    static std::unique_ptr<Workbook> open(const std::string& filename);
    
    /**
     * @brief 构造函数
     * @param filename 文件名
     */
    explicit Workbook(const std::string& filename);
    
    ~Workbook();
    
    // 禁用拷贝，允许移动
    Workbook(const Workbook&) = delete;
    Workbook& operator=(const Workbook&) = delete;
    Workbook(Workbook&&) = default;
    Workbook& operator=(Workbook&&) = default;
    
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
    
    // ========== 工作表管理 ==========
    
    /**
     * @brief 添加工作表
     * @param name 工作表名称
     * @return 工作表指针
     */
    Worksheet* addWorksheet(const std::string& name = "");
    
    /**
     * @brief 根据索引获取工作表
     * @param index 工作表索引（从0开始）
     * @return 工作表指针，如果索引无效则返回nullptr
     */
    Worksheet* getWorksheet(size_t index) const;
    
    /**
     * @brief 根据名称获取工作表
     * @param name 工作表名称
     * @return 工作表指针，如果不存在则返回nullptr
     */
    Worksheet* getWorksheet(const std::string& name) const;
    
    /**
     * @brief 获取工作表数量
     * @return 工作表数量
     */
    size_t getWorksheetCount() const;
    
    /**
     * @brief 重命名工作表
     * @param index 工作表索引
     * @param new_name 新名称
     * @return 是否成功
     */
    bool renameWorksheet(size_t index, const std::string& new_name);
    
    /**
     * @brief 删除工作表
     * @param index 工作表索引
     * @return 是否成功
     */
    bool deleteWorksheet(size_t index);
    
    /**
     * @brief 移动工作表位置
     * @param from_index 源位置
     * @param to_index 目标位置
     * @return 是否成功
     */
    bool moveWorksheet(size_t from_index, size_t to_index);
    
    // ========== 跨工作簿操作 ==========
    
    /**
     * @brief 从另一个工作簿复制样式
     * @param source_workbook 源工作簿
     * @return 样式传输上下文（用于ID映射）
     */
    std::unique_ptr<StyleTransferContext> copyStylesFrom(
        const Workbook& source_workbook);
    
    /**
     * @brief 复制工作表到另一个工作簿
     * @param source_worksheet 源工作表
     * @param new_name 新工作表名称（空则使用源名称）
     * @return 新创建的工作表指针
     */
    Worksheet* copyWorksheetFrom(const Worksheet& source_worksheet, 
                                const std::string& new_name = "");
    
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
     * @brief 关闭工作簿
     */
    void close();
    
    /**
     * @brief 获取文件名
     * @return 文件名
     */
    std::string getFilename() const;
    
    /**
     * @brief 检查工作簿是否已修改
     * @return 是否已修改
     */
    bool isModified() const;
    
    // ========== 工作簿属性 ==========
    
    /**
     * @brief 设置文档属性
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
     * @brief 设置应用程序名称
     * @param application 应用程序名称
     */
    void setApplication(const std::string& application);
    
    /**
     * @brief 设置默认日期格式
     * @param format 日期格式字符串
     */
    void setDefaultDateFormat(const std::string& format);
    
    // ========== 性能和统计 ==========
    
    /**
     * @brief 获取样式去重统计
     * @return 去重统计信息
     */
    FormatRepository::DeduplicationStats getStyleStats() const;
    
    /**
     * @brief 获取内存使用估算
     * @return 内存使用字节数
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 优化工作簿（压缩样式、清理未使用资源）
     * @return 优化的项目数
     */
    size_t optimize();
};

}} // namespace fastexcel::core