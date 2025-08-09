#pragma once

#include "fastexcel/core/Path.hpp"
#include "fastexcel/read/ReadWorkbook.hpp"
#include "fastexcel/read/ReadWorksheet.hpp"
#include "fastexcel/edit/EditSession.hpp"
#include "fastexcel/edit/RowWriter.hpp"
#include <memory>
#include <string>

namespace fastexcel {

/**
 * @brief FastExcel 统一工厂类
 * 
 * 提供语义化的API接口，彻底解决状态管理混乱问题
 * 
 * 使用示例：
 * ```cpp
 * // 只读访问
 * auto readonly = FastExcel::openForReading("report.xlsx");
 * auto ws = readonly->getWorksheet("Data");
 * auto value = ws->readString(0, 0);
 * 
 * // 编辑访问
 * auto editable = FastExcel::openForEditing("report.xlsx");
 * auto ws = editable->getWorksheetForEdit("Data");
 * ws->writeString(0, 0, "Modified");
 * editable->save();
 * 
 * // 创建新文件
 * auto newfile = FastExcel::createWorkbook("new.xlsx");
 * auto ws = newfile->addWorksheet("Sheet1");
 * ws->writeString(0, 0, "Hello");
 * newfile->save();
 * ```
 */
class FastExcel {
public:
    /**
     * @brief 创建新的Excel文件（可编辑）
     * @param path 文件路径
     * @return 可编辑工作簿，失败返回nullptr
     * 
     * 使用场景：
     * - 创建全新的Excel文件
     * - 从空白开始构建工作簿
     */
    static std::unique_ptr<edit::EditSession> createWorkbook(const core::Path& path) {
        try {
            return edit::EditSession::createNew(path);
        } catch (const std::exception& e) {
            LOG_ERROR("Failed to create workbook: {}, error: {}", path.string(), e.what());
            return nullptr;
        }
    }
    
    /**
     * @brief 只读方式打开Excel文件
     * @param path 文件路径  
     * @return 只读工作簿，失败返回nullptr
     * 
     * 使用场景：
     * - 查看Excel文件内容
     * - 数据分析和报告
     * - 不需要修改文件的场景
     * 
     * 特点：
     * - 轻量级：内存占用小
     * - 高性能：优化的读取路径
     * - 安全：编译时防止误修改
     */
    static std::unique_ptr<read::ReadWorkbook> openForReading(const core::Path& path) {
        try {
            if (!path.exists()) {
                LOG_ERROR("File not found: {}", path.string());
                return nullptr;
            }
            
            return std::make_unique<read::ReadWorkbook>(path);
        } catch (const std::exception& e) {
            LOG_ERROR("Failed to open file for reading: {}, error: {}", path.string(), e.what());
            return nullptr;
        }
    }
    
    /**
     * @brief 编辑方式打开Excel文件
     * @param path 文件路径
     * @return 可编辑工作簿，失败返回nullptr
     * 
     * 使用场景：
     * - 修改现有Excel文件
     * - 增量编辑操作
     * - 需要保存更改的场景
     * 
     * 特点：
     * - 完整功能：支持所有编辑操作
     * - 增量更新：只重新生成修改部分
     * - 变更追踪：精确跟踪修改状态
     */
    static std::unique_ptr<edit::EditSession> openForEditing(const core::Path& path) {
        try {
            if (!path.exists()) {
                LOG_ERROR("File not found: {}", path.string());
                return nullptr;
            }
            
            return edit::EditSession::fromExistingFile(path);
        } catch (const std::exception& e) {
            LOG_ERROR("Failed to open file for editing: {}, error: {}", path.string(), e.what());
            return nullptr;
        }
    }
    
    /**
     * @brief 从只读工作簿开始编辑
     * @param read_workbook 只读工作簿
     * @param target_path 目标文件路径
     * @return 编辑会话，失败返回nullptr
     * 
     * 使用场景：
     * - 先读取分析，后决定编辑
     * - 另存为新文件
     * - 保护原文件的编辑
     */
    static std::unique_ptr<edit::EditSession> beginEdit(
        const read::ReadWorkbook& read_workbook,
        const core::Path& target_path) {
        try {
            return edit::EditSession::beginEdit(read_workbook, target_path);
        } catch (const std::exception& e) {
            LOG_ERROR("Failed to begin edit session: {}, error: {}", 
                     target_path.string(), e.what());
            return nullptr;
        }
    }
    
    /**
     * @brief 智能打开：根据需求自动选择模式
     * @param path 文件路径
     * @param mode 访问模式
     * @return 对应模式的工作簿
     * 
     * 适用于动态场景，根据运行时条件选择访问模式
     */
    static std::unique_ptr<read::IReadOnlyWorkbook> open(
        const core::Path& path, 
        read::WorkbookAccessMode mode) {
        
        switch (mode) {
            case read::WorkbookAccessMode::READ_ONLY:
                return openForReading(path);
                
            case read::WorkbookAccessMode::EDITABLE:
                return openForEditing(path);
                
            case read::WorkbookAccessMode::CREATE_NEW:
                return createWorkbook(path);
                
            default:
                LOG_ERROR("Unsupported access mode: {}", static_cast<int>(mode));
                return nullptr;
        }
    }
    
    /**
     * @brief 检查文件是否为有效的Excel文件
     * @param path 文件路径
     * @return 是否为有效的Excel文件
     */
    static bool isValidExcelFile(const core::Path& path) {
        if (!path.exists()) {
            return false;
        }
        
        try {
            reader::XLSXReader reader(path);
            auto result = reader.open();
            reader.close();
            return result == core::ErrorCode::Ok;
        } catch (...) {
            return false;
        }
    }
    
    /**
     * @brief 获取Excel文件基本信息（不加载完整内容）
     * @param path 文件路径
     * @return 文件信息，失败返回空结构体
     */
    struct ExcelFileInfo {
        std::vector<std::string> worksheet_names;
        reader::WorkbookMetadata metadata;
        size_t estimated_size = 0;
        bool is_valid = false;
    };
    
    static ExcelFileInfo getFileInfo(const core::Path& path) {
        ExcelFileInfo info;
        
        try {
            if (!path.exists()) {
                return info;
            }
            
            reader::XLSXReader reader(path);
            auto result = reader.open();
            if (result == core::ErrorCode::Ok) {
                reader.getWorksheetNames(info.worksheet_names);
                reader.getMetadata(info.metadata);
                info.estimated_size = path.fileSize();
                info.is_valid = true;
                reader.close();
            }
        } catch (const std::exception& e) {
            LOG_DEBUG("Failed to get file info: {}, error: {}", path.string(), e.what());
        }
        
        return info;
    }
    
    /**
     * @brief 创建高性能行写入器
     * @param session 编辑会话
     * @param worksheet_name 工作表名称
     * @return 行写入器
     * 
     * 使用场景：
     * - 大量数据导出
     * - 流式数据写入
     * - 内存受限环境
     */
    static std::unique_ptr<edit::RowWriter> createRowWriter(
        edit::EditSession* session,
        const std::string& worksheet_name) {
        
        if (!session) {
            LOG_ERROR("Session is null");
            return nullptr;
        }
        
        return session->createRowWriter(worksheet_name);
    }
    
    /**
     * @brief 批量转换Excel文件
     * @param input_paths 输入文件路径列表
     * @param output_dir 输出目录
     * @param converter 转换函数
     * @return 成功转换的文件数
     */
    template<typename Converter>
    static int batchConvert(
        const std::vector<core::Path>& input_paths,
        const core::Path& output_dir,
        Converter converter) {
        
        int success_count = 0;
        
        for (const auto& input_path : input_paths) {
            try {
                auto readonly = openForReading(input_path);
                if (!readonly) continue;
                
                std::string output_name = input_path.stem() + "_converted.xlsx";
                core::Path output_path = output_dir / output_name;
                
                auto editable = createWorkbook(output_path);
                if (!editable) continue;
                
                // 调用用户提供的转换函数
                if (converter(*readonly, *editable)) {
                    if (editable->save()) {
                        success_count++;
                        LOG_INFO("Converted: {} -> {}", 
                                input_path.string(), output_path.string());
                    }
                }
            } catch (const std::exception& e) {
                LOG_ERROR("Failed to convert {}: {}", 
                         input_path.string(), e.what());
            }
        }
        
        return success_count;
    }
    
    /**
     * @brief 获取版本信息
     * @return 版本字符串
     */
    static std::string getVersion() {
        return "3.0.0 - Clean State Architecture";
    }
    
    /**
     * @brief 获取特性列表
     * @return 支持的特性
     */
    static std::vector<std::string> getFeatures() {
        return {
            "Read-only mode with compile-time safety",
            "Lazy loading and streaming",
            "Two-level cache system",
            "Incremental save with DirtyManager",
            "SIMD-accelerated parsing",
            "Parallel compression",
            "Content hashing and deduplication",
            "Memory pool optimization"
        };
    }
};

// 便捷的命名空间别名
namespace read = fastexcel::read;
namespace edit = fastexcel::edit;

} // namespace fastexcel

// 全局便捷函数（可选）
namespace xlsx {
    using fastexcel::FastExcel;
    
    inline auto read(const std::string& path) {
        return FastExcel::openForReading(fastexcel::core::Path(path));
    }
    
    inline auto edit(const std::string& path) {
        return FastExcel::openForEditing(fastexcel::core::Path(path));
    }
    
    inline auto create(const std::string& path) {
        return FastExcel::createWorkbook(fastexcel::core::Path(path));
    }
}
