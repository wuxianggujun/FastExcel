#pragma once

#include "fastexcel/core/Path.hpp"
// 暂时禁用新架构组件以解决编译问题
// #include "fastexcel/read/ReadWorkbook.hpp"
// #include "fastexcel/read/ReadWorksheet.hpp"
// #include "fastexcel/edit/EditSession.hpp"
// #include "fastexcel/edit/RowWriter.hpp"
#include <memory>
#include <string>

namespace fastexcel {

/**
 * @brief FastExcel 统一工厂类
 * 
 * 提供语义化的API接口，彻底解决状态管理混乱问题
 * 
 * 注意：新架构组件暂时禁用以解决编译问题
 */
class FastExcel {
public:
    // 暂时禁用新架构方法以解决编译问题
    /*
    static std::unique_ptr<edit::EditSession> createWorkbook(const core::Path& path);
    static std::unique_ptr<read::ReadWorkbook> openForReading(const core::Path& path);
    static std::unique_ptr<edit::EditSession> openForEditing(const core::Path& path);
    static std::unique_ptr<edit::EditSession> beginEdit(const read::ReadWorkbook& read_workbook, const core::Path& target_path);
    static std::unique_ptr<read::IReadOnlyWorkbook> open(const core::Path& path, read::WorkbookAccessMode mode);
    static bool isValidExcelFile(const core::Path& path);
    static ExcelFileInfo getFileInfo(const core::Path& path);
    */
    
    /**
     * @brief 获取版本信息
     * @return 版本字符串
     */
    static std::string getVersion() {
        return "3.0.0 - Clean State Architecture (Core Components Only)";
    }
    
    /**
     * @brief 获取特性列表
     * @return 支持的特性
     */
    static std::vector<std::string> getFeatures() {
        return {
            "Core Excel processing components",
            "XML generation pipeline",
            "File management system",
            "Style and formatting system",
            "Memory pool optimization",
            "Compression support"
        };
    }
};

} // namespace fastexcel

// 便捷的命名空间别名（在 fastexcel 命名空间外部定义）
// 暂时注释掉以避免编译错误
// namespace fxl_read = fastexcel::read;
// namespace fxl_edit = fastexcel::edit;

// 全局便捷函数（可选）
// 暂时注释掉以避免编译错误
/*
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
*/