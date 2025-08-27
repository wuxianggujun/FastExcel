/**
 * @file ReadOnlyXLSXReader.hpp
 * @brief 只读模式专用XLSX解析器 - 直接解析到列式存储
 */

#pragma once

#include "fastexcel/core/ErrorCode.hpp"
#include "fastexcel/core/ColumnarStorageManager.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/WorkbookTypes.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <memory>
#include <vector>
#include <string>
#include <unordered_map>

namespace fastexcel {
namespace reader {

/**
 * @brief 只读模式XLSX解析器工作表信息
 */
struct ReadOnlyWorksheetInfo {
    std::string name;
    std::shared_ptr<core::ColumnarStorageManager> storage_manager;
    int first_row{0};
    int first_col{0};  
    int last_row{0};
    int last_col{0};
    
    ReadOnlyWorksheetInfo(std::string sheet_name, 
                         std::shared_ptr<core::ColumnarStorageManager> storage,
                         int fr, int fc, int lr, int lc)
        : name(std::move(sheet_name)), storage_manager(storage),
          first_row(fr), first_col(fc), last_row(lr), last_col(lc) {}
};

/**
 * @brief 专用于只读模式的XLSX解析器
 * 直接解析XLSX文件到列式存储，完全避免创建Cell和Workbook对象
 */
class ReadOnlyXLSXReader {
private:
    core::Path path_;
    const core::WorkbookOptions* options_;
    std::unique_ptr<archive::ZipArchive> zip_archive_;
    std::unordered_map<int, std::string_view> shared_strings_;  // 使用 string_view 提升性能
    std::vector<ReadOnlyWorksheetInfo> worksheet_infos_;
    
    // 内部解析方法
    core::ErrorCode parseSharedStrings();
    core::ErrorCode parseWorkbook();
    core::ErrorCode parseWorksheet(const std::string& worksheet_path, 
                                  const std::string& worksheet_name,
                                  std::shared_ptr<core::ColumnarStorageManager> storage);
    
    // 辅助方法
    std::string extractXMLFromZip(const std::string& path);
    
public:
    explicit ReadOnlyXLSXReader(const core::Path& file_path, 
                               const core::WorkbookOptions* options = nullptr);
    ~ReadOnlyXLSXReader() = default;
    
    // 禁用拷贝
    ReadOnlyXLSXReader(const ReadOnlyXLSXReader&) = delete;
    ReadOnlyXLSXReader& operator=(const ReadOnlyXLSXReader&) = delete;
    
    /**
     * @brief 解析XLSX文件到列式存储
     * @return 解析结果
     */
    core::ErrorCode parse();
    
    /**
     * @brief 获取解析后的工作表信息
     */
    const std::vector<ReadOnlyWorksheetInfo>& getWorksheetInfos() const {
        return worksheet_infos_;
    }
    
    /**
     * @brief 移动工作表信息所有权
     */
    std::vector<ReadOnlyWorksheetInfo> takeWorksheetInfos() {
        return std::move(worksheet_infos_);
    }
    
    /**
     * @brief 获取共享字符串表
     */
    const std::unordered_map<int, std::string_view>& getSharedStrings() const {
        return shared_strings_;
    }
};

}} // namespace fastexcel::reader