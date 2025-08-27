#include "ReadOnlyWorkbook.hpp"
#include "fastexcel/reader/ColumnarXLSXReader.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <stdexcept>

namespace fastexcel {
namespace core {
namespace columnar {

ReadOnlyWorkbook::ReadOnlyWorkbook(const ReadOnlyOptions& options)
    : options_(options) {
}

std::unique_ptr<ReadOnlyWorkbook> ReadOnlyWorkbook::openReadOnly(const std::string& filename,
                                                                const ReadOnlyOptions& options) {
    auto workbook = std::make_unique<ReadOnlyWorkbook>(options);
    workbook->filename_ = filename;
    
    try {
        // 使用ColumnarXLSXReader解析
        reader::ColumnarXLSXReader reader(options);
        return reader.parse(filename);
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to open file in read-only mode: {} - {}", filename, e.what());
        return nullptr;
    }
}

ReadOnlyWorksheet* ReadOnlyWorkbook::addWorksheet(const std::string& name) {
    auto worksheet = std::make_unique<ReadOnlyWorksheet>(name, sst_.get(), format_repo_.get());
    ReadOnlyWorksheet* result = worksheet.get();
    
    worksheet_name_index_[name] = worksheets_.size();
    worksheets_.push_back(std::move(worksheet));
    
    return result;
}

void ReadOnlyWorkbook::setSharedStringTable(std::unique_ptr<SharedStringTable> sst) {
    sst_ = std::move(sst);
    
    // 更新所有工作表的SST引用
    for ([[maybe_unused]] auto& ws : worksheets_) {
        // ReadOnlyWorksheet需要提供updateSST方法
        // ws->updateSST(sst_.get());
    }
}

void ReadOnlyWorkbook::setFormatRepository(std::unique_ptr<FormatRepository> format_repo) {
    format_repo_ = std::move(format_repo);
    
    // 更新所有工作表的格式仓库引用
    for ([[maybe_unused]] auto& ws : worksheets_) {
        // ReadOnlyWorksheet需要提供updateFormatRepository方法
        // ws->updateFormatRepository(format_repo_.get());
    }
}

}}} // namespace fastexcel::core::columnar