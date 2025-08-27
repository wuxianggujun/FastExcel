#include "ColumnarXLSXReader.hpp"
#include "ColumnarWorksheetParser.hpp"
#include "fastexcel/core/SharedStringTable.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include <stdexcept>

namespace fastexcel {
namespace reader {

ColumnarXLSXReader::ColumnarXLSXReader(const core::columnar::ReadOnlyOptions& options)
    : options_(options) {
}

std::unique_ptr<core::columnar::ReadOnlyWorkbook> ColumnarXLSXReader::parse(const std::string& filename) {
    try {
        // 创建ZipReader
        zip_reader_ = std::make_unique<archive::ZipReader>(core::Path(filename));
        if (!zip_reader_->open()) {
            FASTEXCEL_LOG_ERROR("Failed to open ZIP file: {}", filename);
            return nullptr;
        }
        
        // 创建只读工作簿
        auto workbook = std::make_unique<core::columnar::ReadOnlyWorkbook>(options_);
        
        // 解析步骤
        if (!parseSharedStrings(workbook.get())) {
            FASTEXCEL_LOG_ERROR("Failed to parse shared strings");
            return nullptr;
        }
        
        if (!parseWorkbook(workbook.get())) {
            FASTEXCEL_LOG_ERROR("Failed to parse workbook");
            return nullptr;
        }
        
        if (!parseWorksheets(workbook.get())) {
            FASTEXCEL_LOG_ERROR("Failed to parse worksheets");
            return nullptr;
        }
        
        return workbook;
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Exception during columnar parsing: {}", e.what());
        return nullptr;
    }
}

bool ColumnarXLSXReader::parseStream(const std::string& filename, 
                                    core::columnar::ReadOnlyWorkbook* workbook) {
    // 流式解析实现
    // TODO: 实现流式解析逻辑
    return parse(filename) != nullptr;
}

bool ColumnarXLSXReader::parseSharedStrings(core::columnar::ReadOnlyWorkbook* workbook) {
    // TODO: 实现共享字符串表解析
    // 创建SharedStringTable并设置到workbook
    
    std::string sst_content;
    auto error = zip_reader_->extractFile("xl/sharedStrings.xml", sst_content);
    if (error != archive::ZipError::Ok) {
        // 没有共享字符串表是正常的
        FASTEXCEL_LOG_DEBUG("No shared strings table found");
        return true;
    }
    
    // 解析SST内容
    auto sst = std::make_unique<core::SharedStringTable>();
    // TODO: 实现SST解析
    
    workbook->setSharedStringTable(std::move(sst));
    return true;
}

bool ColumnarXLSXReader::parseWorkbook(core::columnar::ReadOnlyWorkbook* workbook) {
    // TODO: 解析workbook.xml获取工作表信息
    return true;
}

bool ColumnarXLSXReader::parseWorksheets(core::columnar::ReadOnlyWorkbook* workbook) {
    // 获取所有工作表路径
    auto worksheet_paths = getWorksheetPaths();
    
    for (const auto& path : worksheet_paths) {
        std::string sheet_name = getWorksheetName(path);
        auto* worksheet = workbook->addWorksheet(sheet_name);
        
        if (!parseWorksheet(path, worksheet)) {
            FASTEXCEL_LOG_ERROR("Failed to parse worksheet: {}", path);
            return false;
        }
    }
    
    return true;
}

bool ColumnarXLSXReader::parseWorksheet(const std::string& worksheet_path,
                                       core::columnar::ReadOnlyWorksheet* worksheet) {
    // 使用ColumnarWorksheetParser解析工作表
    ColumnarWorksheetParser parser;
    
    // TODO: 获取共享字符串映射
    std::unordered_map<int, std::string> shared_strings;
    
    return parser.parseToColumnar(zip_reader_.get(), worksheet_path, 
                                 worksheet, shared_strings, options_);
}

std::vector<std::string> ColumnarXLSXReader::getWorksheetPaths() const {
    // TODO: 从ZIP文件中获取所有工作表路径
    std::vector<std::string> paths;
    
    // 简单实现：假设有sheet1.xml
    paths.push_back("xl/worksheets/sheet1.xml");
    
    return paths;
}

std::string ColumnarXLSXReader::getWorksheetName(const std::string& worksheet_path) const {
    // TODO: 从路径提取工作表名称，或者从workbook.xml获取
    
    // 简单实现：从路径提取
    size_t pos = worksheet_path.find_last_of('/');
    if (pos != std::string::npos) {
        std::string filename = worksheet_path.substr(pos + 1);
        size_t dot = filename.find_last_of('.');
        if (dot != std::string::npos) {
            return filename.substr(0, dot);
        }
    }
    
    return "Sheet1";  // 默认名称
}

}} // namespace fastexcel::reader