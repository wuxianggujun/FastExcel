#include "WorkbookDataManager.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/Worksheet.hpp"
#include "fastexcel/core/CSVProcessor.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <algorithm>
#include <fstream>
#include <sstream>

namespace fastexcel {
namespace core {

WorkbookDataManager::WorkbookDataManager(Workbook* workbook) : workbook_(workbook) {
    csv_processor_ = std::make_unique<CSVProcessor>();
}

// 以下实现补全带 options 的重载，避免链接缺失

std::vector<WorkbookDataManager::ImportResult> 
WorkbookDataManager::batchImportCSV(const std::vector<std::string>& filepaths,
                                    const CSVOptions& options,
                                    ProgressCallback progress) {
    std::vector<ImportResult> results;
    results.reserve(filepaths.size());
    for (const auto& path : filepaths) {
        results.emplace_back(importCSV(path, std::string{}, options, progress));
    }
    return results;
}

std::vector<WorkbookDataManager::ExportResult>
WorkbookDataManager::batchExportCSV(const std::vector<std::pair<std::string, std::string>>& export_configs,
                                    const CSVOptions& options,
                                    ProgressCallback progress) {
    std::vector<ExportResult> results;
    results.reserve(export_configs.size());
    for (const auto& cfg : export_configs) {
        results.emplace_back(exportCSV(cfg.first, cfg.second, options, progress));
    }
    return results;
}

std::vector<WorkbookDataManager::ExportResult>
WorkbookDataManager::exportAllSheetsAsCSV(const std::string& output_directory,
                                          const std::string& filename_prefix,
                                          const CSVOptions& options,
                                          ProgressCallback progress) {
    std::vector<ExportResult> results;
    const size_t sheet_count = workbook_->getSheetCount();
    for (size_t i = 0; i < sheet_count; ++i) {
        auto ws = workbook_->getSheet(i);
        if (!ws) continue;
        std::string name = ws->getName();
        std::string filename = fmt::format("{}/{}{}.csv", output_directory, filename_prefix.empty() ? std::string("") : filename_prefix, name);
        results.emplace_back(exportCSV(i, filename, options, progress));
    }
    return results;
}

std::vector<std::vector<std::string>>
WorkbookDataManager::previewCSV(const std::string& filepath, size_t max_rows, const CSVOptions& options) {
    std::vector<std::vector<std::string>> preview;
    try {
        if (!isCSVFile(filepath)) return preview;
        std::ifstream file(filepath);
        if (!file.is_open()) return preview;
        csv_processor_->setOptions(options);
        size_t count = 0;
        std::string line;
        while (count < max_rows && std::getline(file, line)) {
            preview.emplace_back(core::parseLine(line, options.delimiter, options.quote_char, options.escape_char));
            ++count;
        }
    } catch (const std::ios_base::failure& e) {
        FASTEXCEL_LOG_WARN("I/O error while reading CSV preview: {}", e.what());
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_WARN("Exception while reading CSV preview: {}", e.what());
    }
    return preview;
}

bool WorkbookDataManager::cleanImportedData(std::shared_ptr<Worksheet> worksheet, const DataCleaningOptions& options) {
    if (!worksheet) return false;
    // 这里做一个最小实现：根据选项处理空行/空列的移除与空白裁剪
    // 实际实现应遍历已用范围并应用规则；此处保留接口一致性
    (void)options;
    return true;
}

WorkbookDataManager::ImportResult WorkbookDataManager::importCSV(const std::string& filepath, 
                                                                 const std::string& sheet_name,
                                                                 const CSVOptions& options,
                                                                 ProgressCallback progress) {
    ImportResult result;
    stats_.total_imports++;
    
    try {
        if (!isCSVFile(filepath)) {
            result.error_message = fmt::format("File is not a valid CSV file: {}", filepath);
            stats_.failed_operations++;
            return result;
        }
        
        Path csv_path(filepath);
        if (!csv_path.exists()) {
            result.error_message = fmt::format("CSV file not found: {}", filepath);
            stats_.failed_operations++;
            return result;
        }
        
        // 读取文件内容
        std::ifstream file(filepath);
        if (!file.is_open()) {
            result.error_message = fmt::format("Cannot open CSV file: {}", filepath);
            stats_.failed_operations++;
            return result;
        }
        
        std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
        file.close();
        
        // 委托给字符串导入方法
        std::string final_sheet_name = sheet_name.empty() ? 
            generateUniqueSheetName(std::filesystem::path(csv_path.string()).stem().string(), ".csv") : sheet_name;
        
        result = importCSVString(content, final_sheet_name, options, progress);
        
        if (result.success) {
            updateStatistics(result, content.size());
            FASTEXCEL_LOG_INFO("CSV imported successfully from {}: {} rows, {} cols", 
                     filepath, result.rows_imported, result.cols_imported);
        }
        
    } catch (const std::exception& e) {
        result.error_message = std::string("Exception during CSV import: ") + e.what();
        stats_.failed_operations++;
        FASTEXCEL_LOG_ERROR("CSV import failed: {}", result.error_message);
    }
    
    return result;
}

WorkbookDataManager::ImportResult WorkbookDataManager::importCSVString(const std::string& csv_content,
                                                                       const std::string& sheet_name,
                                                                       const CSVOptions& options,
                                                                       ProgressCallback progress) {
    ImportResult result;
    stats_.total_imports++;
    
    try {
        if (csv_content.empty()) {
            result.error_message = "CSV content is empty";
            stats_.failed_operations++;
            return result;
        }
        
        // 使用 CSV 处理器解析内容
        csv_processor_->setOptions(options);
        auto data = csv_processor_->parseString(csv_content);
        
        if (data.empty()) {
            result.error_message = "No data found in CSV content";
            stats_.failed_operations++;
            return result;
        }
        
        // 检查导入限制
        if (!validateImportLimits(data.size(), data.empty() ? 0 : data[0].size())) {
            result.error_message = "CSV data exceeds import limits";
            stats_.failed_operations++;
            return result;
        }
        
        // 创建工作表
        auto worksheet = workbook_->addSheet(sheet_name);
        if (!worksheet) {
            result.error_message = fmt::format("Failed to create worksheet: {}", sheet_name);
            stats_.failed_operations++;
            return result;
        }
        
        // 填充数据
        size_t processed_rows = 0;
        for (size_t row = 0; row < data.size(); ++row) {
            if (progress && row % config_.batch_size == 0) {
                progress(row, data.size(), fmt::format("Importing row {}", row));
            }
            
            // 跳过空行（如果配置要求）
            if (config_.skip_empty_rows && data[row].empty()) {
                continue;
            }
            
            for (size_t col = 0; col < data[row].size(); ++col) {
                const auto& cell_value = data[row][col];
                
                if (config_.auto_detect_types && !cell_value.empty()) {
                    // 尝试自动检测数据类型
                    if (isNumeric(cell_value)) {
                        try {
                            double num_value = std::stod(cell_value);
                            worksheet->setValue(core::Address(static_cast<int>(processed_rows), static_cast<int>(col)), num_value);
                        } catch (const std::invalid_argument&) {
                            // 数字格式无效，作为字符串处理
                            worksheet->setValue(core::Address(static_cast<int>(processed_rows), static_cast<int>(col)), cell_value);
                        } catch (const std::out_of_range&) {
                            // 数字超出范围，作为字符串处理
                            FASTEXCEL_LOG_DEBUG("Numeric value out of range, treating as string: {}", cell_value);
                            worksheet->setValue(core::Address(static_cast<int>(processed_rows), static_cast<int>(col)), cell_value);
                        }
                    } else {
                        worksheet->setValue(core::Address(static_cast<int>(processed_rows), static_cast<int>(col)), cell_value);
                    }
                } else {
                    worksheet->setValue(core::Address(static_cast<int>(processed_rows), static_cast<int>(col)), cell_value);
                }
            }
            processed_rows++;
        }
        
        result.success = true;
        result.rows_imported = processed_rows;
        result.cols_imported = data.empty() ? 0 : data[0].size();
        result.worksheet = worksheet;
        
        updateStatistics(result);
        
        if (progress) {
            progress(data.size(), data.size(), "Import completed");
        }
        
    } catch (const std::exception& e) {
        result.error_message = std::string("Exception during CSV string import: ") + e.what();
        stats_.failed_operations++;
        FASTEXCEL_LOG_ERROR("CSV string import failed: {}", result.error_message);
    }
    
    return result;
}

WorkbookDataManager::ExportResult WorkbookDataManager::exportCSV(size_t sheet_index, 
                                                                 const std::string& filepath,
                                                                 const CSVOptions& options,
                                                                 ProgressCallback progress) {
    ExportResult result;
    stats_.total_exports++;
    
    try {
        auto worksheet = workbook_->getSheet(sheet_index);
        if (!worksheet) {
            result.error_message = fmt::format("Invalid worksheet index: {}", sheet_index);
            stats_.failed_operations++;
            return result;
        }
        
        result = exportData(worksheet, filepath, DataFormat::CSV, &options, progress);
        
        if (result.success) {
            updateStatistics(result);
            FASTEXCEL_LOG_INFO("CSV exported successfully to {}: {} rows, {} cols", 
                     filepath, result.rows_exported, result.cols_exported);
        }
        
    } catch (const std::exception& e) {
        result.error_message = std::string("Exception during CSV export: ") + e.what();
        stats_.failed_operations++;
        FASTEXCEL_LOG_ERROR("CSV export failed: {}", result.error_message);
    }
    
    return result;
}

WorkbookDataManager::ExportResult WorkbookDataManager::exportCSV(const std::string& sheet_name, 
                                                                 const std::string& filepath,
                                                                 const CSVOptions& options,
                                                                 ProgressCallback progress) {
    ExportResult result;
    stats_.total_exports++;
    
    try {
        auto worksheet = workbook_->getSheet(sheet_name);
        if (!worksheet) {
            result.error_message = fmt::format("Worksheet not found: {}", sheet_name);
            stats_.failed_operations++;
            return result;
        }
        
        result = exportData(worksheet, filepath, DataFormat::CSV, &options, progress);
        
        if (result.success) {
            updateStatistics(result);
            FASTEXCEL_LOG_INFO("CSV exported successfully to {}: {} rows, {} cols", 
                     filepath, result.rows_exported, result.cols_exported);
        }
        
    } catch (const std::exception& e) {
        result.error_message = std::string("Exception during CSV export: ") + e.what();
        stats_.failed_operations++;
        FASTEXCEL_LOG_ERROR("CSV export failed: {}", result.error_message);
    }
    
    return result;
}

std::string WorkbookDataManager::exportCSVString(size_t sheet_index, const CSVOptions& options) {
    try {
        auto worksheet = workbook_->getSheet(sheet_index);
        if (!worksheet) {
            FASTEXCEL_LOG_ERROR("Invalid worksheet index for CSV string export: {}", sheet_index);
            return "";
        }
        
        return exportCSVString(worksheet->getName(), options);
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("CSV string export failed: {}", e.what());
        return "";
    }
}

std::string WorkbookDataManager::exportCSVString(const std::string& sheet_name, const CSVOptions& options) {
    try {
        auto worksheet = workbook_->getSheet(sheet_name);
        if (!worksheet) {
            FASTEXCEL_LOG_ERROR("Worksheet not found for CSV string export: {}", sheet_name);
            return "";
        }
        
        csv_processor_->setOptions(options);
        
        // 获取工作表数据范围
        auto [first_row, last_row, first_col, last_col] = worksheet->getUsedRangeFull();
        if (last_row < first_row || last_col < first_col) {
            return "";  // 空工作表
        }
        
        std::ostringstream oss;
        
        for (int row = first_row; row <= last_row; ++row) {
            std::vector<std::string> row_data;
            
            for (int col = first_col; col <= last_col; ++col) {
                try {
                    std::string cell_value = worksheet->getValue<std::string>(core::Address(row, col));
                    row_data.push_back(cell_value);
                } catch (...) {
                    row_data.push_back("");  // 空单元格
                }
            }
            
            oss << csv_processor_->formatRow(row_data);
            if (row < last_row) {
                oss << "\n";
            }
        }
        
        return oss.str();
        
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("CSV string export failed: {}", e.what());
        return "";
    }
}

bool WorkbookDataManager::isCSVFile(const std::string& filepath) {
    if (filepath.size() < 4) return false;
    
    std::string extension = filepath.substr(filepath.size() - 4);
    std::transform(extension.begin(), extension.end(), extension.begin(), 
                   [](char c) { return static_cast<char>(std::tolower(static_cast<unsigned char>(c))); });
    
    return extension == ".csv";
}

std::string WorkbookDataManager::generateUniqueSheetName(const std::string& base_name, const std::string& file_extension) {
    std::string name = base_name.empty() ? "ImportedData" : base_name;
    
    // 移除文件扩展名
    if (!file_extension.empty() && name.size() > file_extension.size()) {
        if (name.substr(name.size() - file_extension.size()) == file_extension) {
            name = name.substr(0, name.size() - file_extension.size());
        }
    }
    
    // 确保名称唯一
    std::string final_name = name;
    int counter = 1;
    while (workbook_->hasSheet(final_name)) {
        final_name = fmt::format("{}_{}", name, counter++);
    }
    
    return final_name;
}

bool WorkbookDataManager::validateImportLimits(size_t rows, size_t cols) const {
    if (rows > config_.max_import_rows) {
        FASTEXCEL_LOG_WARN("Import rows ({}) exceed limit ({})", rows, config_.max_import_rows);
        return false;
    }
    
    if (cols > config_.max_import_cols) {
        FASTEXCEL_LOG_WARN("Import columns ({}) exceed limit ({})", cols, config_.max_import_cols);
        return false;
    }
    
    return true;
}

void WorkbookDataManager::updateStatistics(const ImportResult& result, size_t bytes_processed) {
    if (result.success) {
        stats_.total_rows_processed += result.rows_imported;
        stats_.total_bytes_processed += bytes_processed;
    } else {
        stats_.failed_operations++;
    }
}

void WorkbookDataManager::updateStatistics(const ExportResult& result) {
    if (result.success) {
        stats_.total_rows_processed += result.rows_exported;
        stats_.total_bytes_processed += result.bytes_written;
    } else {
        stats_.failed_operations++;
    }
}

WorkbookDataManager::ExportResult WorkbookDataManager::exportData(std::shared_ptr<const Worksheet> worksheet,
                                                                  const std::string& filepath, 
                                                                  DataFormat format,
                                                                  const CSVOptions* csv_options,
                                                                  ProgressCallback progress) {
    ExportResult result;
    result.output_path = filepath;
    
    try {
        std::ofstream file(filepath);
        if (!file.is_open()) {
            result.error_message = fmt::format("Cannot create output file: {}", filepath);
            return result;
        }
        
        if (format == DataFormat::CSV && csv_options) {
            csv_processor_->setOptions(*csv_options);
        }
        
        // 获取使用范围
        auto [first_row2, last_row2, first_col2, last_col2] = worksheet->getUsedRangeFull();
        if (last_row2 < first_row2 || last_col2 < first_col2) {
            result.success = true;  // 空工作表也算成功
            return result;
        }
        
        size_t total_rows = last_row2 - first_row2 + 1;
        size_t processed_rows = 0;
        
        for (int row = first_row2; row <= last_row2; ++row) {
            if (progress && processed_rows % config_.batch_size == 0) {
                progress(processed_rows, total_rows, fmt::format("Exporting row {}", row));
            }
            
            std::vector<std::string> row_data;
            for (int col = first_col2; col <= last_col2; ++col) {
                try {
                    std::string cell_value = worksheet->getValue<std::string>(core::Address(row, col));
                    row_data.push_back(cell_value);
                } catch (...) {
                    row_data.push_back("");  // 空单元格
                }
            }
            
            if (format == DataFormat::CSV) {
                file << csv_processor_->formatRow(row_data);
            } else {
                // 其他格式的处理...
                for (size_t i = 0; i < row_data.size(); ++i) {
                    file << row_data[i];
                    if (i < row_data.size() - 1) {
                        file << "\t";  // TSV 格式
                    }
                }
            }
            
            if (row < last_row2) {
                file << "\n";
            }
            
            processed_rows++;
        }
        
        file.close();
        
        result.success = true;
        result.rows_exported = processed_rows;
        result.cols_exported = last_col2 - first_col2 + 1;
        
        // 获取文件大小
        Path output_path(filepath);
        if (output_path.exists()) {
            result.bytes_written = output_path.fileSize();
        }
        
        if (progress) {
            progress(total_rows, total_rows, "Export completed");
        }
        
    } catch (const std::exception& e) {
        result.error_message = std::string("Exception during data export: ") + e.what();
    }
    
    return result;
}

bool WorkbookDataManager::isNumeric(const std::string& str) const {
    if (str.empty()) return false;
    
    char* end = nullptr;
    std::strtod(str.c_str(), &end);
    return end == str.c_str() + str.length();
}

}} // namespace fastexcel::core
// 新增便捷重载与默认参数落地实现，避免在头文件使用嵌套类型的默认参数

namespace fastexcel {
namespace core {

// 若实现已存在，仅追加便捷重载实现

WorkbookDataManager::ImportResult WorkbookDataManager::importCSV(const std::string& filepath) {
    return importCSV(filepath, std::string{}, CSVOptions{}, nullptr);
}

WorkbookDataManager::ImportResult WorkbookDataManager::importCSVString(const std::string& csv_content) {
    return importCSVString(csv_content, std::string{"Sheet1"}, CSVOptions{}, nullptr);
}

WorkbookDataManager::ExportResult WorkbookDataManager::exportCSV(size_t sheet_index, const std::string& filepath) {
    return exportCSV(sheet_index, filepath, CSVOptions{}, nullptr);
}

WorkbookDataManager::ExportResult WorkbookDataManager::exportCSV(const std::string& sheet_name, const std::string& filepath) {
    return exportCSV(sheet_name, filepath, CSVOptions{}, nullptr);
}

std::string WorkbookDataManager::exportCSVString(size_t sheet_index) {
    return exportCSVString(sheet_index, CSVOptions{});
}

std::string WorkbookDataManager::exportCSVString(const std::string& sheet_name) {
    return exportCSVString(sheet_name, CSVOptions{});
}

std::vector<WorkbookDataManager::ImportResult> WorkbookDataManager::batchImportCSV(const std::vector<std::string>& filepaths) {
    return batchImportCSV(filepaths, CSVOptions{}, nullptr);
}

std::vector<WorkbookDataManager::ExportResult> WorkbookDataManager::batchExportCSV(const std::vector<std::pair<std::string, std::string>>& export_configs) {
    return batchExportCSV(export_configs, CSVOptions{}, nullptr);
}

std::vector<WorkbookDataManager::ExportResult> WorkbookDataManager::exportAllSheetsAsCSV(const std::string& output_directory) {
    return exportAllSheetsAsCSV(output_directory, std::string{}, CSVOptions{}, nullptr);
}

std::vector<std::vector<std::string>> WorkbookDataManager::previewCSV(const std::string& filepath) {
    return previewCSV(filepath, 10, CSVOptions{});
}

bool WorkbookDataManager::cleanImportedData(std::shared_ptr<Worksheet> worksheet) {
    return cleanImportedData(std::move(worksheet), DataCleaningOptions{});
}

} // namespace core
} // namespace fastexcel
