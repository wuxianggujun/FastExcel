#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Exception.hpp"
#include <sstream>

namespace fastexcel {
namespace core {

ExcelStructureGenerator::ExcelStructureGenerator(const Workbook* workbook, std::unique_ptr<IFileWriter> writer)
    : workbook_(workbook), writer_(std::move(writer)) {
    if (!workbook_) {
        FASTEXCEL_THROW_PARAM("Workbook cannot be null");
    }
    if (!writer_) {
        FASTEXCEL_THROW_PARAM("FileWriter cannot be null");
    }
}

ExcelStructureGenerator::~ExcelStructureGenerator() = default;

bool ExcelStructureGenerator::generate() {
    if (!workbook_ || !writer_) {
        LOG_ERROR("ExcelStructureGenerator not properly initialized");
        return false;
    }
    
    LOG_INFO("Starting Excel structure generation using {}", writer_->getTypeName());
    reportProgress("Initializing", 0, 100);
    
    try {
        // 1. 生成基础文件
        reportProgress("Generating basic files", 10, 100);
        if (!generateBasicFiles()) {
            LOG_ERROR("Failed to generate basic Excel files");
            return false;
        }
        
        // 2. 生成工作表文件
        reportProgress("Generating worksheets", 50, 100);
        if (!generateWorksheets()) {
            LOG_ERROR("Failed to generate worksheet files");
            return false;
        }
        
        // 3. 最终化
        reportProgress("Finalizing", 90, 100);
        if (!finalize()) {
            LOG_ERROR("Failed to finalize Excel structure generation");
            return false;
        }
        
        reportProgress("Completed", 100, 100);
        
        auto stats = writer_->getStats();
        LOG_INFO("Excel structure generation completed successfully: {} files ({} batch, {} streaming), {} total bytes",
                stats.files_written, stats.batch_files, stats.streaming_files, stats.total_bytes);
        
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during Excel structure generation: {}", e.what());
        return false;
    }
}

void ExcelStructureGenerator::setProgressCallback(std::function<void(const std::string&, int, int)> callback) {
    progress_callback_ = callback;
    options_.enable_progress_callback = true;
}

IFileWriter::WriteStats ExcelStructureGenerator::getWriterStats() const {
    return writer_ ? writer_->getStats() : IFileWriter::WriteStats{};
}

std::string ExcelStructureGenerator::getGeneratorType() const {
    return writer_ ? writer_->getTypeName() : "Unknown";
}

bool ExcelStructureGenerator::generateBasicFiles() {
    // 预估基础文件数量：8个基础文件 + 可能的自定义属性文件
    size_t estimated_files = 8;
    if (!workbook_->getDocumentProperties().title.empty() || 
        !workbook_->getCustomProperty("").empty()) { // 简化检查
        estimated_files++;
    }
    
    LOG_DEBUG("Generating {} basic Excel files", estimated_files);
    
    // 生成 [Content_Types].xml
    std::string content_types_xml;
    workbook_->generateContentTypesXML([&content_types_xml](const char* data, size_t size) {
        content_types_xml.append(data, size);
    });
    if (!writer_->writeFile("[Content_Types].xml", content_types_xml)) {
        return false;
    }
    
    // 生成 _rels/.rels
    std::string rels_xml;
    workbook_->generateRelsXML([&rels_xml](const char* data, size_t size) {
        rels_xml.append(data, size);
    });
    if (!writer_->writeFile("_rels/.rels", rels_xml)) {
        return false;
    }
    
    // 生成 docProps/app.xml
    std::string app_xml;
    workbook_->generateDocPropsAppXML([&app_xml](const char* data, size_t size) {
        app_xml.append(data, size);
    });
    if (!writer_->writeFile("docProps/app.xml", app_xml)) {
        return false;
    }
    
    // 生成 docProps/core.xml
    std::string core_xml;
    workbook_->generateDocPropsCoreXML([&core_xml](const char* data, size_t size) {
        core_xml.append(data, size);
    });
    if (!writer_->writeFile("docProps/core.xml", core_xml)) {
        return false;
    }
    
    // 生成自定义属性（如果有）
    std::string custom_xml;
    workbook_->generateDocPropsCustomXML([&custom_xml](const char* data, size_t size) {
        custom_xml.append(data, size);
    });
    if (!custom_xml.empty()) {
        if (!writer_->writeFile("docProps/custom.xml", custom_xml)) {
            return false;
        }
    }
    
    // 生成 xl/_rels/workbook.xml.rels
    std::string workbook_rels_xml;
    workbook_->generateWorkbookRelsXML([&workbook_rels_xml](const char* data, size_t size) {
        workbook_rels_xml.append(data, size);
    });
    if (!writer_->writeFile("xl/_rels/workbook.xml.rels", workbook_rels_xml)) {
        return false;
    }
    
    // 生成 xl/workbook.xml
    std::string workbook_xml;
    workbook_->generateWorkbookXML([&workbook_xml](const char* data, size_t size) {
        workbook_xml.append(data, size);
    });
    if (!writer_->writeFile("xl/workbook.xml", workbook_xml)) {
        return false;
    }
    
    // 生成 xl/styles.xml
    std::string styles_xml;
    workbook_->generateStylesXML([&styles_xml](const char* data, size_t size) {
        styles_xml.append(data, size);
    });
    if (!writer_->writeFile("xl/styles.xml", styles_xml)) {
        return false;
    }
    
    // 注意：sharedStrings.xml 将在所有工作表生成后，在 finalize() 中生成
    // 这样可以确保共享字符串表已经被填充
    
    // 生成 xl/theme/theme1.xml
    std::string theme_xml;
    workbook_->generateThemeXML([&theme_xml](const char* data, size_t size) {
        theme_xml.append(data, size);
    });
    if (!writer_->writeFile("xl/theme/theme1.xml", theme_xml)) {
        return false;
    }
    
    LOG_DEBUG("Successfully generated basic Excel files");
    return true;
}

bool ExcelStructureGenerator::generateWorksheets() {
    size_t worksheet_count = workbook_->getWorksheetCount();
    if (worksheet_count == 0) {
        LOG_WARN("No worksheets to generate");
        return true;
    }
    
    LOG_DEBUG("Generating {} worksheets", worksheet_count);
    
    for (size_t i = 0; i < worksheet_count; ++i) {
        auto worksheet = workbook_->getWorksheet(i);
        if (!worksheet) {
            LOG_ERROR("Worksheet {} is null", i);
            return false;
        }
        
        std::string worksheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
        
        // 智能决策：根据工作表大小选择生成方式
        if (shouldUseStreamingForWorksheet(worksheet)) {
            LOG_DEBUG("Using streaming mode for large worksheet: {}", worksheet->getName());
            if (!generateWorksheetStreaming(worksheet, worksheet_path)) {
                LOG_ERROR("Failed to generate worksheet {} in streaming mode", worksheet->getName());
                return false;
            }
        } else {
            LOG_DEBUG("Using batch mode for small worksheet: {}", worksheet->getName());
            if (!generateWorksheetBatch(worksheet, worksheet_path)) {
                LOG_ERROR("Failed to generate worksheet {} in batch mode", worksheet->getName());
                return false;
            }
        }
        
        // 生成工作表关系文件（如果有超链接等）
        std::string worksheet_rels_xml;
        worksheet->generateRelsXML([&worksheet_rels_xml](const char* data, size_t size) {
            worksheet_rels_xml.append(data, size);
        });
        if (!worksheet_rels_xml.empty()) {
            std::string rels_path = "xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels";
            if (!writer_->writeFile(rels_path, worksheet_rels_xml)) {
                LOG_ERROR("Failed to generate worksheet relations file: {}", rels_path);
                return false;
            }
        }
        
        // 报告进度
        int progress = 50 + static_cast<int>((i + 1) * 40 / worksheet_count);
        reportProgress("Generating worksheets", progress, 100);
    }
    
    LOG_DEBUG("Successfully generated all worksheets");
    return true;
}

bool ExcelStructureGenerator::finalize() {
    // 生成共享字符串文件（如果启用）
    // 这必须在所有工作表生成之后进行，因为工作表生成时会填充共享字符串表
    if (workbook_->getOptions().use_shared_strings) {
        LOG_DEBUG("Generating shared strings XML");
        std::string shared_strings_xml;
        workbook_->generateSharedStringsXML([&shared_strings_xml](const char* data, size_t size) {
            shared_strings_xml.append(data, size);
        });
        
        // 即使是空的共享字符串表也要生成文件，因为Content_Types.xml和workbook.xml.rels已经引用了它
        if (!writer_->writeFile("xl/sharedStrings.xml", shared_strings_xml)) {
            LOG_ERROR("Failed to write shared strings file");
            return false;
        }
        LOG_DEBUG("Shared strings XML generated successfully");
    }
    
    // 对于批量模式，需要调用flush
    if (auto batch_writer = dynamic_cast<BatchFileWriter*>(writer_.get())) {
        LOG_DEBUG("Flushing batch writer");
        return batch_writer->flush();
    }
    
    // 流式模式不需要特殊的最终化操作
    LOG_DEBUG("Finalization completed for streaming writer");
    return true;
}

bool ExcelStructureGenerator::shouldUseStreamingForWorksheet(const std::shared_ptr<const Worksheet>& worksheet) const {
    if (!worksheet) {
        return false;
    }
    
    // 估算工作表大小
    size_t estimated_size = estimateWorksheetSize(worksheet);
    
    // 根据阈值决定
    bool use_streaming = estimated_size > options_.streaming_threshold;
    
    LOG_DEBUG("Worksheet '{}' estimated size: {} cells, using {} mode",
             worksheet->getName(), estimated_size, use_streaming ? "streaming" : "batch");
    
    return use_streaming;
}

bool ExcelStructureGenerator::generateWorksheetStreaming(const std::shared_ptr<const Worksheet>& worksheet, const std::string& path) {
    if (!writer_->openStreamingFile(path)) {
        return false;
    }
    
    // 使用工作表的流式XML生成
    bool success = true;
    worksheet->generateXML([this, &success](const char* data, size_t size) {
        if (success && !writer_->writeStreamingChunk(data, size)) {
            success = false;
        }
    });
    
    if (!writer_->closeStreamingFile()) {
        success = false;
    }
    
    return success;
}

bool ExcelStructureGenerator::generateWorksheetBatch(const std::shared_ptr<const Worksheet>& worksheet, const std::string& path) {
    std::string worksheet_xml;
    worksheet->generateXML([&worksheet_xml](const char* data, size_t size) {
        worksheet_xml.append(data, size);
    });
    
    // 可选：验证XML
    if (options_.validate_xml && !validateXML(worksheet_xml, path)) {
        LOG_WARN("XML validation failed for worksheet: {}", path);
    }
    
    return writer_->writeFile(path, worksheet_xml);
}

void ExcelStructureGenerator::reportProgress(const std::string& stage, int current, int total) {
    if (options_.enable_progress_callback && progress_callback_) {
        progress_callback_(stage, current, total);
    }
}

size_t ExcelStructureGenerator::estimateWorksheetSize(const std::shared_ptr<const Worksheet>& worksheet) const {
    if (!worksheet) {
        return 0;
    }
    
    // 如果工作表支持优化模式，使用精确的单元格计数
    if (worksheet->isOptimizeMode()) {
        return worksheet->getCellCount();
    }
    
    // 否则使用使用范围估算
    auto [max_row, max_col] = worksheet->getUsedRange();
    if (max_row < 0 || max_col < 0) {
        return 0;
    }
    
    // 估算实际有数据的单元格数量（不是整个矩形区域）
    size_t estimated_cells = 0;
    for (int row = 0; row <= max_row && row < 1000; ++row) { // 限制扫描范围以提高性能
        for (int col = 0; col <= max_col && col < 100; ++col) {
            if (worksheet->hasCellAt(row, col)) {
                estimated_cells++;
            }
        }
    }
    
    // 如果只扫描了部分区域，按比例估算
    if (max_row >= 1000 || max_col >= 100) {
        double scan_ratio = std::min(1000.0 / (max_row + 1), 100.0 / (max_col + 1));
        estimated_cells = static_cast<size_t>(estimated_cells / scan_ratio);
    }
    
    return estimated_cells;
}

bool ExcelStructureGenerator::validateXML(const std::string& xml_content, const std::string& file_path) const {
    // 简单的XML验证：检查基本结构
    if (xml_content.empty()) {
        LOG_WARN("Empty XML content for file: {}", file_path);
        return false;
    }
    
    // 检查XML声明
    if (xml_content.find("<?xml") == std::string::npos) {
        LOG_WARN("Missing XML declaration in file: {}", file_path);
        return false;
    }
    
    // 检查基本的标签匹配（简化版本）
    size_t open_tags = 0;
    size_t close_tags = 0;
    for (size_t i = 0; i < xml_content.length(); ++i) {
        if (xml_content[i] == '<') {
            if (i + 1 < xml_content.length() && xml_content[i + 1] == '/') {
                close_tags++;
            } else if (i + 1 < xml_content.length() && xml_content[i + 1] != '?') {
                // 检查是否是自闭合标签
                size_t tag_end = xml_content.find('>', i);
                if (tag_end != std::string::npos && tag_end > 0 && xml_content[tag_end - 1] != '/') {
                    open_tags++;
                }
            }
        }
    }
    
    if (open_tags != close_tags) {
        LOG_WARN("Mismatched XML tags in file: {} (open: {}, close: {})", file_path, open_tags, close_tags);
        return false;
    }
    
    return true;
}

}} // namespace fastexcel::core