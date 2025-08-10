#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/xml/DocPropsXMLGenerator.hpp"
#include <sstream>
#include <chrono>

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
    
    auto start_time = std::chrono::high_resolution_clock::now();
    LOG_INFO("Starting Excel structure generation using {}", writer_->getTypeName());
    reportProgress("Initializing", 0, 100);
    
    try {
        // 检查内存限制
        if (options_.max_memory_limit > 0) {
            size_t estimated_memory = estimateWorksheetSize(nullptr) * workbook_->getWorksheetCount();
            if (estimated_memory > options_.max_memory_limit) {
                LOG_WARN("Estimated memory {} exceeds limit {}, forcing streaming mode",
                        estimated_memory, options_.max_memory_limit);
                // 强制使用流式模式
                options_.streaming_threshold = 0;
            }
        }
        
    // 1. 生成基础文件
    reportProgress("Generating basic files", 10, 100);
    auto basic_start = std::chrono::high_resolution_clock::now();
    if (!generateBasicFiles()) {
            LOG_ERROR("Failed to generate basic Excel files");
            return false;
        }
        perf_stats_.basic_files_time = std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::high_resolution_clock::now() - basic_start);
        
        // 2. 生成工作表文件
        reportProgress("Generating worksheets", 50, 100);
        auto worksheets_start = std::chrono::high_resolution_clock::now();
        if (!generateWorksheets()) {
            LOG_ERROR("Failed to generate worksheet files");
            return false;
        }
        perf_stats_.worksheets_time = std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::high_resolution_clock::now() - worksheets_start);
        
        // 3. 最终化
        reportProgress("Finalizing", 90, 100);
        auto finalize_start = std::chrono::high_resolution_clock::now();
        if (!finalize()) {
            LOG_ERROR("Failed to finalize Excel structure generation");
            return false;
        }
        perf_stats_.finalize_time = std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::high_resolution_clock::now() - finalize_start);
        
        reportProgress("Completed", 100, 100);
        
        // 记录总时间和峰值内存
        perf_stats_.total_time = std::chrono::duration_cast<std::chrono::milliseconds>(
            std::chrono::high_resolution_clock::now() - start_time);
        
        // 估算峰值内存使用
        if (auto batch_writer = dynamic_cast<BatchFileWriter*>(writer_.get())) {
            perf_stats_.peak_memory_usage = batch_writer->getEstimatedMemoryUsage();
        }
        
        auto stats = writer_->getStats();
        LOG_INFO("Excel structure generation completed successfully: {} files ({} batch, {} streaming), {} total bytes",
                stats.files_written, stats.batch_files, stats.streaming_files, stats.total_bytes);
        LOG_INFO("Performance: Total time {}ms (basic: {}ms, worksheets: {}ms, finalize: {}ms), Peak memory: {} bytes",
                perf_stats_.total_time.count(), perf_stats_.basic_files_time.count(),
                perf_stats_.worksheets_time.count(), perf_stats_.finalize_time.count(),
                perf_stats_.peak_memory_usage);
        
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
    LOG_DEBUG("Generating basic Excel files via orchestrator");

    // 创建UnifiedXMLGenerator实例
    auto xml_generator = xml::UnifiedXMLGenerator::fromWorkbook(workbook_);
    if (!xml_generator) {
        LOG_ERROR("Failed to create UnifiedXMLGenerator from workbook");
        return false;
    }

    // 构建需要生成的基础部件列表（排除 worksheets 与 sharedStrings）
    std::vector<std::string> parts;
    if (workbook_->shouldGenerateContentTypes()) {
        parts.emplace_back("[Content_Types].xml");
    }
    if (workbook_->shouldGenerateRootRels()) {
        parts.emplace_back("_rels/.rels");
    }
    if (workbook_->shouldGenerateDocPropsApp()) {
        parts.emplace_back("docProps/app.xml");
    }
    if (workbook_->shouldGenerateDocPropsCore()) {
        parts.emplace_back("docProps/core.xml");
    }
    if (workbook_->shouldGenerateDocPropsCustom()) {
        parts.emplace_back("docProps/custom.xml");
    }
    if (workbook_->shouldGenerateWorkbookCore()) {
        parts.emplace_back("xl/_rels/workbook.xml.rels");
        parts.emplace_back("xl/workbook.xml");
    }
    if (workbook_->shouldGenerateStyles()) {
        parts.emplace_back("xl/styles.xml");
    }
    if (workbook_->shouldGenerateTheme()) {
        parts.emplace_back("xl/theme/theme1.xml");
    }

    if (!xml_generator->generateParts(*writer_, parts)) {
        LOG_ERROR("Failed to generate basic parts via orchestrator");
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

    // 统一的 orchestrator（用于生成 per-sheet rels）
    auto xml_generator = xml::UnifiedXMLGenerator::fromWorkbook(workbook_);

    for (size_t i = 0; i < worksheet_count; ++i) {
        auto worksheet = workbook_->getWorksheet(i);
        if (!worksheet) {
            LOG_ERROR("Worksheet {} is null", i);
            return false;
        }
        
        std::string worksheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";

        // 若为透传编辑且该sheet未变更，则跳过生成，保留透传版本
        if (!workbook_->shouldGenerateSheet(i)) {
            LOG_DEBUG("Skip generating sheet{} due to pass-through mode", i + 1);
            continue;
        }
        
        // 统一通过 orchestrator 生成工作表 XML（内部自动采用流式写入）
        if (!xml_generator || !xml_generator->generateParts(*writer_, {worksheet_path})) {
            LOG_ERROR("Failed to generate worksheet via orchestrator: {}", worksheet_path);
            return false;
        }
        
        // 生成工作表关系文件（如有超链接等），交由 orchestrator 封装生成
        if (workbook_->shouldGenerateSheetRels(i) && xml_generator) {
            std::string rels_path = "xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels";
            if (!xml_generator->generateParts(*writer_, {rels_path})) {
                LOG_ERROR("Failed to generate worksheet relations file via orchestrator: {}", rels_path);
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
    LOG_DEBUG("ExcelStructureGenerator::finalize() called");
    
    // 生成共享字符串文件（如果启用）
    // 这必须在所有工作表生成之后进行，因为工作表生成时会填充共享字符串表
    bool should_generate = workbook_->shouldGenerateSharedStrings();
    LOG_DEBUG("workbook_->shouldGenerateSharedStrings() = {}", should_generate);
    
    if (should_generate) {
        LOG_DEBUG("Generating shared strings XML via orchestrator");

        auto xml_generator = xml::UnifiedXMLGenerator::fromWorkbook(workbook_);
        if (!xml_generator) {
            LOG_ERROR("Failed to create UnifiedXMLGenerator for shared strings");
            return false;
        }
        if (!xml_generator->generateParts(*writer_, {"xl/sharedStrings.xml"})) {
            LOG_ERROR("Failed to write shared strings file via orchestrator");
            return false;
        }
        LOG_DEBUG("Shared strings XML generated successfully");
    } else {
        LOG_DEBUG("Skipping SharedStrings generation");
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

void ExcelStructureGenerator::reportProgress(const std::string& stage, int current, int total) {
    if (options_.enable_progress_callback && progress_callback_) {
        progress_callback_(stage, current, total);
    }
}


bool ExcelStructureGenerator::generateFileWithCallback(const std::string& path,
    std::function<void(const std::function<void(const char*, size_t)>&)> generator) {
    
    // 检查是否可以使用流式写入
    if (auto streaming_writer = dynamic_cast<StreamingFileWriter*>(writer_.get())) {
        // 流式模式：直接流式写入，不缓存
        if (!writer_->openStreamingFile(path)) {
            return false;
        }
        
        generator([this](const char* data, size_t size) {
            writer_->writeStreamingChunk(data, size);
        });
        
        return writer_->closeStreamingFile();
    } else {
        // 批量模式：仍需要缓存到字符串
        std::string content;
        generator([&content](const char* data, size_t size) {
            content.append(data, size);
        });
        
        return writer_->writeFile(path, content);
    }
}

bool ExcelStructureGenerator::generateFileWithCallbackIfNotEmpty(const std::string& path,
    std::function<void(const std::function<void(const char*, size_t)>&)> generator) {
    
    // 先检查内容是否为空
    std::string test_content;
    generator([&test_content](const char* data, size_t size) {
        test_content.append(data, size);
    });
    
    if (test_content.empty()) {
        return true; // 内容为空，跳过生成
    }
    
    // 内容不为空，使用正常的生成流程
    return generateFileWithCallback(path, generator);
}

}} // namespace fastexcel::core
