#include "fastexcel/core/WorkbookCoordinator.hpp"
#include "fastexcel/core/ResourceManager.hpp"
#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/DirtyManager.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/utils/Logger.hpp"
#include <chrono>
#include <set>

namespace fastexcel {
namespace core {

WorkbookCoordinator::WorkbookCoordinator(Workbook* workbook)
    : workbook_(workbook)
    , resource_manager_(std::make_unique<ResourceManager>()) {
    if (!workbook_) {
        throw std::invalid_argument("Workbook cannot be null");
    }
}

WorkbookCoordinator::~WorkbookCoordinator() {
    clearCache();
}

// ========== 核心保存流程 ==========

bool WorkbookCoordinator::save(const std::string& filename, const SaveStrategy& strategy) {
    auto start_time = std::chrono::steady_clock::now();
    
    // 初始化资源管理器
    if (!initializeResourceManager(filename)) {
        FASTEXCEL_LOG_ERROR("Failed to initialize ResourceManager for: " + filename);
        return false;
    }
    
    // 验证保存前状态
    if (!validateBeforeSave()) {
        FASTEXCEL_LOG_ERROR("Validation failed before save");
        return false;
    }
    
    // 执行保存
    bool success = performSave(strategy);
    
    // 更新统计信息
    auto end_time = std::chrono::steady_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    updateStatistics(stats_.files_written, stats_.bytes_written, duration.count());
    
    if (success) {
        FASTEXCEL_LOG_INFO("Successfully saved workbook to: " + filename);
    } else {
        FASTEXCEL_LOG_ERROR("Failed to save workbook to: " + filename);
    }
    
    return success;
}

bool WorkbookCoordinator::saveAs(const std::string& new_filename, const SaveStrategy& strategy) {
    // 另存为逻辑：创建新的资源管理器
    resource_manager_ = std::make_unique<ResourceManager>(Path(new_filename), ResourceManager::Mode::WRITE_NEW);
    return save(new_filename, strategy);
}

bool WorkbookCoordinator::saveIncremental(const DirtyManager* dirty_manager) {
    if (!dirty_manager) {
        FASTEXCEL_LOG_WARN("DirtyManager is null, falling back to full save");
        return save(workbook_->getFilename());
    }
    
    // 检查是否需要增量保存
    if (!shouldUseIncremental(dirty_manager)) {
        FASTEXCEL_LOG_DEBUG("Incremental save not beneficial, performing full save");
        return save(workbook_->getFilename());
    }
    
    // 确定需要重新生成的部件
    auto parts_to_generate = determinePartsToGenerate(dirty_manager);
    
    if (parts_to_generate.empty()) {
        FASTEXCEL_LOG_DEBUG("No dirty parts to save");
        return true;
    }
    
    FASTEXCEL_LOG_DEBUG("Incremental save: {} parts to regenerate", parts_to_generate.size());
    
    // 创建文件写入器
    auto writer = createFileWriter(false);
    if (!writer) {
        FASTEXCEL_LOG_ERROR("Failed to create file writer");
        return false;
    }
    
    // 生成特定的XML文件
    return generateSpecificXML(*writer, parts_to_generate);
}

// ========== XML生成协调 ==========

bool WorkbookCoordinator::generateAllXML(IFileWriter& writer) {
    auto generator = getOrCreateXMLGenerator();
    if (!generator) {
        FASTEXCEL_LOG_ERROR("Failed to create XML generator");
        return false;
    }
    
    // 使用现有的UnifiedXMLGenerator生成所有XML
    bool success = generator->generateAll(writer);
    
    if (success) {
        auto write_stats = writer.getStats();
        stats_.files_written = write_stats.files_written;
        stats_.bytes_written = write_stats.total_bytes;
        FASTEXCEL_LOG_DEBUG("Generated {} XML files, {} bytes", stats_.files_written, stats_.bytes_written);
    }
    
    return success;
}

bool WorkbookCoordinator::generateSpecificXML(IFileWriter& writer, const std::vector<std::string>& parts) {
    auto generator = getOrCreateXMLGenerator();
    if (!generator) {
        FASTEXCEL_LOG_ERROR("Failed to create XML generator");
        return false;
    }
    
    // 使用现有的UnifiedXMLGenerator生成指定部件
    bool success = generator->generateParts(writer, parts);
    
    if (success) {
        stats_.files_written += parts.size();
        FASTEXCEL_LOG_DEBUG("Generated {} specific XML parts", parts.size());
    }
    
    return success;
}

::fastexcel::xml::UnifiedXMLGenerator* WorkbookCoordinator::getOrCreateXMLGenerator() {
    if (!xml_generator_cache_) {
        stats_.cache_misses++;
        
        // 从workbook创建XML生成器
        xml_generator_cache_ = ::fastexcel::xml::UnifiedXMLGenerator::fromWorkbook(workbook_);
        
        if (xml_generator_cache_) {
            FASTEXCEL_LOG_DEBUG("Created new XML generator (cache miss)");
        }
    } else {
        stats_.cache_hits++;
        FASTEXCEL_LOG_DEBUG("Using cached XML generator (cache hit)");
    }
    
    return xml_generator_cache_.get();
}

// ========== 资源管理协调 ==========

bool WorkbookCoordinator::prepareForEditing(const std::string& original_path) {
    if (!resource_manager_) {
        FASTEXCEL_LOG_ERROR("ResourceManager is null");
        return false;
    }
    
    // 准备编辑模式
    return resource_manager_->prepareForEditing(
        Path(workbook_->getFilename()), 
        original_path
    );
}

bool WorkbookCoordinator::performPassthroughCopy(const std::string& source_path) {
    if (!resource_manager_) {
        FASTEXCEL_LOG_ERROR("ResourceManager is null");
        return false;
    }
    
    // 执行智能透传复制
    return resource_manager_->smartPassthrough(
        Path(source_path),
        true,  // preserve_media
        true   // preserve_vba
    );
}

// ========== 文件写入器工厂 ==========

std::unique_ptr<IFileWriter> WorkbookCoordinator::createFileWriter(bool use_streaming) {
    // 从ResourceManager获取FileManager
    if (!resource_manager_) {
        FASTEXCEL_LOG_ERROR("ResourceManager is null, cannot create file writer");
        return nullptr;
    }
    
    archive::FileManager* file_manager = resource_manager_->getFileManager();
    if (!file_manager) {
        FASTEXCEL_LOG_ERROR("FileManager is null, cannot create file writer");
        return nullptr;
    }
    
    if (use_streaming) {
        FASTEXCEL_LOG_DEBUG("Creating streaming file writer");
        return std::make_unique<StreamingFileWriter>(file_manager);
    } else {
        FASTEXCEL_LOG_DEBUG("Creating batch file writer");
        return std::make_unique<BatchFileWriter>(file_manager);
    }
}

// ========== 性能优化 ==========

void WorkbookCoordinator::warmupCache() {
    // 预创建XML生成器
    if (!xml_generator_cache_) {
        getOrCreateXMLGenerator();
    }
    
    FASTEXCEL_LOG_DEBUG("Cache warmed up");
}

void WorkbookCoordinator::clearCache() {
    xml_generator_cache_.reset();
    stats_.cache_hits = 0;
    stats_.cache_misses = 0;
    
    FASTEXCEL_LOG_DEBUG("Cache cleared");
}

void WorkbookCoordinator::optimizeMemory() {
    // 清理不必要的缓存
    clearCache();
    
    // 触发共享字符串表优化
    if (workbook_->getSharedStringTable()) {
        // 可以在这里添加字符串表压缩逻辑
    }
    
    FASTEXCEL_LOG_DEBUG("Memory optimized");
}

// ========== 私有辅助方法 ==========

bool WorkbookCoordinator::initializeResourceManager(const std::string& filename) {
    if (!resource_manager_) {
        resource_manager_ = std::make_unique<ResourceManager>();
    }
    
    // 根据工作簿状态确定模式
    ResourceManager::Mode mode = ResourceManager::Mode::WRITE_NEW;
    
    if (workbook_->isEditMode()) {
        mode = ResourceManager::Mode::EDIT_EXISTING;
        
        // 准备编辑模式
        if (!resource_manager_->prepareForEditing(Path(filename), workbook_->getOriginalPackagePath())) {
            return false;
        }
    }
    
    // 设置压缩级别
    resource_manager_->setCompressionLevel(config_.compression_level);
    
    return resource_manager_->open(true);
}

bool WorkbookCoordinator::performSave(const SaveStrategy& strategy) {
    // 决定使用哪种写入器
    bool use_streaming = strategy.use_streaming || 
                         shouldUseStreaming(workbook_->getEstimatedSize());
    
    auto writer = createFileWriter(use_streaming);
    if (!writer) {
        FASTEXCEL_LOG_ERROR("Failed to create file writer");
        return false;
    }
    
    // 如果需要保留资源，执行透传复制
    if (strategy.preserve_resources && workbook_->isEditMode()) {
        if (!performPassthroughCopy(workbook_->getOriginalPackagePath())) {
            FASTEXCEL_LOG_WARN("Failed to perform passthrough copy, continuing anyway");
        }
    }
    
    // 生成XML
    bool success = false;
    
    if (strategy.incremental && workbook_->getDirtyManager()) {
        // 增量保存
        auto parts = determinePartsToGenerate(workbook_->getDirtyManager());
        success = generateSpecificXML(*writer, parts);
    } else {
        // 完整保存
        success = generateAllXML(*writer);
    }
    
    if (success) {
        // 使用ResourceManager进行原子保存
        ResourceManager::SaveStrategy rm_strategy;
        rm_strategy.use_temp_file = true;
        rm_strategy.atomic_replace = true;
        rm_strategy.preserve_backup = false;
        
        success = resource_manager_->atomicSave(workbook_, rm_strategy);
    }
    
    return success;
}

bool WorkbookCoordinator::validateBeforeSave() {
    // 验证工作簿状态
    if (!workbook_) {
        FASTEXCEL_LOG_ERROR("Workbook is null");
        return false;
    }
    
    // 验证至少有一个工作表
    if (workbook_->getSheetCount() == 0) {
        FASTEXCEL_LOG_ERROR("Workbook has no sheets");
        return false;
    }
    
    // 验证资源管理器
    if (!resource_manager_) {
        FASTEXCEL_LOG_ERROR("ResourceManager is not initialized");
        return false;
    }
    
    return true;
}

void WorkbookCoordinator::updateStatistics(size_t files, size_t bytes, size_t time_ms) {
    stats_.files_written = files;
    stats_.bytes_written = bytes;
    stats_.time_ms = time_ms;
    
    FASTEXCEL_LOG_DEBUG("Save statistics: {} files, {} bytes, {} ms", files, bytes, time_ms);
}

std::vector<std::string> WorkbookCoordinator::determinePartsToGenerate(const DirtyManager* dirty_manager) {
    std::vector<std::string> parts;
    
    if (!dirty_manager) {
        return parts;
    }
    
    // 根据脏标记确定需要重新生成的部件
    // 使用getChanges方法获取所有变更
    auto changes = dirty_manager->getChanges();
    std::set<std::string> unique_parts;
    
    for (const auto& change : changes.getChanges()) {
        unique_parts.insert(change.part);
    }
    
    // 将set转换为vector
    for (const auto& part : unique_parts) {
        parts.push_back(part);
    }
    
    // 总是重新生成关系和内容类型
    parts.push_back("[Content_Types].xml");
    parts.push_back("_rels/.rels");
    
    return parts;
}

// ========== 智能决策方法 ==========

bool WorkbookCoordinator::shouldUseStreaming(size_t estimated_size) const {
    // 如果估计大小超过50MB，使用流式写入
    const size_t STREAMING_THRESHOLD = 50 * 1024 * 1024;
    return estimated_size > STREAMING_THRESHOLD;
}

bool WorkbookCoordinator::shouldUseIncremental(const DirtyManager* dirty_manager) const {
    if (!dirty_manager) {
        return false;
    }
    
    // 如果脏部件少于总部件的30%，使用增量保存
    size_t total_parts = workbook_->getSheetCount() + 10; // 工作表数 + 固定部件数
    size_t dirty_parts = dirty_manager->getDirtyCount();
    
    return (dirty_parts > 0) && (dirty_parts < total_parts * 0.3);
}

int WorkbookCoordinator::determineOptimalCompressionLevel(size_t file_size) const {
    // 根据文件大小智能选择压缩级别
    if (file_size < 1024 * 1024) {       // < 1MB
        return 9;  // 最大压缩
    } else if (file_size < 10 * 1024 * 1024) {  // < 10MB
        return 6;  // 默认压缩
    } else {
        return 3;  // 快速压缩
    }
}

}} // namespace fastexcel::core
