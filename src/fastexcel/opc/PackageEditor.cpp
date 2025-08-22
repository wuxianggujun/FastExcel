#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/opc/PackageEditor.hpp"
#include "fastexcel/opc/PackageEditorManager.hpp"
#include "fastexcel/tracking/StandardChangeTracker.hpp"
#include "fastexcel/opc/ZipRepackWriter.hpp"
#include "fastexcel/archive/ZipReader.hpp"
#include "fastexcel/archive/ZipArchive.hpp"
#include "fastexcel/core/IFileWriter.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/core/Workbook.hpp"
#include <sstream>
#include <unordered_set>
#include <algorithm>

namespace fastexcel {
namespace opc {

// ===== PackageEditor 实现 =====

PackageEditor::PackageEditor() = default;
PackageEditor::~PackageEditor() = default;

// 工厂方法
std::unique_ptr<PackageEditor> PackageEditor::open(const core::Path& xlsx_path) {
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    
    // 创建ZIP读取器
    auto zip_reader = std::make_unique<archive::ZipReader>(xlsx_path);
    if (!zip_reader->open()) {
        FASTEXCEL_LOG_ERROR("Failed to open ZIP file: {}", xlsx_path.string());
        return nullptr;
    }
    
    editor->initializeServices(std::move(zip_reader), nullptr);
    FASTEXCEL_LOG_INFO("Opened Excel package: {}", xlsx_path.string());
    return editor;
}

std::unique_ptr<PackageEditor> PackageEditor::fromWorkbook(core::Workbook* workbook) {
    if (!workbook) {
        FASTEXCEL_LOG_ERROR("Cannot create PackageEditor from null Workbook");
        return nullptr;
    }
    
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    editor->initializeServices(nullptr, workbook);
    
    FASTEXCEL_LOG_INFO("Created PackageEditor from Workbook with {} sheets", 
             workbook->getSheetNames().size());
    return editor;
}

std::unique_ptr<PackageEditor> PackageEditor::create() {
    auto editor = std::unique_ptr<PackageEditor>(new PackageEditor());
    
    // 创建新的工作簿
    auto workbook = std::make_unique<core::Workbook>(core::Path("new_workbook.xlsx"));
    if (!workbook->open()) {
        FASTEXCEL_LOG_ERROR("Failed to create new Workbook");
        return nullptr;
    }
    
    // 添加默认工作表
    workbook->addSheet("Sheet1");
    
    editor->initializeServices(nullptr, workbook.release());
    FASTEXCEL_LOG_INFO("Created new Excel package with default sheet");
    return editor;
}

bool PackageEditor::initializeServices(std::unique_ptr<archive::ZipReader> zip_reader, core::Workbook* workbook) {
    workbook_ = workbook;
    
    // 初始化包管理器 - 使用独立的PackageEditorManager类
    package_manager_ = std::make_unique<PackageEditorManager>(std::move(zip_reader));
    
    // 初始化统一XML生成器
    if (workbook_) {
        xml_generator_ = xml::UnifiedXMLGenerator::fromWorkbook(workbook_);
    } else {
        xml_generator_ = xml::XMLGeneratorFactory::createLightweightGenerator();
    }
    
    // 初始化变更跟踪器 - 使用独立的StandardChangeTracker类
    change_tracker_ = std::make_unique<tracking::StandardChangeTracker>();
    
    // 如果是从工作簿创建，标记初始dirty状态
    if (workbook_) {
        change_tracker_->markPartDirty("xl/workbook.xml");
        change_tracker_->markPartDirty("xl/styles.xml");
        change_tracker_->markPartDirty("[Content_Types].xml");
        change_tracker_->markPartDirty("_rels/.rels");
        change_tracker_->markPartDirty("xl/_rels/workbook.xml.rels");
        
        // 为每个工作表标记dirty
        auto sheet_names = workbook_->getSheetNames();
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
            change_tracker_->markPartDirty(sheet_path);
        }
        
        if (workbook_->getOptions().use_shared_strings) {
            change_tracker_->markPartDirty("xl/sharedStrings.xml");
        }
    }
    
    initialized_ = true;
    return true;
}

void PackageEditor::detectChanges() {
    if (!workbook_ || !change_tracker_) return;
    
    // 检测工作簿是否有修改
    if (workbook_->isModified()) {
        FASTEXCEL_LOG_DEBUG("Detected workbook modifications, marking dirty parts");
        
        // 标记核心部件
        change_tracker_->markPartDirty("xl/workbook.xml");
        change_tracker_->markPartDirty("xl/styles.xml");
        
        // 检查工作表修改
        auto sheet_names = workbook_->getSheetNames();
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            std::string sheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
            change_tracker_->markPartDirty(sheet_path);
        }
        
        // 如果使用共享字符串
        if (workbook_->getOptions().use_shared_strings) {
            change_tracker_->markPartDirty("xl/sharedStrings.xml");
        }
        
        // 元数据文件
        change_tracker_->markPartDirty("[Content_Types].xml");
        change_tracker_->markPartDirty("_rels/.rels");
        change_tracker_->markPartDirty("xl/_rels/workbook.xml.rels");
    }
}

void PackageEditor::markPartDirty(const std::string& part) {
    if (change_tracker_) {
        change_tracker_->markPartDirty(part);
    }
}

bool PackageEditor::commit(const core::Path& target_path) {
    detectChanges();
    
    if (!change_tracker_->hasChanges()) {
        FASTEXCEL_LOG_INFO("No changes detected, fast copy to: {}", target_path.string());
        // 快速路径：直接复制
        // TODO: 实现快速复制逻辑
        return true;
    }
    
    FASTEXCEL_LOG_INFO("Committing {} dirty parts to: {}", 
             change_tracker_->getDirtyParts().size(), target_path.string());
    
    // 设置写入路径
    if (!package_manager_->openForWriting(target_path)) {
        FASTEXCEL_LOG_ERROR("Failed to open package for writing: {}", target_path.string());
        return false;
    }
    
    // 为所有dirty部件生成内容并写入
    auto dirty_parts = change_tracker_->getDirtyParts();
    for (const auto& part : dirty_parts) {
        std::string content = generatePart(part);
        if (content.empty() && isRequiredPart(part)) {
            FASTEXCEL_LOG_ERROR("Failed to generate required part: {}", part);
            return false;
        }
        
        if (!content.empty()) {
            if (!package_manager_->writePart(part, content)) {
                FASTEXCEL_LOG_ERROR("Failed to write part: {}", part);
                return false;
            }
        }
    }
    
    // 提交所有更改
    bool success = package_manager_->commit();
    
    if (success) {
        change_tracker_->clearAll();
        FASTEXCEL_LOG_INFO("Successfully committed changes to: {}", target_path.string());
    }
    
    return success;
}

std::string PackageEditor::generatePart(const std::string& path) const {
    if (!xml_generator_) {
        FASTEXCEL_LOG_ERROR("No XML generator available for part: {}", path);
        return "";
    }
    
    FASTEXCEL_LOG_DEBUG("Generating part: {}", path);

    // 以 IFileWriter 方式生成到字符串
    struct StringWriter : public core::IFileWriter {
        std::string target;
        std::string content;
        bool writeFile(const std::string& path, const std::string& data) override { (void)path; content = data; return true; }
        bool openStreamingFile(const std::string& path) override { (void)path; content.clear(); return true; }
        bool writeStreamingChunk(const char* data, size_t size) override { content.append(data, size); return true; }
        bool closeStreamingFile() override { return true; }
        std::string getTypeName() const override { return "StringWriter"; }
        WriteStats getStats() const override { return {}; }
    } sw;

    if (!xml_generator_->generateParts(sw, {path})) {
        FASTEXCEL_LOG_WARN("No generator for part: {}", path);
        return "";
    }
    return sw.content;
}

std::string PackageEditor::extractSheetNameFromPath(const std::string& path) const {
    // 从 "xl/worksheets/sheet1.xml" 提取 sheet ID，然后映射到名称
    const std::string prefix = "xl/worksheets/sheet";
    if (path.find(prefix) != 0) return "";
    
    size_t xml_pos = path.find(".xml");
    if (xml_pos == std::string::npos) return "";
    
    std::string id_str = path.substr(prefix.length(), xml_pos - prefix.length());
    if (id_str.empty()) return "";
    
    try {
        int sheet_id = std::stoi(id_str);
        if (workbook_) {
            auto sheet_names = workbook_->getSheetNames();
            if (sheet_id > 0 && sheet_id <= static_cast<int>(sheet_names.size())) {
                return sheet_names[sheet_id - 1];
            }
        }
    } catch (const std::exception& e) {
        FASTEXCEL_LOG_ERROR("Failed to parse sheet ID from path: {} - {}", path, e.what());
    }
    
    return "";
}



bool PackageEditor::isRequiredPart(const std::string& part) const {
    // 定义必需的部件
    static const std::unordered_set<std::string> required_parts = {
        "xl/workbook.xml",
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels"
    };
    
    return required_parts.count(part) > 0 || 
           part.find("xl/worksheets/sheet") == 0;
}

bool PackageEditor::isDirty() const {
    return change_tracker_ && change_tracker_->hasChanges();
}

std::vector<std::string> PackageEditor::getDirtyParts() const {
    if (change_tracker_) {
        return change_tracker_->getDirtyParts();
    }
    return {};
}

bool PackageEditor::save() {
    // 目前简化实现，后续可以优化为就地更新
    if (source_path_.empty()) {
        FASTEXCEL_LOG_ERROR("No source path specified for save operation");
        return false;
    }
    
    return commit(source_path_);
}

PackageEditor::ChangeStats PackageEditor::getChangeStats() const {
    ChangeStats stats;
    if (change_tracker_) {
        auto dirty_parts = change_tracker_->getDirtyParts();
        stats.modified_parts = dirty_parts.size();
        // TODO: 计算其他统计信息
    }
    return stats;
}

std::vector<std::string> PackageEditor::getSheetNames() const {
    if (workbook_) {
        return workbook_->getSheetNames();
    }
    return {};
}

std::vector<std::string> PackageEditor::getAllParts() const {
    if (package_manager_) {
        return package_manager_->listParts();
    }
    return {};
}

bool PackageEditor::validateXML(const std::string& xml_content) const {
    // 基本验证：检查是否为有效的XML结构
    if (xml_content.empty()) return false;
    
    // 简单检查：确保有开始和结束标签
    return xml_content.find("<?xml") != std::string::npos &&
           xml_content.find(">") != std::string::npos;
}

bool PackageEditor::initialize(const core::Path& xlsx_path) {
    source_path_ = xlsx_path;
    
    // 创建ZIP读取器
    auto zip_reader = std::make_unique<archive::ZipReader>(xlsx_path);
    if (!zip_reader->open()) {
        FASTEXCEL_LOG_ERROR("Failed to open ZIP file: {}", xlsx_path.string());
        return false;
    }
    
    return initializeServices(std::move(zip_reader), nullptr);
}

bool PackageEditor::initializeFromWorkbook() {
    if (!workbook_) {
        FASTEXCEL_LOG_ERROR("No workbook provided for initialization");
        return false;
    }
    
    return initializeServices(nullptr, workbook_);
}

void PackageEditor::logOperationStats() const {
    if (change_tracker_) {
        auto dirty_parts = change_tracker_->getDirtyParts();
        FASTEXCEL_LOG_INFO("Operation stats: {} dirty parts", dirty_parts.size());
    }
}

bool PackageEditor::isValidSheetName(const std::string& name) {
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查禁止的字符
    static const std::string forbidden_chars = "[]\\/*?:";
    for (char c : forbidden_chars) {
        if (name.find(c) != std::string::npos) {
            return false;
        }
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    return true;
}

bool PackageEditor::isValidCellRef(int row, int col) {
    return row >= 1 && row <= MAX_ROWS && col >= 1 && col <= MAX_COLS;
}

}} // namespace fastexcel::opc
