#include "fastexcel/core/Workbook.hpp"
#include "fastexcel/core/BatchFileWriter.hpp"
#include "fastexcel/core/CustomPropertyManager.hpp"
#include "fastexcel/core/DefinedNameManager.hpp"
#include "fastexcel/core/ExcelStructureGenerator.hpp"
#include "fastexcel/core/Exception.hpp"
#include "fastexcel/core/Path.hpp"
#include "fastexcel/core/StreamingFileWriter.hpp"
#include "fastexcel/core/StyleTransferContext.hpp"
#include "fastexcel/reader/XLSXReader.hpp"
#include "fastexcel/theme/Theme.hpp"
#include "fastexcel/theme/ThemeParser.hpp"
#include "fastexcel/utils/LogConfig.hpp"
#include "fastexcel/utils/Logger.hpp"
#include "fastexcel/utils/TimeUtils.hpp"
#include "fastexcel/xml/Relationships.hpp"
#include "fastexcel/xml/SharedStrings.hpp"
#include "fastexcel/xml/StyleSerializer.hpp"
#include "fastexcel/xml/UnifiedXMLGenerator.hpp"
#include "fastexcel/xml/XMLStreamWriter.hpp"

#include <algorithm>
#include <ctime>
#include <fstream>
#include <functional>
#include <iomanip>
#include <sstream>

namespace fastexcel {
namespace core {

// ========== DocumentProperties 实现 ==========

DocumentProperties::DocumentProperties() {
    // 使用 TimeUtils 获取当前时间
    created_time = utils::TimeUtils::getCurrentTime();
    modified_time = created_time;
}

// ========== Workbook 实现 ==========

std::unique_ptr<Workbook> Workbook::create(const Path& path) {
    return std::make_unique<Workbook>(path);
}

Workbook::Workbook(const Path& path) : filename_(path.string()) {
    // 检查是否为内存模式（任何以::memory::开头的路径）
    if (path.string().find("::memory::") == 0) {
        // 内存模式：不创建FileManager，保持纯内存操作
        file_manager_ = nullptr;
        LOG_DEBUG("Created workbook in memory mode: {}", filename_);
    } else {
        // 文件模式：创建FileManager处理文件操作
        file_manager_ = std::make_unique<archive::FileManager>(path);
    }
    
    format_repo_ = std::make_unique<FormatRepository>();
    // 初始化共享字符串表
    shared_string_table_ = std::make_unique<SharedStringTable>();
    
    // 初始化管理器
    custom_property_manager_ = std::make_unique<CustomPropertyManager>();
    defined_name_manager_ = std::make_unique<DefinedNameManager>();
    
    // 初始化智能脏数据管理器
    dirty_manager_ = std::make_unique<DirtyManager>();
    dirty_manager_->setIsNewFile(!path.exists()); // 如果文件不存在，则是新文件
    
    // 设置默认文档属性
    doc_properties_.author = "FastExcel";
    doc_properties_.company = "FastExcel Library";
}

Workbook::~Workbook() {
    close();
}

// ========== 文件操作 ==========

bool Workbook::open() {
    if (is_open_) {
        return true;
    }
    
    // 内存模式直接标记为已打开，无需文件操作
    if (!file_manager_) {
        is_open_ = true;
        LOG_DEBUG("Memory workbook opened: {}", filename_);
        return true;
    }
    
    // 文件模式需要打开FileManager
    is_open_ = file_manager_->open(true);
    if (is_open_) {
        LOG_INFO("Workbook opened: {}", filename_);
    }
    
    return is_open_;
}

bool Workbook::save() {
    if (!is_open_) {
        FASTEXCEL_THROW_OP("Workbook is not open");
    }
    
    // 🔧 修复：检查只读标志
    if (read_only_) {
        LOG_ERROR("Cannot save: workbook is opened in read-only mode");
        return false;
    }
    
    try {
        // 🔧 优化：使用 DirtyManager 获取最优保存策略
        if (!dirty_manager_) {
            LOG_ERROR("DirtyManager not initialized");
            return false;
        }
        
        auto strategy = dirty_manager_->getOptimalStrategy();
        LOG_INFO("Using save strategy: {} ({})",
                 static_cast<int>(strategy),
                 strategy == SaveStrategy::NONE ? "No changes" :
                 strategy == SaveStrategy::PURE_CREATE ? "Pure create" :
                 strategy == SaveStrategy::MINIMAL_UPDATE ? "Minimal update" :
                 strategy == SaveStrategy::SMART_EDIT ? "Smart edit" :
                 strategy == SaveStrategy::FULL_REBUILD ? "Full rebuild" : "Unknown");
        
        // 根据策略决定保存方式
        switch(strategy) {
            case SaveStrategy::NONE:
                LOG_INFO("No changes detected, skipping save");
                return true;
                
            case SaveStrategy::MINIMAL_UPDATE:
                // 仅更新修改的部分
                return saveIncremental();
                
            case SaveStrategy::SMART_EDIT:
                // 智能选择批量或流式
                {
                    size_t estimated_memory = estimateMemoryUsage();
                    bool use_streaming = estimated_memory > options_.auto_mode_memory_threshold;
                    LOG_INFO("Smart edit mode: using {} mode (estimated memory: {}MB, threshold: {}MB)",
                            use_streaming ? "streaming" : "batch",
                            estimated_memory / (1024*1024),
                            options_.auto_mode_memory_threshold / (1024*1024));
                    return saveWithFullGeneration(use_streaming);
                }
                
            case SaveStrategy::PURE_CREATE:
            case SaveStrategy::FULL_REBUILD:
            default:
                // 完全重建
                return saveWithFullGeneration(options_.mode == WorkbookMode::STREAMING);
        }
        
    } catch (const std::exception& e) {
        LOG_ERROR("Failed to save workbook: {}", e.what());
        return false;
    }
}

// 新增：增量保存方法
bool Workbook::saveIncremental() {
    LOG_INFO("Performing incremental save");
    
    // 使用 TimeUtils 更新修改时间
    doc_properties_.modified_time = utils::TimeUtils::getCurrentTime();
    
    auto changes = dirty_manager_->getChanges();
    if (changes.isEmpty()) {
        LOG_INFO("No changes to save");
        return true;
    }
    
    // 设置ZIP压缩级别
    if (file_manager_ && file_manager_->isOpen()) {
        file_manager_->setCompressionLevel(options_.compression_level);
    }
    
    // 仅重新生成修改的部分
    for (const auto& change : changes.getChanges()) {
        LOG_DEBUG("Processing change: part={}, level={}",
                 change.part, static_cast<int>(change.level));
        
        if (change.part.find("worksheet") != std::string::npos) {
            // 提取工作表索引
            size_t index = 0;
            if (sscanf(change.part.c_str(), "xl/worksheets/sheet%zu.xml", &index) == 1) {
                if (index > 0 && index <= worksheets_.size()) {
                    LOG_DEBUG("Regenerating worksheet {}", index);
                    // 这里应该只重新生成特定的工作表
                    // 暂时使用完整生成，后续优化
                    return saveWithFullGeneration(false);
                }
            }
        } else if (change.part == "xl/styles.xml") {
            // 仅重新生成样式文件
            LOG_DEBUG("Regenerating styles");
            return saveWithFullGeneration(false);
        }
    }
    
    // 清除脏标记
    dirty_manager_->clear();
    LOG_INFO("Incremental save completed");
    return true;
}

// 新增：完整生成保存方法
bool Workbook::saveWithFullGeneration(bool use_streaming) {
    LOG_INFO("Performing full generation save (streaming: {})", use_streaming);
    
    // 使用 TimeUtils 更新修改时间
    doc_properties_.modified_time = utils::TimeUtils::getCurrentTime();
    
    // 设置ZIP压缩级别
    if (file_manager_ && file_manager_->isOpen()) {
        if (!file_manager_->setCompressionLevel(options_.compression_level)) {
            LOG_WARN("Failed to set compression level to {}", options_.compression_level);
        } else {
            LOG_ZIP_DEBUG("Set ZIP compression level to {}", options_.compression_level);
        }
    }
    
    // 🔧 修复SharedStrings生成逻辑：移除手动收集，依赖工作表XML生成时自动添加
    // 清空共享字符串列表，让工作表XML生成时自动填充
    if (options_.use_shared_strings) {
        LOG_DEBUG("SharedStrings enabled - SST will be populated during worksheet XML generation");
        if (shared_string_table_) shared_string_table_->clear();
    } else {
        LOG_DEBUG("SharedStrings disabled for performance");
        if (shared_string_table_) shared_string_table_->clear();
    }
    
    // 编辑模式下，先将原包中未被我们生成的条目拷贝过来（绘图、图片、打印设置等）
    if (opened_from_existing_ && preserve_unknown_parts_ && !original_package_path_.empty() && file_manager_ && file_manager_->isOpen()) {
        // 我们将跳过这些前缀（由生成逻辑负责写入/覆盖）
        // 透传阶段：不跳过任何前缀，先复制全部条目；后续生成阶段会覆盖我们需要更新的部件
        std::vector<std::string> skip_prefixes = { };
        file_manager_->copyFromExistingPackage(core::Path(original_package_path_), skip_prefixes);
    }

    // 生成Excel文件结构（会覆盖我们管理的核心部件）
    if (!generateWithGenerator(use_streaming)) {
        LOG_ERROR("Failed to generate Excel structure");
        return false;
    }
    
    // 清除脏标记
    if (dirty_manager_) {
        dirty_manager_->clear();
    }
    
    LOG_INFO("Workbook saved successfully: {}", filename_);
    return true;

bool Workbook::saveAs(const std::string& filename) {
    // 🔧 修复：saveAs 允许只读文件另存为新文件
    // 但需要检查目标文件的写权限
    Path target_path(filename);
    if (target_path.exists() && !target_path.isWritable()) {
        LOG_ERROR("Cannot save as: target file is not writable: {}", filename);
        return false;
    }
    
    std::string old_filename = filename_;
    std::string original_source = original_package_path_;
    bool was_from_existing = opened_from_existing_;
    bool was_read_only = read_only_;

    // 检查是否保存到同一个文件
    bool is_same_file = (filename == old_filename) || (filename == original_source);
    
    // 如果保存到不同文件，解除只读限制
    if (!is_same_file && read_only_) {
        read_only_ = false;
        LOG_INFO("Read-only restriction removed for save as operation");
    }
    
    if (is_same_file && was_from_existing && !original_source.empty()) {
        // 如果保存到同一个文件，需要先复制原文件到临时位置
        LOG_INFO("Saving to same file, creating temporary backup for resource preservation");
        
        // 创建临时文件路径
        std::string temp_backup = original_source + ".tmp_backup";
        core::Path source_path(original_source);
        core::Path temp_path(temp_backup);
        
        // 复制原文件到临时位置
        try {
            if (temp_path.exists()) {
                temp_path.remove();
            }
            source_path.copyTo(temp_path);
            original_package_path_ = temp_backup;  // 更新源路径为临时文件
            LOG_DEBUG("Created temporary backup: {}", temp_backup);
        } catch (const std::exception& e) {
            LOG_ERROR("Failed to create temporary backup: {}", e.what());
            return false;
        }
    }

    filename_ = filename;
    
    // 重新创建文件管理器
    file_manager_ = std::make_unique<archive::FileManager>(core::Path(filename));
    
    if (!file_manager_->open(true)) {
        // 恢复原文件名
        filename_ = old_filename;
        file_manager_ = std::make_unique<archive::FileManager>(core::Path(old_filename));
        
        // 如果创建了临时文件，删除它
        if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
            core::Path temp_path(original_package_path_);
            if (temp_path.exists()) {
                temp_path.remove();
            }
            original_package_path_ = original_source;  // 恢复原路径
        }
        return false;
    }

    // 在另存为场景下，如果当前工作簿是从现有包打开的，那么保留 original_package_path_ 用于拷贝未修改部件
    opened_from_existing_ = was_from_existing;
    // original_package_path_ 已经在上面设置好了（可能是临时文件或原始文件）
    
    bool save_result = save();
    
    // 如果保存失败，恢复只读状态
    if (!save_result && was_read_only) {
        read_only_ = was_read_only;
    }
    
    // 清理临时文件（如果有）
    if (is_same_file && original_package_path_.find(".tmp_backup") != std::string::npos) {
        core::Path temp_path(original_package_path_);
        if (temp_path.exists()) {
            temp_path.remove();
            LOG_DEBUG("Removed temporary backup: {}", original_package_path_);
        }
        original_package_path_ = original_source;  // 恢复原路径
    }
    
    return save_result;
}

bool Workbook::close() {
    if (is_open_) {
        // 内存模式只需要重置状态
        if (!file_manager_) {
            is_open_ = false;
            LOG_DEBUG("Memory workbook closed: {}", filename_);
        } else {
            // 文件模式需要关闭FileManager
            file_manager_->close();
            is_open_ = false;
            LOG_INFO("Workbook closed: {}", filename_);
        }
    }
    return true;
}

// ========== 工作表管理 ==========

std::shared_ptr<Worksheet> Workbook::addWorksheet(const std::string& name) {
    if (!is_open_) {
        FASTEXCEL_THROW_OP("Workbook is not open");
    }
    
    std::string sheet_name;
    if (name.empty()) {
        sheet_name = generateUniqueSheetName("Sheet1");
    } else {
        // 检查名称是否已存在，如果存在则生成唯一名称
        if (getWorksheet(name) != nullptr) {
            sheet_name = generateUniqueSheetName(name);
        } else {
            sheet_name = name;
        }
    }
    
    if (!validateSheetName(sheet_name)) {
        LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.push_back(worksheet);
    
    // 关键修复：如果这是第一个工作表，自动设置为激活状态
    if (worksheets_.size() == 1) {
        worksheet->setTabSelected(true);
        LOG_DEBUG("Added worksheet: {} (activated as first sheet)", sheet_name);
    } else {
        LOG_DEBUG("Added worksheet: {}", sheet_name);
    }
    
    return worksheet;
}

std::shared_ptr<Worksheet> Workbook::insertWorksheet(size_t index, const std::string& name) {
    if (!is_open_) {
        LOG_ERROR("Workbook is not open");
        return nullptr;
    }
    
    if (index > worksheets_.size()) {
        index = worksheets_.size();
    }
    
    std::string sheet_name = name.empty() ? generateUniqueSheetName("Sheet1") : name;
    
    if (!validateSheetName(sheet_name)) {
        LOG_ERROR("Invalid sheet name: {}", sheet_name);
        return nullptr;
    }
    
    auto worksheet = std::make_shared<Worksheet>(sheet_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    worksheets_.insert(worksheets_.begin() + index, worksheet);
    
    LOG_DEBUG("Inserted worksheet: {} at index {}", sheet_name, index);
    return worksheet;
}

bool Workbook::removeWorksheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        worksheets_.erase(it);
        LOG_DEBUG("Removed worksheet: {}", name);
        return true;
    }
    
    return false;
}

bool Workbook::removeWorksheet(size_t index) {
    if (index < worksheets_.size()) {
        std::string name = worksheets_[index]->getName();
        worksheets_.erase(worksheets_.begin() + index);
        LOG_DEBUG("Removed worksheet: {} at index {}", name, index);
        return true;
    }
    
    return false;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(const std::string& name) {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(), 
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<Worksheet> Workbook::getWorksheet(size_t index) {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(const std::string& name) const {
    auto it = std::find_if(worksheets_.begin(), worksheets_.end(),
                          [&name](const std::shared_ptr<Worksheet>& ws) {
                              return ws->getName() == name;
                          });
    
    if (it != worksheets_.end()) {
        return *it;
    }
    
    return nullptr;
}

std::shared_ptr<const Worksheet> Workbook::getWorksheet(size_t index) const {
    if (index < worksheets_.size()) {
        return worksheets_[index];
    }
    return nullptr;
}

std::vector<std::string> Workbook::getWorksheetNames() const {
    std::vector<std::string> names;
    names.reserve(worksheets_.size());
    
    for (const auto& worksheet : worksheets_) {
        names.push_back(worksheet->getName());
    }
    
    return names;
}

bool Workbook::renameWorksheet(const std::string& old_name, const std::string& new_name) {
    auto worksheet = getWorksheet(old_name);
    if (!worksheet) {
        return false;
    }
    
    if (!validateSheetName(new_name)) {
        return false;
    }
    
    worksheet->setName(new_name);
    LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
    return true;
}

bool Workbook::moveWorksheet(size_t from_index, size_t to_index) {
    if (from_index >= worksheets_.size() || to_index >= worksheets_.size()) {
        return false;
    }
    
    if (from_index == to_index) {
        return true;
    }
    
    auto worksheet = worksheets_[from_index];
    worksheets_.erase(worksheets_.begin() + from_index);
    
    if (to_index > from_index) {
        to_index--;
    }
    
    worksheets_.insert(worksheets_.begin() + to_index, worksheet);
    
    LOG_DEBUG("Moved worksheet from index {} to {}", from_index, to_index);
    return true;
}

std::shared_ptr<Worksheet> Workbook::copyWorksheet(const std::string& source_name, const std::string& new_name) {
    auto source_worksheet = getWorksheet(source_name);
    if (!source_worksheet) {
        return nullptr;
    }
    
    if (!validateSheetName(new_name)) {
        return nullptr;
    }
    
    // 创建新工作表
    auto new_worksheet = std::make_shared<Worksheet>(new_name, std::shared_ptr<Workbook>(this, [](Workbook*){}), next_sheet_id_++);
    
    // 这里应该实现深拷贝逻辑
    // 简化版本，实际需要复制所有单元格、格式、设置等
    
    worksheets_.push_back(new_worksheet);
    
    LOG_DEBUG("Copied worksheet: {} -> {}", source_name, new_name);
    return new_worksheet;
}

void Workbook::setActiveWorksheet(size_t index) {
    // 取消所有工作表的选中状态
    for (auto& worksheet : worksheets_) {
        worksheet->setTabSelected(false);
    }
    
    // 设置指定工作表为活动状态
    if (index < worksheets_.size()) {
        worksheets_[index]->setTabSelected(true);
    }
}

// ========== 样式管理 ==========

int Workbook::addStyle(const FormatDescriptor& style) {
    return format_repo_->addFormat(style);
}

int Workbook::addStyle(const StyleBuilder& builder) {
    auto format = builder.build();
    return format_repo_->addFormat(format);
}

std::shared_ptr<const FormatDescriptor> Workbook::getStyle(int style_id) const {
    if (!is_open_) {
        return nullptr;
    }
    
    // 从格式仓储中根据ID获取格式描述符
    return format_repo_->getFormat(style_id);
}

int Workbook::getDefaultStyleId() const {
    return format_repo_->getDefaultFormatId();
}

bool Workbook::isValidStyleId(int style_id) const {
    return format_repo_->isValidFormatId(style_id);
}

const FormatRepository& Workbook::getStyleRepository() const {
    return *format_repo_;
}

void Workbook::setThemeXML(const std::string& theme_xml) {
    theme_xml_ = theme_xml;
    theme_dirty_ = true; // 外部显式设置主题XML视为编辑
    LOG_DEBUG("设置自定义主题XML ({} 字节)", theme_xml_.size());
    // 尝试解析为结构化主题对象
    if (!theme_xml_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_);
        if (parsed) {
            theme_ = std::move(parsed);
            LOG_DEBUG("主题XML已解析为对象: {}", theme_->getName());
        } else {
            LOG_WARN("主题XML解析失败，保留原始XML");
        }
    }
}

const std::string& Workbook::getThemeXML() const {
    return theme_xml_;
}

void Workbook::setOriginalThemeXML(const std::string& theme_xml) {
    theme_xml_original_ = theme_xml;
    LOG_DEBUG("保存原始主题XML ({} 字节)", theme_xml_original_.size());
    // 同步解析一次，便于后续编辑
    if (!theme_xml_original_.empty()) {
        auto parsed = theme::ThemeParser::parseFromXML(theme_xml_original_);
        if (parsed) {
            theme_ = std::move(parsed);
            LOG_DEBUG("原始主题XML已解析为对象: {}", theme_->getName());
        }
    }
}

void Workbook::setTheme(const theme::Theme& theme) {
    theme_ = std::make_unique<theme::Theme>(theme);
    // 同步XML缓存
    theme_xml_ = theme_->toXML();
    theme_dirty_ = true;
}

void Workbook::setThemeName(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->setName(name);
    theme_xml_.clear(); // 让生成时重新序列化
    theme_dirty_ = true;
}

void Workbook::setThemeColor(theme::ThemeColorScheme::ColorType type, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->colors().setColor(type, color);
    theme_xml_.clear();
    theme_dirty_ = true;
}

bool Workbook::setThemeColorByName(const std::string& name, const core::Color& color) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    bool ok = theme_->colors().setColorByName(name, color);
    if (ok) { theme_xml_.clear(); theme_dirty_ = true; }
    return ok;
}

void Workbook::setThemeMajorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMajorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMajorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontLatin(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontLatin(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontEastAsia(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontEastAsia(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

void Workbook::setThemeMinorFontComplex(const std::string& name) {
    if (!theme_) theme_ = std::make_unique<theme::Theme>();
    theme_->fonts().setMinorFontComplex(name);
    theme_xml_.clear();
    theme_dirty_ = true;
}

// ========== 自定义属性 ==========

void Workbook::setCustomProperty(const std::string& name, const std::string& value) {
    custom_property_manager_->setProperty(name, value);
}

void Workbook::setCustomProperty(const std::string& name, double value) {
    custom_property_manager_->setProperty(name, value);
}

void Workbook::setCustomProperty(const std::string& name, bool value) {
    custom_property_manager_->setProperty(name, value);
}

std::string Workbook::getCustomProperty(const std::string& name) const {
    return custom_property_manager_->getProperty(name);
}

bool Workbook::removeCustomProperty(const std::string& name) {
    return custom_property_manager_->removeProperty(name);
}

std::unordered_map<std::string, std::string> Workbook::getCustomProperties() const {
    return custom_property_manager_->getAllProperties();
}

// ========== 定义名称 ==========

void Workbook::defineName(const std::string& name, const std::string& formula, const std::string& scope) {
    defined_name_manager_->define(name, formula, scope);
}

std::string Workbook::getDefinedName(const std::string& name, const std::string& scope) const {
    return defined_name_manager_->get(name, scope);
}

bool Workbook::removeDefinedName(const std::string& name, const std::string& scope) {
    return defined_name_manager_->remove(name, scope);
}

// ========== VBA项目 ==========

bool Workbook::addVbaProject(const std::string& vba_project_path) {
    // 检查文件是否存在
    std::ifstream file(vba_project_path, std::ios::binary);
    if (!file.is_open()) {
        LOG_ERROR("VBA project file not found: {}", vba_project_path);
        return false;
    }
    
    vba_project_path_ = vba_project_path;
    has_vba_ = true;
    
    LOG_INFO("Added VBA project: {}", vba_project_path);
    return true;
}

// ========== 工作簿保护 ==========

void Workbook::protect(const std::string& password, bool lock_structure, bool lock_windows) {
    protected_ = true;
    protection_password_ = password;
    lock_structure_ = lock_structure;
    lock_windows_ = lock_windows;
}

void Workbook::unprotect() {
    protected_ = false;
    protection_password_.clear();
    lock_structure_ = false;
    lock_windows_ = false;
}

// ========== 工作簿选项 ==========

void Workbook::setCalcOptions(bool calc_on_load, bool full_calc_on_load) {
    options_.calc_on_load = calc_on_load;
    options_.full_calc_on_load = full_calc_on_load;
}

// ========== 生成控制判定（使用DirtyManager智能管理） ==========

bool Workbook::shouldGenerateContentTypes() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("[Content_Types].xml");
}

bool Workbook::shouldGenerateRootRels() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("_rels/.rels");
}

bool Workbook::shouldGenerateWorkbookCore() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("xl/workbook.xml");
}

bool Workbook::shouldGenerateStyles() const {
    // 始终生成样式文件，保证包内引用一致性：
    // - workbook.xml 和 [Content_Types].xml 总是包含对 xl/styles.xml 的引用
    // - 如不生成，将导致包缺少被引用的部件，Excel 打开会提示修复
    // 样式文件很小，生成最小可用样式的成本可以忽略
    return true;
}

bool Workbook::shouldGenerateTheme() const {
    if (!dirty_manager_) return true;
    // 主题特殊处理：如果有主题内容则需要生成
    if (!theme_xml_.empty() || !theme_xml_original_.empty() || theme_) {
        return true;
    }
    return dirty_manager_->shouldUpdate("xl/theme/theme1.xml");
}

bool Workbook::shouldGenerateSharedStrings() const {
    LOG_DEBUG("shouldGenerateSharedStrings() called - analyzing conditions");
    
    if (!options_.use_shared_strings) {
        LOG_DEBUG("SharedStrings generation disabled by options_.use_shared_strings = false");
        return false; // 未启用SST
    }
    LOG_DEBUG("options_.use_shared_strings = true, SharedStrings enabled");
    
    if (!dirty_manager_) {
        LOG_DEBUG("No dirty manager, SharedStrings generation enabled (default true)");
        return true;
    }
    LOG_DEBUG("DirtyManager exists, checking shouldUpdate for xl/sharedStrings.xml");
    
    bool should_update = dirty_manager_->shouldUpdate("xl/sharedStrings.xml");
    LOG_DEBUG("DirtyManager shouldUpdate for SharedStrings: {}", should_update);
    
    // 🔧 关键修复：如果SharedStringTable有内容但DirtyManager说不需要更新，强制生成
    if (shared_string_table_) {
        size_t string_count = shared_string_table_->getStringCount();
        LOG_DEBUG("SharedStringTable contains {} strings", string_count);
        
        if (string_count > 0 && !should_update) {
            LOG_DEBUG("🔧 FORCE GENERATION: SharedStringTable has {} strings but DirtyManager says no update needed", string_count);
            LOG_DEBUG("🔧 This happens when target file exists but we're creating new content with strings");
            LOG_DEBUG("🔧 Forcing SharedStrings generation to avoid missing sharedStrings.xml");
            return true; // 强制生成
        }
    } else {
        LOG_DEBUG("SharedStringTable is null");
    }
    
    return should_update;
}

bool Workbook::shouldGenerateDocPropsCore() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/core.xml");
}

bool Workbook::shouldGenerateDocPropsApp() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/app.xml");
}

bool Workbook::shouldGenerateDocPropsCustom() const {
    if (!dirty_manager_) return true;
    return dirty_manager_->shouldUpdate("docProps/custom.xml");
}

bool Workbook::shouldGenerateSheet(size_t index) const {
    if (!dirty_manager_) return true;
    std::string sheetPart = "xl/worksheets/sheet" + std::to_string(index + 1) + ".xml";
    return dirty_manager_->shouldUpdate(sheetPart);
}

bool Workbook::shouldGenerateSheetRels(size_t index) const {
    if (!dirty_manager_) return true;
    std::string sheetRelsPart = "xl/worksheets/_rels/sheet" + std::to_string(index + 1) + ".xml.rels";
    return dirty_manager_->shouldUpdate(sheetRelsPart);
}

// ========== 共享字符串管理 ==========

int Workbook::addSharedString(const std::string& str) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addString(str));
}

int Workbook::addSharedStringWithIndex(const std::string& str, int original_index) {
    if (!shared_string_table_) shared_string_table_ = std::make_unique<SharedStringTable>();
    return static_cast<int>(shared_string_table_->addStringWithId(str, original_index));
}

int Workbook::getSharedStringIndex(const std::string& str) const {
    if (!shared_string_table_) return -1;
    return static_cast<int>(shared_string_table_->getStringId(str));
}

const SharedStringTable* Workbook::getSharedStringTable() const {
    return shared_string_table_.get();
}

// ========== 内部方法 ==========

bool Workbook::generateExcelStructure() {
    // 智能选择生成模式：根据数据量和内存使用情况自动决定
    size_t estimated_memory = estimateMemoryUsage();
    size_t total_cells = getTotalCellCount();
    
    bool use_streaming = false;
    
    // 新的决策逻辑：基于WorkbookMode
    switch (options_.mode) {
        case WorkbookMode::AUTO:
            // 自动模式：根据数据量智能选择
            if (total_cells > options_.auto_mode_cell_threshold ||
                estimated_memory > options_.auto_mode_memory_threshold) {
                use_streaming = true;
                LOG_INFO("Auto-selected streaming mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            } else {
                use_streaming = false;
                LOG_INFO("Auto-selected batch mode: {} cells, {}MB estimated memory (thresholds: {} cells, {}MB)",
                        total_cells, estimated_memory / (1024*1024),
                        options_.auto_mode_cell_threshold, options_.auto_mode_memory_threshold / (1024*1024));
            }
            break;
            
        case WorkbookMode::BATCH:
            // 强制批量模式
            use_streaming = false;
            LOG_INFO("Using forced batch mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
            
        case WorkbookMode::STREAMING:
            // 强制流式模式
            use_streaming = true;
            LOG_INFO("Using forced streaming mode: {} cells, {}MB estimated memory",
                    total_cells, estimated_memory / (1024*1024));
            break;
    }
    
    // 如果设置了constant_memory，强制使用流式模式
    if (options_.constant_memory) {
        use_streaming = true;
        LOG_INFO("Constant memory mode enabled, forcing streaming mode");
    }
    
    return generateWithGenerator(use_streaming);
}



void Workbook::generateWorkbookXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator，消除重复代码
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateWorkbookXML(callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for workbook XML generation");
    }
}

void Workbook::generateStylesXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateStylesXML(callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for styles XML generation");
    }
}

void Workbook::generateSharedStringsXML(const std::function<void(const char*, size_t)>& callback) const {
    LOG_DEBUG("generateSharedStringsXML called");
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        LOG_DEBUG("UnifiedXMLGenerator created successfully for SharedStrings");
        generator->generateSharedStringsXML(callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for shared strings XML generation");
    }
}

void Workbook::generateWorksheetXML(const std::shared_ptr<Worksheet>& worksheet, const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateWorksheetXML(worksheet.get(), callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for worksheet XML generation");
        // 备选方案：使用原来的方法
        worksheet->generateXML(callback);
    }
}

void Workbook::generateDocPropsAppXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateDocPropsXML("app", callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for app properties XML generation");
    }
}

void Workbook::generateDocPropsCoreXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateDocPropsXML("core", callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for core properties XML generation");
    }
}

void Workbook::generateDocPropsCustomXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateDocPropsXML("custom", callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for custom properties XML generation");
    }
}

void Workbook::generateContentTypesXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateContentTypesXML(callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for content types XML generation");
    }
}

void Workbook::generateRelsXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateRelationshipsXML("root", callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for root relationships XML generation");
    }
}

void Workbook::generateWorkbookRelsXML(const std::function<void(const char*, size_t)>& callback) const {
    // 重构：委托给UnifiedXMLGenerator
    auto generator = xml::UnifiedXMLGenerator::fromWorkbook(this);
    if (generator) {
        generator->generateRelationshipsXML("workbook", callback);
    } else {
        LOG_ERROR("Failed to create UnifiedXMLGenerator for workbook relationships XML generation");
    }
}

void Workbook::generateThemeXML(const std::function<void(const char*, size_t)>& callback) const {
    // 优先级：
    // 1) 如果未脏且有原始XML，保真写回原始
    if (!theme_dirty_ && !theme_xml_original_.empty()) {
        LOG_DEBUG("使用原始主题XML保真写回 ({} 字节)", theme_xml_original_.size());
        callback(theme_xml_original_.c_str(), theme_xml_original_.size());
        return;
    }
    // 2) 有显式设置的自定义XML，直接写出
    if (!theme_xml_.empty()) {
        LOG_DEBUG("使用自定义主题XML ({} 字节)", theme_xml_.size());
        callback(theme_xml_.c_str(), theme_xml_.size());
        return;
    }
    // 3) 有结构化主题对象，则序列化
    if (theme_) {
        const std::string xml = theme_->toXML();
        callback(xml.c_str(), xml.size());
        return;
    }
    // 4) 回退：生成最小可用默认主题
    theme::Theme default_theme("Office");
    const std::string xml = default_theme.toXML();
    callback(xml.c_str(), xml.size());
}

// ========== 格式管理内部方法 ==========


// ========== 辅助函数 ==========

std::string Workbook::generateUniqueSheetName(const std::string& base_name) const {
    // 如果base_name不存在，直接返回
    if (getWorksheet(base_name) == nullptr) {
        return base_name;
    }
    
    // 如果base_name是"Sheet1"，从"Sheet2"开始尝试
    if (base_name == "Sheet1") {
        int counter = 2;
        std::string name = "Sheet" + std::to_string(counter);
        while (getWorksheet(name) != nullptr) {
            name = "Sheet" + std::to_string(++counter);
        }
        return name;
    }
    
    // 对于其他base_name，添加数字后缀
    int suffix_counter = 1;
    std::string name = base_name + std::to_string(suffix_counter);
    while (getWorksheet(name) != nullptr) {
        name = base_name + std::to_string(++suffix_counter);
    }
    
    return name;
}

bool Workbook::validateSheetName(const std::string& name) const {
    // 检查长度
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // 检查非法字符
    const std::string invalid_chars = ":\\/?*[]";
    if (name.find_first_of(invalid_chars) != std::string::npos) {
        return false;
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 不检查是否已存在，因为这个方法也被用于验证新名称
    // 重复名称的检查应该在调用方处理
    
    return true;
}

void Workbook::collectSharedStrings() {
    if (!shared_string_table_) {
        shared_string_table_ = std::make_unique<SharedStringTable>();
    } else {
        shared_string_table_->clear();
    }
    
    for (const auto& worksheet : worksheets_) {
        // 这里需要访问工作表的单元格来收集字符串
        // 简化版本，实际实现需要遍历所有字符串单元格
        auto [max_row, max_col] = worksheet->getUsedRange();
        
        for (int row = 0; row <= max_row; ++row) {
            for (int col = 0; col <= max_col; ++col) {
                if (worksheet->hasCellAt(row, col)) {
                    const auto& cell = worksheet->getCell(row, col);
                    if (cell.isString()) {
                        addSharedString(cell.getStringValue());
                    }
                }
            }
        }
    }
}


std::string Workbook::getWorksheetPath(int sheet_id) const {
    return "xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::getWorksheetRelPath(int sheet_id) const {
    return "worksheets/sheet" + std::to_string(sheet_id) + ".xml";
}

std::string Workbook::formatTime(const std::tm& time) const {
    // 使用 TimeUtils 进行时间格式化
    return utils::TimeUtils::formatTimeISO8601(time);
}

std::string Workbook::hashPassword(const std::string& password) const {
    // 简化的密码哈希实现
    // 实际应该使用Excel的密码哈希算法
    std::hash<std::string> hasher;
    size_t hash = hasher(password);
    
    std::ostringstream oss;
    oss << std::hex << std::uppercase << hash;
    return oss.str();
}

void Workbook::setHighPerformanceMode(bool enable) {
    if (enable) {
        LOG_INFO("Enabling ultra high performance mode (beyond defaults)");
        
        // 进一步优化：使用无压缩模式排除压缩算法影响
        options_.compression_level = 0;  // 无压缩
        
        // 更大的缓冲区
        options_.row_buffer_size = 10000;
        options_.xml_buffer_size = 8 * 1024 * 1024;  // 8MB
        
        // 使用AUTO模式，让系统根据数据量自动选择
        options_.mode = WorkbookMode::AUTO;
        options_.use_shared_strings = true;
        
        // 调整自动模式阈值，更倾向于使用批量模式以获得更好的性能
        options_.auto_mode_cell_threshold = 2000000;  // 200万单元格
        options_.auto_mode_memory_threshold = 200 * 1024 * 1024;  // 200MB
        
        LOG_INFO("Ultra high performance mode configured: Mode=AUTO, Compression=OFF, RowBuffer={}, XMLBuffer={}MB",
                options_.row_buffer_size, options_.xml_buffer_size / (1024*1024));
    } else {
        LOG_INFO("Using standard high performance mode (default settings)");
        
        // 恢复到默认的高性能设置
        options_.mode = WorkbookMode::AUTO;           // 默认自动模式
        options_.use_shared_strings = true;           // 默认启用以匹配Excel格式
        options_.row_buffer_size = 5000;              // 默认较大缓冲
        options_.compression_level = 6;               // 恢复到默认的中等压缩
        options_.xml_buffer_size = 4 * 1024 * 1024;  // 默认4MB
        
        // 恢复默认阈值
        options_.auto_mode_cell_threshold = 1000000;     // 100万单元格
        options_.auto_mode_memory_threshold = 100 * 1024 * 1024; // 100MB
    }
}



// 辅助方法：XML转义
std::string Workbook::escapeXML(const std::string& text) const {
    std::string result;
    result.reserve(static_cast<size_t>(text.size() * 1.2));
    
    for (char c : text) {
        switch (c) {
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '&': result += "&amp;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default:
                if (c < 0x20 && c != 0x09 && c != 0x0A && c != 0x0D) {
                    continue; // 跳过无效控制字符
                }
                result.push_back(c);
                break;
        }
    }
    
    return result;
}

// ========== 工作簿编辑功能实现 ==========

std::unique_ptr<Workbook> Workbook::open(const Path& path) {
    try {
        // 使用Path的内置文件检查
        if (!path.exists()) {
            LOG_ERROR("File not found for editing: {}", path.string());
            return nullptr;
        }
        
        // 使用XLSXReader读取现有文件
        reader::XLSXReader reader(path);
        auto result = reader.open();
        if (result != core::ErrorCode::Ok) {
            LOG_ERROR("Failed to open XLSX file for reading: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 加载工作簿
        std::unique_ptr<core::Workbook> loaded_workbook;
        result = reader.loadWorkbook(loaded_workbook);
        reader.close();
        
        if (result != core::ErrorCode::Ok) {
            LOG_ERROR("Failed to load workbook from file: {}, error code: {}", path.string(), static_cast<int>(result));
            return nullptr;
        }
        
        // 标记来源以便保存时进行未修改部件的保真写回
        if (loaded_workbook) {
            loaded_workbook->opened_from_existing_ = true;
            loaded_workbook->original_package_path_ = path.string();
            
            // 🔧 修复：检查文件权限，设置只读标志
            if (!path.isWritable()) {
                loaded_workbook->read_only_ = true;
                LOG_INFO("File opened in read-only mode (no write permission): {}", path.string());
            } else {
                loaded_workbook->read_only_ = false;
                LOG_INFO("File opened in read-write mode: {}", path.string());
            }
            
            // 在只读模式下优化设置
            if (loaded_workbook->read_only_) {
                // 只读模式不需要共享字符串表（不会修改）
                loaded_workbook->options_.use_shared_strings = false;
                // 强制使用流式模式以减少内存占用
                loaded_workbook->options_.mode = WorkbookMode::STREAMING;
                LOG_DEBUG("Read-only optimizations applied: SST disabled, streaming mode forced");
            }
        }
        
        LOG_INFO("Successfully loaded workbook for editing: {}", path.string());
        return loaded_workbook;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception while loading workbook for editing: {}, error: {}", path.string(), e.what());
        return nullptr;
    }
}

bool Workbook::refresh() {
    if (!is_open_) {
        LOG_ERROR("Cannot refresh: workbook is not open");
        return false;
    }
    
    try {
        // 保存当前状态
        std::string current_filename = filename_;
        bool was_open = is_open_;
        
        // 关闭当前工作簿
        close();
        
        // 重新加载
        Path current_path(current_filename);
        auto refreshed_workbook = open(current_path);
        if (!refreshed_workbook) {
            LOG_ERROR("Failed to refresh workbook: {}", current_filename);
            return false;
        }
        
        // 替换当前内容
        worksheets_ = std::move(refreshed_workbook->worksheets_);
        format_repo_ = std::move(refreshed_workbook->format_repo_);
        doc_properties_ = refreshed_workbook->doc_properties_;
        custom_property_manager_ = std::move(refreshed_workbook->custom_property_manager_);
        defined_name_manager_ = std::move(refreshed_workbook->defined_name_manager_);
        
        // 恢复打开状态
        if (was_open) {
            open();
        }
        
        LOG_INFO("Successfully refreshed workbook: {}", current_filename);
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during workbook refresh: {}", e.what());
        return false;
    }
}

bool Workbook::mergeWorkbook(const std::unique_ptr<Workbook>& other_workbook, const MergeOptions& options) {
    if (!other_workbook) {
        LOG_ERROR("Cannot merge: other workbook is null");
        return false;
    }
    
    if (!is_open_) {
        LOG_ERROR("Cannot merge: current workbook is not open");
        return false;
    }
    
    try {
        int merged_count = 0;
        
        // 合并工作表
        if (options.merge_worksheets) {
            for (const auto& other_worksheet : other_workbook->worksheets_) {
                std::string new_name = options.name_prefix + other_worksheet->getName();
                
                // 检查名称冲突
                if (getWorksheet(new_name) != nullptr) {
                    if (options.overwrite_existing) {
                        removeWorksheet(new_name);
                        LOG_INFO("Removed existing worksheet for merge: {}", new_name);
                    } else {
                        new_name = generateUniqueSheetName(new_name);
                        LOG_INFO("Generated unique name for merge: {}", new_name);
                    }
                }
                
                // 创建新工作表并复制内容
                auto new_worksheet = addWorksheet(new_name);
                if (new_worksheet) {
                    // 这里需要实现深拷贝逻辑
                    // 简化版本：复制基本属性
                    merged_count++;
                    LOG_DEBUG("Merged worksheet: {} -> {}", other_worksheet->getName(), new_name);
                }
            }
        }
        
        // 合并格式
        if (options.merge_formats) {
            // 将其他工作簿的格式仓储合并到当前格式仓储
            // 遍历其他工作簿的所有格式并添加到当前仓储中（自动去重）
            for (const auto& format_item : *other_workbook->format_repo_) {
                format_repo_->addFormat(*format_item.format);
            }
            LOG_DEBUG("Merged formats from other workbook");
        }
        
        // 合并文档属性
        if (options.merge_properties) {
            if (!other_workbook->doc_properties_.title.empty()) {
                doc_properties_.title = other_workbook->doc_properties_.title;
            }
            if (!other_workbook->doc_properties_.author.empty()) {
                doc_properties_.author = other_workbook->doc_properties_.author;
            }
            if (!other_workbook->doc_properties_.subject.empty()) {
                doc_properties_.subject = other_workbook->doc_properties_.subject;
            }
            if (!other_workbook->doc_properties_.company.empty()) {
                doc_properties_.company = other_workbook->doc_properties_.company;
            }
            
            // 合并自定义属性
            for (const auto& prop : other_workbook->custom_property_manager_->getAllDetailedProperties()) {
                setCustomProperty(prop.name, prop.value);
            }
            
            LOG_DEBUG("Merged document properties");
        }
        
        LOG_INFO("Successfully merged workbook: {} worksheets, {} formats",
                merged_count, other_workbook->format_repo_->getFormatCount());
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during workbook merge: {}", e.what());
        return false;
    }
}

bool Workbook::exportWorksheets(const std::vector<std::string>& worksheet_names, const std::string& output_filename) {
    if (worksheet_names.empty()) {
        LOG_ERROR("No worksheets specified for export");
        return false;
    }
    
    try {
        // 创建新工作簿
        auto export_workbook = create(Path(output_filename));
        if (!export_workbook->open()) {
            LOG_ERROR("Failed to create export workbook: {}", output_filename);
            return false;
        }
        
        // 复制指定的工作表
        int exported_count = 0;
        for (const std::string& name : worksheet_names) {
            auto source_worksheet = getWorksheet(name);
            if (!source_worksheet) {
                LOG_WARN("Worksheet not found for export: {}", name);
                continue;
            }
            
            auto new_worksheet = export_workbook->addWorksheet(name);
            if (new_worksheet) {
                // 这里需要实现深拷贝逻辑
                // 简化版本：复制基本属性
                exported_count++;
                LOG_DEBUG("Exported worksheet: {}", name);
            }
        }
        
        // 复制文档属性
        export_workbook->doc_properties_ = doc_properties_;
        // 复制自定义属性
        for (const auto& prop : custom_property_manager_->getAllDetailedProperties()) {
            export_workbook->setCustomProperty(prop.name, prop.value);
        }
        
        // 保存导出的工作簿
        bool success = export_workbook->save();
        export_workbook->close();
        
        if (success) {
            LOG_INFO("Successfully exported {} worksheets to: {}", exported_count, output_filename);
        } else {
            LOG_ERROR("Failed to save exported workbook: {}", output_filename);
        }
        
        return success;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during worksheet export: {}", e.what());
        return false;
    }
}

int Workbook::batchRenameWorksheets(const std::unordered_map<std::string, std::string>& rename_map) {
    int renamed_count = 0;
    
    for (const auto& [old_name, new_name] : rename_map) {
        if (renameWorksheet(old_name, new_name)) {
            renamed_count++;
            LOG_DEBUG("Renamed worksheet: {} -> {}", old_name, new_name);
        } else {
            LOG_WARN("Failed to rename worksheet: {} -> {}", old_name, new_name);
        }
    }
    
    LOG_INFO("Batch rename completed: {} worksheets renamed", renamed_count);
    return renamed_count;
}

int Workbook::batchRemoveWorksheets(const std::vector<std::string>& worksheet_names) {
    int removed_count = 0;
    
    for (const std::string& name : worksheet_names) {
        if (removeWorksheet(name)) {
            removed_count++;
            LOG_DEBUG("Removed worksheet: {}", name);
        } else {
            LOG_WARN("Failed to remove worksheet: {}", name);
        }
    }
    
    LOG_INFO("Batch remove completed: {} worksheets removed", removed_count);
    return removed_count;
}

bool Workbook::reorderWorksheets(const std::vector<std::string>& new_order) {
    if (new_order.size() != worksheets_.size()) {
        LOG_ERROR("New order size ({}) doesn't match worksheet count ({})",
                 new_order.size(), worksheets_.size());
        return false;
    }
    
    try {
        std::vector<std::shared_ptr<Worksheet>> reordered_worksheets;
        reordered_worksheets.reserve(worksheets_.size());
        
        // 按新顺序重新排列工作表
        for (const std::string& name : new_order) {
            auto worksheet = getWorksheet(name);
            if (!worksheet) {
                LOG_ERROR("Worksheet not found in reorder list: {}", name);
                return false;
            }
            reordered_worksheets.push_back(worksheet);
        }
        
        // 替换工作表列表
        worksheets_ = std::move(reordered_worksheets);
        
        LOG_INFO("Successfully reordered {} worksheets", worksheets_.size());
        return true;
        
    } catch (const std::exception& e) {
        LOG_ERROR("Exception during worksheet reordering: {}", e.what());
        return false;
    }
}

int Workbook::findAndReplaceAll(const std::string& find_text, const std::string& replace_text,
                                const FindReplaceOptions& options) {
    int total_replacements = 0;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        int replacements = worksheet->findAndReplace(find_text, replace_text,
                                                   options.match_case, options.match_entire_cell);
        total_replacements += replacements;
        
        if (replacements > 0) {
            LOG_DEBUG("Found and replaced {} occurrences in worksheet: {}",
                     replacements, worksheet->getName());
        }
    }
    
    LOG_INFO("Global find and replace completed: {} total replacements", total_replacements);
    return total_replacements;
}

std::vector<std::tuple<std::string, int, int>> Workbook::findAll(const std::string& search_text,
                                                                 const FindReplaceOptions& options) {
    std::vector<std::tuple<std::string, int, int>> results;
    
    for (const auto& worksheet : worksheets_) {
        // 检查工作表过滤器
        if (!options.worksheet_filter.empty()) {
            bool found = std::find(options.worksheet_filter.begin(), options.worksheet_filter.end(),
                                 worksheet->getName()) != options.worksheet_filter.end();
            if (!found) {
                continue; // 跳过不在过滤器中的工作表
            }
        }
        
        auto worksheet_results = worksheet->findCells(search_text, options.match_case, options.match_entire_cell);
        
        // 将结果添加到总结果中，包含工作表名称
        for (const auto& [row, col] : worksheet_results) {
            results.emplace_back(worksheet->getName(), row, col);
        }
        
        if (!worksheet_results.empty()) {
            LOG_DEBUG("Found {} matches in worksheet: {}", worksheet_results.size(), worksheet->getName());
        }
    }
    
    LOG_INFO("Global search completed: {} total matches found", results.size());
    return results;
}

Workbook::WorkbookStats Workbook::getStatistics() const {
    WorkbookStats stats;
    
    stats.total_worksheets = worksheets_.size();
    stats.total_formats = format_repo_->getFormatCount();
    
    // 计算总单元格数和内存使用
    for (const auto& worksheet : worksheets_) {
        size_t cell_count = worksheet->getCellCount();
        stats.total_cells += cell_count;
        stats.worksheet_cell_counts[worksheet->getName()] = cell_count;
        
        if (worksheet->isOptimizeMode()) {
            stats.memory_usage += worksheet->getMemoryUsage();
        }
    }
    
    // 估算工作簿本身的内存使用
    stats.memory_usage += sizeof(Workbook);
    stats.memory_usage += worksheets_.capacity() * sizeof(std::shared_ptr<Worksheet>);
    stats.memory_usage += format_repo_->getMemoryUsage();
    stats.memory_usage += custom_property_manager_->size() * sizeof(CustomProperty);
    stats.memory_usage += defined_name_manager_->size() * sizeof(DefinedName);
    
    return stats;
}

// ========== 智能模式选择辅助方法 ==========

size_t Workbook::estimateMemoryUsage() const {
    size_t total_memory = 0;
    
    // 估算工作表内存使用
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_memory += worksheet->getMemoryUsage();
        } else {
            // 估算标准模式的内存使用
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                size_t cell_count = (max_row + 1) * (max_col + 1);
                total_memory += cell_count * 100; // 估算每个单元格100字节
            }
        }
    }
    
    // 估算格式池内存
    total_memory += format_repo_->getMemoryUsage();
    
    // 估算共享字符串内存
    if (shared_string_table_) {
        total_memory += shared_string_table_->getMemoryUsage();
    }
    
    // 估算XML生成时的临时内存（约为数据的2-3倍）
    total_memory *= 3;
    
    return total_memory;
}

size_t Workbook::getTotalCellCount() const {
    size_t total_cells = 0;
    
    for (const auto& worksheet : worksheets_) {
        if (worksheet->isOptimizeMode()) {
            total_cells += worksheet->getCellCount();
        } else {
            auto [max_row, max_col] = worksheet->getUsedRange();
            if (max_row >= 0 && max_col >= 0) {
                // 估算实际有数据的单元格数量（不是整个矩形区域）
                size_t estimated_cells = 0;
                for (int row = 0; row <= max_row; ++row) {
                    for (int col = 0; col <= max_col; ++col) {
                        if (worksheet->hasCellAt(row, col)) {
                            estimated_cells++;
                        }
                    }
                }
                total_cells += estimated_cells;
            }
        }
    }
    
    return total_cells;
}

std::unique_ptr<StyleTransferContext> Workbook::copyStylesFrom(const Workbook& source_workbook) {
    LOG_DEBUG("开始从源工作簿复制样式数据");
    
    // 创建样式传输上下文
    auto transfer_context = std::make_unique<StyleTransferContext>(*source_workbook.format_repo_, *format_repo_);
    
    // 预加载所有映射以触发批量复制
    transfer_context->preloadAllMappings();
    
    auto stats = transfer_context->getTransferStats();
    LOG_DEBUG("完成样式复制，传输了{}个格式，去重了{}个", 
             stats.transferred_count, stats.deduplicated_count);
    
    // 🔧 关键修复：自动复制主题XML以保持颜色和字体一致性
    const std::string& source_theme = source_workbook.getThemeXML();
    if (!source_theme.empty()) {
        // 只有当前工作簿没有自定义主题时才复制源主题
        if (theme_xml_.empty()) {
            theme_xml_ = source_theme;
            LOG_DEBUG("自动复制主题XML ({} 字节)", theme_xml_.size());
        } else {
            LOG_DEBUG("当前工作簿已有自定义主题，保持现有主题不变");
        }
    } else {
        LOG_DEBUG("源工作簿无自定义主题，保持默认主题");
    }
    
    return transfer_context;
}

FormatRepository::DeduplicationStats Workbook::getStyleStats() const {
    return format_repo_->getDeduplicationStats();
}

bool Workbook::generateWithGenerator(bool use_streaming_writer) {
    if (!file_manager_) {
        LOG_ERROR("FileManager is null - cannot write workbook");
        return false;
    }
    std::unique_ptr<IFileWriter> writer;
    if (use_streaming_writer) {
        writer = std::make_unique<StreamingFileWriter>(file_manager_.get());
    } else {
        writer = std::make_unique<BatchFileWriter>(file_manager_.get());
    }
    ExcelStructureGenerator generator(this, std::move(writer));
    return generator.generate();
}

bool Workbook::isModified() const {
    // 检查DirtyManager是否有修改标记
    if (dirty_manager_ && dirty_manager_->hasDirtyData()) {
        return true;
    }
    
    // 检查主题是否被修改
    if (theme_dirty_) {
        return true;
    }
    
    // 检查工作表是否有修改（如果有hasChanges方法）
    for (const auto& worksheet : worksheets_) {
        if (worksheet) {
            // TODO: 检查Worksheet是否有hasChanges或类似方法
            // if (worksheet->hasChanges()) {
            //     return true;
            // }
        }
    }
    
    return false;
}

}} // namespace fastexcel::core
